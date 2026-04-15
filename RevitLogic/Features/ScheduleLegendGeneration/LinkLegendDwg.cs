using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitLogic.Features.ScheduleLegendGeneration
{
    /// <summary>
    /// Schedule Legend Generation 的 Node04。
    /// 負責依 Node01 的圖紙/明細表位置與 Node03 的 DWG 對應結果，
    /// 將圖例 DWG 放置到目標圖紙上。
    /// 本類別會處理既有對應 CAD 的刪除、放置起點計算、逐列排版、
    /// DWG link/import 嘗試，以及放置後的位置移動。
    /// 本類別不負責明細表解析、族群名稱辨識或 DWG 檔名搜尋；
    /// 這些邏輯應由前置 node 提供。
    /// 維護時若圖例位置異常，優先檢查 bbox 左上基準點、offset / row height 換算、
    /// 以及移動後座標是否與預期一致。
    /// </summary>
    public static class LinkLegendDwg
    {
        private const double MmToFt = 1.0 / 304.8;

        /// <summary>Node04 參數，先寫死；之後可改為從 UI 或設定讀取。</summary>
        public sealed class LinkLegendDwgOptions
        {
            public double OffsetXMm { get; set; } = 11;
            public double RowHeightMm { get; set; } = 6.7;
            public bool DeleteOld { get; set; } = true;
            public double? AbsXMm { get; set; }
            public double? AbsYMm { get; set; }
            public string DwgUnits { get; set; } = "mm";
            public double OffsetYMm { get; set; } = 21;
        }

        /// <summary>單筆放置結果。</summary>
        public sealed class PlacedRecord
        {
            public string ScheduleName { get; set; }
            public string GroupName { get; set; }
            public string DwgPath { get; set; }
            public string Mode { get; set; }  // "link" | "import"
            public string ImportInstanceId { get; set; }
        }

        /// <summary>執行結果。</summary>
        public sealed class LinkLegendDwgResult
        {
            public bool Ok { get; set; }
            public string Error { get; set; }
            public string SheetId { get; set; }
            public string SheetNumber { get; set; }
            public string SheetName { get; set; }
            public int DeletedCount { get; set; }
            public int PlacedCount { get; set; }
            public int LinkOk { get; set; }
            public int LinkFail { get; set; }
            public int ImportOk { get; set; }
            public int ImportFail { get; set; }
            public List<PlacedRecord> Placed { get; set; }
            public List<string> Warnings { get; set; }
        }

        /// <summary>
        /// 依 Node01 與 Node03 結果，在圖紙上刪除舊的對應 CAD 連結後，依序放置 DWG（Link 優先，失敗改 Import，ThisViewOnly）。
        /// </summary>
        public static LinkLegendDwgResult Place(
            Document doc,
            UIDocument uidoc,
            FindViewSheetSchedule.FindViewSheetScheduleResult node01,
            FindLegendDwg.FindLegendDwgResult node03,
            LinkLegendDwgOptions options = null)
        {
            var result = new LinkLegendDwgResult { Placed = new List<PlacedRecord>(), Warnings = new List<string>() };
            options = options ?? new LinkLegendDwgOptions();

            if (doc == null)
            {
                result.Ok = false;
                result.Error = "doc is null";
                return result;
            }
            if (node01 == null || !node01.Ok || node01.PickedSheet == null)
            {
                result.Ok = false;
                result.Error = "Node01 invalid or missing picked_sheet";
                return result;
            }
            if (node03?.Records == null)
            {
                result.Ok = false;
                result.Error = "Node03 missing records";
                return result;
            }

            ViewSheet sheet = ResolveSheet(doc, uidoc, node01);
            if (sheet == null)
            {
                result.Ok = false;
                result.Error = "Cannot resolve ViewSheet. Activate the target sheet tab and run again.";
                return result;
            }

            result.SheetId = sheet.Id.ToString();
            result.SheetNumber = sheet.SheetNumber ?? "";
            result.SheetName = sheet.Name ?? "";

            var ssiBySchedName = new Dictionary<string, ScheduleSheetInstance>(StringComparer.OrdinalIgnoreCase);
            foreach (Element el in new FilteredElementCollector(doc, sheet.Id).OfClass(typeof(ScheduleSheetInstance)))
            {
                if (el is ScheduleSheetInstance ssi)
                {
                    ViewSchedule vs = doc.GetElement(ssi.ScheduleId) as ViewSchedule;
                    if (vs != null && !string.IsNullOrEmpty(vs.Name))
                        ssiBySchedName[vs.Name] = ssi;
                }
            }

            var targetPathsLower = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in node03.Records)
            {
                if (r.DwgPaths != null && r.DwgPaths.Count > 0 && !string.IsNullOrEmpty(r.DwgPaths[0]))
                    targetPathsLower.Add(r.DwgPaths[0].Trim());
            }

            double offsetXFt = options.OffsetXMm * MmToFt;
            double rowHFt = options.RowHeightMm * MmToFt;
            double offsetYFt = options.OffsetYMm * MmToFt;
            ImportUnit? importUnit = ParseDwgUnits(options.DwgUnits);

            using (Transaction tx = new Transaction(doc, "LinkLegendDwg"))
            {
                tx.Start();

                if (options.DeleteOld && targetPathsLower.Count > 0)
                {
                    foreach (Element el in new FilteredElementCollector(doc, sheet.Id).OfClass(typeof(ImportInstance)))
                    {
                        if (!(el is ImportInstance inst)) continue;
                        string path = TryGetLinkPath(doc, inst);
                        if (string.IsNullOrEmpty(path)) continue;
                        if (targetPathsLower.Contains(path.Trim()))
                        {
                            try
                            {
                                doc.Delete(inst.Id);
                                result.DeletedCount++;
                            }
                            catch { }
                        }
                    }
                }

                var recordsBySchedule = node03.Records
                    .Where(r => !string.IsNullOrWhiteSpace(r.ScheduleName))
                    .GroupBy(r => r.ScheduleName.Trim(), StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

                foreach (var kv in recordsBySchedule)
                {
                    string schedName = kv.Key;
                    List<FindLegendDwg.DwgMatchRecord> recs = kv.Value;
                    if (!ssiBySchedName.TryGetValue(schedName, out ScheduleSheetInstance ssi))
                    {
                        result.Warnings.Add("Schedule not found on sheet: " + schedName);
                        continue;
                    }

                    XYZ start;
                    if (options.AbsXMm.HasValue && options.AbsYMm.HasValue)
                    {
                        start = new XYZ(options.AbsXMm.Value * MmToFt, options.AbsYMm.Value * MmToFt, 0);
                    }
                    else
                    {
                        XYZ basePt = TopLeftFromBbox(ssi, sheet);
                        if (basePt == null)
                        {
                            result.Warnings.Add("Cannot get schedule bbox: " + schedName);
                            continue;
                        }
                        start = new XYZ(basePt.X + offsetXFt, basePt.Y, 0);
                    }

                    for (int i = 0; i < recs.Count; i++)
                    {
                        var r = recs[i];
                        if (string.Equals(r.Status, "found", StringComparison.OrdinalIgnoreCase) == false)
                        {
                            result.Warnings.Add("Missing DWG: " + schedName + " / " + (r.GroupName ?? ""));
                            continue;
                        }
                        string dwgPath = (r.DwgPaths != null && r.DwgPaths.Count > 0) ? r.DwgPaths[0] : null;
                        if (string.IsNullOrEmpty(dwgPath))
                        {
                            result.Warnings.Add("No dwg_paths: " + schedName + " / " + (r.GroupName ?? ""));
                            continue;
                        }
                        if (!File.Exists(dwgPath))
                        {
                            result.Warnings.Add("DWG not exists: " + dwgPath);
                            continue;
                        }

                        XYZ pt = new XYZ(start.X, start.Y - offsetYFt - i * rowHFt, 0);
                        bool ok = PlaceDwgOnSheet(doc, sheet, dwgPath, pt, importUnit, out ElementId eid, out string mode, out string err);
                        if (ok && eid != null && eid != ElementId.InvalidElementId)
                        {
                            if (string.Equals(mode, "link", StringComparison.OrdinalIgnoreCase))
                                result.LinkOk++;
                            else
                                result.ImportOk++;
                            result.Placed.Add(new PlacedRecord
                            {
                                ScheduleName = schedName,
                                GroupName = r.GroupName ?? "",
                                DwgPath = dwgPath,
                                Mode = mode ?? "",
                                ImportInstanceId = eid.ToString()
                            });
                        }
                        else
                        {
                            if (string.Equals(mode, "link", StringComparison.OrdinalIgnoreCase))
                                result.LinkFail++;
                            else
                                result.ImportFail++;
                            result.Warnings.Add("Place failed: " + dwgPath + " | " + (err ?? ""));
                        }
                    }
                }

                tx.Commit();
            }

            result.PlacedCount = result.Placed.Count;
            result.Ok = true;
            result.Error = null;
            return result;
        }

        private static ViewSheet ResolveSheet(Document doc, UIDocument uidoc, FindViewSheetSchedule.FindViewSheetScheduleResult node01)
        {
            if (node01.PickedSheet != null && !string.IsNullOrEmpty(node01.PickedSheet.SheetId))
            {
                if (long.TryParse(node01.PickedSheet.SheetId, out long idVal))
                {
                    Element el = doc.GetElement(new ElementId(idVal));
                    if (el is ViewSheet vs)
                        return vs;
                }
            }
            if (uidoc?.ActiveView is ViewSheet av)
                return av;
            if (node01.PickedSheet != null && !string.IsNullOrEmpty(node01.PickedSheet.SheetNumber))
            {
                ViewSheet byNum = new FilteredElementCollector(doc).OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .FirstOrDefault(s => s.SheetNumber == node01.PickedSheet.SheetNumber);
                if (byNum != null) return byNum;
            }
            return null;
        }

        private static XYZ TopLeftFromBbox(ScheduleSheetInstance ssi, ViewSheet sheet)
        {
            try
            {
                BoundingBoxXYZ bb = ssi.get_BoundingBox(sheet);
                if (bb == null) return null;
                return new XYZ(bb.Min.X, bb.Max.Y, 0);
            }
            catch { return null; }
        }

        private static string TryGetLinkPath(Document doc, ImportInstance inst)
        {
            try
            {
                Element typeEl = doc.GetElement(inst.GetTypeId());
                if (!(typeEl is CADLinkType cadType)) return null;
                ExternalFileReference extRef = ExternalFileUtils.GetExternalFileReference(doc, cadType.Id);
                if (extRef == null) return null;
                ModelPath mp = extRef.GetAbsolutePath();
                if (mp == null) return null;
                return ModelPathUtils.ConvertModelPathToUserVisiblePath(mp);
            }
            catch { return null; }
        }

        private static ImportUnit? ParseDwgUnits(string u)
        {
            if (string.IsNullOrWhiteSpace(u)) return null;
            switch (u.Trim().ToLowerInvariant())
            {
                case "auto":
                case "default": return null;
                case "mm":
                case "millimeter":
                case "millimeters": return ImportUnit.Millimeter;
                case "cm":
                case "centimeter": return ImportUnit.Centimeter;
                case "m":
                case "meter": return ImportUnit.Meter;
                case "ft":
                case "foot": return ImportUnit.Foot;
                case "in":
                case "inch": return ImportUnit.Inch;
                default: return null;
            }
        }

        private static bool PlaceDwgOnSheet(Document doc, ViewSheet sheetView, string dwgPath, XYZ targetPt,
            ImportUnit? importUnit, out ElementId outEid, out string usedMode, out string err)
        {
            outEid = ElementId.InvalidElementId;
            usedMode = "link";
            err = null;

            var opt = new DWGImportOptions();
            opt.ThisViewOnly = true;
            opt.Placement = ImportPlacement.Origin;
            if (importUnit.HasValue)
                opt.Unit = importUnit.Value;

            try
            {
                ElementId eid = ElementId.InvalidElementId;
                if (doc.Link(dwgPath, opt, sheetView, out eid) && eid != null && eid != ElementId.InvalidElementId)
                {
                    if (UnpinAndMoveTo(doc, sheetView, eid, targetPt))
                    {
                        outEid = eid;
                        usedMode = "link";
                        return true;
                    }
                }
            }
            catch (Exception ex) { err = ex.Message; }

            try
            {
                ElementId eid = ElementId.InvalidElementId;
                if (doc.Import(dwgPath, opt, sheetView, out eid) && eid != null && eid != ElementId.InvalidElementId)
                {
                    if (UnpinAndMoveTo(doc, sheetView, eid, targetPt))
                    {
                        outEid = eid;
                        usedMode = "import";
                        return true;
                    }
                }
            }
            catch (Exception ex) { err = err ?? ex.Message; }

            usedMode = "import";
            err = err ?? "Link and Import failed.";
            return false;
        }

        private static void TryUnpin(Element e)
        {
            try
            {
                if (e?.Pinned == true)
                    e.Pinned = false;
            }
            catch { }
        }

        private static XYZ CurrentPointOnSheet(Document doc, ElementId eid, View sheetView)
        {
            try
            {
                Element e = doc.GetElement(eid);
                if (e?.Location is LocationPoint lp)
                    return lp.Point;
                if (e != null)
                {
                    BoundingBoxXYZ bb = e.get_BoundingBox(sheetView);
                    if (bb != null)
                        return new XYZ((bb.Min.X + bb.Max.X) * 0.5, (bb.Min.Y + bb.Max.Y) * 0.5, 0);
                }
            }
            catch { }
            return null;
        }

        private static bool UnpinAndMoveTo(Document doc, View sheetView, ElementId eid, XYZ targetPt)
        {
            try
            {
                Element e = doc.GetElement(eid);
                if (e != null) TryUnpin(e);
                XYZ cur = CurrentPointOnSheet(doc, eid, sheetView) ?? new XYZ(0, 0, 0);
                XYZ move = new XYZ(targetPt.X - cur.X, targetPt.Y - cur.Y, 0);
                ElementTransformUtils.MoveElement(doc, eid, move);
                return true;
            }
            catch { return false; }
        }
    }
}
