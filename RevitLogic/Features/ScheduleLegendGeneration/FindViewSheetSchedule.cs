using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitLogic.Features.ScheduleLegendGeneration
{
    /// <summary>
    /// Schedule Legend Generation / Node01：
    /// 從目前作用中的 ViewSheet 取得圖紙資訊，並收集該圖紙上所有
    /// ScheduleSheetInstance 的基本資料與邊界資訊。
    /// 
    /// 本類別負責：
    /// - 驗證目前 ActiveView 是否為 ViewSheet
    /// - 讀取圖紙基本資訊
    /// - 列出圖紙上的明細表放置實例
    /// - 取得各實例的 bounding box 與 base_left_top（feet）
    /// - 整理並回傳標準化結果 DTO
    /// 
    /// 本類別不負責：
    /// - 明細表內容解析
    /// - 族群名稱辨識
    /// - DWG 搜尋或圖紙放置
    /// 
    /// 維護注意：
    /// - 邊界與座標皆為 Revit 內部單位 feet
    /// - 若後續圖例放置位置異常，先確認本類別輸出的 bbox / base_left_top 是否正確
    /// </summary>
    public static class FindViewSheetSchedule
    {
        /// <summary>圖紙基本資訊。</summary>
        public sealed class SheetInfo
        {
            public string SheetId { get; set; }
            public string SheetNumber { get; set; }
            public string SheetName { get; set; }
        }

        /// <summary>圖紙上單一明細表放置的資訊（含邊界，單位 feet）。</summary>
        public sealed class ScheduleOnSheetInfo
        {
            public string ScheduleName { get; set; }
            public string ScheduleViewId { get; set; }
            public string SheetInstanceId { get; set; }
            public double BboxMinX { get; set; }
            public double BboxMinY { get; set; }
            public double BboxMaxX { get; set; }
            public double BboxMaxY { get; set; }
            public double BaseLeftTopX { get; set; }
            public double BaseLeftTopY { get; set; }
        }

        /// <summary>取得圖紙上明細表邊界的結果。</summary>
        public sealed class FindViewSheetScheduleResult
        {
            public bool Ok { get; set; }
            public string PickReason { get; set; }
            public string Error { get; set; }
            public SheetInfo PickedSheet { get; set; }
            public List<ScheduleOnSheetInfo> SchedulesOnPickedSheet { get; set; }
            public List<string> FailedInstances { get; set; }
        }

        /// <summary>
        /// 以目前作用中視圖決定圖紙：若 ActiveView 為 ViewSheet 則使用該圖紙，否則回傳失敗。
        /// 並列出該圖紙上所有 ScheduleSheetInstance 及其在圖紙上的邊界（feet）。
        /// </summary>
        public static FindViewSheetScheduleResult Find(UIApplication uiapp)
        {
            var result = new FindViewSheetScheduleResult
            {
                SchedulesOnPickedSheet = new List<ScheduleOnSheetInfo>(),
                FailedInstances = new List<string>()
            };

            if (uiapp == null)
            {
                result.Ok = false;
                result.Error = "uiapp is null";
                return result;
            }

            UIDocument uidoc = uiapp.ActiveUIDocument;
            if (uidoc == null)
            {
                result.Ok = false;
                result.Error = "沒有 ActiveUIDocument（可能沒有開任何模型）";
                return result;
            }

            Document doc = uidoc.Document;
            View activeView = uidoc.ActiveView;
            if (activeView == null)
            {
                result.Ok = false;
                result.Error = "無法取得目前視圖（ActiveView 為 null）";
                return result;
            }

            ViewSheet pickedSheet = null;
            if (activeView is ViewSheet vs)
            {
                pickedSheet = vs;
                result.PickReason = "ActiveView is ViewSheet";
            }
            else
            {
                result.Ok = false;
                result.PickReason = "目前視圖不是圖紙。請先開啟圖紙 tab，並點一下圖紙畫面後再執行。";
                result.Error = "目前無法取得可用的 ViewSheet。請確認有開啟圖紙 tab，並點一下圖紙畫面後再跑。";
                return result;
            }

            result.PickedSheet = new SheetInfo
            {
                SheetId = pickedSheet.Id.ToString(),
                SheetNumber = pickedSheet.SheetNumber ?? "",
                SheetName = pickedSheet.Name ?? ""
            };

            var instances = new FilteredElementCollector(doc, pickedSheet.Id)
                .OfClass(typeof(ScheduleSheetInstance))
                .Cast<ScheduleSheetInstance>()
                .ToList();

            foreach (ScheduleSheetInstance ssi in instances)
            {
                ViewSchedule scheduleView = doc.GetElement(ssi.ScheduleId) as ViewSchedule;
                string scheduleName = scheduleView?.Name ?? "<Unknown>";

                BoundingBoxXYZ bb = null;
                try
                {
                    bb = ssi.get_BoundingBox(pickedSheet);
                }
                catch (System.Exception ex)
                {
                    result.FailedInstances.Add($"{scheduleName} (instance {ssi.Id}): get_BoundingBox failed - {ex.Message}");
                    continue;
                }

                if (bb == null)
                {
                    result.FailedInstances.Add($"{scheduleName} (instance {ssi.Id}): BoundingBox is None (maybe not placed / not resolvable in this view)");
                    continue;
                }

                XYZ minPt = bb.Min;
                XYZ maxPt = bb.Max;
                double baseLeftTopX = minPt.X;
                double baseLeftTopY = maxPt.Y;

                result.SchedulesOnPickedSheet.Add(new ScheduleOnSheetInfo
                {
                    ScheduleName = scheduleName,
                    ScheduleViewId = ssi.ScheduleId.ToString(),
                    SheetInstanceId = ssi.Id.ToString(),
                    BboxMinX = minPt.X,
                    BboxMinY = minPt.Y,
                    BboxMaxX = maxPt.X,
                    BboxMaxY = maxPt.Y,
                    BaseLeftTopX = baseLeftTopX,
                    BaseLeftTopY = baseLeftTopY
                });
            }

            result.SchedulesOnPickedSheet = result.SchedulesOnPickedSheet
                .OrderBy(s => s.ScheduleName)
                .ToList();
            result.Ok = true;
            return result;
        }
    }
}
