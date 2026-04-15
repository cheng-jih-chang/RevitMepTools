using System;
using System.Linq;
using System.Reflection;
using System.Text;
using Autodesk.Revit.UI;

namespace RevitLogic.Features.ScheduleLegendGeneration
{
    /// <summary>
    /// Schedule Legend Generation 的主流程協調器（orchestrator）。
    ///
    /// 【角色】
    /// 此檔案位於圖例生成功能的 workflow / service 層，
    /// 負責串接整個「圖紙明細表 → 族群名稱解析 → DWG 搜尋 → 圖紙放置」流程。
    /// 它的主要工作是控制執行順序、整合各節點結果、整理輸出訊息，
    /// 而不是實作各節點的底層判斷規則。
    ///
    /// 【本檔負責】
    /// 1. 取得目前圖紙與圖紙上的明細表清單。
    /// 2. 呼叫明細表解析邏輯，取得各明細表中的族群名稱。
    /// 3. 依族群名稱到指定根目錄搜尋對應的 DWG 檔案。
    /// 4. 以固定放置參數呼叫 DWG 放置邏輯，將圖例放到圖紙上。
    /// 5. 彙整每個節點的成功、失敗、警告與摘要，回傳為可顯示的文字報告。
    ///
    /// 【本檔不負責】
    /// 1. 不負責明細表標題列、欄位或族群值的解析規則。
    /// 2. 不負責 DWG 檔名比對、搜尋策略與重複判定邏輯。
    /// 3. 不負責 DWG 在圖紙上的幾何定位、列高計算或實際 Revit link/import 細節。
    /// 4. 不應在此檔案加入過多 node 內部演算法；各節點細節應維持在各自的 class 中。
    ///
    /// 【流程位置】
    /// 上游通常由 EntryPoint / Command 呼叫本 service 的 Execute(...)。
    /// 下游依序呼叫：
    /// - FindViewSheetSchedule
    /// - FindLegendName
    /// - FindLegendDwg
    /// - LinkLegendDwg
    ///
    /// 【輸入】
    /// - UIApplication uiapp：提供目前 Revit UI / 文件 / 圖紙上下文。
    ///
    /// 【輸出】
    /// - string：回傳多行文字報告，內容包含圖紙資訊、明細表摘要、族群解析結果、
    ///   DWG 搜尋結果、放置結果、錯誤與警告訊息。
    ///
    /// 【重要假設 / 維護注意】
    /// - 目前 DWG 根目錄由 Revit Ribbon UI（RevitMepAddinHost.App.LegendDwgRootFolderText）提供。
    /// - 目前 DWG 放置參數（例如 OffsetXMm、RowHeightMm、OffsetYMm、DwgUnits）也先寫死在本檔。
    /// - 若未來要支援不同專案 / 使用者設定，應優先將這些值抽離成設定來源或 options，
    ///   不要把更多條件分支直接堆進 Execute(...）。
    /// - 若 Execute(...) 持續變長，優先抽離「報告組裝」或「各 node 結果格式化」邏輯，
    ///   但保留本檔作為高層流程入口。
    ///
    /// 【AI / 接手者閱讀順序】
    /// 1. 先看 Execute(...)，理解整體流程順序。
    /// 2. 再分別查看 FindViewSheetSchedule / FindLegendName / FindLegendDwg / LinkLegendDwg 的實作。
    /// 3. 若要除錯，先判斷問題屬於哪個 node，再往下追該 node 的資料來源與判斷規則。
    /// </summary>
    public class ScheduleLegendGenerationService
    {
        public string Execute(UIApplication uiapp)
        {
            if (uiapp == null) return "uiapp is null";

            string legendDwgRootFolder = GetLegendDwgRootFolderTextFromUi();
            if (string.IsNullOrWhiteSpace(legendDwgRootFolder))
                return "Legend DWG 根目錄未設定。請先在 Ribbon 的 ScheduleLegendGeneration 旁輸入路徑並按 Enter，或從下拉選單選擇預設路徑。";

            FindViewSheetSchedule.FindViewSheetScheduleResult sheetResult = FindViewSheetSchedule.Find(uiapp);

            var sb = new StringBuilder();
            if (!sheetResult.Ok)
            {
                sb.AppendLine("無法取得圖紙或明細表。");
                sb.AppendLine("原因: " + (sheetResult.PickReason ?? ""));
                if (!string.IsNullOrEmpty(sheetResult.Error))
                    sb.AppendLine("錯誤: " + sheetResult.Error);
                return sb.ToString();
            }

            sb.AppendLine("圖紙: " + sheetResult.PickedSheet.SheetNumber + " - " + sheetResult.PickedSheet.SheetName);
            sb.AppendLine("取得方式: " + sheetResult.PickReason);
            sb.AppendLine("圖紙上明細表數量: " + sheetResult.SchedulesOnPickedSheet.Count);
            sb.AppendLine();
            sb.AppendLine("--- 明細表清單（邊界單位: feet）---");

            foreach (var s in sheetResult.SchedulesOnPickedSheet)
            {
                sb.AppendLine("  " + s.ScheduleName);
                sb.AppendLine("    視圖ID: " + s.ScheduleViewId + ", 放置實例ID: " + s.SheetInstanceId);
                sb.AppendLine("    bbox: (" + s.BboxMinX + "," + s.BboxMinY + ") ~ (" + s.BboxMaxX + "," + s.BboxMaxY + ")");
                sb.AppendLine("    base_left_top: (" + s.BaseLeftTopX + "," + s.BaseLeftTopY + ")");
            }

            if (sheetResult.FailedInstances != null && sheetResult.FailedInstances.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("--- 無法取得邊界的實例 ---");
                foreach (string fail in sheetResult.FailedInstances)
                    sb.AppendLine("  " + fail);
            }

            // Node02: 解析每張明細表的標題列與「族群」值
            var doc = uiapp.ActiveUIDocument?.Document;
            if (doc != null && sheetResult.SchedulesOnPickedSheet != null && sheetResult.SchedulesOnPickedSheet.Count > 0)
            {
                FindLegendName.FindLegendNameResult legendResult = FindLegendName.Find(doc, sheetResult.SchedulesOnPickedSheet);

                sb.AppendLine();
                sb.AppendLine("--- 族群解析（Node02 FindLegendName）---");
                if (!string.IsNullOrEmpty(legendResult.Error))
                    sb.AppendLine("錯誤: " + legendResult.Error);
                if (legendResult.Results != null && legendResult.Results.Count > 0)
                {
                    foreach (var r in legendResult.Results)
                    {
                        sb.AppendLine("  " + r.ScheduleName);
                        sb.AppendLine("    標題列: " + r.PickedSection + " row " + r.PickedRow + ", score=" + r.PickedRowScore);
                        sb.AppendLine("    族群欄索引: " + (r.ColIndexGroup.HasValue ? r.ColIndexGroup.Value.ToString() : "—"));
                        sb.AppendLine("    說明欄索引: " + (r.ColIndexDesc.HasValue ? r.ColIndexDesc.Value.ToString() : "—"));
                        if (r.GroupValues != null && r.GroupValues.Count > 0)
                            sb.AppendLine("    族群值(" + r.GroupValues.Count + "): " + string.Join(", ", r.GroupValues.Take(10)) + (r.GroupValues.Count > 10 ? " ..." : ""));
                        else
                            sb.AppendLine("    族群值: (無)");
                    }
                }
                if (legendResult.Errors != null && legendResult.Errors.Count > 0)
                {
                    sb.AppendLine();
                    sb.AppendLine("  解析錯誤:");
                    foreach (string err in legendResult.Errors)
                        sb.AppendLine("    " + err);
                }

                // Node03: 依族群名稱在根目錄下找對應 .dwg
                FindLegendDwg.FindLegendDwgResult dwgResult = FindLegendDwg.Find(legendResult.GroupMapBySchedule, legendDwgRootFolder);

                sb.AppendLine();
                sb.AppendLine("--- DWG 對應（Node03 FindLegendDwg）---");
                sb.AppendLine("根目錄: " + (dwgResult.Root ?? legendDwgRootFolder));
                if (!dwgResult.Ok)
                {
                    sb.AppendLine("錯誤: " + (dwgResult.Error ?? ""));
                }
                else
                {
                    sb.AppendLine("根目錄下 .dwg 總數: " + dwgResult.DwgTotalFoundUnderRoot);
                    if (dwgResult.SummaryBySchedule != null && dwgResult.SummaryBySchedule.Count > 0)
                    {
                        foreach (var kv in dwgResult.SummaryBySchedule)
                        {
                            var sum = kv.Value;
                            sb.AppendLine("  " + kv.Key + " => 族群數 " + sum.GroupsTotal + ", 找到 " + sum.Found + ", 缺 " + sum.Missing + ", 重複 " + sum.Duplicates);
                        }
                    }
                    if (dwgResult.Records != null && dwgResult.Records.Count > 0)
                    {
                        sb.AppendLine("  明細(前 15 筆):");
                        foreach (var rec in dwgResult.Records.Take(15))
                        {
                            string paths = rec.DwgPaths != null && rec.DwgPaths.Count > 0
                                ? string.Join("; ", rec.DwgPaths.Take(2)) + (rec.DwgPaths.Count > 2 ? " ..." : "")
                                : "—";
                            sb.AppendLine("    [" + rec.Status + "] " + rec.ScheduleName + " / " + rec.GroupName + " => " + (rec.ExpectedFilenameTry ?? "") + " | " + paths);
                        }
                        if (dwgResult.Records.Count > 15)
                            sb.AppendLine("    ... 共 " + dwgResult.Records.Count + " 筆");
                    }
                }

                // Node04: 將 DWG 放置到圖紙上（參數先寫死：OffsetXMm=11, RowHeightMm=6.7, DwgUnits=mm, OffsetYMm=21）
                if (dwgResult.Ok && dwgResult.Records != null && dwgResult.Records.Count > 0)
                {
                    var linkOptions = new LinkLegendDwg.LinkLegendDwgOptions
                    {
                        OffsetXMm = 11,
                        RowHeightMm = 6.7,
                        DeleteOld = true,
                        AbsXMm = null,
                        AbsYMm = null,
                        DwgUnits = "mm",
                        OffsetYMm = 21
                    };
                    LinkLegendDwg.LinkLegendDwgResult linkResult = LinkLegendDwg.Place(
                        doc, uiapp.ActiveUIDocument, sheetResult, dwgResult, linkOptions);

                    sb.AppendLine();
                    sb.AppendLine("--- 圖紙放置 DWG（Node04 LinkLegendDwg）---");
                    if (!string.IsNullOrEmpty(linkResult.Error))
                        sb.AppendLine("錯誤: " + linkResult.Error);
                    else
                    {
                        sb.AppendLine("圖紙: " + linkResult.SheetNumber + " - " + linkResult.SheetName);
                        sb.AppendLine("已刪除舊連結: " + linkResult.DeletedCount + "，已放置: " + linkResult.PlacedCount);
                        sb.AppendLine("Link 成功/失敗: " + linkResult.LinkOk + " / " + linkResult.LinkFail + "，Import 成功/失敗: " + linkResult.ImportOk + " / " + linkResult.ImportFail);
                        if (linkResult.Warnings != null && linkResult.Warnings.Count > 0)
                        {
                            sb.AppendLine("警告:");
                            foreach (string w in linkResult.Warnings.Take(10))
                                sb.AppendLine("  " + w);
                            if (linkResult.Warnings.Count > 10)
                                sb.AppendLine("  ... 共 " + linkResult.Warnings.Count + " 則");
                        }
                    }
                }
            }

            sb.AppendLine();
            sb.AppendLine("(後續可依此清單進行圖例生成)");
            return sb.ToString();
        }

        private static string GetLegendDwgRootFolderTextFromUi()
        {
            // 以反射取值，避免 RevitLogic 對 RevitMepAddinHost 產生編譯期相依。
            Type appType = Type.GetType("RevitMepAddinHost.App, RevitMepAddinHost", false);
            if (appType == null)
                return null;

            PropertyInfo rootFolderProperty = appType.GetProperty(
                "LegendDwgRootFolderText",
                BindingFlags.Public | BindingFlags.Static);

            if (rootFolderProperty == null)
                return null;

            object value = rootFolderProperty.GetValue(null, null);
            return value == null ? null : value.ToString();
        }
    }
}
