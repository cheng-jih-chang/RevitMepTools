using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace RevitLogic.Features.ScheduleLegendGeneration
{
    /// <summary>
    /// Schedule Legend Generation 的 Node02。
    /// 負責解析 Node01 取得的各張明細表，找出最可能的標題列、
    /// 定位「族群」與「說明」欄位，並依 UI 順序收集所有族群值。
    /// 本類別只處理明細表內容解析與欄位辨識，
    /// 不負責圖紙明細表定位、DWG 搜尋或圖紙放置邏輯。
    /// </summary>
    public static class FindLegendName
    {
        private const string TargetColName = "族群";
        private const string DescColName = "說明";

        /// <summary>單一明細表的標題/族群解析結果。</summary>
        public sealed class ScheduleLegendResult
        {
            public string ScheduleName { get; set; }
            public string ScheduleViewId { get; set; }
            public string PickedSection { get; set; }
            public int PickedRow { get; set; }
            public int PickedRowScore { get; set; }
            public int? ColIndexGroup { get; set; }
            public int? ColIndexDesc { get; set; }
            public List<string> GroupValues { get; set; }
        }

        /// <summary>FindLegendName 的整體結果。</summary>
        public sealed class FindLegendNameResult
        {
            public bool Ok { get; set; }
            public string Error { get; set; }
            public List<ScheduleLegendResult> Results { get; set; }
            public Dictionary<string, List<string>> GroupMapBySchedule { get; set; }
            public List<string> Errors { get; set; }
        }

        /// <summary>
        /// 依 Node01 的明細表清單，解析每張明細表的標題列與「族群」值清單。
        /// </summary>
        /// <param name="doc">Revit 文件。</param>
        /// <param name="schedulesOnSheet">Node01 回傳的圖紙上明細表清單（至少需含 ScheduleViewId、ScheduleName）。</param>
        public static FindLegendNameResult Find(Document doc, IList<FindViewSheetSchedule.ScheduleOnSheetInfo> schedulesOnSheet)
        {
            var result = new FindLegendNameResult
            {
                Results = new List<ScheduleLegendResult>(),
                GroupMapBySchedule = new Dictionary<string, List<string>>(),
                Errors = new List<string>()
            };

            if (doc == null)
            {
                result.Ok = false;
                result.Error = "doc is null";
                return result;
            }

            if (schedulesOnSheet == null || schedulesOnSheet.Count == 0)
            {
                result.Ok = true;
                result.Error = null;
                return result;
            }

            foreach (var meta in schedulesOnSheet)
            {
                string scheduleName = meta.ScheduleName ?? "<Unknown>";
                string scheduleViewIdStr = meta.ScheduleViewId;
                if (string.IsNullOrEmpty(scheduleViewIdStr))
                {
                    result.Errors.Add($"{scheduleName}: missing schedule_view_id");
                    continue;
                }

                if (!long.TryParse(scheduleViewIdStr, out long idVal))
                {
                    result.Errors.Add($"{scheduleName}: invalid schedule_view_id");
                    continue;
                }

                Element el = doc.GetElement(new ElementId(idVal));
                if (!(el is ViewSchedule vs))
                {
                    result.Errors.Add($"{scheduleName}: not ViewSchedule");
                    continue;
                }

                try
                {
                    TableData td = vs.GetTableData();
                    if (td == null)
                    {
                        result.Errors.Add($"{scheduleName}: GetTableData() is null");
                        continue;
                    }

                    TableSectionData header = td.GetSectionData(SectionType.Header);
                    TableSectionData body = td.GetSectionData(SectionType.Body);
                    if (header == null || body == null)
                    {
                        result.Errors.Add($"{scheduleName}: Header or Body section is null");
                        continue;
                    }

                    int ncolsH = header.NumberOfColumns;
                    int ncolsB = body.NumberOfColumns;
                    var candidates = new List<(string section, int row, string[] texts, int score, int nonEmpty, string joined)>();

                    int scanH = Math.Min(12, header.NumberOfRows);
                    for (int r = 0; r < scanH; r++)
                    {
                        string[] texts = GetRowTexts(vs, SectionType.Header, r, ncolsH);
                        ScoreHeadingRow(texts, out int score, out int nonEmpty, out string joined);
                        candidates.Add(("Header", r, texts, score, nonEmpty, joined));
                    }

                    int scanB = Math.Min(5, body.NumberOfRows);
                    for (int r = 0; r < scanB; r++)
                    {
                        string[] texts = GetRowTexts(vs, SectionType.Body, r, ncolsB);
                        ScoreHeadingRow(texts, out int score, out int nonEmpty, out string joined);
                        candidates.Add(("Body", r, texts, score - 2, nonEmpty, joined));
                    }

                    var best = candidates.OrderByDescending(c => c.score).FirstOrDefault();
                    if (best.texts == null)
                    {
                        result.Errors.Add($"{scheduleName}: no candidates");
                        continue;
                    }

                    string secName = best.section;
                    int bestRow = best.row;
                    string[] bestTexts = best.texts;
                    int bestScore = best.score;

                    int? colIdxGroup = FindColIndex(bestTexts, TargetColName);
                    int? colIdxDesc = FindColIndex(bestTexts, DescColName);

                    var groups = new List<string>();
                    int startRow = (secName == "Body") ? bestRow + 1 : 0;
                    if (colIdxGroup.HasValue)
                    {
                        for (int r = startRow; r < body.NumberOfRows; r++)
                        {
                            string txt = vs.GetCellText(SectionType.Body, r, colIdxGroup.Value);
                            if (txt == null) txt = "";
                            string t = txt.Trim();
                            if (string.IsNullOrEmpty(t)) continue;
                            groups.Add(t);
                        }
                    }

                    result.GroupMapBySchedule[scheduleName] = groups;
                    result.Results.Add(new ScheduleLegendResult
                    {
                        ScheduleName = scheduleName,
                        ScheduleViewId = vs.Id.ToString(),
                        PickedSection = secName,
                        PickedRow = bestRow,
                        PickedRowScore = bestScore,
                        ColIndexGroup = colIdxGroup,
                        ColIndexDesc = colIdxDesc,
                        GroupValues = groups
                    });
                }
                catch (Exception ex)
                {
                    result.Errors.Add($"{scheduleName}: {ex.Message}");
                }
            }

            result.Ok = true;
            result.Error = null;
            return result;
        }

        private static string[] GetRowTexts(ViewSchedule vs, SectionType sectionType, int row, int ncols)
        {
            var list = new string[ncols];
            for (int c = 0; c < ncols; c++)
            {
                string t = vs.GetCellText(sectionType, row, c);
                list[c] = t ?? "";
            }
            return list;
        }

        private static void ScoreHeadingRow(string[] texts, out int score, out int nonEmpty, out string joined)
        {
            nonEmpty = 0;
            var parts = new List<string>();
            foreach (string t in texts)
            {
                if (!string.IsNullOrWhiteSpace(t))
                {
                    nonEmpty++;
                    parts.Add(t.Trim());
                }
            }
            joined = string.Join(" ", parts);
            score = nonEmpty;
            if (joined.Contains(TargetColName)) score += 100;
            if (joined.Contains(DescColName)) score += 80;
            if (nonEmpty <= 1) score -= 10;
        }

        private static int? FindColIndex(string[] texts, string name)
        {
            for (int i = 0; i < texts.Length; i++)
            {
                if (texts[i] == name) return i;
            }
            for (int i = 0; i < texts.Length; i++)
            {
                if (!string.IsNullOrEmpty(texts[i]) && texts[i].Contains(name)) return i;
            }
            return null;
        }
    }
}
