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
        private const double DefaultCharWidthEstimateFt = 0.011;
        private const int MinCalibrationTextLength = 16;

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
            public List<GroupItemInfo> GroupItems { get; set; }
            public int BodyRowCount { get; set; }
            public List<double> BodyRowHeights { get; set; }
            public double BodyTotalRowHeight { get; set; }
            public double HeaderTotalRowHeight { get; set; }
            public double AllTotalRowHeight { get; set; }
            public double CalibratedCharWidthFt { get; set; }
            public int CalibratedCharsPerLine { get; set; }
            public bool UsedAutoCalibration { get; set; }
            public List<string> CalibrationDebugLines { get; set; }
        }

        /// <summary>單筆族群資訊（含 body row 與視覺行數）。</summary>
        public sealed class GroupItemInfo
        {
            public string Value { get; set; }
            public int BodyRowIndex { get; set; }
            public double BodyRowHeight { get; set; }
            public double GroupColumnWidth { get; set; }
            public int TextLength { get; set; }
            public int EstimatedCharsPerLine { get; set; }
            public int VisualLineCount { get; set; }
            public double CalibratedCharWidthFt { get; set; }
            public int CalibratedCharsPerLine { get; set; }
            public bool UsedAutoCalibration { get; set; }
        }

        private sealed class CharWidthCalibrationResult
        {
            public double CalibratedCharWidthFt { get; set; }
            public int CalibratedCharsPerLine { get; set; }
            public bool UsedAutoCalibration { get; set; }
            public List<string> DebugLines { get; set; }
        }

        /// <summary>FindLegendName 的整體結果。</summary>
        public sealed class FindLegendNameResult
        {
            public bool Ok { get; set; }
            public string Error { get; set; }
            public List<ScheduleLegendResult> Results { get; set; }
            public Dictionary<string, List<string>> GroupMapBySchedule { get; set; }
            public Dictionary<string, List<GroupItemInfo>> GroupItemsBySchedule { get; set; }
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
                GroupItemsBySchedule = new Dictionary<string, List<GroupItemInfo>>(),
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
                    var bodyRowHeights = new List<double>();

                    double headerTotalRowHeight = 0;
                    for (int hr = 0; hr < header.NumberOfRows; hr++)
                        headerTotalRowHeight += GetSectionRowHeight(header, hr);

                    double bodyTotalRowHeight = 0;
                    for (int br = 0; br < body.NumberOfRows; br++)
                    {
                        double h = GetSectionRowHeight(body, br);
                        bodyRowHeights.Add(h);
                        bodyTotalRowHeight += h;
                    }

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
                    var groupItems = new List<GroupItemInfo>();
                    var rawGroupItems = new List<(string value, int bodyRowIndex, double bodyRowHeight, int textLength, double groupColumnWidth)>();
                    double groupColumnWidth = colIdxGroup.HasValue ? GetColumnWidth(body, colIdxGroup.Value) : 0;
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
                            rawGroupItems.Add((txt, r, GetSectionRowHeight(body, r), t.Length, groupColumnWidth));
                        }
                    }

                    CharWidthCalibrationResult calibration = CalibrateCharWidthFromSamples(
                        rawGroupItems.Select(x => (x.textLength, x.groupColumnWidth)));

                    foreach (var raw in rawGroupItems)
                    {
                        string normalized = (raw.value ?? "").Trim();
                        int lineCount = EstimateVisualLineCountByCalibratedCharsPerLine(
                            normalized,
                            calibration.CalibratedCharsPerLine,
                            out int estimatedCharsPerLine);
                        groupItems.Add(new GroupItemInfo
                        {
                            Value = raw.value,
                            BodyRowIndex = raw.bodyRowIndex,
                            BodyRowHeight = raw.bodyRowHeight,
                            GroupColumnWidth = raw.groupColumnWidth,
                            TextLength = raw.textLength,
                            EstimatedCharsPerLine = estimatedCharsPerLine,
                            VisualLineCount = lineCount,
                            CalibratedCharWidthFt = calibration.CalibratedCharWidthFt,
                            CalibratedCharsPerLine = calibration.CalibratedCharsPerLine,
                            UsedAutoCalibration = calibration.UsedAutoCalibration
                        });
                    }

                    result.GroupMapBySchedule[scheduleName] = groups;
                    result.GroupItemsBySchedule[scheduleName] = groupItems;
                    result.Results.Add(new ScheduleLegendResult
                    {
                        ScheduleName = scheduleName,
                        ScheduleViewId = vs.Id.ToString(),
                        PickedSection = secName,
                        PickedRow = bestRow,
                        PickedRowScore = bestScore,
                        ColIndexGroup = colIdxGroup,
                        ColIndexDesc = colIdxDesc,
                        GroupValues = groups,
                        GroupItems = groupItems,
                        BodyRowCount = body.NumberOfRows,
                        BodyRowHeights = bodyRowHeights,
                        BodyTotalRowHeight = bodyTotalRowHeight,
                        HeaderTotalRowHeight = headerTotalRowHeight,
                        AllTotalRowHeight = headerTotalRowHeight + bodyTotalRowHeight,
                        CalibratedCharWidthFt = calibration.CalibratedCharWidthFt,
                        CalibratedCharsPerLine = calibration.CalibratedCharsPerLine,
                        UsedAutoCalibration = calibration.UsedAutoCalibration,
                        CalibrationDebugLines = calibration.DebugLines
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

        private static CharWidthCalibrationResult CalibrateCharWidthFromSamples(IEnumerable<(int textLength, double groupColumnWidth)> samples)
        {
            var debugLines = new List<string>();
            var charsPerLineCandidates = new List<int>();
            var charWidthCandidates = new List<double>();
            var sampleList = samples == null
                ? new List<(int textLength, double groupColumnWidth)>()
                : samples.ToList();

            if (sampleList.Count > 0)
            {
                foreach (var sample in sampleList)
                {
                    int inferredExpectedLineCount = InferExpectedLineCountFromTextLength(sample.textLength);
                    if (sample.textLength < MinCalibrationTextLength || sample.groupColumnWidth <= 0 || inferredExpectedLineCount < 2)
                        continue;

                    int candidateCharsPerLine = (int)Math.Ceiling(sample.textLength / (double)inferredExpectedLineCount);
                    if (candidateCharsPerLine <= 0)
                        continue;

                    double candidateCharWidthFt = sample.groupColumnWidth / candidateCharsPerLine;
                    if (candidateCharWidthFt <= 0)
                        continue;

                    charsPerLineCandidates.Add(candidateCharsPerLine);
                    charWidthCandidates.Add(candidateCharWidthFt);
                    debugLines.Add("[CalibrationSample] textLength=" + sample.textLength
                        + " | inferredLines=" + inferredExpectedLineCount
                        + " | candidateCharsPerLine=" + candidateCharsPerLine
                        + " | candidateCharWidthFt=" + candidateCharWidthFt.ToString("0.######"));
                }
            }

            if (charsPerLineCandidates.Count >= 2)
            {
                double avgCharsPerLine = charsPerLineCandidates.Average();
                int calibratedCharsPerLine = Math.Max(1, (int)Math.Round(avgCharsPerLine, MidpointRounding.AwayFromZero));
                double calibratedCharWidthFt = charWidthCandidates.Count > 0
                    ? charWidthCandidates.Average()
                    : DefaultCharWidthEstimateFt;

                return new CharWidthCalibrationResult
                {
                    CalibratedCharWidthFt = calibratedCharWidthFt,
                    CalibratedCharsPerLine = calibratedCharsPerLine,
                    UsedAutoCalibration = true,
                    DebugLines = debugLines
                };
            }

            double? fallbackGroupColumnWidth = sampleList
                .Select(x => x.groupColumnWidth)
                .Where(w => w > 0)
                .Cast<double?>()
                .FirstOrDefault();
            int fallbackCharsPerLine = 12;
            if (fallbackGroupColumnWidth.HasValue && fallbackGroupColumnWidth.Value > 0 && DefaultCharWidthEstimateFt > 0)
            {
                fallbackCharsPerLine = Math.Max(
                    1,
                    (int)Math.Round(
                        fallbackGroupColumnWidth.Value / DefaultCharWidthEstimateFt,
                        MidpointRounding.AwayFromZero));
            }
            debugLines.Add("[CalibrationFallback] columnWidth="
                + (fallbackGroupColumnWidth.HasValue ? fallbackGroupColumnWidth.Value.ToString("0.######") : "(none)")
                + " | defaultCharWidthFt=" + DefaultCharWidthEstimateFt.ToString("0.######")
                + " | fallbackCharsPerLine=" + fallbackCharsPerLine);

            return new CharWidthCalibrationResult
            {
                CalibratedCharWidthFt = DefaultCharWidthEstimateFt,
                CalibratedCharsPerLine = fallbackCharsPerLine,
                UsedAutoCalibration = false,
                DebugLines = debugLines
            };
        }

        private static int EstimateVisualLineCountByCalibratedCharsPerLine(
            string text,
            int calibratedCharsPerLine,
            out int estimatedCharsPerLine)
        {
            string normalizedText = (text ?? "").Trim();
            int textLength = normalizedText.Length;
            if (textLength == 0)
            {
                estimatedCharsPerLine = 0;
                return 1;
            }
            if (calibratedCharsPerLine <= 0)
            {
                estimatedCharsPerLine = 0;
                return 1;
            }

            estimatedCharsPerLine = Math.Max(1, calibratedCharsPerLine);
            int lineCount = (int)Math.Ceiling(textLength / (double)estimatedCharsPerLine);
            return Math.Max(1, lineCount);
        }

        private static int InferExpectedLineCountFromTextLength(int textLength)
        {
            if (textLength <= 18) return 1;
            if (textLength <= 30) return 2;
            return 3;
        }

        private static double GetColumnWidth(TableSectionData body, int colIndex)
        {
            if (body == null || colIndex < 0)
                return 0;
            try
            {
                return body.GetColumnWidth(colIndex);
            }
            catch
            {
                try
                {
                    var method = body.GetType().GetMethod("GetColumnWidth", new[] { typeof(int) });
                    if (method != null)
                    {
                        object value = method.Invoke(body, new object[] { colIndex });
                        if (value is double d) return d;
                    }
                }
                catch { }

                return 0;
            }
        }

        private static double GetSectionRowHeight(TableSectionData section, int rowIndex)
        {
            if (section == null) return 0;
            try
            {
                return section.GetRowHeight(rowIndex);
            }
            catch
            {
                try
                {
                    var method = section.GetType().GetMethod("GetRowHeight", new[] { typeof(int) });
                    if (method != null)
                    {
                        object value = method.Invoke(section, new object[] { rowIndex });
                        if (value is double d) return d;
                    }
                }
                catch { }

                return 0;
            }
        }
    }
}
