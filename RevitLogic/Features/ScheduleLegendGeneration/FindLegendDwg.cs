using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RevitLogic.Features.ScheduleLegendGeneration
{
    /// <summary>
    /// 依族群名稱在指定根目錄下遞迴搜尋對應的 .dwg 檔（族群名.dwg），產出每筆族群的對應路徑與狀態。
    /// 功能對應 Dynamo Schedule2Legend/Node03_FindLegendDwg.py。
    /// </summary>
    public static class FindLegendDwg
    {
        private static readonly char[] InvalidFilenameChars = { '\\', '/', ':', '*', '?', '"', '<', '>', '|' };

        /// <summary>單一明細表的 DWG 對應摘要。</summary>
        public sealed class ScheduleSummary
        {
            public int GroupsTotal { get; set; }
            public int Found { get; set; }
            public int Missing { get; set; }
            public int Duplicates { get; set; }
        }

        /// <summary>單一族群名稱的 DWG 對應結果。</summary>
        public sealed class DwgMatchRecord
        {
            public string ScheduleName { get; set; }
            public string GroupName { get; set; }
            public string ExpectedFilenameTry { get; set; }
            public string Status { get; set; }  // "found" | "duplicate" | "missing"
            public List<string> DwgPaths { get; set; }
        }

        /// <summary>FindLegendDwg 的整體結果。</summary>
        public sealed class FindLegendDwgResult
        {
            public bool Ok { get; set; }
            public string Error { get; set; }
            public string Root { get; set; }
            public int DwgTotalFoundUnderRoot { get; set; }
            public Dictionary<string, ScheduleSummary> SummaryBySchedule { get; set; }
            public List<DwgMatchRecord> Records { get; set; }
        }

        /// <summary>
        /// 以 Node02 的 group_map_by_schedule 與根目錄路徑，找出每個族群名稱對應的 .dwg 檔案。
        /// </summary>
        /// <param name="groupMapBySchedule">明細表名稱 → 族群名稱清單（Node02 的 group_map_by_schedule）。</param>
        /// <param name="rootFolder">搜尋 .dwg 的根目錄；若為 null 或空則使用預設寫死路徑。</param>
        public static FindLegendDwgResult Find(
            IReadOnlyDictionary<string, List<string>> groupMapBySchedule,
            string rootFolder = null)
        {
            var result = new FindLegendDwgResult
            {
                SummaryBySchedule = new Dictionary<string, ScheduleSummary>(),
                Records = new List<DwgMatchRecord>()
            };

            if (groupMapBySchedule == null || groupMapBySchedule.Count == 0)
            {
                result.Ok = true;
                result.Root = rootFolder ?? "";
                result.Error = null;
                return result;
            }

            string root = !string.IsNullOrWhiteSpace(rootFolder)
                ? rootFolder.Trim()
                : @"C:\Users\sunny\NoteSystem\PublicNotes\CivilWorkMaterials\REVIT\元件\00_使用中圖例";
            // @"D:/PublicNotes/CivilWorkMaterials/REVIT/元件/00_使用中圖例"
            // @"C:\Users\sunny\NoteSystem\PublicNotes\CivilWorkMaterials\REVIT\元件\00_使用中圖例"

            if (!Directory.Exists(root))
            {
                result.Ok = false;
                result.Error = "Folder not found: " + root;
                result.Root = root;
                return result;
            }

            result.Root = root;

            // 遞迴建立檔名(小寫) → 完整路徑清單 的索引
            var dwgIndex = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var allDwgPaths = new List<string>();
            try
            {
                foreach (string path in Directory.GetFiles(root, "*.dwg", SearchOption.AllDirectories))
                {
                    allDwgPaths.Add(path);
                    string fn = Path.GetFileName(path);
                    if (string.IsNullOrEmpty(fn)) continue;
                    string key = fn.ToLowerInvariant();
                    if (!dwgIndex.TryGetValue(key, out var list))
                    {
                        list = new List<string>();
                        dwgIndex[key] = list;
                    }
                    list.Add(path);
                }
            }
            catch (Exception ex)
            {
                result.Ok = false;
                result.Error = "Error indexing folder: " + ex.Message;
                return result;
            }

            result.DwgTotalFoundUnderRoot = allDwgPaths.Count;

            foreach (var kv in groupMapBySchedule)
            {
                string scheduleName = kv.Key ?? "";
                List<string> groups = kv.Value;
                if (groups == null) continue;

                var summary = new ScheduleSummary { GroupsTotal = 0, Found = 0, Missing = 0, Duplicates = 0 };

                foreach (string g in groups)
                {
                    string gname = (g ?? "").Trim();
                    if (string.IsNullOrEmpty(gname)) continue;

                    summary.GroupsTotal++;
                    var candidateFilenames = GetCandidateFilenames(gname);
                    List<string> matched = null;
                    string usedFilename = null;

                    foreach (string cand in candidateFilenames)
                    {
                        string key = cand.ToLowerInvariant();
                        if (dwgIndex.TryGetValue(key, out var paths) && paths != null && paths.Count > 0)
                        {
                            matched = paths;
                            usedFilename = cand;
                            break;
                        }
                    }

                    string status;
                    if (matched != null && matched.Count > 0)
                    {
                        summary.Found++;
                        if (matched.Count > 1)
                        {
                            summary.Duplicates++;
                            status = "duplicate";
                        }
                        else
                            status = "found";
                    }
                    else
                    {
                        summary.Missing++;
                        status = "missing";
                        usedFilename = candidateFilenames.Count > 0 ? candidateFilenames[0] : "";
                    }

                    result.Records.Add(new DwgMatchRecord
                    {
                        ScheduleName = scheduleName,
                        GroupName = gname,
                        ExpectedFilenameTry = usedFilename ?? "",
                        Status = status,
                        DwgPaths = matched ?? new List<string>()
                    });
                }

                result.SummaryBySchedule[scheduleName] = summary;
            }

            result.Ok = true;
            result.Error = null;
            return result;
        }

        private static string SanitizeFilename(string name)
        {
            if (string.IsNullOrEmpty(name)) return name;
            string out_ = name;
            foreach (char ch in InvalidFilenameChars)
                out_ = out_.Replace(ch, '_');
            return out_;
        }

        private static List<string> GetCandidateFilenames(string groupName)
        {
            string g = (groupName ?? "").Trim();
            var cands = new List<string>();
            if (g.Length == 0) return cands;

            cands.Add(g + ".dwg");
            cands.Add(SanitizeFilename(g) + ".dwg");
            cands.Add(g.TrimEnd(' ', '.') + ".dwg");
            cands.Add(SanitizeFilename(g.TrimEnd(' ', '.')) + ".dwg");

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var uniq = new List<string>();
            foreach (string x in cands)
            {
                string k = x.ToLowerInvariant();
                if (seen.Contains(k)) continue;
                seen.Add(k);
                uniq.Add(x);
            }
            return uniq;
        }
    }
}
