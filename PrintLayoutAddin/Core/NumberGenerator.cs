using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PrintLayoutAddin.Core
{
    // ============================================================
    // Parameter objects for Simple and Advanced generation modes.
    // Both modes ultimately produce a List<string> of frame codes
    // that SttAssigner walks through in polyline-vertex order.
    // ============================================================

    public class SimpleGenerationParams
    {
        public string Prefix = "";
        public string Suffix = "";
        public int Start = 1;
        public int Step = 1;
        public HashSet<int> Skip = new HashSet<int>();
        public int Padding = 2;
        /// <summary>
        /// How many codes to generate. Usually = number of frames the user is targeting.
        /// </summary>
        public int Count = 0;
    }

    public class AdvancedGenerationParams
    {
        public string Prefix = "";
        public int Start = 1;
        public int End = 1;
        public int Step = 1;
        public int Padding = 4;
        public string Separator = ".";
        public string DefaultSubSuffix = "A";
        /// <summary>
        /// main-number -> ordered list of sub-suffixes. Example: 4204 -> ["A", "B"].
        /// Main numbers not present here get a single <c>DefaultSubSuffix</c>.
        /// </summary>
        public Dictionary<int, List<string>> Exceptions = new Dictionary<int, List<string>>();
    }

    public class ValidationResult
    {
        public bool Ok => Errors.Count == 0;
        public List<string> Errors = new List<string>();
        public List<string> Warnings = new List<string>();

        public override string ToString()
        {
            if (Ok && Warnings.Count == 0) return "OK";
            var parts = new List<string>();
            if (Errors.Count > 0) parts.Add("Errors:\n  " + string.Join("\n  ", Errors));
            if (Warnings.Count > 0) parts.Add("Warnings:\n  " + string.Join("\n  ", Warnings));
            return string.Join("\n\n", parts);
        }
    }

    public static class NumberGenerator
    {
        // ------------------------------------------------------------
        // Simple mode: classic linear numbering with optional Skip list.
        // Backward-compatible with the original SttOptions behaviour.
        // ------------------------------------------------------------
        /// <summary>Alias matching the upgrade-spec naming.</summary>
        public static List<string> GenerateSimpleNumbers(SimpleGenerationParams p) => GenerateSimple(p);

        public static List<string> GenerateSimple(SimpleGenerationParams p)
        {
            if (p == null) throw new ArgumentNullException(nameof(p));
            if (p.Step == 0) throw new ArgumentException("Step must not be zero.");
            if (p.Count < 0) throw new ArgumentException("Count must be non-negative.");

            var result = new List<string>(p.Count);
            int n = p.Start;
            // advance past any Skip values that happen to equal the very first number
            while (p.Skip.Contains(n)) n += p.Step;

            for (int i = 0; i < p.Count; i++)
            {
                string numStr = p.Padding > 0 ? n.ToString("D" + p.Padding) : n.ToString();
                result.Add((p.Prefix ?? "") + numStr + (p.Suffix ?? ""));

                // compute the next candidate, skipping Skip values
                n += p.Step;
                while (p.Skip.Contains(n)) n += p.Step;
            }
            return result;
        }

        // ------------------------------------------------------------
        // Advanced mode: Main number in [Start .. End] (inclusive).
        // For each main number, emit (Prefix)(MainPadded)(Separator)(SubSuffix),
        // where SubSuffix is DefaultSubSuffix unless the main number is listed
        // in Exceptions, in which case every suffix in that list is emitted in
        // the order given.
        // ------------------------------------------------------------
        /// <summary>Alias matching the upgrade-spec naming.</summary>
        public static List<string> GenerateAdvancedNumbers(AdvancedGenerationParams p) => GenerateAdvanced(p);

        public static List<string> GenerateAdvanced(AdvancedGenerationParams p)
        {
            if (p == null) throw new ArgumentNullException(nameof(p));
            if (p.Step == 0) throw new ArgumentException("Step must not be zero.");

            var result = new List<string>();
            // support both increasing and decreasing ranges
            bool ascending = p.End >= p.Start;
            if (ascending && p.Step < 0) throw new ArgumentException("Step direction mismatched with Start/End.");
            if (!ascending && p.Step > 0) throw new ArgumentException("Step direction mismatched with Start/End.");

            for (int main = p.Start;
                 ascending ? main <= p.End : main >= p.End;
                 main += p.Step)
            {
                string mainStr = p.Padding > 0 ? main.ToString("D" + p.Padding) : main.ToString();
                if (p.Exceptions != null && p.Exceptions.TryGetValue(main, out var suffixes) && suffixes != null && suffixes.Count > 0)
                {
                    foreach (var s in suffixes)
                        result.Add((p.Prefix ?? "") + mainStr + (p.Separator ?? "") + (s ?? ""));
                }
                else
                {
                    result.Add((p.Prefix ?? "") + mainStr + (p.Separator ?? "") + (p.DefaultSubSuffix ?? ""));
                }
            }
            return result;
        }

        // ------------------------------------------------------------
        // Parse a multiline exception rules text in format:
        //   4204 = A, B
        //   4205=A,B,C
        // Whitespace around "=" and "," is tolerated. Blank lines and lines
        // starting with '#' are skipped (treated as comments).
        // Any malformed line is reported in <paramref name="parseErrors"/>
        // as "Line N: <reason>" so the UI can show it inline.
        // ------------------------------------------------------------
        public static Dictionary<int, List<string>> ParseExceptionRules(string text, out List<string> parseErrors)
        {
            parseErrors = new List<string>();
            var dict = new Dictionary<int, List<string>>();
            if (string.IsNullOrWhiteSpace(text)) return dict;

            var lines = text.Replace("\r\n", "\n").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                var raw = lines[i];
                var line = raw.Trim();
                if (line.Length == 0) continue;
                if (line.StartsWith("#")) continue;

                // Split on the FIRST '=' only — suffixes themselves shouldn't contain '=' but be defensive.
                int eq = line.IndexOf('=');
                if (eq < 0)
                {
                    parseErrors.Add($"Line {i + 1}: missing '=' separator.");
                    continue;
                }
                var leftRaw = line.Substring(0, eq).Trim();
                var rightRaw = line.Substring(eq + 1).Trim();

                if (!int.TryParse(leftRaw, out int main))
                {
                    parseErrors.Add($"Line {i + 1}: '{leftRaw}' is not an integer main number.");
                    continue;
                }
                if (rightRaw.Length == 0)
                {
                    parseErrors.Add($"Line {i + 1}: suffix list is empty.");
                    continue;
                }

                var suffixes = rightRaw
                    .Split(',')
                    .Select(s => s.Trim())
                    .ToList();

                if (suffixes.Any(s => s.Length == 0))
                {
                    parseErrors.Add($"Line {i + 1}: contains an empty suffix (check stray commas).");
                    continue;
                }

                if (dict.ContainsKey(main))
                {
                    parseErrors.Add($"Line {i + 1}: main number {main} appears more than once.");
                    continue;
                }
                dict[main] = suffixes;
            }
            return dict;
        }

        // ------------------------------------------------------------
        // Validate a final code list before applying to frames.
        // Checks: empty list, empty codes, duplicate codes (unless allowed),
        // optional expected count match.
        // ------------------------------------------------------------
        public static ValidationResult ValidateNumberList(
            List<string> codes,
            int? expectedCount = null,
            bool allowDuplicates = false)
        {
            var r = new ValidationResult();
            if (codes == null || codes.Count == 0)
            {
                r.Errors.Add("Code list is empty.");
                return r;
            }

            for (int i = 0; i < codes.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(codes[i]))
                    r.Errors.Add($"Code #{i + 1} is empty.");
            }

            if (!allowDuplicates)
            {
                var dup = codes
                    .GroupBy(c => c ?? "", StringComparer.OrdinalIgnoreCase)
                    .Where(g => !string.IsNullOrEmpty(g.Key) && g.Count() > 1)
                    .Select(g => g.Key)
                    .ToList();
                foreach (var d in dup)
                    r.Errors.Add($"Duplicate code: '{d}'.");
            }

            if (expectedCount.HasValue && expectedCount.Value != codes.Count)
            {
                r.Errors.Add(
                    $"Count mismatch: expected {expectedCount.Value} code(s), got {codes.Count}.");
            }

            return r;
        }

        // ------------------------------------------------------------
        // Parse a comma/space/semicolon-separated list of integers (for the
        // legacy "Skip" field). Invalid tokens are silently ignored — matches
        // the old behaviour so saved registry values still work.
        // ------------------------------------------------------------
        public static HashSet<int> ParseSkipList(string text)
        {
            var set = new HashSet<int>();
            if (string.IsNullOrWhiteSpace(text)) return set;
            foreach (var part in text.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (int.TryParse(part.Trim(), out var n)) set.Add(n);
            }
            return set;
        }
    }
}
