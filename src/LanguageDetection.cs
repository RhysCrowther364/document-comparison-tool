using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace LocalDocumentCompare
{
    public enum LanguageStatus { Known, Unknown, Mixed }

    public sealed class SourceLanguage
    {
        public string Code { get; set; }
        public string DisplayName { get; set; }
        public LanguageStatus Status { get; set; }
    }

    /// <summary>
    /// Conservative language checks against actual visible text. Text is used
    /// only in memory; results contain language tags, never excerpts. Windows
    /// ELS runs locally and does not require OCR packs, downloads or a service.
    /// Its ranked candidates are not confidence probabilities. Agreement across
    /// independent, distributed samples is a heuristic, not proof of language.
    /// </summary>
    public static class LanguageDetection
    {
        private const int MaximumWindows = 9;
        private const int MaximumWindowCharacters = 1600;
        private static readonly Regex Words = new Regex(@"[\p{L}][\p{L}\p{M}'-]*", RegexOptions.CultureInvariant);

        public static SourceLanguage Detect(string text)
        {
            if (String.IsNullOrWhiteSpace(text)) return Unknown();
            try
            {
                // Do not infer language from the document's spelling locale or
                // headings. Inspect bounded samples spanning the supplied scope.
                int windows = Math.Min(MaximumWindows, Math.Max(3, text.Length / 800));
                var votes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                int recognized = 0;
                bool uncertainSample = false;
                using (NativeDetector detector = NativeDetector.Create())
                {
                    if (detector == null) return Unknown();
                    for (int index = 0; index < windows; index++)
                    {
                        int first = (int)((long)text.Length * index / windows);
                        int last = (int)((long)text.Length * (index + 1) / windows);
                        if (last - first > MaximumWindowCharacters)
                        {
                            if (index == windows - 1) first = last - MaximumWindowCharacters;
                            else
                            {
                                if (index != 0) first += (last - first - MaximumWindowCharacters) / 2;
                                last = first + MaximumWindowCharacters;
                            }
                        }
                        // Avoid cutting UTF-16 surrogate pairs at sample edges.
                        if (first > 0 && Char.IsLowSurrogate(text[first]) && Char.IsHighSurrogate(text[first - 1])) first++;
                        if (last < text.Length && last > first && Char.IsHighSurrogate(text[last - 1]) && Char.IsLowSurrogate(text[last])) last--;
                        if (last <= first) continue;
                        string sample = text.Substring(first, last - first).Normalize(NormalizationForm.FormC);
                        if (!HasEvidence(sample)) continue;
                        string code = NormalizeCode(detector.TopLanguage(sample));
                        if (code == null) { uncertainSample = true; continue; }
                        recognized++;
                        int count;
                        votes.TryGetValue(code, out count);
                        votes[code] = count + 1;
                    }
                }
                // Require repeated evidence for each side before calling a
                // document mixed. A single disagreeing sample stays unknown.
                if (votes.Count(v => v.Value >= 2) >= 2)
                    return new SourceLanguage { Code = "", DisplayName = "Mixed languages", Status = LanguageStatus.Mixed };
                if (uncertainSample || recognized < 3 || votes.Count != 1) return Unknown();
                string known = votes.Keys.First();
                string displayName;
                try { displayName = CultureInfo.GetCultureInfo(known).EnglishName; }
                catch (CultureNotFoundException) { return Unknown(); }
                return new SourceLanguage { Code = known, DisplayName = displayName, Status = LanguageStatus.Known };
            }
            catch (DllNotFoundException) { return Unknown(); }
            catch (EntryPointNotFoundException) { return Unknown(); }
            catch (BadImageFormatException) { return Unknown(); }
            catch (ExternalException) { return Unknown(); }
            catch (ArgumentException) { return Unknown(); }
        }

        private static bool HasEvidence(string sample)
        {
            int letters = 0, unspaced = 0, visible = 0;
            foreach (char value in sample)
            {
                if (!Char.IsWhiteSpace(value)) visible++;
                if (Char.IsLetter(value))
                {
                    letters++;
                    if (value >= '\u3040' && value <= '\u30ff' || value >= '\u3400' && value <= '\u9fff' ||
                        value >= '\uac00' && value <= '\ud7af' || value >= '\u0e00' && value <= '\u0e7f') unspaced++;
                }
            }
            if (letters < 120 || letters * 2 < visible) return false;
            if (unspaced * 2 >= letters) return true;
            MatchCollection words = Words.Matches(sample);
            if (words.Count < 24) return false;
            var distinct = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match word in words) distinct.Add(word.Value);
            return distinct.Count >= 12;
        }

        internal static string NormalizeCode(string value)
        {
            if (String.IsNullOrWhiteSpace(value)) return null;
            // Regional/script variants are the same source language for this
            // comparison guard (for example en-US/en-GB and zh-Hans/zh-Hant).
            string code = value.Split('-')[0].ToLowerInvariant();
            if (code == "und" || code == "zxx" || !Regex.IsMatch(code, "^[a-z]{2,3}$")) return null;
            return code;
        }

        private static SourceLanguage Unknown()
        {
            return new SourceLanguage { Code = "", DisplayName = "Undetermined", Status = LanguageStatus.Unknown };
        }

        private sealed class NativeDetector : IDisposable
        {
            private IntPtr services;

            public static NativeDetector Create()
            {
                IntPtr guid = Marshal.AllocHGlobal(16);
                IntPtr found = IntPtr.Zero;
                try
                {
                    Marshal.StructureToPtr(new Guid("CF7E00B1-909B-4d95-A8F4-611F7C377702"), guid, false);
                    var options = new EnumOptions { Size = (UIntPtr)Marshal.SizeOf(typeof(EnumOptions)), Guid = guid };
                    uint count;
                    int result = MappingGetServices(ref options, out found, out count);
                    if (result != 0 || count != 1) return null;
                    var detector = new NativeDetector { services = found };
                    found = IntPtr.Zero;
                    return detector;
                }
                finally
                {
                    if (found != IntPtr.Zero) MappingFreeServices(found);
                    Marshal.FreeHGlobal(guid);
                }
            }

            public string TopLanguage(string sample)
            {
                var bag = new PropertyBag { Size = (UIntPtr)Marshal.SizeOf(typeof(PropertyBag)) };
                int result = MappingRecognizeText(services, sample, (uint)sample.Length, 0, IntPtr.Zero, ref bag);
                if (result != 0) return null;
                try
                {
                    // Microsoft's detector returns one range containing a
                    // UTF-16 MULTI_SZ ranked list. No confidence score exists.
                    if (bag.Count != 1 || bag.Ranges == IntPtr.Zero) return null;
                    var range = (DataRange)Marshal.PtrToStructure(bag.Ranges, typeof(DataRange));
                    if (range.Data == IntPtr.Zero || range.DataSize < 4 || range.DataSize > 65536 || range.DataSize % 2 != 0) return null;
                    string candidates = Marshal.PtrToStringUni(range.Data, (int)range.DataSize / 2);
                    int end = candidates.IndexOf('\0');
                    return end > 0 ? candidates.Substring(0, end) : null;
                }
                finally { MappingFreePropertyBag(ref bag); }
            }

            public void Dispose()
            {
                if (services != IntPtr.Zero)
                {
                    IntPtr old = services;
                    services = IntPtr.Zero;
                    MappingFreeServices(old);
                }
            }
        }

        // Use pointer-sized fields for SIZE_T and pointer members; default
        // sequential alignment matches the Windows x86 and x64 ELS headers.
        [StructLayout(LayoutKind.Sequential)]
        private struct EnumOptions
        {
            public UIntPtr Size;
            public IntPtr Category, InputLanguage, OutputLanguage, InputScript, OutputScript, InputContentType, OutputContentType, Guid;
            public uint Flags;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct PropertyBag
        {
            public UIntPtr Size;
            public IntPtr Ranges;
            public uint Count;
            public IntPtr ServiceData;
            public uint ServiceDataSize;
            public IntPtr CallerData;
            public uint CallerDataSize;
            public IntPtr Context;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DataRange
        {
            public uint Start, End;
            public IntPtr Description;
            public uint DescriptionLength;
            public IntPtr Data;
            public uint DataSize;
            public IntPtr ContentType, ActionIds;
            public uint ActionsCount;
            public IntPtr ActionDisplayNames;
        }

        [DllImport("elscore.dll")]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        private static extern int MappingGetServices(ref EnumOptions options, out IntPtr services, out uint count);

        [DllImport("elscore.dll", CharSet = CharSet.Unicode)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        private static extern int MappingRecognizeText(IntPtr service, [MarshalAs(UnmanagedType.LPWStr)] string text,
            uint length, uint index, IntPtr options, ref PropertyBag bag);

        [DllImport("elscore.dll")]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        private static extern int MappingFreePropertyBag(ref PropertyBag bag);

        [DllImport("elscore.dll")]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        private static extern int MappingFreeServices(IntPtr services);
    }
}
