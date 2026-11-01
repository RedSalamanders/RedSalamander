using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

public static class RedSalamanderNlsOracle
{
    private const uint Upper = 0x200;
    private const int Less = 1, Equal = 2, Greater = 3;
    [DllImport("kernel32.dll", EntryPoint="LCMapStringEx", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)]
    private static extern int MapQuery(string locale, uint flags, IntPtr src, int cchSrc, IntPtr dst, int cchDst,
                                       IntPtr version, IntPtr reserved, IntPtr sort);
    [DllImport("kernel32.dll", EntryPoint="LCMapStringEx", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)]
    private static extern int MapFill(string locale, uint flags, IntPtr src, int cchSrc, IntPtr dst, int cchDst,
                                      IntPtr version, IntPtr reserved, IntPtr sort);
    [DllImport("kernel32.dll", EntryPoint="CompareStringOrdinal", ExactSpelling=true, SetLastError=true)]
    private static extern int Compare(IntPtr lhs, int cchLhs, IntPtr rhs, int cchRhs,
                                      [MarshalAs(UnmanagedType.Bool)] bool ignoreCase);

    private sealed class NativeSort : IComparer<int>
    {
        private readonly IntPtr _base;
        public int Errors;
        public NativeSort(IntPtr pinnedBase) { _base = pinnedBase; }
        public int Compare(int lhs, int rhs)
        {
            int result = RedSalamanderNlsOracle.Compare(IntPtr.Add(_base, lhs * 2), 1,
                                                       IntPtr.Add(_base, rhs * 2), 1, true);
            if (result == 0)
            {
                System.Threading.Interlocked.Increment(ref Errors);
                return lhs.CompareTo(rhs);
            }
            return result == Equal ? lhs.CompareTo(rhs) : result == Less ? -1 : result == Greater ? 1 : lhs.CompareTo(rhs);
        }
    }

    private static string U(char c) { return "U+" + ((int)c).ToString("X4", CultureInfo.InvariantCulture); }
    private static string Esc(string s)
    {
        var b = new StringBuilder();
        for (int i = 0; i < s.Length; ++i) { if (i != 0) b.Append(' '); b.Append(U(s[i])); }
        return b.ToString();
    }
    private static int CompareStrings(string lhs, string rhs)
    {
        char[] a = lhs.ToCharArray(), b = rhs.ToCharArray();
        GCHandle ap = GCHandle.Alloc(a, GCHandleType.Pinned), bp = GCHandle.Alloc(b, GCHandleType.Pinned);
        try { return Compare(ap.AddrOfPinnedObject(), a.Length, bp.AddrOfPinnedObject(), b.Length, true); }
        finally { bp.Free(); ap.Free(); }
    }
    private static Dictionary<string, object> Map(string value)
    {
        char[] source = value.ToCharArray(), dest = new char[Math.Max(1, value.Length * 3)];
        GCHandle sp = GCHandle.Alloc(source, GCHandleType.Pinned), dp = GCHandle.Alloc(dest, GCHandleType.Pinned);
        try
        {
            int required = MapQuery("", Upper, sp.AddrOfPinnedObject(), source.Length, IntPtr.Zero, 0, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            int queryError = required == 0 ? Marshal.GetLastWin32Error() : 0;
            int written = 0, fillError = 0;
            if (required > 0)
            {
                if (dest.Length < required) { Array.Resize(ref dest, required); dp.Free(); dp = GCHandle.Alloc(dest, GCHandleType.Pinned); }
                written = MapFill("", Upper, sp.AddrOfPinnedObject(), source.Length, dp.AddrOfPinnedObject(), required,
                                 IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                fillError = written == 0 ? Marshal.GetLastWin32Error() : 0;
            }
            string mapped = written > 0 ? new string(dest, 0, written) : "";
            return new Dictionary<string, object>
            {
                ["sourceUtf16Units"] = source.Length, ["required"] = required, ["written"] = written,
                ["acceptedByProductionRule"] = required == source.Length && written == required,
                ["mappedUtf16Units"] = Esc(mapped), ["queryError"] = queryError, ["fillError"] = fillError
            };
        }
        finally { dp.Free(); sp.Free(); }
    }
    private static Dictionary<string, object> Pair(string name, string lhs, string rhs)
    {
        var left = Map(lhs); var right = Map(rhs); int native = CompareStrings(lhs, rhs);
        bool accepted = (bool)left["acceptedByProductionRule"] && (bool)right["acceptedByProductionRule"];
        bool keyEqual = accepted && (string)left["mappedUtf16Units"] == (string)right["mappedUtf16Units"];
        bool nativeEqual = native == Equal;
        return new Dictionary<string, object>
        {
            ["name"] = name, ["leftUtf16"] = Esc(lhs), ["rightUtf16"] = Esc(rhs),
            ["compareStringOrdinal"] = native, ["nativeEqual"] = nativeEqual,
            ["leftMap"] = left, ["rightMap"] = right, ["bothKeysAccepted"] = accepted,
            ["keyEqualWhenAccepted"] = keyEqual, ["parityChecked"] = accepted,
            ["parity"] = !accepted || nativeEqual == keyEqual
        };
    }

    public static Dictionary<string, object> Run()
    {
        const int n = 65536;
        char[] units = new char[n], keys = new char[n];
        bool[] accepted = new bool[n];
        int[] order = new int[n];
        int mapErrors = 0, fillErrors = 0, rejectedByLength = 0;
        GCHandle up = GCHandle.Alloc(units, GCHandleType.Pinned), kp = GCHandle.Alloc(keys, GCHandleType.Pinned);
        try
        {
            IntPtr ub = up.AddrOfPinnedObject(), kb = kp.AddrOfPinnedObject();
            for (int c = 0; c < n; ++c)
            {
                units[c] = (char)c; order[c] = c;
                IntPtr src = IntPtr.Add(ub, c * 2);
                int required = MapQuery("", Upper, src, 1, IntPtr.Zero, 0, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                if (required != 1) { if (required == 0) ++mapErrors; else ++rejectedByLength; continue; }
                if (MapFill("", Upper, src, 1, IntPtr.Add(kb, c * 2), 1, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero) != 1)
                { ++fillErrors; continue; }
                accepted[c] = true;
            }

            var nativeSort = new NativeSort(ub);
            Array.Sort(order, nativeSort);
            var classMismatches = new List<Dictionary<string, object>>();
            int classes = 0, nontrivialClasses = 0, acceptedUnits = 0;
            for (int start = 0; start < n;)
            {
                int end = start + 1;
                while (end < n && Compare(IntPtr.Add(ub, order[start] * 2), 1, IntPtr.Add(ub, order[end] * 2), 1, true) == Equal) ++end;
                ++classes; if (end - start > 1) ++nontrivialClasses;
                bool haveKey = false; char firstKey = '\0'; int firstUnit = -1;
                for (int p = start; p < end; ++p)
                {
                    int c = order[p]; if (!accepted[c]) continue; ++acceptedUnits;
                    if (!haveKey) { haveKey = true; firstKey = keys[c]; firstUnit = c; }
                    else if (keys[c] != firstKey && classMismatches.Count < 40)
                        classMismatches.Add(new Dictionary<string, object>
                        {
                            ["classFirst"] = U((char)firstUnit), ["other"] = U((char)c),
                            ["firstKey"] = U(firstKey), ["otherKey"] = U(keys[c])
                        });
                }
                start = end;
            }

            var unitsByKey = new Dictionary<int, List<int>>();
            for (int c = 0; c < n; ++c)
            {
                if (!accepted[c]) continue;
                int key = keys[c];
                if (!unitsByKey.TryGetValue(key, out List<int> members))
                {
                    members = new List<int>();
                    unitsByKey.Add(key, members);
                }
                members.Add(c);
            }
            var collisions = new List<Dictionary<string, object>>();
            int collisionChecks = 0;
            foreach (List<int> members in unitsByKey.Values)
            {
                for (int left = 0; left < members.Count; ++left)
                {
                    for (int right = left + 1; right < members.Count; ++right)
                    {
                        int a = members[left], b = members[right];
                        ++collisionChecks;
                        int relation = Compare(IntPtr.Add(ub, a * 2), 1, IntPtr.Add(ub, b * 2), 1, true);
                        if (relation != Equal && collisions.Count < 40)
                            collisions.Add(new Dictionary<string, object>
                            {
                                ["first"] = U((char)a), ["other"] = U((char)b), ["key"] = U(keys[a]),
                                ["nativeRelation"] = relation
                            });
                    }
                }
            }

            var pairs = new List<Dictionary<string, object>>
            {
                Pair("ASCII-case", "a", "A"),
                Pair("Greek-final-sigma", "\u03A3", "\u03C2"),
                Pair("Greek-final-sigma-context", "\u039F\u03A3", "\u03BF\u03C2"),
                Pair("Greek-normal-sigma-context", "\u039F\u03A3", "\u03BF\u03C3"),
                Pair("Turkish-dotless-i", "I", "\u0131"),
                Pair("Turkish-dotted-I", "\u0130", "i"),
                Pair("sharp-s-expansion", "\u00DF", "SS"),
                Pair("capital-sharp-s", "\u1E9E", "\u00DF"),
                Pair("kelvin-sign", "K", "\u212A"),
                Pair("canonical-composition-control", "A\u030A", "\u00C5"),
                Pair("Deseret-supplementary", "\U00010428", "\U00010400"),
                Pair("Deseret-context", "x\U00010428z", "X\U00010400Z"),
                Pair("unpaired-high-surrogate", "\uD801", "\uD801"),
                Pair("unpaired-surrogate-distinct", "\uD801", "\uD802")
            };
            int contextualMismatches = 0;
            foreach (var p in pairs) if (!(bool)p["parity"]) ++contextualMismatches;
            return new Dictionary<string, object>
            {
                ["utf16SingleUnitCount"] = n, ["acceptedUnits"] = acceptedUnits,
                ["requiredLengthOtherThanOne"] = rejectedByLength, ["mapQueryErrors"] = mapErrors,
                ["mapFillErrors"] = fillErrors, ["compareErrorsDuringSort"] = nativeSort.Errors,
                ["nativeEquivalenceClasses"] = classes, ["nontrivialNativeEquivalenceClasses"] = nontrivialClasses,
                ["acceptedKeyEquivalenceClassMismatches"] = classMismatches,
                ["acceptedKeyCollisionChecks"] = collisionChecks, ["acceptedKeyRawUnequalCollisions"] = collisions,
                ["contextPairs"] = pairs, ["contextPairMismatches"] = contextualMismatches,
                ["pass"] = mapErrors == 0 && fillErrors == 0 && nativeSort.Errors == 0 &&
                          classMismatches.Count == 0 && collisions.Count == 0 && contextualMismatches == 0
            };
        }
        finally { kp.Free(); up.Free(); }
    }
}