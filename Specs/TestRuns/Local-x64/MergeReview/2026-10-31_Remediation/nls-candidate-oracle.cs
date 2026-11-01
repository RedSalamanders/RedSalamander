using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

public static class RedSalamanderCandidateNlsOracle
{
    private const uint Upper = 0x200;
    private const int Less = 1, Equal = 2, Greater = 3;
    [DllImport("kernel32.dll", EntryPoint="LCMapStringEx", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)]
    private static extern int MapQuery(string locale, uint flags, IntPtr src, int cchSrc, IntPtr dst, int cchDst, IntPtr version, IntPtr reserved, IntPtr sort);
    [DllImport("kernel32.dll", EntryPoint="LCMapStringEx", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)]
    private static extern int MapFill(string locale, uint flags, IntPtr src, int cchSrc, IntPtr dst, int cchDst, IntPtr version, IntPtr reserved, IntPtr sort);
    [DllImport("kernel32.dll", EntryPoint="CompareStringOrdinal", ExactSpelling=true, SetLastError=true)]
    private static extern int Compare(IntPtr lhs, int cchLhs, IntPtr rhs, int cchRhs, [MarshalAs(UnmanagedType.Bool)] bool ignoreCase);

    private sealed class NativeSort : IComparer<int>
    {
        private readonly IntPtr _base;
        public int Errors;
        public NativeSort(IntPtr baseAddress) { _base = baseAddress; }
        public int Compare(int lhs, int rhs)
        {
            int c = RedSalamanderCandidateNlsOracle.Compare(IntPtr.Add(_base, lhs * 2), 1, IntPtr.Add(_base, rhs * 2), 1, true);
            if (c == 0) { System.Threading.Interlocked.Increment(ref Errors); return lhs.CompareTo(rhs); }
            return c == Equal ? lhs.CompareTo(rhs) : c == Less ? -1 : c == Greater ? 1 : lhs.CompareTo(rhs);
        }
    }
    private static string U(char c) => "U+" + ((int)c).ToString("X4", CultureInfo.InvariantCulture);
    private static string Esc(string s) { var b = new StringBuilder(); for (int i=0;i<s.Length;i++) { if(i!=0)b.Append(' '); b.Append(U(s[i])); } return b.ToString(); }
    private static int CompareStrings(string a, string b)
    {
        char[] ac=a.ToCharArray(), bc=b.ToCharArray(); GCHandle ap=GCHandle.Alloc(ac,GCHandleType.Pinned), bp=GCHandle.Alloc(bc,GCHandleType.Pinned);
        try { return Compare(ap.AddrOfPinnedObject(), ac.Length, bp.AddrOfPinnedObject(), bc.Length, true); }
        finally { bp.Free(); ap.Free(); }
    }
    private static bool TryCandidateKey(string s, out string key, out string reason)
    {
        char[] source=s.ToCharArray(), folded=new char[source.Length];
        for (int i=0;i<source.Length;i++)
        {
            char unit=source[i], mapped=unit;
            if (unit >= 'a' && unit <= 'z') mapped=(char)(unit-'a'+'A');
            else if (unit >= 0x80 && (unit < 0xD800 || unit > 0xDFFF))
            {
                GCHandle sp=GCHandle.Alloc(source,GCHandleType.Pinned), dp=GCHandle.Alloc(folded,GCHandleType.Pinned);
                try
                {
                    int need=MapQuery("",Upper,IntPtr.Add(sp.AddrOfPinnedObject(),i*2),1,IntPtr.Zero,0,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero);
                    if (need != 1) { key=""; reason="single-unit-mapping-length-or-error"; return false; }
                    if (MapFill("",Upper,IntPtr.Add(sp.AddrOfPinnedObject(),i*2),1,IntPtr.Add(dp.AddrOfPinnedObject(),i*2),1,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero) != 1)
                    { key=""; reason="single-unit-map-fill-error"; return false; }
                    mapped=folded[i];
                    if (Compare(IntPtr.Add(sp.AddrOfPinnedObject(),i*2),1,IntPtr.Add(dp.AddrOfPinnedObject(),i*2),1,true) != Equal)
                    { key=""; reason="single-unit-native-equivalence-reject"; return false; }
                }
                finally { dp.Free(); sp.Free(); }
            }
            folded[i]=mapped;
        }
        if (CompareStrings(s,new string(folded)) != Equal) { key=""; reason="whole-component-native-equivalence-reject"; return false; }
        key=new string(folded); reason="accepted"; return true;
    }
    private static Dictionary<string,object> Pair(string name,string a,string b,bool requireBoth)
    {
        bool ak=TryCandidateKey(a,out string ka,out string ar), bk=TryCandidateKey(b,out string kb,out string br);
        int native=CompareStrings(a,b); bool nativeEq=native==Equal, both=ak&&bk, keyEq=both&&String.Equals(ka,kb,StringComparison.Ordinal);
        return new Dictionary<string,object>{{"name",name},{"leftUtf16",Esc(a)},{"rightUtf16",Esc(b)},{"nativeRelation",native},{"nativeEqual",nativeEq},{"leftAccepted",ak},{"leftReason",ar},{"leftKey",ak?Esc(ka):"<none>"},{"rightAccepted",bk},{"rightReason",br},{"rightKey",bk?Esc(kb):"<none>"},{"bothKeysAccepted",both},{"keyEqualWhenAccepted",keyEq},{"parityChecked",both},{"parity",!both||nativeEq==keyEq},{"requiredAvailabilitySatisfied",!requireBoth||both}};
    }
    public static Dictionary<string,object> Run()
    {
        const int n=65536; char[] units=new char[n]; string[] keys=new string[n]; bool[] accepted=new bool[n]; int[] order=new int[n]; int rejects=0;
        GCHandle up=GCHandle.Alloc(units,GCHandleType.Pinned);
        try
        {
            IntPtr ub=up.AddrOfPinnedObject();
            for(int c=0;c<n;c++) { units[c]=(char)c; order[c]=c; accepted[c]=TryCandidateKey(new string((char)c,1),out keys[c],out _); if(!accepted[c])rejects++; }
            var sorter=new NativeSort(ub); Array.Sort(order,sorter);
            var classMismatch=new List<Dictionary<string,object>>(); int classes=0, nontrivial=0, acceptedInClasses=0, classesWithUnavailable=0;
            for(int start=0;start<n;)
            {
                int end=start+1; while(end<n && Compare(IntPtr.Add(ub,order[start]*2),1,IntPtr.Add(ub,order[end]*2),1,true)==Equal) end++;
                classes++; if(end-start>1)nontrivial++; bool hasKey=false, hasMissing=false; string firstKey=""; int first=-1;
                for(int p=start;p<end;p++) { int c=order[p]; if(!accepted[c]) {hasMissing=true;continue;} acceptedInClasses++; if(!hasKey){hasKey=true;firstKey=keys[c];first=c;} else if(!String.Equals(firstKey,keys[c],StringComparison.Ordinal)&&classMismatch.Count<40) classMismatch.Add(new Dictionary<string,object>{{"first",U((char)first)},{"other",U((char)c)},{"firstKey",Esc(firstKey)},{"otherKey",Esc(keys[c])}}); }
                if(hasMissing)classesWithUnavailable++; start=end;
            }
            var byKey=new Dictionary<string,List<int>>(StringComparer.Ordinal);
            for(int c=0;c<n;c++) if(accepted[c]) { if(!byKey.TryGetValue(keys[c],out List<int> members)){members=new List<int>();byKey.Add(keys[c],members);} members.Add(c); }
            var collisions=new List<Dictionary<string,object>>(); int collisionChecks=0;
            foreach(var members in byKey.Values) for(int i=0;i<members.Count;i++) for(int j=i+1;j<members.Count;j++) { int a=members[i],b=members[j];collisionChecks++;int relation=Compare(IntPtr.Add(ub,a*2),1,IntPtr.Add(ub,b*2),1,true);if(relation!=Equal&&collisions.Count<40)collisions.Add(new Dictionary<string,object>{{"left",U((char)a)},{"right",U((char)b)},{"key",Esc(keys[a])},{"nativeRelation",relation}}); }
            var pairs=new List<Dictionary<string,object>>{
                Pair("ASCII-case","a","A",true),
                Pair("Greek-final-sigma","\u03A3","\u03C2",false),
                Pair("Turkish-dotless-i","I","\u0131",false),
                Pair("Turkish-dotted-I","\u0130","i",false),
                Pair("sharp-s-expansion","\u00DF","SS",false),
                Pair("kelvin-sign","K","\u212A",false),
                Pair("Deseret-supplementary","\U00010428","\U00010400",true),
                Pair("Deseret-context","x\U00010428z","X\U00010400Z",true),
                Pair("Deseret-same-lower-control","\U00010428","\U00010428",true),
                Pair("emoji-identical","\U0001F4C1","\U0001F4C1",true),
                Pair("emoji-distinct","\U0001F4C1","\U0001F4C2",true),
                Pair("emoji-context","x\U0001F4C1z","X\U0001F4C1Z",true),
                Pair("unpaired-high-surrogate","\uD801","\uD801",true),
                Pair("unpaired-surrogate-distinct","\uD801","\uD802",true)
            };
            int badPairs=0; foreach(var p in pairs) if(!(bool)p["parity"]||!(bool)p["requiredAvailabilitySatisfied"])badPairs++;
            int invalidComparison=0; foreach(int c in order) if(!accepted[c]) { int self=Compare(IntPtr.Add(ub,c*2),1,IntPtr.Add(ub,c*2),1,true); if(self!=Equal)invalidComparison++; }
            return new Dictionary<string,object>{{"unitCount",n},{"acceptedUnits",n-rejects},{"rejectedUnits",rejects},{"nativeEquivalenceClasses",classes},{"nontrivialNativeEquivalenceClasses",nontrivial},{"acceptedUnitsWithinClasses",acceptedInClasses},{"classesContainingUnavailableKeys",classesWithUnavailable},{"acceptedKeyWithinNativeClassMismatches",classMismatch},{"keyCollisionPairChecks",collisionChecks},{"acceptedKeyRawUnequalCollisions",collisions},{"nativeSortComparisonErrors",sorter.Errors},{"unavailableUnitSelfComparisonErrors",invalidComparison},{"representativePairs",pairs},{"representativePairFailures",badPairs},{"pass",sorter.Errors==0&&invalidComparison==0&&classMismatch.Count==0&&collisions.Count==0&&badPairs==0}};
        }
        finally { up.Free(); }
    }
}
