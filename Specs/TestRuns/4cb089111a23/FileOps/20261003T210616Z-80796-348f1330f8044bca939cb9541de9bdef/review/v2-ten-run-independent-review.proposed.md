# Independent closed v2 review — scoped whole-product observations

Read only `v2-pairs-20261003T210306Z-487089b39880442e8158b96b95fc97dd/performance-comparison-v2.json` (391045bytes), containing10 captures,12 timing scenarios and13 resource comparisons. Root reports the serial collection and corrected comparator completed. No native/raw/output trees, source/PE hashing, tests, builds or desktop actions were performed. No performance-comparison-v2.md exists; JSON is the analysis input. Independently recomputed75 aggregate median/percentage/resource-delta checks: zero discrepancies. Individual native samples are not embedded in this compact analysis; its per-capture medians are the finest timing values reviewed here.

Suggested disposition: retain a successful v2 collection and narrowly labeled whole-product fixed-fixture aggregate comparison, with all six positive timing medians and handle+1 explicitly reviewed. Do not claim universal/worst-case nonregression, zero resource retention or native-change-only attribution. Applying the existing acceptance threshold and resolving resource/outlier interpretation remain root decisions; this review grants no waiver or automatic pass. Original v1 adverses/nonqualification remain separate and unaccepted. Do not pool, replace or rebaseline them with v2.

Baseline receipt2fb4d7aabad65f136080922bd94c2999d10e6d13f85350b39b2806632c9405c9/source73e48c8e89dd4061b6a681835bbd03d2ff24796526b6989492b4a141a71b93dd, candidate receipt256d18be002b07877b2e00cd18181d35fc30a3f3a702e6c436b8f8fc1ff35de7/source2ba4d8c2e89fc3fd4760f03251d8b69eb1cfc0d24338bf727a55964bc85d1fc0 are retained analysis identities, not fresh re-attestations. Compiler digest matches270491a7822f6732b3fab0f322710a3d063368ace4d2ee50f8bea88e1401cb50, but vcpkg differs: baseline7ffe9209f80ebb665be746dd1dad72edf5fc0e2257115ea63e316b1aac2dd932 versus candidated14885142b565a1097b95a5d2c1d35613a0f578e54f84a9c36111e16a56f9eb5. Any result includes dependencies and other whole-product changes; it cannot isolate I27 native changes. Fixtureb0848be7ab47d0d4fd4804b858e8a7729242a3ad8ba03468a151cac0878377d4, machine4cb089111a23. Typed controls declare144DPI, popup/issues1140x720pixels, scan720x1080pixels, exact eight renders, workload tokens3899185283365874287 and10980240383857445081. Comparator equality of those controls is reported; this compact review does not independently replay or audit all native witnesses. The reviewed blank candidate commit exception remains missing native metadata, never an imputed commit.

Timing estimator is the median of five per-capture medians, separately for each variant. It is not the median of signed pair differences, a pooled native-sample median, p95, confidence interval or FPS. The five captures contain5 samples/scenario except churn32 samples/capture. Odd pairs run baseline first, even pairs candidate first as indicated by retained run IDs. Run IDs are not report-end timestamps; this compact JSON alone does not prove process nonoverlap or quiet closure. Root-owned controller/archive evidence supplies those checks.

All12 scenarios and all60 signed paired per-capture median differences follow. Positive values remain adverse observations even in rows with a favorable aggregate median. Units microseconds; pair columns1..5 in chronological pair order.

| Scenario | Baseline->candidate aggregate median us | Aggregate delta% | Signed pair1/2/3/4/5 us |
|---|---:|---:|---|
|GraphOpenRedraw8Us|66651->66678|+0.0405095|+27 /-2772 /-18 /-69479 /+5895|
|GraphCollapsedRedraw8Us|66663->66670|+0.0105006|+37 /+4932 /+153 /-135834 /-99|
|ChurnUpdate8Us|66678->66686.5|+0.0127478|+621.5 /+80 /+63.5 /-3184 /-245|
|IssuesDetailRedraw8Us|66166->65984|-0.275066|+252 /-2571 /+46 /+670 /-318|
|Scan clean compare|66033->66240|+0.313480|-1296 /+524 /+193 /+207 /+6260|
|Scan clean change-case|66552->66412|-0.210362|-140 /+1947 /-454 /-5879 /-143|
|Scan clean change-attributes|66443->62322|-6.202309|-9186 /-4060 /-78 /-12824 /-691|
|Scan clean file-list|66552->66375|-0.265957|-137 /-263 /-13155 /-177 /+2498|
|Scan changing compare|66449->66359|-0.135442|+2336 /-265 /-3817 /-643 /-90|
|Scan changing change-case|66349->66489|+0.211005|+15221 /-9837 /+98 /-4125 /+4170|
|Scan changing change-attributes|64132->66362|+3.477203|-7251 /-5229 /+2230 /+5426 /-2780|
|Scan changing file-list|67993->66520|-2.166399|-11826 /-973 /+864 /+11276 /-19887|

The largest aggregate adverse is changing attributes+2230us/+3.477203%; raw per-capture baseline medians=[73779,58477,64132,60963,67888], candidate=[66528,53248,66362,66389,65108]. Only pairs3/4 worsen (+3.4772%/+8.9005%); this does not establish a consistent causal slowdown. It also does not prove absence of one. Clean attributes is the only timing scenario favorable in all five pairs.

Important retained outliers: graph-open baseline pair4=136036 versus candidate66557us (-51.074%); collapsed baseline pair4=192556 versus candidate56722us (-70.543%). Their dramatic gains do not make candidate universal improvement. Candidate graph-open pair5=72536 versus66641us (+8.84591%); collapsed pair2=71897 versus66965us (+7.36504%). Clean compare pair5=66339 versus60079us (+10.4196%). Changing case pair1=66489 versus51268us (+29.6891%, largest paired percentage); changing file-list pair4=79269 versus67993us (+16.5841%, +11276us). Changing compare pair1+3.64493%, clean case pair2+2.91101%, clean file-list pair5+3.76613%, issue detail pair4+1.01260%, churn maximum pair1+0.940683%. Every original pair observation is retained; none is dropped as noise or replaced by the new median. Maximum per-capture median overall is baseline collapsed192556us; candidate maximum is changing file-list79269us. These are capture medians, not maxima of individual native samples.

All13 resource metrics and all65 paired signed differences:

| Resource | Baseline->candidate median | Signed pair1/2/3/4/5 differences |
|---|---:|---|
|peak observed private bytes|275791872->271691776|-507904 /-4206592 /-7585792 /+4214784 /-3080192|
|peak observed working bytes|194408448->192151552|+471040 /-2400256 /-2256896 /-450560 /-2387968|
|closed-minus-cold private bytes|-6602752->-7000064|+761856 /-397312 /+3424256 /-15736832 /-2445312|
|scenario UI CPU us|2484375->2359375|-109375 /-484375 /+234375 /+250000 /+140625|
|scenario process CPU us|3140625->2984375|-31250 /-734375 /-187500 /+593750 /+640625|
|peak observed handles|2687->2687|+40 /0 /0 /-18 /-2|
|peak observed GDI|59->59|0 /0 /0 /0 /0|
|peak observed USER|78->78|0 /0 /0 /-1 /0|
|closed-minus-cold handles|36->37|+1 /+1 /0 /-230 /+2|
|closed-minus-cold hosts|0->0|0 /0 /0 /0 /0|
|closed-minus-cold GDI|0->0|0 /0 /0 /0 /0|
|closed-minus-cold USER|3->3|0 /0 /0 /-2 /0|
|closed-minus-cold attachments|0->0|0 /0 /0 /0 /0|

Private peak median improves4100096bytes/3.910MiB (-1.486663%), working2256896bytes (-1.160904%). Both are maxima of labelled observations, not continuous process/GPU peaks. Candidate private observation peaks=[275283968,272519168,268963840,271691776,271581184]; baseline=[275791872,276725760,276549632,267476992,274661376]. Candidate peak pair4 is+4214784bytes despite favorable median; maximum across captures baseline276725760 versus candidate275283968. Working pair1 is+471040bytes despite favorable median; maximum across captures195010560 versus193110016bytes. These adverse samples remain relevant to a resource-budget decision.

Candidate closed-minus-cold private is negative in every run[-9187328,-7000064,-3358720,-7196672,-4993024]; baseline=[-9949184,-6602752,-6782976,+8540160,-2547712]. Candidate releases less than baseline in pairs1/3 by761856/3424256bytes, while baseline pair4 shows positive8540160bytes. Cold/closed deltas are not automatic leaks or automatic recovery proof. The compact analysis does not expose the absolute cold/closed levels needed to explain that baseline outlier or any lifetime boundary.

CPU aggregate medians improve UI125000us (-5.031447%) and process156250us (-4.975124%). UI worsens pairs3/4/5 by+11.0294%/+10.0629%/+5.11364%; process worsens pairs4/5 by+18.9055%/+19.9029%. Candidate UI totals=[2296875,2187500,2359375,2734375,2890625], baseline=[2406250,2671875,2125000,2484375,2750000]. Candidate process=[2984375,2703125,2734375,3734375,3859375], baseline=[3015625,3437500,2921875,3140625,3218750]. Maximum candidate UI2890625/process3859375 exceed baseline maxima2750000/3437500. These are scenario CPU totals, not elapsed-normalized CPU rates or statistically significant gains; compact analysis lacks scenario wall duration for such normalization. OS granularity remains acknowledged, with no positive result waived merely by that label.

Closed handles are genuinely nonzero relative to cold. Baseline raw deltas=[36,36,37,265,35], candidate=[37,37,37,35,37]; signed median comparison+1 is not a zero-growth result. Baseline pair4+265 is a preserved outlier; candidate pair1 peak2727 versus baseline2687 (+40), while later peak medians match. Peak counts alone cannot explain cold->closed retention. The JSON contains only peak levels and closed-minus-cold deltas, not absolute cold/closed handle levels or handle ownership categories. Therefore no claim that+1 is entirely a starting-level difference, a proven leak, or zero retained resources is justified here. Root should inspect the retained absolute cold/warm/closed levels and ownership/lifetime witnesses before closing a handle-retention gate; repeated identical cycles/post-cleanup observations are required if that ambiguity remains. Hosts/attachments/GDI deltas are zero in every capture, GDI peak59, USER closed delta3 in all candidates versus baseline[3,3,3,5,3]; USER is also not zero cold->closed retention.

Possible remaining gaps: apply existing timing/resource gates to these explicit aggregate and paired observations; explain unusual baseline pair4/positive candidate pairs without discarding any evidence; resolve handle and USER cold/closed level semantics. If a gate concerns isolated native optimization, use equal dependency closure for a separately bound comparison or explicitly restrict the claim to whole-product behavior. A longer/more repeated paired follow-up can address variability if needed, after current frozen evidence is bound; this review does not direct a build or test now. There is no p95, presented FPS, driver allocation/GPU peak, process-exit lifetime or all-DPI/all-viewport qualification. Resource observations do not replace interaction/gallery/native completeness gates.

Analysis retains ANALYSIS_ONLY/NOT_QUALIFIED, global_is_clean=false for every capture and265 external C issues. This review does not certify historical-tree bytes or global cleanliness. Original v1 adverse evidence and earlier failed collections remain frozen with their original status. Root owns scoped acceptance and remaining plan disposition.
