# C0 owned-publication regression

Run `c0-owned-publication-regression-20260905`, Debug,
`FileOpsFamily_R3OwnedPublication`: 7/7 passed, no skips.
Build receipt `167683804c5f713ab2845ffc11459409a91d27ef68482d9d88506ae353893145`;
zero warnings/errors. Base `58f22d0e` plus working changes is runner-bound.

The five behavior cases cover identity-less conditional replacement, writer proof,
publication faults, missing expected-occupant tokens, and destination-name refusal.
Raw results, trace, and runner aggregate are retained unchanged.

The full performance stream remains in this run's owned test-root artifact folder:
8419986 bytes; 23751 nonblank rows; SHA-256
8ec5e3f15b084f146ebab422db77a5dc64275a633eb75f32e253c7988a2c872f.
The compact metrics retain all 107 rows with exact metric
`FileOps.Bridge.PublicationUs` in original order. Original line numbers:
430, 763, 932, 1131, 1408, 1650, 1878, 2106, 2334, 2562, 2834, 3056, 3284, 3509, 3806, 4034, 4262, 4515, 4768, 5011, 5239, 5510, 5683, 5880, 6105, 6321, 6532, 6733, 6944, 7160, 7374, 7638, 7853, 8071, 8288, 8505, 8721, 8937, 9199, 9367, 9588, 9819, 10015, 10236, 10467, 10668, 10879, 11081, 11343, 11534, 11818, 12034, 12270, 12458, 12679, 12881, 13088, 13290, 13487, 13694, 13891, 14098, 14315, 14502, 14699, 14901, 15083, 15265, 15467, 15669, 15851, 16033, 16272, 16448, 16662, 16881, 17080, 17279, 17498, 17697, 17916, 18118, 18380, 18557, 18756, 18975, 19179, 19413, 19617, 19877, 20053, 20272, 20491, 20705, 20929, 21143, 21347, 21566, 21778, 22028, 22225, 22449, 22663, 22865, 23069, 23283, 23575.
No speedup/percentile claim or final closeout is made from this focused regression.

The disk audit reports 23 retained earlier run directories outside its allow-list;
they were not deleted. This does not establish a clean test-root audit.
