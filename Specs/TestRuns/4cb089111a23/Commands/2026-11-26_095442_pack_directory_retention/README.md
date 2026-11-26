# Pack retains directory selections in ZIP and 7z

The test-enabled x64 Release build at `afdbd637dadd4f2bdcbcacfacd71b55297a5d966` completed with zero warnings and errors. Its verified receipt was `9b43c245164302c717435677952314684d7eecc360d5720576cde7ac04de7388`.

Run `20261126T085442Z-70836-e99fb3e158c14679bdc7e451781ccb61` passed the exact archive Commands case without activation in 2,192 ms with zero failures or skips. It selected each writer explicitly. The ZIP and 7z archives each contained all 16 planned entries, including the nested file and empty directory. Both returned `S_FALSE` for retained sources and started no cleanup prompt or Delete. The originals kept their checked bytes. The ZIP worker took 4,600 µs and the 7z worker took 6,889 µs.

This selection verifies source retention under the current exact-only cleanup contract. Directory selections lack a proven descendant cleanup set.
