# I66 OneDrive Personal live fixture preflight

Integrated x64 Debug head `a6ecbe9a2` built with zero warnings and errors.
Receipt `5a31700b8c32f0232836483d588471be299e2282a8f720d33f74536bf2c6f7c7`.

- The exact non-activating Graph loopback case passed setup, host Rename and
  cleanup in `20261125T133618Z-70668-71612e80240a47bb947774782083ba73`.
  The standalone plugin suite passed in
  `20261125T133835Z-36816-250f9ee4148942bfbbfb1e010daa21de`.
- Initial configured diagnostics `20261125T133053Z-73028-10a8655f8626455981b44b62bab5340a`
  and `20261125T133236Z-48852-cbbc041d39e94234843aeec780e0bcdd` exited during
  startup with no case result. The launcher had opened its logs inside the app's
  artifact rotation path. The corrected launch wrote its process logs to a
  sibling artifact folder. An unseeded control reached the exact case and
  reported its expected profile skip in
  `20261125T133427Z-44124-50ab011fea7449aaa1e5d89e61ea7b56`.
- The first corrected configured read-only run
  `20261125T133503Z-85836-05eee2567d424761abb2b5f24d64bf43` failed with
  `0x8007052E`. The user renewed saved sign-in in Connection Manager.
  The exact read-only rerun `20261125T134004Z-64848-aeabc524b7aa42d8885d6e9ffb23c5fe`
  passed without activation and verified the approved fixture. The exception for
  Windows Hello applied to the isolated test profile. The installed settings
  digest matched before and after, and the isolated settings file was removed before
  each archive.
- The broad configured File Operations run
  `20261125T134032Z-88548-e470bcd602474da485cb1e1685b3fc49` created its owned
  child, verified the probe file and enumerated the remote pane. The performance
  rows show that Copy was admitted before the launcher timed out. Native case
  results are missing. The preserved traces identify the remaining test child.
  The visible fixture is retained for recovery, and Rename and cleanup have no
  pass result from this run.

The dedicated live Rename scenario is in progress. Hosted platform acceptance,
other configured services and the broader OneDrive workflow remain open.
