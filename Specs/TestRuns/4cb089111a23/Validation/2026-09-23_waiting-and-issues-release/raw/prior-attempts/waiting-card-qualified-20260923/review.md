# Bounded P3 review

Root implementation with Sol High read-only reviews; no competing writer/build lane.

- Root verified Queue.cpp admitted-only queuePaused producer and State.cpp admission/started publication before narrowing the compact predicate.
- Height is 104 DIP plus the existing 4 DIP path gap; optional total/concurrency/source rows match actual paint increments.
- Expanded admission state and waiting-row changes snap with their content; unrelated motion remains unchanged.
- Corrected inherited queue-arrow flags in isolated fixtures; snapshots assert the actual arrows.
- Added seven same-HWND/task phases, including first-frame known/zero-total/concurrency row changes, with lower and upper waiting gaps.
- Fixed the two C5219 test-only DPI conversions found by the first build.
- A formatting cleanup accidentally decoded three existing Unicode characters with the system code page in the primary checkout. Restored exact originals before this freeze; all removed non-ASCII lines were checked against Git. No affected source was built or committed.

Qualification is pending; this source review is not native input/UIA, Full or resource acceptance.

The first gallery failed only four waiting-card assertions applied to the footer-only scenario116, which intentionally has no cards. The final test limits those assertions to expanded-card scenes; existing footer-only coverage remains. All first-run actual waiting-card and transition geometry checks passed; the failed run is retained in full.
