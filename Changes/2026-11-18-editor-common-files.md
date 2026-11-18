- `.gitattributes` and `.gitignore` start with the common blocks that every RedSalamanders repository now
  shares, and `.editorconfig` is the shared file. The `.gitattributes` block stores text with LF and checks out
  C and C++ sources, PowerShell, and Visual Studio files with CRLF and Markdown, JSON, and YAML with LF; the
  repository's own lines follow it unchanged and still take precedence. No tracked file needs renormalizing:
  every file with CRLF or mixed line endings in the index is already `-text`.
  This is step CV-6-05 of the `RedHome` convergence plan (register D-03, D-04, and the `.editorconfig` part of B-12).
