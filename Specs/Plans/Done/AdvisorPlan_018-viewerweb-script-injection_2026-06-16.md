# Advisor Plan 018 - ViewerWeb script injection / CSP

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/018-viewerweb-script-injection.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/018-viewerweb-script-injection.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 018: ViewerWeb — stop untrusted file content from breaking out of inline `<script>` (script injection in JSON/JSONL/Markdown viewers)

> **Frozen historical instructions (do not execute)**: Follow step by step; run every verification before continuing.
> Historical only. Root plans/ was deleted; do not execute.
>
> **Drift check (run first)**: `git diff --stat b274022d9..HEAD -- Plugins/ViewerWeb/ViewerWeb.cpp`
> If changed, compare excerpts to live code first; mismatch ⇒ STOP.

## Status

- **Priority**: P1 (security — script injection on intentionally-untrusted inputs)
- **Effort**: S–M
- **Risk**: LOW (escaping is strictly additive; CSP is additive)
- **Depends on**: none
- **Category**: security
- **Planned at**: commit `b274022d9`, 2026-06-16

## Why this matters

ViewerWeb renders JSON / JSONL / Markdown files by embedding the file's text into **inline
`<script>` string literals**. The escaper `EscapeJavaScriptStringUtf8` escapes quotes,
backslashes and control chars but **not `<` or the token `</script>`**. The HTML tokenizer ends
a `<script>` element on the byte sequence `</script` regardless of JS string-literal quoting, so
a `.json`/`.jsonl`/`.md` file whose content contains
`</script><img src=x onerror=...>` closes the script and injects arbitrary HTML/JS. The document
runs at the synthetic origin `https://viewer.redsalamander.invalid` with scripting enabled and no
CSP, and `NavigationStarting` only blocks top-level navigations — not `fetch()`/`Image()`
sub-resource requests — so injected script can exfiltrate the rendered document to a remote
server. Opening a file is the single most normal action in a file manager; this is
stored-XSS-equivalent on inputs the viewer explicitly targets as untrusted. (Note: the find-query
path is already escaped and is **not** part of this finding.)

## Current state

```cpp
// ViewerWeb.cpp:4981-5009  — escaper that misses '<' and '</script>'
[[nodiscard]] std::string EscapeJavaScriptStringUtf8(std::string_view text) noexcept
{
    std::string out; out.reserve(text.size() + text.size()/8);
    for (const char ch : text)
    {
        const auto u = static_cast<uint8_t>(ch);
        switch (ch)
        {
            case '\\': out += "\\\\"; break;
            case '\'': out += "\\'";  break;
            case '\"': out += "\\\""; break;
            case '\r': out += "\\r";  break;
            case '\n': out += "\\n";  break;
            case '\t': out += "\\t";  break;
            default:
                if (u < 0x20) out += std::format("\\x{:02X}", static_cast<unsigned int>(u));
                else          out.push_back(ch);    // '<' falls here, unescaped
                break;
        }
    }
    return out;
}
```

Injection sites (all concatenate the escaped content into an inline `<script>`):
- `ViewerWeb.cpp:4363` — Pretty JSON: `html += "code.textContent='" + escapedJson + "';";`
- `ViewerWeb.cpp:4432` — Tree JSON: `html += "const jsonText='" + escapedJson + "';";`
- `ViewerWeb.cpp:4496` — Markdown: `html += "const src='" + escapedMarkdown + "';";`
- `ViewerWeb.cpp:1147-1161` — JSONL per-entry fields (ts/level/category/message/summary/json),
  embedded into a JS object/array literal via the same escaper.

No `<meta http-equiv="Content-Security-Policy">` is emitted in any of these templates (confirm:
`grep -n "Content-Security-Policy" Plugins/ViewerWeb/ViewerWeb.cpp` returns nothing).

There may also be a wide-string variant `EscapeJavaScriptString` (the find-query uses one). Check:
`grep -n "EscapeJavaScriptString" Plugins/ViewerWeb/ViewerWeb.cpp`.

## Commands you will need

| Purpose | Command | Expected |
|---|---|---|
| Build plugin | `.\build.ps1 -ProjectName ViewerWeb` | exit 0, 0 warnings/errors |
| Full build | `.\build.ps1` | exit 0 |
| Tests | `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` | all pass |

## Scope

**In scope**: `Plugins/ViewerWeb/ViewerWeb.cpp` (escaper(s) + the JSON/JSONL/Markdown templates).
Optionally `Plugins/ViewerWeb/resource.h` + the `.rc` only if you add a user-facing string (not
required here).
**Out of scope**: the find-query path (already mitigated); WebView2 settings / navigation policy
(separate concern); the async-load race (plan 019).

## Git workflow

- Branch: `advisor/018-viewerweb-script-injection`
- Message: `fix(ViewerWeb): neutralize </script> in inline JS and add CSP to generated docs`
- Do NOT push/PR unless instructed.

## Steps

### Step 1: Escape `<` (and JS line separators) in the JS-string escaper(s)

In `EscapeJavaScriptStringUtf8`, escape `<` so `</script>` can never appear literally inside a JS
string literal. The minimal, robust change is to emit `\x3C` for `<`:

```cpp
case '<': out += "\\x3C"; break;   // neutralizes </script>, <!-- , <![CDATA[
```

Also handle the two JS line-terminator code points that break string literals — U+2028 and U+2029
— which arrive as UTF-8 byte sequences. Since the escaper iterates bytes, the simplest correct
approach is: when you see the lead byte `0xE2` followed by `0x80` and `0xA8`/`0xA9`, emit ` `
/` `. If multi-byte lookahead is awkward in the per-`char` loop, an acceptable alternative is
to switch the JSON path to also pass `YYJSON_WRITE_ESCAPE_UNICODE` (already set at `:4307`, which
ASCII-escapes all non-ASCII, covering U+2028/9 for JSON) and document that the Markdown/JSONL raw
text still needs the lead-byte handling. Prefer the explicit U+2028/9 handling so all three paths
are safe.

If a wide `EscapeJavaScriptString` variant exists and is used for any content (not just the
already-safe find-query), apply the same `<` → `\x3C` escaping there.

**Verify**: `.\build.ps1 -ProjectName ViewerWeb` → exit 0.

### Step 2: Add a restrictive CSP meta tag to the generated JSON/JSONL/Markdown documents

As defense-in-depth, emit a `<meta http-equiv="Content-Security-Policy">` in the `<head>` of each
generated document. Because these templates rely on **inline** scripts, a nonce-based policy is
the clean option, but the high-leverage, low-risk win is to forbid network egress and external
resources so even a successful breakout cannot exfiltrate:

```
default-src 'none';
script-src 'unsafe-inline' https://viewer.redsalamander.invalid;
style-src 'unsafe-inline' https://viewer.redsalamander.invalid;
img-src data: https://viewer.redsalamander.invalid;
font-src https://viewer.redsalamander.invalid;
connect-src 'none';
```

Add it once to the shared `<head>` portion of the JSON (pretty + tree), JSONL, and Markdown
templates. Confirm the legitimate inline scripts and the bundled `hljs`/`JSONEditor`/`markdownit`
assets still load (they are served from the synthetic origin via `WebResourceRequested`, so
allow that origin in `script-src`/`style-src`/`img-src`/`font-src`). `connect-src 'none'` blocks
`fetch`/XHR/WebSocket exfiltration.

**Verify**: `.\build.ps1` → exit 0, then a manual smoke (optional): open a normal JSON and a
normal Markdown file in the app and confirm they still render correctly (CSP didn't break the
legit inline scripts/assets). If you cannot run the app, rely on the self-test in Step 3.

### Step 3: Add a self-test that feeds a breakout payload

Add a case to `Tests/ViewerPETests/ViewerPETests.cpp` (the viewer integration harness; it already
drives ViewerWeb) that opens a JSON (or JSONL/Markdown) file whose content contains the literal
`</script><script>window.__pwn=1</script>` and asserts the injected script did **not** execute —
e.g. via the WebView2 (`ExecuteScript`) checking `window.__pwn` is undefined, or by asserting the
document text shows the payload as data, not as a new element. If the harness lacks an
`ExecuteScript` hook, at minimum add a unit-level test of `EscapeJavaScriptStringUtf8` asserting
that its output of an input containing `</script>` contains no substring `</script` (you may need
to expose the function via an internal header or a small test shim; keep it in the test project).

**Verify**: `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` → all pass, including the new case.

## Done criteria

ALL must hold:

- (archived, not live)  `EscapeJavaScriptStringUtf8` (and any wide variant used for content) emits no literal `<`;
      `</script>` in input cannot appear in output (verified by the new test).
- (archived, not live)  U+2028/U+2029 in input cannot terminate a JS string literal in output.
- (archived, not live)  Each generated JSON/JSONL/Markdown document includes a `Content-Security-Policy` meta with
      `connect-src 'none'`.
- (archived, not live)  Normal JSON and Markdown files still render correctly (smoke or self-test).
- (archived, not live)  A self-test feeds a `</script>` breakout payload and confirms no script execution.
- (archived, not live)  `.\build.ps1 -ProjectName ViewerWeb` and `.\build.ps1` exit 0, 0 warnings.
- (archived, not live)  `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` passes.
- (archived, not live)  `plans/README.md` updated.

## STOP conditions

- Excerpts don't match live code (drift).
- The CSP breaks the legitimate inline scripts/bundled assets and you cannot find an origin/nonce
  combination that both renders normal files and blocks egress — in that case land Step 1
  (escaping, which fully closes the injection) and Step 3, and report the CSP as a follow-up
  rather than shipping a broken viewer.
- You discover a host-object bridge (`AddHostObjectToScript`) or `add_WebMessageReceived` handler
  was added since this plan — that raises the severity to critical (RCE surface); report before
  proceeding so the bridge is reviewed.

## Maintenance notes

- Any new inline-script template that embeds file/user content must use the hardened escaper and
  carry the CSP. Prefer moving payloads out of inline literals into a
  `<script type="application/json">` block read via `textContent` (immune to `</script>` breakout)
  if these templates are revisited.
- Reviewer: confirm every `"... '" + escaped* + "' ..."` concatenation now feeds the hardened
  escaper, and that no template forgot the CSP.
- This plan does not change the deliberate "ViewerWeb renders untrusted HTML as a browser" design
  (that is by-design per `Specs/Plugins/Plugins_ViewerWeb.md`); it only stops *data* files from
  becoming *code*.
