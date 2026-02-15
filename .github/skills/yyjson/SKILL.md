---
name: yyjson
description: Safe yyjson usage patterns for Red Salamander C++23 (UTF-16). Use when parsing or serializing JSON/JSON5 with yyjson, adding keys/strings to `yyjson_mut_doc`, managing `yyjson_doc` lifetimes, or auditing code for yyjson ownership and cleanup bugs (use-after-free/leaks).
---

# yyjson

## DO / DON'T (quick rules)

- ✅ DO treat `yyjson_doc*`/`yyjson_mut_doc*` as the owner of all `yyjson_val*` pointers you read from it.
- ✅ DO free `yyjson_doc*` with `yyjson_doc_free()` and `yyjson_mut_doc*` with `yyjson_mut_doc_free()` (prefer RAII).
- ✅ DO free buffers returned by `yyjson_write*()` / `yyjson_mut_write*()` with `free()` (or `alc->free` if you used a custom allocator).
- ✅ DO copy/convert `yyjson_get_str()` results before freeing the owning doc.
- ✅ DO check all pointer-returning yyjson allocations (`*_new`, `yyjson_mut_obj/arr`, `yyjson_mut_*str*`, `yyjson_mut_*number*`) and propagate `E_OUTOFMEMORY`/failure.
- ✅ DO check `*_add_*` return values when the output JSON must be complete/correct (treat `false` as out-of-memory in practice).
- ❌ DON'T store `yyjson_val*` or any `const char*` returned by `yyjson_get_str()` beyond the doc lifetime.
- ❌ DON'T use `delete`/`delete[]` on yyjson allocations.
- ❌ DON'T pass temporary strings to APIs that do not copy key/value strings.
- ❌ DON'T use `YYJSON_READ_INSITU` unless you *really* want the input buffer to be modified and to outlive the parsed doc.

## Mutable API: key/value ownership (most common bugs)

### Critical rule: many `yyjson_mut_obj_add_*` APIs do not copy strings

yyjson mutable builders often *borrow* pointers:

- `yyjson_mut_obj_add_*` convenience functions do **not** copy the **key**.
- For string values, `..._str` / `..._strn` do **not** copy the **value** either.
- Only `..._strcpy` / `..._strncpy` copy the *value bytes* into the doc.

This is safe only when the borrowed key/value outlives the doc (string literals, static storage, or buffers you keep alive).

### ❌ DON'T: Pass temporary strings to non-copy APIs

```cpp
std::string temp = GetSomeString();
yyjson_mut_obj_add_str(doc, root, "key", temp.c_str()); // WRONG: value not copied
```

### ✅ DO: Copy C++ strings into the doc

```cpp
std::string temp = GetSomeString();
yyjson_mut_obj_add_strcpy(doc, root, "key", temp.c_str()); // value copied
```

### ✅ DO: Copy known-length UTF-8

```cpp
std::string utf8 = Utf8FromUtf16(wide);
yyjson_mut_obj_add_strncpy(doc, root, "key", utf8.data(), utf8.size()); // value copied
```

### ❌ DON'T: Use dynamic object keys with `yyjson_mut_obj_add_*` convenience APIs

`yyjson_mut_obj_add_*` functions do not copy the key:

```cpp
std::string keyUtf8 = Utf8FromUtf16(dynamicKey);
yyjson_mut_obj_add_strcpy(doc, root, keyUtf8.c_str(), "value"); // WRONG: key not copied
```

### ✅ DO: For dynamic keys, allocate the key as a yyjson value and use `yyjson_mut_obj_add()`

```cpp
std::string keyUtf8 = Utf8FromUtf16(dynamicKey);
yyjson_mut_val* key = yyjson_mut_strncpy(doc, keyUtf8.data(), keyUtf8.size());
yyjson_mut_val* val = yyjson_mut_strcpy(doc, "value");
yyjson_mut_obj_add(root, key, val);
```

### Quick reference (mutable builder strings)

| Function | Copies key? | Copies value bytes? | Use when |
|----------|-------------|---------------------|----------|
| `yyjson_mut_obj_add_str` | ❌ No | ❌ No | Only when key+value have guaranteed lifetime (literals/static) |
| `yyjson_mut_obj_add_strcpy` | ❌ No | ✅ Yes | Key is a literal/static; value is a temporary/C++ string |
| `yyjson_mut_obj_add_strncpy` | ❌ No | ✅ Yes | Key is a literal/static; value is known-length UTF-8 |
| `yyjson_mut_obj_add_val` | ❌ No | n/a | Key is a literal/static; value is a `yyjson_mut_val*` owned by the doc |
| `yyjson_mut_obj_add` | ✅ (via `yyjson_mut_val*`) | ✅ (via `yyjson_mut_val*`) | Dynamic keys and/or fully doc-owned strings |
| `yyjson_mut_str` / `yyjson_mut_strn` | n/a | ❌ No | Only for literals/static buffers |
| `yyjson_mut_strcpy` / `yyjson_mut_strncpy` | n/a | ✅ Yes | Preferred for any temporary/converted UTF-8 |

## Read API: parsing & type safety

- Prefer `yyjson_read_opts()` when you need error details (`yyjson_read_err`).
- `yyjson_read_opts()` takes a mutable `char*`; pass a mutable buffer (`std::string`, `std::vector<char>`), not a string literal.
- Use project defaults when reading user-facing config: `YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM`.
- Always check types (`yyjson_is_*`) before reading (`yyjson_get_*`).
- Treat `yyjson_get_str()` as borrowed; convert to `std::wstring`/`std::string` immediately.

## Memory management (recommended RAII patterns)

Prefer `wil::unique_any` or `wil::scope_exit` to make cleanup unconditional:

```cpp
using unique_yyjson_doc     = wil::unique_any<yyjson_doc*, decltype(&yyjson_doc_free), yyjson_doc_free>;
using unique_yyjson_mut_doc = wil::unique_any<yyjson_mut_doc*, decltype(&yyjson_mut_doc_free), yyjson_mut_doc_free>;
using unique_malloc_string  = wil::unique_any<char*, decltype(&free), free>;
```

## Write API: serialization

- Use `yyjson_write_err` with `yyjson_write_opts()` / `yyjson_mut_write_opts()` when you need diagnostics.
- Free output with `free()` (or `alc->free` if you passed an allocator).
- Use the returned `len` when building `std::string` (don’t assume a null terminator is sufficient).

## Audit checklist (use when reviewing a change)

- Search for non-copy string APIs: `yyjson_mut_obj_add_str*`, `yyjson_mut_str*`.
- Verify any non-copy call sites only use string literals/static buffers.
- Verify dynamic keys never flow into `yyjson_mut_obj_add_*` convenience APIs.
- Verify `yyjson_doc_free()` / `yyjson_mut_doc_free()` and write-buffer frees happen on all paths (prefer RAII).
