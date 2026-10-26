#pragma once

#include "YyjsonHelpers.h"
#include <format>
#include <string>

namespace ViewerSpaceExport
{
struct Row final
{
    std::string name;
    std::string path;
    uint64_t bytes      = 0u;
    uint64_t shownBytes = 0u;
    uint64_t fileCount  = 0u;
    int64_t modified    = 0;
    bool countKnown     = false;
    bool directory      = false;
    bool synthetic      = false;
    bool excluded       = false;
    bool partial        = false;
    bool unreadable     = false;
    bool unmeasured     = false;
};

// Viewer export policy: quoting alone does not make a CSV cell a spreadsheet literal.
// JSON retains the exact original strings; CSV records when a protective apostrophe was added.
[[nodiscard]] inline bool NeedsSpreadsheetLiteral(std::string_view value) noexcept
{
    return ! value.empty() && std::string_view("=+-@\t\r\n").find(value.front()) != std::string_view::npos;
}

[[nodiscard]] inline std::string CsvText(std::string_view value)
{
    std::string result = "\"";
    if (NeedsSpreadsheetLiteral(value))
        result.push_back('\'');
    for (char ch : value)
    {
        result.push_back(ch);
        if (ch == '"')
            result.push_back('"');
    }
    result.push_back('"');
    return result;
}

[[nodiscard]] inline std::string CsvRow(const Row& row)
{
    return std::format("{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}\r\n",
                       CsvText(row.name),
                       CsvText(row.path),
                       row.directory,
                       row.synthetic,
                       row.bytes,
                       row.shownBytes,
                       row.countKnown ? std::to_string(row.fileCount) : std::string{},
                       row.modified > 0 ? std::to_string(row.modified) : std::string{},
                       row.excluded,
                       row.partial,
                       row.unreadable,
                       row.unmeasured,
                       NeedsSpreadsheetLiteral(row.name),
                       NeedsSpreadsheetLiteral(row.path),
                       "");
}

[[nodiscard]] inline std::string JsonRow(const Row& row)
{
    Common::Json::UniqueMutableDocument document{yyjson_mut_doc_new(nullptr)};
    if (! document)
        std::terminate();
    auto* object = yyjson_mut_obj(document.get());
    yyjson_mut_doc_set_root(document.get(), object);
    yyjson_mut_obj_add_strcpy(document.get(), object, "name", row.name.c_str());
    if (row.synthetic)
        yyjson_mut_obj_add_null(document.get(), object, "path");
    else
        yyjson_mut_obj_add_strcpy(document.get(), object, "path", row.path.c_str());
    yyjson_mut_obj_add_bool(document.get(), object, "directory", row.directory);
    yyjson_mut_obj_add_bool(document.get(), object, "synthetic", row.synthetic);
    yyjson_mut_obj_add_uint(document.get(), object, "logicalBytes", row.bytes);
    yyjson_mut_obj_add_uint(document.get(), object, "shownBytes", row.shownBytes);
    if (row.countKnown)
        yyjson_mut_obj_add_uint(document.get(), object, "fileCount", row.fileCount);
    else
        yyjson_mut_obj_add_null(document.get(), object, "fileCount");
    if (row.modified > 0)
        yyjson_mut_obj_add_sint(document.get(), object, "modified100ns", row.modified);
    else
        yyjson_mut_obj_add_null(document.get(), object, "modified100ns");
    yyjson_mut_obj_add_null(document.get(), object, "allocationBytes");
    yyjson_mut_obj_add_bool(document.get(), object, "excluded", row.excluded);
    yyjson_mut_obj_add_bool(document.get(), object, "partial", row.partial);
    yyjson_mut_obj_add_bool(document.get(), object, "unreadable", row.unreadable);
    yyjson_mut_obj_add_bool(document.get(), object, "unmeasured", row.unmeasured);
    size_t length = 0u;
    Common::Json::UniqueMallocString text{yyjson_mut_write(document.get(), YYJSON_WRITE_NOFLAG, &length)};
    if (! text)
        std::terminate();
    return std::string(text.get(), length);
}
} // namespace ViewerSpaceExport
