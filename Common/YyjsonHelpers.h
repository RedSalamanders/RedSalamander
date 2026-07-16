#pragma once

#include "StringConversion.h"

#include <charconv>
#include <cstdint>
#include <cstdlib>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <string_view>

#pragma warning(push)
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

namespace Common::Json
{
struct ImmutableDocumentDeleter
{
    void operator()(yyjson_doc* document) const noexcept
    {
        yyjson_doc_free(document);
    }
};

struct MutableDocumentDeleter
{
    void operator()(yyjson_mut_doc* document) const noexcept
    {
        yyjson_mut_doc_free(document);
    }
};

struct MallocStringDeleter
{
    void operator()(char* text) const noexcept
    {
        std::free(text);
    }
};

using UniqueDocument        = std::unique_ptr<yyjson_doc, ImmutableDocumentDeleter>;
using UniqueMutableDocument = std::unique_ptr<yyjson_mut_doc, MutableDocumentDeleter>;
using UniqueMallocString    = std::unique_ptr<char, MallocStringDeleter>;

enum class MemberRequirement : uint8_t
{
    Optional,
    Required,
};

enum class NumericStringPolicy : uint8_t
{
    Reject,
    Allow,
};

enum class BooleanIntegerPolicy : uint8_t
{
    Reject,
    AllowZeroAndNonzero,
};

enum class UnsignedIntegerPolicy : uint8_t
{
    AllowNonnegativeSigned,
    RequireUnsignedStorage,
};

enum class MemberStatus : uint8_t
{
    Value,
    MissingOptional,
    MissingRequired,
    Null,
    WrongType,
    OutOfRange,
    InvalidUtf8,
};

template <typename T> struct MemberResult
{
    MemberStatus status = MemberStatus::MissingOptional;
    T value{};

    [[nodiscard]] bool HasValue() const noexcept
    {
        return status == MemberStatus::Value;
    }
};

[[nodiscard]] inline const yyjson_val* FindMember(const yyjson_val* object, const char* key) noexcept
{
    yyjson_val* mutableObject = const_cast<yyjson_val*>(object);
    if (! mutableObject || ! key || ! yyjson_is_obj(mutableObject))
    {
        return nullptr;
    }
    return yyjson_obj_get(mutableObject, key);
}

[[nodiscard]] inline MemberStatus MissingStatus(MemberRequirement requirement) noexcept
{
    return requirement == MemberRequirement::Required ? MemberStatus::MissingRequired : MemberStatus::MissingOptional;
}

[[nodiscard]] inline MemberResult<std::string_view> GetStringMember(const yyjson_val* object,
                                                                    const char* key,
                                                                    MemberRequirement requirement) noexcept
{
    yyjson_val* value = const_cast<yyjson_val*>(FindMember(object, key));
    if (! value)
    {
        return {.status = MissingStatus(requirement)};
    }
    if (yyjson_is_null(value))
    {
        return {.status = MemberStatus::Null};
    }
    if (! yyjson_is_str(value))
    {
        return {.status = MemberStatus::WrongType};
    }

    const char* text = yyjson_get_str(value);
    return text ? MemberResult<std::string_view>{.status = MemberStatus::Value, .value = std::string_view{text, yyjson_get_len(value)}}
                : MemberResult<std::string_view>{.status = MemberStatus::WrongType};
}

[[nodiscard]] inline MemberResult<std::wstring> GetUtf16StringMemberStrict(const yyjson_val* object,
                                                                           const char* key,
                                                                           MemberRequirement requirement) noexcept
{
    const MemberResult<std::string_view> utf8 = GetStringMember(object, key, requirement);
    if (! utf8.HasValue())
    {
        return {.status = utf8.status};
    }

    const std::optional<std::wstring> converted = Common::Strings::TryUtf16FromUtf8Strict(utf8.value);
    if (! converted.has_value())
    {
        return {.status = MemberStatus::InvalidUtf8};
    }
    return {.status = MemberStatus::Value, .value = converted.value()};
}

[[nodiscard]] inline MemberResult<bool> GetBoolMember(const yyjson_val* object,
                                                      const char* key,
                                                      MemberRequirement requirement,
                                                      BooleanIntegerPolicy integerPolicy = BooleanIntegerPolicy::Reject) noexcept
{
    yyjson_val* value = const_cast<yyjson_val*>(FindMember(object, key));
    if (! value)
    {
        return {.status = MissingStatus(requirement)};
    }
    if (yyjson_is_null(value))
    {
        return {.status = MemberStatus::Null};
    }
    if (! yyjson_is_bool(value))
    {
        if (integerPolicy == BooleanIntegerPolicy::AllowZeroAndNonzero && yyjson_is_sint(value))
        {
            return {.status = MemberStatus::Value, .value = yyjson_get_sint(value) != 0};
        }
        if (integerPolicy == BooleanIntegerPolicy::AllowZeroAndNonzero && yyjson_is_uint(value))
        {
            return {.status = MemberStatus::Value, .value = yyjson_get_uint(value) != 0};
        }
        return {.status = MemberStatus::WrongType};
    }
    return {.status = MemberStatus::Value, .value = yyjson_get_bool(value)};
}

template <typename T>
[[nodiscard]] inline std::optional<T> ParseIntegerString(std::string_view text) noexcept
{
    T value{};
    const std::from_chars_result parsed = std::from_chars(text.data(), text.data() + text.size(), value, 10);
    if (parsed.ec != std::errc{} || parsed.ptr != text.data() + text.size())
    {
        return std::nullopt;
    }
    return value;
}

[[nodiscard]] inline MemberResult<int64_t> GetInt64Member(const yyjson_val* object,
                                                          const char* key,
                                                          MemberRequirement requirement,
                                                          NumericStringPolicy stringPolicy = NumericStringPolicy::Reject) noexcept
{
    yyjson_val* value = const_cast<yyjson_val*>(FindMember(object, key));
    if (! value)
    {
        return {.status = MissingStatus(requirement)};
    }
    if (yyjson_is_null(value))
    {
        return {.status = MemberStatus::Null};
    }
    if (yyjson_is_sint(value))
    {
        return {.status = MemberStatus::Value, .value = yyjson_get_sint(value)};
    }
    if (yyjson_is_uint(value))
    {
        const uint64_t unsignedValue = yyjson_get_uint(value);
        if (unsignedValue > static_cast<uint64_t>((std::numeric_limits<int64_t>::max)()))
        {
            return {.status = MemberStatus::OutOfRange};
        }
        return {.status = MemberStatus::Value, .value = static_cast<int64_t>(unsignedValue)};
    }
    if (stringPolicy == NumericStringPolicy::Allow && yyjson_is_str(value))
    {
        const char* text = yyjson_get_str(value);
        const std::optional<int64_t> parsed = text ? ParseIntegerString<int64_t>(std::string_view{text, yyjson_get_len(value)}) : std::nullopt;
        return parsed.has_value() ? MemberResult<int64_t>{.status = MemberStatus::Value, .value = parsed.value()}
                                  : MemberResult<int64_t>{.status = MemberStatus::OutOfRange};
    }
    return {.status = MemberStatus::WrongType};
}

[[nodiscard]] inline MemberResult<uint64_t> GetUInt64Member(const yyjson_val* object,
                                                            const char* key,
                                                            MemberRequirement requirement,
                                                            NumericStringPolicy stringPolicy = NumericStringPolicy::Reject,
                                                            UnsignedIntegerPolicy integerPolicy = UnsignedIntegerPolicy::AllowNonnegativeSigned) noexcept
{
    yyjson_val* value = const_cast<yyjson_val*>(FindMember(object, key));
    if (! value)
    {
        return {.status = MissingStatus(requirement)};
    }
    if (yyjson_is_null(value))
    {
        return {.status = MemberStatus::Null};
    }
    if (yyjson_is_uint(value))
    {
        return {.status = MemberStatus::Value, .value = yyjson_get_uint(value)};
    }
    if (yyjson_is_sint(value))
    {
        if (integerPolicy == UnsignedIntegerPolicy::RequireUnsignedStorage)
        {
            return {.status = MemberStatus::WrongType};
        }
        const int64_t signedValue = yyjson_get_sint(value);
        if (signedValue < 0)
        {
            return {.status = MemberStatus::OutOfRange};
        }
        return {.status = MemberStatus::Value, .value = static_cast<uint64_t>(signedValue)};
    }
    if (stringPolicy == NumericStringPolicy::Allow && yyjson_is_str(value))
    {
        const char* text = yyjson_get_str(value);
        const std::optional<uint64_t> parsed = text ? ParseIntegerString<uint64_t>(std::string_view{text, yyjson_get_len(value)}) : std::nullopt;
        return parsed.has_value() ? MemberResult<uint64_t>{.status = MemberStatus::Value, .value = parsed.value()}
                                   : MemberResult<uint64_t>{.status = MemberStatus::OutOfRange};
    }
    return {.status = MemberStatus::WrongType};
}
} // namespace Common::Json
