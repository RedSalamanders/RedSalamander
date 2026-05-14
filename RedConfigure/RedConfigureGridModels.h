#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include "DxUi.h"
#include "RedConfigureSession.h"
#include "RedConfigureUiHelpers.h"
#include "SettingsStore.h"
#include "resource.h"

#include <algorithm>
#include <cstdint>
#include <iterator>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace RedConfigure::Ui
{
using RedSalamander::DxUi::GridCellData;
using RedSalamander::DxUi::GridColumnDesc;
using RedSalamander::DxUi::GridColumnKind;
using RedSalamander::DxUi::GridRowStyle;
using RedSalamander::DxUi::GridRowTone;
using RedSalamander::DxUi::IDxGridModel;

class InventoryGridModel final : public IDxGridModel
{
public:
    InventoryGridModel(HINSTANCE instance, const RedConfigure::RedConfigureSession& session) : _instance(instance), _session(session)
    {
        _columns.push_back(GridColumnDesc{
            .id = L"kind", .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_KIND), .widthDip = 130.0f, .minWidthDip = 90.0f, .kind = GridColumnKind::Text, .sortable = false, .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"owner",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_OWNER),
            .widthDip = 150.0f,
            .minWidthDip = 90.0f,
            .kind = GridColumnKind::Text,
            .sortable = false,
            .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"id", .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_ID), .widthDip = 160.0f, .minWidthDip = 90.0f, .kind = GridColumnKind::Text, .sortable = false, .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"source",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_SOURCE),
            .widthDip = 360.0f,
            .minWidthDip = 160.0f,
            .kind = GridColumnKind::Text,
            .sortable = false,
            .multiline = false});
    }

    InventoryGridModel(const InventoryGridModel&)            = delete;
    InventoryGridModel& operator=(const InventoryGridModel&) = delete;
    InventoryGridModel(InventoryGridModel&&)                 = delete;
    InventoryGridModel& operator=(InventoryGridModel&&)      = delete;

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _session.GetInventoryEntries().size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return columnIndex < _columns.size() ? _columns[columnIndex] : GridColumnDesc{};
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        const auto rows = _session.GetInventoryEntries();
        if (rowIndex >= rows.size())
        {
            return;
        }

        const RedConfigure::InventoryEntry& row = rows[rowIndex];
        switch (columnIndex)
        {
            case 0u: outCell.text = LocalizableKindText(_instance, row.kind); break;
            case 1u: outCell.text = row.ownerName; break;
            case 2u: outCell.text = row.itemId.empty() ? row.resourceId : row.itemId; break;
            case 3u: outCell.text = row.sourceText; break;
            default: break;
        }
        outCell.tooltipText = outCell.text;
        outCell.multiline   = false;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return static_cast<uint64_t>(rowIndex) + 1u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId == 0u)
        {
            return std::nullopt;
        }

        const size_t rowIndex = static_cast<size_t>(rowId - 1u);
        return rowIndex < _session.GetInventoryEntries().size() ? std::optional<size_t>(rowIndex) : std::nullopt;
    }

private:
    HINSTANCE _instance = nullptr;
    const RedConfigure::RedConfigureSession& _session;
    std::vector<GridColumnDesc> _columns;
};

class TranslationGridModel final : public IDxGridModel
{
public:
    TranslationGridModel(HINSTANCE instance, const RedConfigure::RedConfigureSession& session) : _instance(instance), _session(session)
    {
        _columns.push_back(GridColumnDesc{
            .id = L"id", .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_ID), .widthDip = 160.0f, .minWidthDip = 90.0f, .kind = GridColumnKind::Text, .sortable = true, .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"source",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_SOURCE),
            .widthDip = 360.0f,
            .minWidthDip = 160.0f,
            .kind = GridColumnKind::Text,
            .sortable = true,
            .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"target",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_TARGET),
            .widthDip = 360.0f,
            .minWidthDip = 160.0f,
            .kind = GridColumnKind::Text,
            .sortable = true,
            .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"status",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_STATUS),
            .widthDip = 150.0f,
            .minWidthDip = 110.0f,
            .kind = GridColumnKind::Text,
            .sortable = true,
            .multiline = false});
    }

    TranslationGridModel(const TranslationGridModel&)            = delete;
    TranslationGridModel& operator=(const TranslationGridModel&) = delete;
    TranslationGridModel(TranslationGridModel&&)                 = delete;
    TranslationGridModel& operator=(TranslationGridModel&&)      = delete;

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _viewRows.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return columnIndex < _columns.size() ? _columns[columnIndex] : GridColumnDesc{};
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        const auto rows = _session.GetTranslations();
        const std::optional<size_t> sessionRow = ResolveSessionRow(rowIndex);
        if (! sessionRow || sessionRow.value() >= rows.size())
        {
            return;
        }

        const RedConfigure::TranslationEntry& row = rows[sessionRow.value()];
        switch (columnIndex)
        {
            case 0u: outCell.text = row.id; break;
            case 1u: outCell.text = row.sourceText; break;
            case 2u: outCell.text = row.targetText; break;
            case 3u: outCell.text = PlaceholderStatusText(_instance, row.validation.status); break;
            default: break;
        }
        outCell.tooltipText = outCell.text;
        outCell.multiline   = false;
    }

    [[nodiscard]] GridRowStyle GetRowStyle(size_t rowIndex) const override
    {
        const auto rows = _session.GetTranslations();
        const std::optional<size_t> sessionRow = ResolveSessionRow(rowIndex);
        if (! sessionRow || sessionRow.value() >= rows.size() || rows[sessionRow.value()].validation.status == RedConfigure::Localization::PlaceholderStatus::Ok)
        {
            return {};
        }

        return GridRowStyle{.tone = GridRowTone::Warning};
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        const std::optional<size_t> sessionRow = ResolveSessionRow(rowIndex);
        return sessionRow ? static_cast<uint64_t>(sessionRow.value()) + 1u : 0u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId == 0u)
        {
            return std::nullopt;
        }

        const size_t sessionRow = static_cast<size_t>(rowId - 1u);
        const auto it = std::find(_viewRows.begin(), _viewRows.end(), sessionRow);
        return it != _viewRows.end() ? std::optional<size_t>(static_cast<size_t>(std::distance(_viewRows.begin(), it))) : std::nullopt;
    }

    void SetViewRows(std::vector<size_t> viewRows)
    {
        _viewRows = std::move(viewRows);
    }

    [[nodiscard]] std::optional<size_t> ResolveSessionRow(size_t viewRow) const noexcept
    {
        return viewRow < _viewRows.size() ? std::optional<size_t>(_viewRows[viewRow]) : std::nullopt;
    }

private:
    HINSTANCE _instance = nullptr;
    const RedConfigure::RedConfigureSession& _session;
    std::vector<GridColumnDesc> _columns;
    std::vector<size_t> _viewRows;
};

class ThemeColorGridModel final : public IDxGridModel
{
public:
    ThemeColorGridModel(HINSTANCE instance, const RedConfigure::RedConfigureSession& session) : _instance(instance), _session(session)
    {
        _columns.push_back(GridColumnDesc{
            .id = L"key",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_LABEL_COLOR_KEY),
            .widthDip = 210.0f,
            .minWidthDip = 140.0f,
            .kind = GridColumnKind::Text,
            .sortable = false,
            .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"effective",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_EFFECTIVE_VALUE),
            .widthDip = 120.0f,
            .minWidthDip = 90.0f,
            .kind = GridColumnKind::Text,
            .sortable = false,
            .multiline = false});
        _columns.push_back(GridColumnDesc{
            .id = L"authored",
            .title = LoadAppString(_instance, IDS_REDCONFIGURE_COL_AUTHORED_VALUE),
            .widthDip = 190.0f,
            .minWidthDip = 120.0f,
            .kind = GridColumnKind::Text,
            .sortable = false,
            .multiline = false});
    }

    ThemeColorGridModel(const ThemeColorGridModel&)            = delete;
    ThemeColorGridModel& operator=(const ThemeColorGridModel&) = delete;
    ThemeColorGridModel(ThemeColorGridModel&&)                 = delete;
    ThemeColorGridModel& operator=(ThemeColorGridModel&&)      = delete;

    [[nodiscard]] size_t GetRowCount() const noexcept override
    {
        return _keys.size();
    }

    [[nodiscard]] size_t GetColumnCount() const noexcept override
    {
        return _columns.size();
    }

    [[nodiscard]] GridColumnDesc GetColumn(size_t columnIndex) const override
    {
        return columnIndex < _columns.size() ? _columns[columnIndex] : GridColumnDesc{};
    }

    void GetCellData(size_t rowIndex, size_t columnIndex, GridCellData& outCell) const override
    {
        outCell = {};
        if (rowIndex >= _keys.size())
        {
            return;
        }

        const std::wstring& key = _keys[rowIndex];
        switch (columnIndex)
        {
            case 0u: outCell.text = key; break;
            case 1u:
                if (const std::optional<uint32_t> color = _session.GetThemePreviewModel().GetEffectiveColor(key))
                {
                    outCell.text           = Common::Settings::FormatColor(color.value());
                    outCell.hasSwatchValue = true;
                    outCell.swatchArgb     = color.value();
                }
                break;
            case 2u: outCell.text = _session.GetThemePreviewModel().GetAuthoredColorText(key); break;
            default: break;
        }

        outCell.tooltipText = outCell.text.empty() ? key : outCell.text;
        outCell.multiline   = false;
    }

    [[nodiscard]] uint64_t GetStableRowId(size_t rowIndex) const noexcept override
    {
        return static_cast<uint64_t>(rowIndex) + 1u;
    }

    [[nodiscard]] std::optional<size_t> FindRowByStableId(uint64_t rowId) const noexcept override
    {
        if (rowId == 0u)
        {
            return std::nullopt;
        }

        const size_t rowIndex = static_cast<size_t>(rowId - 1u);
        return rowIndex < _keys.size() ? std::optional<size_t>(rowIndex) : std::nullopt;
    }

    void SetKeys(std::vector<std::wstring> keys)
    {
        _keys = std::move(keys);
    }

    [[nodiscard]] std::optional<std::wstring> GetKeyAt(size_t rowIndex) const
    {
        return rowIndex < _keys.size() ? std::optional<std::wstring>(_keys[rowIndex]) : std::nullopt;
    }

    [[nodiscard]] std::optional<size_t> FindKey(std::wstring_view key) const noexcept
    {
        for (size_t index = 0u; index < _keys.size(); ++index)
        {
            if (_keys[index] == key)
            {
                return index;
            }
        }
        return std::nullopt;
    }

private:
    HINSTANCE _instance = nullptr;
    const RedConfigure::RedConfigureSession& _session;
    std::vector<GridColumnDesc> _columns;
    std::vector<std::wstring> _keys;
};
}
