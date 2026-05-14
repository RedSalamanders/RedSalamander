#include "pch.h"

#include "FolderViewColumnLayout.h"

#include <array>
#include <vector>

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace
{
[[nodiscard]] FolderViewColumnLayout::Input MakeInput(std::span<const FolderViewColumnLayout::ItemTextMetrics> items,
                                                      bool includeDetails,
                                                      bool includeMetadata) noexcept
{
    return FolderViewColumnLayout::Input{
        .clientWidthDip      = 1000.0f,
        .clientHeightDip     = 32.0f,
        .tileHeightDip       = 10.0f,
        .rowSpacingDip       = 1.0f,
        .iconSizeDip         = 16.0f,
        .iconTextGapDip      = 4.0f,
        .horizontalPaddingDip = 8.0f,
        .columnSpacingDip    = 6.0f,
        .textWidthSafetyDip  = 2.0f,
        .includeDetailsLine  = includeDetails,
        .includeMetadataLine = includeMetadata,
        .items               = items,
    };
}
}

namespace PerformanceTests2
{
TEST_CLASS(FolderViewColumnLayoutTests)
{
public:
    TEST_METHOD(VisibleColumnWidths_DifferentColumnsDoNotShareGlobalMax)
    {
        std::array<FolderViewColumnLayout::ItemTextMetrics, 12> items{};
        for (auto& item : items)
        {
            item.labelWidthDip = 20.0f;
        }

        items[3].labelWidthDip = 240.0f;
        items[4].labelWidthDip = 260.0f;
        items[5].labelWidthDip = 250.0f;
        items[8].labelWidthDip = 75.0f;

        const FolderViewColumnLayout::Result result = FolderViewColumnLayout::Resolve(MakeInput(items, false, false));

        Assert::AreEqual(3, result.rowsPerColumn);
        Assert::AreEqual(static_cast<size_t>(4), result.columns.size());
        Assert::IsTrue(result.columns[1].widthDip > result.columns[0].widthDip + 200.0f);
        Assert::IsTrue(result.columns[1].widthDip > result.columns[2].widthDip + 150.0f);
        Assert::IsTrue(result.columns[3].widthDip < result.columns[1].widthDip);
        Assert::IsTrue(result.columns[0].widthDip < result.maxColumnWidthDip);

        for (size_t index = 1; index < result.columns.size(); ++index)
        {
            const float expectedLeft = result.columns[index - 1u].RightDip() + 6.0f;
            Assert::IsTrue(std::abs(result.columns[index].leftDip - expectedLeft) < 0.01f);
        }
    }

    TEST_METHOD(VisibleColumnWidths_DetailedAndMetadataLinesStayColumnLocal)
    {
        std::array<FolderViewColumnLayout::ItemTextMetrics, 8> items{};
        for (auto& item : items)
        {
            item.labelWidthDip   = 20.0f;
            item.detailsWidthDip = 30.0f;
            item.metadataWidthDip = 25.0f;
        }

        items[4].detailsWidthDip  = 170.0f;
        items[6].metadataWidthDip = 210.0f;

        const FolderViewColumnLayout::Result result = FolderViewColumnLayout::Resolve(MakeInput(items, true, true));

        Assert::AreEqual(3, result.rowsPerColumn);
        Assert::AreEqual(static_cast<size_t>(3), result.columns.size());
        Assert::IsTrue(result.columns[1].widthDip > result.columns[0].widthDip + 120.0f);
        Assert::IsTrue(result.columns[2].widthDip > result.columns[1].widthDip + 30.0f);
        Assert::IsTrue(result.columns[0].widthDip < result.columns[1].widthDip);
    }

    TEST_METHOD(ScrollStops_FirstRightSkipsInitialLeftGap)
    {
        const std::array<FolderViewColumnLayout::Column, 4> columns = {{
            {.startIndex = 0, .itemCount = 3, .leftDip = 18.0f, .widthDip = 120.0f},
            {.startIndex = 3, .itemCount = 3, .leftDip = 156.0f, .widthDip = 300.0f},
            {.startIndex = 6, .itemCount = 3, .leftDip = 474.0f, .widthDip = 120.0f},
            {.startIndex = 9, .itemCount = 3, .leftDip = 612.0f, .widthDip = 120.0f},
        }};

        Assert::AreEqual(156.0f, FolderViewColumnLayout::ResolveNextScrollStop(0.0f, 500.0f, columns), 0.01f);
        Assert::AreEqual(156.0f, FolderViewColumnLayout::ResolveNextScrollStop(18.0f, 500.0f, columns), 0.01f);
        Assert::AreEqual(474.0f, FolderViewColumnLayout::ResolveNextScrollStop(156.0f, 500.0f, columns), 0.01f);
    }

    TEST_METHOD(ScrollStops_FirstLeftRestoresInitialLeftGap)
    {
        const std::array<FolderViewColumnLayout::Column, 4> columns = {{
            {.startIndex = 0, .itemCount = 3, .leftDip = 18.0f, .widthDip = 120.0f},
            {.startIndex = 3, .itemCount = 3, .leftDip = 156.0f, .widthDip = 300.0f},
            {.startIndex = 6, .itemCount = 3, .leftDip = 474.0f, .widthDip = 120.0f},
            {.startIndex = 9, .itemCount = 3, .leftDip = 612.0f, .widthDip = 120.0f},
        }};

        Assert::AreEqual(0.0f, FolderViewColumnLayout::ResolvePreviousScrollStop(156.0f, 500.0f, columns), 0.01f);
        Assert::AreEqual(0.0f, FolderViewColumnLayout::ResolvePreviousScrollStop(18.0f, 500.0f, columns), 0.01f);
        Assert::AreEqual(156.0f, FolderViewColumnLayout::ResolvePreviousScrollStop(474.0f, 500.0f, columns), 0.01f);
    }
};
}
