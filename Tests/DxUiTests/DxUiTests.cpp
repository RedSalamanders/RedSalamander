#define REDSAL_DEFINE_TRACE_PROVIDER
#include "DxUiTestHelpers.h"
#include "Ui/AnimationDispatcher.h"

#include <optional>
#include <string>
#include <string_view>

void RunGridTests();
void RunThemeTests();
void RunControlTests();
void RunComboBoxTests();
void RunWindowHostTests();
void RunTreeTests();
void RunTextFieldTests();
void RunNativeTextInputTests();
void RunMultilineTextTests();
void RunReadOnlyTests();
void RunTooltipTests();
void RunRenderingTests();
void RunAnimationTests();
void RunAccessibilityTests();
void RunMenuTests();
void RunNewControlTests();
void RunGalleryGenerator(const std::filesystem::path& outputPath);

int wmain(int argc, wchar_t** argv)
{
    std::optional<std::wstring> suiteFilter;
    std::optional<std::filesystem::path> perfJsonlPath;
    std::optional<std::filesystem::path> galleryOutputPath;
    bool writeBaselines = false;
    for (int argIndex = 1; argIndex < argc; ++argIndex)
    {
        const std::wstring_view arg                  = argv[argIndex] ? std::wstring_view(argv[argIndex]) : std::wstring_view{};
        constexpr std::wstring_view kSuitePrefix     = L"--suite=";
        constexpr std::wstring_view kPerfJsonlPrefix = L"--perf-jsonl=";
        constexpr std::wstring_view kGalleryPrefix   = L"--gallery-output=";
        if (arg.rfind(kSuitePrefix, 0) == 0)
        {
            if (arg.size() == kSuitePrefix.size())
            {
                std::wcerr << L"Missing suite name for --suite.\n";
                return 2;
            }
            suiteFilter = std::wstring(arg.substr(kSuitePrefix.size()));
            continue;
        }
        if (arg == L"--write-baselines")
        {
            writeBaselines = true;
            continue;
        }
        if (arg.rfind(kPerfJsonlPrefix, 0) == 0)
        {
            if (arg.size() == kPerfJsonlPrefix.size())
            {
                std::wcerr << L"Missing perf JSONL path for --perf-jsonl.\n";
                return 2;
            }
            perfJsonlPath = std::filesystem::path(arg.substr(kPerfJsonlPrefix.size()));
            continue;
        }
        if (arg.rfind(kGalleryPrefix, 0) == 0)
        {
            if (arg.size() == kGalleryPrefix.size())
            {
                std::wcerr << L"Missing output path for --gallery-output.\n";
                return 2;
            }
            galleryOutputPath = std::filesystem::path(arg.substr(kGalleryPrefix.size()));
            continue;
        }
        if (! arg.empty() && arg[0] == L'-')
        {
            std::wcerr << L"Unknown argument: " << arg << L'\n';
            return 2;
        }
        if (! arg.empty())
        {
            suiteFilter = std::wstring(arg);
            continue;
        }
    }

    SetDxUiWriteBaselines(writeBaselines);
    if (perfJsonlPath.has_value())
    {
        Debug::Perf::ConfigureJsonlOutput(perfJsonlPath.value(), L"DxUiTests", L"Debug");
    }
    const auto perfCleanup = wil::scope_exit([&] { Debug::Perf::ClearJsonlOutput(); });

    const auto shouldRunSuite = [&](const char* name) -> bool
    {
        if (! suiteFilter.has_value())
        {
            return true;
        }

        std::wstring wideName;
        while (*name != '\0')
        {
            wideName.push_back(static_cast<wchar_t>(*name));
            ++name;
        }
        return _wcsicmp(wideName.c_str(), suiteFilter->c_str()) == 0;
    };

    auto runSuite = [](const char* name, void (*fn)())
    {
        std::cerr << "[START] " << name << '\n' << std::flush;
        fn();
        RedSalamander::Ui::AnimationDispatcher::GetInstance().Shutdown();
        std::cerr << "[DONE] " << name << '\n' << std::flush;
    };

    bool ranAnySuite = false;
    if (suiteFilter.has_value() && shouldRunSuite("Gallery"))
    {
        const std::filesystem::path outputPath =
            galleryOutputPath.value_or(FindRepoRootForDxUiTests() / L"Specs" / L"TestRuns" / L"DxUiGallery" / L"DxUiControlGallery.png");
        std::cerr << "[START] Gallery\n" << std::flush;
        RunGalleryGenerator(outputPath);
        RedSalamander::Ui::AnimationDispatcher::GetInstance().Shutdown();
        std::cerr << "[DONE] Gallery\n" << std::flush;
        ranAnySuite = true;
    }
    if (shouldRunSuite("Grid"))
    {
        runSuite("Grid", RunGridTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("Theme"))
    {
        runSuite("Theme", RunThemeTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("Control"))
    {
        runSuite("Control", RunControlTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("Menu"))
    {
        runSuite("Menu", RunMenuTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("NewControls"))
    {
        runSuite("NewControls", RunNewControlTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("TextField"))
    {
        runSuite("TextField", RunTextFieldTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("NativeTextInput"))
    {
        runSuite("NativeTextInput", RunNativeTextInputTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("ComboBox"))
    {
        runSuite("ComboBox", RunComboBoxTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("WindowHost"))
    {
        runSuite("WindowHost", RunWindowHostTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("Tree"))
    {
        runSuite("Tree", RunTreeTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("MultilineText"))
    {
        runSuite("MultilineText", RunMultilineTextTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("ReadOnly"))
    {
        runSuite("ReadOnly", RunReadOnlyTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("Tooltip"))
    {
        runSuite("Tooltip", RunTooltipTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("Rendering"))
    {
        runSuite("Rendering", RunRenderingTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("Animation"))
    {
        runSuite("Animation", RunAnimationTests);
        ranAnySuite = true;
    }
    if (shouldRunSuite("Accessibility"))
    {
        runSuite("Accessibility", RunAccessibilityTests);
        ranAnySuite = true;
    }

    if (! ranAnySuite)
    {
        std::wcerr << L"Unknown suite filter: " << suiteFilter.value_or(L"<empty>") << L'\n';
        return 2;
    }

    std::cout << "All DxUi tests passed.\n";
    return 0;
}
