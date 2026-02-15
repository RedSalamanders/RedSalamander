// MonitorTest.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <chrono>
#include <fcntl.h>
#include <io.h>
#include <iomanip>
#include <iostream>
#include <random>
#include <string>
#include <thread>
#include <vector>

#include <Windows.h>

// Define local ETW provider for MonitorTest
// NOTE: Each executable must have its own TraceLogging provider instance
// to avoid cross-DLL symbol conflicts. The GUID is the same as Common.dll
// so external trace sessions will receive events from both.
#define REDSAL_DEFINE_TRACE_PROVIDER
#include "ExceptionHelpers.h"
#include "Helpers.h"

// number of messages to send in each test
auto constexpr kMaxLoop = 50000;

namespace
{
std::wstring GenerateDiverseMessage(size_t index, Debug::InfoParam::Type type)
{
    static std::random_device rd;
    static std::mt19937 gen(rd());
    static std::uniform_int_distribution<> styleDist(0, 7);

    const std::wstring emojis[] = {L"🎮", L"🌟", L"💻", L"🚀", L"✨", L"🎯", L"🌈", L"💫", L"🔥", L"⚠️",
                                   L"🎪", L"🎨", L"🎭", L"🏆", L"🌺", L"💎", L"🔮", L"⭐", L"🌸", L"💥"};
    const std::wstring words[]  = {L"System",
                                   L"Process",
                                   L"Thread",
                                   L"Application",
                                   L"Monitor",
                                   L"Handler",
                                   L"Function",
                                   L"Method",
                                   L"こんにちは",
                                   L"مع السلامة",
                                   L"Пожалуйста",
                                   L"สวัสดี",
                                   L"你好",
                                   L"Bonjour",
                                   L"Hola"};

    // Type-specific prefixes for easy identification
    const wchar_t* typePrefix = L"[UNKNOWN]";
    switch (type)
    {
        case Debug::InfoParam::Type::Text: typePrefix = L"[TEXT]"; break;
        case Debug::InfoParam::Type::Error: typePrefix = L"[ERROR]"; break;
        case Debug::InfoParam::Type::Warning: typePrefix = L"[WARN]"; break;
        case Debug::InfoParam::Type::Info: typePrefix = L"[INFO]"; break;
        case Debug::InfoParam::Type::Debug: typePrefix = L"[DEBUG]"; break;
        case Debug::InfoParam::Type::All: typePrefix = L"[ALL]"; break;
    }

    int style = styleDist(gen);
    std::wstring msg;

    switch (style)
    {
        case 0: // Short with emoji
            msg = std::format(L"{} {} [{}] Quick message", typePrefix, emojis[index % 20], index);
            break;
        case 1: // Medium with status code
            msg = std::format(L"{} [{}] Processing {} {} operation with status code 0x{:08X}",
                              typePrefix,
                              index,
                              words[index % 15],
                              (type == Debug::InfoParam::Type::Error)     ? L"ERROR"
                              : (type == Debug::InfoParam::Type::Warning) ? L"WARNING"
                                                                          : L"NORMAL",
                              static_cast<unsigned>(index * 0x1000 + static_cast<int>(type)));
            break;
        case 2: // Long with wrapping
            msg = std::format(L"{} [{}] This is a longer message that contains multiple words and should test text wrapping capabilities with {} and {} "
                              L"processing {} data structures",
                              typePrefix,
                              index,
                              words[gen() % 15],
                              words[gen() % 15],
                              words[gen() % 15]);
            break;
        case 3: // Multiline with 2 newlines
            msg = std::format(L"{} [{}] First line\nSecond line with data: {}\nThird line complete", typePrefix, index, words[index % 15]);
            break;
        case 4: // International characters
            msg = std::format(L"{} [{}] Mixed: Hello {} {} {} {} {}", typePrefix, index, words[10], words[11], words[12], words[13], emojis[index % 20]);
            break;
        case 5: // Simulated stack trace with 3 newlines
            msg = std::format(
                L"{} [{}] Exception in {}::{}() at line {}\n  Callstack: main->processData->validateInput\n  Context: {} processing\n  Module: {}.dll",
                typePrefix,
                index,
                words[gen() % 8],
                words[gen() % 8],
                100 + (index % 500),
                words[index % 8],
                words[gen() % 8]);
            break;
        case 6: // Performance metrics
            msg = std::format(L"{} [{}] Performance: {} took ({}ms) | Memory: {}KB | CPU: {}%",
                              typePrefix,
                              index,
                              words[index % 8],
                              static_cast<double>(index % 100) / 10.0,
                              1024 + (index % 4096),
                              15 + (index % 70));
            break;
        case 7: // Single newline embedded
            msg = std::format(L"{} [{}] Line one with {}\nLine two with data: {} complete", typePrefix, index, words[gen() % 8], words[gen() % 8]);
            break;
    }

    return msg;
}
void PrintProgress(size_t current, size_t total, std::wstring_view label, double elapsedSeconds, size_t messagesPerSecond)
{
    const int barWidth   = 50;
    const float progress = static_cast<float>(current) / static_cast<float>(total);
    const int pos        = static_cast<int>(barWidth * progress);

    std::wcout << L"\r" << label << L" [";
    for (int i = 0; i < barWidth; ++i)
    {
        if (i < pos)
            std::wcout << L"#";
        else if (i == pos)
            std::wcout << L">";
        else
            std::wcout << L"-";
    }
    std::wcout << L"] " << static_cast<int>(progress * 100.0f) << L"% (" << current << L"/" << total << L") | ";
    std::wcout << std::fixed << std::setprecision(2) << elapsedSeconds << L"s | ";
    std::wcout << messagesPerSecond << L" msg/s";
    std::wcout << L"    ";
    std::wcout.flush();
}

void RunDiverseMessages(size_t messageCount, std::wstring_view label)
{
    const Debug::InfoParam::Type types[] = {Debug::InfoParam::Type::Text,
                                            Debug::InfoParam::Type::Info,
                                            Debug::InfoParam::Type::Warning,
                                            Debug::InfoParam::Type::Error,
                                            Debug::InfoParam::Type::Debug};

    std::wcout << L"\nStarting " << label << L" - " << messageCount << L" diverse messages\n";

    const auto startTime   = std::chrono::steady_clock::now();
    auto lastUpdateTime    = startTime;
    size_t lastUpdateCount = 0;

    for (size_t i = 0; i < messageCount; ++i)
    {
        const auto type        = types[i % 5];
        const std::wstring msg = GenerateDiverseMessage(i, type);
        Debug::Out(type, L"{}", msg);

        if ((i % 100) == 0 || i == messageCount - 1)
        {
            const auto now             = std::chrono::steady_clock::now();
            const auto elapsed         = std::chrono::duration<double, std::milli>(now - startTime).count();
            const auto sinceLastUpdate = std::chrono::duration<double, std::milli>(now - lastUpdateTime).count();

            size_t messagesPerSecond = 0;
            if (sinceLastUpdate > 0.0)
            {
                const size_t messagesSinceUpdate = (i + 1) - lastUpdateCount;
                messagesPerSecond                = static_cast<size_t>(static_cast<double>(messagesSinceUpdate) / sinceLastUpdate);
            }

            PrintProgress(i + 1, messageCount, label, elapsed, messagesPerSecond);

            lastUpdateTime  = now;
            lastUpdateCount = i + 1;
        }

        if ((i % 512) == 0)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(1));
        }
    }

    const auto endTime       = std::chrono::steady_clock::now();
    const auto totalDuration = std::chrono::duration<double>(endTime - startTime).count();
    const auto avgRate       = static_cast<size_t>(static_cast<double>(messageCount) / totalDuration);

    std::wcout << L"\n" << label << L" completed in " << std::fixed << std::setprecision(2) << totalDuration << L"s (avg: " << avgRate << L" msg/s)\n";
}

[[maybe_unused]] void RunHighRateBurst(size_t messageCount, std::wstring_view label)
{
    std::wcout << L"\nStarting " << label << L" - " << messageCount << L" messages\n";

    const auto startTime   = std::chrono::steady_clock::now();
    auto lastUpdateTime    = startTime;
    size_t lastUpdateCount = 0;

    for (size_t i = 0; i < messageCount; ++i)
    {
        Debug::Out(Debug::InfoParam::Type::Info, L"[{}] High-rate message {}", label, i);

        // Update progress every 100 messages
        if ((i % 100) == 0 || i == messageCount - 1)
        {
            const auto now             = std::chrono::steady_clock::now();
            const auto elapsed         = std::chrono::duration<double>(now - startTime).count();
            const auto sinceLastUpdate = std::chrono::duration<double>(now - lastUpdateTime).count();

            size_t messagesPerSecond = 0;
            if (sinceLastUpdate > 0.0)
            {
                const size_t messagesSinceUpdate = (i + 1) - lastUpdateCount;
                messagesPerSecond                = static_cast<size_t>(static_cast<double>(messagesSinceUpdate) / sinceLastUpdate);
            }

            PrintProgress(i + 1, messageCount, label, elapsed, messagesPerSecond);

            lastUpdateTime  = now;
            lastUpdateCount = i + 1;
        }

        if ((i % 512) == 0)
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(1));
        }
    }

    const auto endTime       = std::chrono::steady_clock::now();
    const auto totalDuration = std::chrono::duration<double>(endTime - startTime).count();
    const auto avgRate       = static_cast<size_t>(static_cast<double>(messageCount) / totalDuration);

    std::wcout << L"\n" << label << L" completed in " << std::fixed << std::setprecision(2) << totalDuration << L"s (avg: " << avgRate << L" msg/s)\n";
}
} // namespace

// Separate function with C++ objects (cannot use __try/__except)
static int RunMonitorTest()
{
    // Enable Unicode output for console
    _setmode(_fileno(stdout), _O_U16TEXT);

    std::wcout << L"=== MonitorTest Starting ===\n";
    std::wcout << L"Generating diverse messages with all severity types\n";

    // Force ETW registration at startup
    const bool registered = Debug::detail::EnsureTraceLoggingRegistered();

    std::wcout << L"ETW Status: ";
    if (registered)
    {
        std::wcout << L"✓ Registered successfully\n";
        std::wcout << L"  Provider GUID: {440c70f6-6c6b-4ff7-9a3f-0b7db411b31a}\n\n";
    }
    else
    {
        std::wcout << L"✗ Registration failed\n";
        std::wcout << L"  Warning: Events will not be sent\n";
        std::wcout << L"  Possible causes:\n";
        std::wcout << L"    • TraceLogging provider definition conflict\n";
        std::wcout << L"    • Run as Administrator and rebuild if needed\n";
        std::wcout << L"    • Check debug output for HRESULT error code\n\n";
    }

    const auto testStartTime = std::chrono::steady_clock::now();

    std::wcout << L"Running ETW-only tests...\n\n";

    std::wcout << L"Test 1: Diverse messages (burst-A)\n";
    RunDiverseMessages(kMaxLoop, L"burst-A");

    std::wcout << L"\nTest 2: Diverse messages (burst-B)\n";
    RunDiverseMessages(kMaxLoop, L"burst-B");

    std::wcout << L"\nTest 3: Diverse messages (mixed)\n";
    RunDiverseMessages(kMaxLoop, L"mixed");

    const auto testEndTime       = std::chrono::steady_clock::now();
    const auto totalTestDuration = std::chrono::duration<double>(testEndTime - testStartTime).count();

    const auto stats = Debug::GetTransportStats();

    std::wcout << L"\n=== ETW Transport Statistics ===\n";
    std::wcout << L"  Total duration:   " << std::fixed << std::setprecision(2) << totalTestDuration << L"s\n";
    std::wcout << L"  ETW written:      " << stats.etwWritten << L"\n";
    std::wcout << L"  ETW failed:       " << stats.etwFailed << L"\n";
    std::wcout << L"  Overall rate:     " << static_cast<size_t>(static_cast<double>(stats.etwWritten) / totalTestDuration) << L" msg/s\n";

    if (stats.etwFailed > 0)
    {
        std::wcout << L"\n⚠️  Warning: " << stats.etwFailed << L" ETW events failed to write\n";
        std::wcout << L"  This may indicate ETW registration issues or insufficient privileges\n";
    }

    std::wcout << L"\n💡 Launch RedSalamanderMonitor.exe to view these messages in real-time!\n";
    std::wcout << L"=========================\n";

    OutputDebugStringW(std::format(L"ETW Transport stats - etwWritten: {}, etwFailed: {}\n", stats.etwWritten, stats.etwFailed).c_str());

    return 0;
}

int main()
{
    // Use SEH to catch all exceptions (no C++ objects in this scope)
    __try
    {
        return RunMonitorTest();
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        // Handle all exceptions including SEH exceptions
        const DWORD exceptionCode    = GetExceptionCode();
        const wchar_t* exceptionName = exception::GetExceptionName(exceptionCode);

        wchar_t errorMsg[512];
        swprintf_s(errorMsg,
                   512,
                   L"Fatal Exception in MonitorTest\n\n"
                   L"Exception: %s (0x%08X)\n\n"
                   L"The application will now terminate.",
                   exceptionName,
                   static_cast<unsigned>(exceptionCode));

        // Try to output to both console and debug output
        std::wcerr << errorMsg << L"\n";
        OutputDebugStringW(errorMsg);
        MessageBoxW(nullptr, errorMsg, L"MonitorTest - Fatal Error", MB_OK | MB_ICONERROR);

        return -1;
    }
}
