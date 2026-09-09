#include "TerminalInternal.h"

#include <atomic>
#include <mutex>
#include <set>

namespace TerminalPluginDetail
{
std::atomic_uint32_t g_terminalObjectCount{0u};
std::atomic_uint32_t g_terminalWindowCount{0u};
std::mutex g_terminalWindowClassMutex;
bool g_terminalWindowClassRegistered = false;
std::mutex g_terminalSchemaMutex;
std::set<std::string> g_terminalSchemas;

} // namespace TerminalPluginDetail
