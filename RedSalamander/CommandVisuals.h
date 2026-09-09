#pragma once

#include "CommandRegistry.h"

#include <string>

[[nodiscard]] std::wstring ResolveCommandVisualText(CommandVisualId visualId, bool fluentIconFontAvailable);
