#pragma once

#if defined(RS_MONITOR_SHOW_INVALID_RECTS)
#define RS_MONITOR_INVALID_RECT_VISUALIZATION_ENABLED 1
#else
#define RS_MONITOR_INVALID_RECT_VISUALIZATION_ENABLED 0
#endif

namespace RedSalamanderMonitor
{
inline constexpr bool kInvalidRectVisualizationEnabled = RS_MONITOR_INVALID_RECT_VISUALIZATION_ENABLED != 0;
}
