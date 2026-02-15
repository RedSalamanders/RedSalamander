---
name: direct2d-rendering
description: Direct2D and DirectWrite graphics rendering patterns for Windows. Use when implementing 2D graphics, text rendering, hardware-accelerated drawing, handling device loss, or working with ID2D1DeviceContext.
metadata:
  author: DualTail
  version: "1.0"
---

# Direct2D/DirectWrite Rendering

## Graphics Stack
- **Direct2D 1.1** - Hardware-accelerated 2D graphics
- **DirectWrite** - Advanced text rendering
- **Direct3D 11** - GPU resources and swap chain
- **DXGI 1.3** - Display management

## Key Patterns

### Creating Resources
```cpp
wil::com_ptr<ID2D1SolidColorBrush> _brush;
THROW_IF_FAILED(_d2dContext->CreateSolidColorBrush(
    D2D1::ColorF(D2D1::ColorF::Black), &_brush));
```

### BeginDraw/EndDraw with RAII
```cpp
{
    HRESULT hr = S_OK;
    {
        _d2dContext->BeginDraw();
        auto endDraw = wil::scope_exit([&] { hr = _d2dContext->EndDraw(); });
        
        _d2dContext->Clear(D2D1::ColorF(1.0f, 1.0f, 1.0f));
        // Drawing operations...
    }
    
    if (hr == D2DERR_RECREATE_TARGET) 
    {
        RecreateDeviceResources();
    }
}
```

### Text with DirectWrite
```cpp
wil::com_ptr<IDWriteTextLayout> layout;
THROW_IF_FAILED(_dwriteFactory->CreateTextLayout(
    text.c_str(), static_cast<UINT32>(text.length()),
    _textFormat.get(), maxWidth, maxHeight, &layout));

_d2dContext->DrawTextLayout(D2D1::Point2F(x, y), layout.get(), _brush.get());
```

## Best Practices

1. **Validate device state** before rendering
2. **Handle D2DERR_RECREATE_TARGET** for device loss
3. **Cache brushes and layouts** - minimize recreation
4. **Use dirty region updates** for performance
5. **Handle DPI changes** - recalculate layouts

## DPI Awareness

```cpp
void OnDpiChanged(UINT dpi) 
{
    _d2dContext->SetDpi(static_cast<float>(dpi), static_cast<float>(dpi));
    IconCache::SetDpi(dpi);
    RecreateDeviceDependentResources();
}
```

## Text Rendering (ColorTextView)

- Use `IDWriteTextLayout` for complex text scenarios
- Implement virtualization for large documents
- Cache frequently used resources (brushes, layouts)
- Handle Unicode properly with UTF-16 encoding
- Optimize for smooth scrolling with async layout updates

## Performance Considerations

- **Async Operations**: Use thread pools for heavy computations
- **Caching**: Implement LRU caches for expensive resources
- **Memory**: Monitor usage, implement cleanup strategies
- **Rendering**: Use offscreen rendering and dirty region updates
- **Measurements**: Profile performance-critical paths

## Device Loss Handling

```cpp
HRESULT hr = _d2dContext->EndDraw();
if (hr == D2DERR_RECREATE_TARGET) {
    // Device lost - recreate all device-dependent resources
    DiscardDeviceResources();
    CreateDeviceResources();
} else if (FAILED(hr)) {
    // Other error - log and handle
    Debug::Error(L"EndDraw failed: {:#x}", hr);
}
```
