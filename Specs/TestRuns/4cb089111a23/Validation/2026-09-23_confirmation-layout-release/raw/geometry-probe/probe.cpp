#include <cstdio>
#include <format>
#include "AlertOverlay.h"
void Fail(std::wstring_view text) { std::wprintf(L"FAIL: %.*s\n", static_cast<int>(text.size()), text.data()); }
void AppendLog(std::wstring_view text) { std::wprintf(L"%.*s\n", static_cast<int>(text.size()), text.data()); }
[[nodiscard]] bool ValidateFileOpsConfirmationLayout()
{
    using namespace RedSalamander::Ui;
    wil::com_ptr<IDWriteFactory> factory;
    if (FAILED(DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(factory.put()))))
        return false;
    for (const float width : {320.0f, 480.0f, 920.0f})
    {
        AlertOverlay overlay;
        overlay.SetScrollingEnabled(true);
        AlertModel model;
        model.title   = L"Confirmer la copie des photographies";
        model.message = L"Vérifiez les emplacements avant de poursuivre cette opération.\n"
                        L"C:\\Photographies\\Archives familiales de plusieurs générations\\Sélection définitive pour archivage.\n"
                        L"D:\\Sauvegardes\\Photographies retouchées en très haute résolution.";
        for (uint32_t id = 1u; id <= 4u; ++id)
            model.options.push_back({id,
                                     L"Vérification après la copie",
                                     {{0u, L"Désactivée"}, {1u, L"Vérifier intégralement chaque fichier après la copie avant de supprimer le fichier source"}},
                                     1u});
        model.buttons = {{10u, L"Copier les fichiers sélectionnés", true}, {11u, L"Annuler cette opération", false}};
        overlay.SetModel(std::move(model));
        if (! overlay.EnsureLayout(factory.get(), width, 480.0f) || ! overlay.DebugAllControlLabelsFit() || (width < 920.0f && overlay.GetScrollMaximum() <= 0.0f))
        {
            Fail(std::format(L"Confirmation layout lost/clipped a control at width {}.", width));
            return false;
        }
        const auto initialLastOption = overlay.GetVisibleControlRect({AlertHitTest::Part::Option, 4u});
        const uint64_t layoutPasses  = overlay.DebugGetLayoutPassCount();
        for (uint32_t repeat = 0u; repeat < 32u; ++repeat)
            static_cast<void>(overlay.EnsureLayout(factory.get(), width, 480.0f));
        if (overlay.DebugGetLayoutPassCount() != layoutPasses)
        {
            Fail(L"Confirmation rebuilt unchanged text/geometry during warm paints.");
            return false;
        }
        for (uint32_t id = 1u; id <= 4u; ++id)
        {
            const AlertHitTest focus{AlertHitTest::Part::Option, id};
            static_cast<void>(overlay.FocusControl(focus));
            const auto rect = overlay.GetVisibleControlRect(focus);
            const auto hit  = overlay.HitTest(D2D1::Point2F((rect.left + rect.right) * 0.5f, (rect.top + rect.bottom) * 0.5f));
            if (rect.bottom <= rect.top || hit.part != focus.part || hit.buttonId != id)
            {
                Fail(L"Confirmation focus did not reveal the correct option hit target.");
                return false;
            }
        }
        static_cast<void>(overlay.ScrollBy(-overlay.GetScrollMaximum()));
        static_cast<void>(overlay.EnsureLayout(factory.get(), width, 480.0f));
        const auto resetLastOption = overlay.GetVisibleControlRect({AlertHitTest::Part::Option, 4u});
        if (overlay.GetScrollOffset() != 0.0f || std::abs(resetLastOption.top - initialLastOption.top) > 0.5f ||
            std::abs(resetLastOption.bottom - initialLastOption.bottom) > 0.5f)
        {
            Fail(std::format(L"Confirmation scroll reset changed visible geometry at width {}: offset {:.1f}, original ({:.1f},{:.1f}), reset ({:.1f},{:.1f}).",
                             width,
                             overlay.GetScrollOffset(),
                             initialLastOption.top,
                             initialLastOption.bottom,
                             resetLastOption.top,
                             resetLastOption.bottom));
            return false;
        }
        const auto footer    = overlay.GetVisibleControlRect({AlertHitTest::Part::Button, 11u});
        const auto footerHit = overlay.HitTest(D2D1::Point2F((footer.left + footer.right) * 0.5f, (footer.top + footer.bottom) * 0.5f));
        if (footerHit.part != AlertHitTest::Part::Button || footerHit.buttonId != 11u)
        {
            Fail(L"Confirmation option hit testing overlapped the pinned footer.");
            return false;
        }
        static_cast<void>(overlay.FocusControl({AlertHitTest::Part::Option, 4u}));
        static_cast<void>(overlay.ActivateOption(4u));
        static_cast<void>(overlay.EnsureLayout(factory.get(), width, 480.0f));
        overlay.RevealFocusedControl();
        if (! overlay.DebugAllControlLabelsFit() || overlay.GetScrollOffset() > overlay.GetScrollMaximum())
            return false;
        static_cast<void>(overlay.EnsureLayout(factory.get(), 920.0f, 1600.0f));
        if (overlay.GetScrollOffset() != 0.0f || overlay.GetScrollMaximum() != 0.0f || ! overlay.DebugAllControlLabelsFit())
        {
            Fail(L"Confirmation resize did not clamp scrolling after content fit.");
            return false;
        }
        static_cast<void>(overlay.EnsureLayout(factory.get(), 320.0f, 80.0f));
        static_cast<void>(overlay.FocusControl({AlertHitTest::Part::Button, 11u}));
        const auto shortAction = overlay.GetVisibleControlRect({AlertHitTest::Part::Button, 11u});
        std::wprintf(L"short: action=(%f,%f,%f,%f), fit=%d, textbottom=%f, offset=%f,max=%f\n",shortAction.left,shortAction.top,shortAction.right,shortAction.bottom,overlay.DebugAllControlLabelsFit(),overlay.GetVisibleTextRect().bottom,overlay.GetScrollOffset(),overlay.GetScrollMaximum());
        if (shortAction.bottom <= shortAction.top || shortAction.top < 0.0f || shortAction.bottom > 80.0f || ! overlay.DebugAllControlLabelsFit() ||
            overlay.GetVisibleTextRect().bottom != 0.0f)
        {
            Fail(L"Short confirmation lost its final action or reported invisible text as visible.");
            return false;
        }
        AppendLog(std::format(
            L"ConfirmationLayout: width {} wrapped labels, all options reachable, footer isolation, resize clamping, 32 warm layout hits passed", width));
    }
    return true;
}


int main(){return ValidateFileOpsConfirmationLayout() ? 0 : 1;}
