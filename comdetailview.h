#pragma once
#include "regview.h"
#include "typelibview.h"

class ComDetailView : public CTabViewImpl<ComDetailView>
{
public:
    using Base = CTabViewImpl<ComDetailView>;

BEGIN_MSG_MAP(ComDetailView)
        MESSAGE_HANDLER(WM_SELCHANGED, OnSelChanged)
        REFLECT_NOTIFY_CODE(TBVN_PAGEACTIVATED)
        CHAIN_MSG_MAP(Base)
        ALT_MSG_MAP(1) // tab control
    END_MSG_MAP()

    DECLARE_WND_CLASS_EX(_T("ComDetailView"), 0, COLOR_APPWORKSPACE)

    bool CreateTabControl();
    void UpdateLayout();
    LRESULT OnSelChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);

private:
    RegistryView m_regView{};
    TypeLibView m_typeLibView{};
    CImageList m_imageList;
};
