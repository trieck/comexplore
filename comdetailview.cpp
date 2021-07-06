#include "StdAfx.h"
#include "comdetailview.h"
#include "resource.h"

bool ComDetailView::CreateTabControl()
{
    auto result = CTabViewImpl<ComDetailView>::CreateTabControl();
    if (!result) {
        return false;
    }

    m_imageList = ImageList_Create(16, 16, ILC_MASK | ILC_COLOR32, 2, 0);

    for (auto icon : { IDI_REGISTRY, IDI_TYPELIB }) {
        auto hIcon = LoadIcon(_Module.GetResourceInstance(), MAKEINTRESOURCE(icon));
        ATLASSERT(hIcon);
        m_imageList.AddIcon(hIcon);
    }

    m_tab.SetImageList(m_imageList);

    return true;
}

void ComDetailView::UpdateLayout()
{
    RECT rect;
    GetClientRect(&rect);

    if (m_tab.IsWindow() && m_tab.IsWindowVisible()) {
        m_tab.SetWindowPos(nullptr, 0, 0, rect.right - rect.left, m_cyTabHeight, SWP_NOZORDER);
    }

    if (m_nActivePage != -1) {
        ::SetWindowPos(GetPageHWND(m_nActivePage), nullptr, 0, m_cyTabHeight, rect.right - rect.left,
                       (rect.bottom - rect.top) - m_cyTabHeight, SWP_NOZORDER);
    }
}

LRESULT ComDetailView::OnSelChanged(UINT, WPARAM, LPARAM lParam, BOOL& /*bHandled*/)
{
    RemoveAllPages();

    auto pdata = reinterpret_cast<LPOBJECTDATA>(lParam);
    if (pdata != nullptr && pdata->guid != GUID_NULL) {
        if (m_regView.Create(*this, rcDefault, nullptr, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN,
                             WS_EX_CLIENTEDGE, 0U, pdata)) {
            AddPage(m_regView, _T("Registry"), 0, pdata);
        }

        if (pdata->type == ObjectType::TYPELIB) {
            if (m_typeLibView.Create(*this, rcDefault, nullptr,
                                     WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, WS_EX_CLIENTEDGE,
                                     0U, pdata)) {
                AddPage(m_typeLibView, _T("Type Library"), 1, pdata);
            }
        }

        const auto nPageCount = GetPageCount();
        if (nPageCount > 0) {
            SetActivePage(0);
        }
    }

    UpdateLayout();

    return 0;
}
