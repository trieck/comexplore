#include "StdAfx.h"
#include "typelibview.h"

TypeLibView::TypeLibView(): m_bMsgHandled(0)
{
}

LRESULT TypeLibView::OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
    auto pcs = reinterpret_cast<LPCREATESTRUCT>(lParam);

    if (!m_tree.Create(*this, rcDefault, nullptr,
                       WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | TVS_HASLINES |
                       TVS_LINESATROOT | TVS_SHOWSELALWAYS, WS_EX_CLIENTEDGE, 0U, pcs->lpCreateParams)) {
        return -1;
    }

    auto bResult = CSplitterWindow::OnCreate(uMsg, wParam, lParam, bHandled);

    SetSplitterPane(SPLIT_PANE_LEFT, m_tree);
    SetSplitterPosPct(50);

    return bResult;
}
