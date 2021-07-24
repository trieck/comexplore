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
                       TVS_LINESATROOT | TVS_SHOWSELALWAYS, 0, 0U, pcs->lpCreateParams)) {
        return -1;
    }

    if (!m_idlView.Create(*this, rcDefault, nullptr,
                          WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL |
                          ES_AUTOHSCROLL | ES_AUTOVSCROLL | ES_MULTILINE 
                          | ES_NOOLEDRAGDROP | ES_READONLY)) {
        return -1;
    }

    auto bResult = CSplitterWindow::OnCreate(uMsg, wParam, lParam, bHandled);

    SetSplitterPane(SPLIT_PANE_LEFT, m_tree);
    SetSplitterPane(SPLIT_PANE_RIGHT, m_idlView);
    SetSplitterPosPct(50);

    return bResult;
}

LRESULT TypeLibView::OnTVSelChanged(LPNMHDR pnmhdr)
{
    if (pnmhdr != nullptr && pnmhdr->hwndFrom == m_tree) {
        auto item = CTreeItem(reinterpret_cast<LPNMTREEVIEW>(pnmhdr)->itemNew.hItem, &m_tree);

        CComPtr<ITypeLib> pTypeLib;
        auto hr = m_tree.GetTypeLib(&pTypeLib);
        if (FAILED(hr)) {
            return 0;
        }

        m_idlView.SendMessage(WM_SELCHANGED, reinterpret_cast<WPARAM>(pTypeLib.p), item.GetData());
    }

    return 0;
}

