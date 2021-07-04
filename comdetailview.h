#pragma once
#include "regview.h"

class ComDetailView : public CTabViewImpl<ComDetailView>
{
BEGIN_MSG_MAP(ComDetailView)
        MESSAGE_HANDLER(WM_SELCHANGED, OnSelChanged)
        CHAIN_MSG_MAP(CTabViewImpl<ComDetailView>)
        ALT_MSG_MAP(1) // tab control
    END_MSG_MAP()

    DECLARE_WND_CLASS_EX(_T("ComDetailView"), 0, COLOR_APPWORKSPACE)

    bool CreateTabControl()
    {
        auto result = CTabViewImpl<ComDetailView>::CreateTabControl();
        if (!result) {
            return false;
        }

        m_imageList = ImageList_Create(16, 16, ILC_MASK | ILC_COLOR32, 1, 0);

        for (auto icon : { IDI_REGISTRY }) {
            auto hIcon = LoadIcon(_Module.GetResourceInstance(), MAKEINTRESOURCE(icon));
            ATLASSERT(hIcon);
            m_imageList.AddIcon(hIcon);
        }

        m_tab.ModifyStyle(TCS_FLATBUTTONS, 0);
        m_tab.SetImageList(m_imageList);

        return true;
    }

    void UpdateLayout()
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

    LRESULT OnSelChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
    {
        auto pdata = reinterpret_cast<LPOBJECTDATA>(lParam);
        if (pdata != nullptr && pdata->guid != GUID_NULL) {
            if (!m_regView.IsWindow()) {
                m_regView.Create(*this, rcDefault, nullptr,
                                 WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | TVS_HASLINES |
                                 TVS_SHOWSELALWAYS, WS_EX_CLIENTEDGE);

                if (m_regView.IsWindow()) {
                    AddPage(m_regView, _T("Registry"), 0);
                }
            }

            auto hWndActive = GetPageHWND(m_nActivePage);
            ::SendMessage(hWndActive, WM_SELCHANGED, 0, lParam);

            ::ShowWindow(hWndActive, SW_SHOW);
            ::UpdateWindow(hWndActive);

            m_tab.ShowWindow(SW_SHOW);
            m_tab.UpdateWindow();

        } else if (m_regView.IsWindow()) {
            m_regView.ShowWindow(SW_HIDE);
            m_regView.UpdateWindow();

            m_tab.ShowWindow(SW_HIDE);
            m_tab.UpdateWindow();
        }

        UpdateLayout();

        return 0;
    }

private:
    RegistryView m_regView{};
    CImageList m_imageList;
};
