#pragma once

#include "comtreeview.h"
#include "aboutdlg.h"
#include "comdetailview.h"

class ObjectPane : public CPaneContainerImpl<ObjectPane>
{
};

class CMainFrame : public CFrameWindowImpl<CMainFrame>,
                   public CUpdateUI<CMainFrame>,
                   public CMessageFilter,
                   public CIdleHandler
{
public:
    DECLARE_FRAME_WND_CLASS(NULL, IDR_MAINFRAME)

    enum { CX_WIDTH = 800, CY_HEIGHT = 600 };

    BOOL PreTranslateMessage(MSG* pMsg) override
    {
        if (m_detailView.PreTranslateMessage(pMsg) != FALSE) {
            return TRUE;
        }

        return CFrameWindowImpl<CMainFrame>::PreTranslateMessage(pMsg);
    }

    BOOL OnIdle() override
    {
        UIUpdateToolBar();
        return FALSE;
    }

    BOOL DefCreate()
    {
        RECT rect = { 0, 0, CX_WIDTH, CY_HEIGHT };
        return CreateEx(nullptr, rect) != nullptr;
    }

    BEGIN_UPDATE_UI_MAP(CMainFrame)
            UPDATE_ELEMENT(ID_VIEW_TOOLBAR, UPDUI_MENUPOPUP)
            UPDATE_ELEMENT(ID_VIEW_STATUS_BAR, UPDUI_MENUPOPUP)
    END_UPDATE_UI_MAP()

BEGIN_MSG_MAP(CMainFrame)
        MSG_WM_CREATE(OnCreate)
        COMMAND_ID_HANDLER(ID_APP_EXIT, OnFileExit)
        COMMAND_ID_HANDLER(ID_FILE_NEW, OnFileNew)
        COMMAND_ID_HANDLER(ID_VIEW_TOOLBAR, OnViewToolBar)
        COMMAND_ID_HANDLER(ID_VIEW_STATUS_BAR, OnViewStatusBar)
        COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
        REFLECT_NOTIFY_CODE(TVN_ITEMEXPANDING)
        REFLECT_NOTIFY_CODE(TVN_DELETEITEM)
        NOTIFY_CODE_HANDLER_EX(TVN_SELCHANGED, OnTVSelChanged)
        CHAIN_MSG_MAP(CUpdateUI<CMainFrame>)
        CHAIN_MSG_MAP(CFrameWindowImpl<CMainFrame>)
    END_MSG_MAP()

    LRESULT OnTVSelChanged(LPNMHDR pnmhdr)
    {
        if (pnmhdr != nullptr && pnmhdr->hwndFrom == m_treeView) {
            auto item = CTreeItem(reinterpret_cast<LPNMTREEVIEW>(pnmhdr)->itemNew.hItem, &m_treeView);
            m_detailView.SendMessage(WM_SELCHANGED, 0, item.GetData());
        }

        return 0;
    }

    LRESULT OnCreate(LPCREATESTRUCT pcs)
    {
        auto hWndCmdBar = m_cmdBar.Create(m_hWnd, rcDefault, nullptr, ATL_SIMPLE_CMDBAR_PANE_STYLE);
        m_cmdBar.AttachMenu(GetMenu());
        m_cmdBar.LoadImages(IDR_MAINFRAME);
        SetMenu(nullptr); // remove old menu

        auto hWndToolBar = CreateSimpleToolBarCtrl(m_hWnd, IDR_MAINFRAME, FALSE, ATL_SIMPLE_TOOLBAR_PANE_STYLE);

        CreateSimpleReBar(ATL_SIMPLE_REBAR_NOBORDER_STYLE);
        AddSimpleReBarBand(hWndCmdBar);
        AddSimpleReBarBand(hWndToolBar, nullptr, TRUE);

        CreateSimpleStatusBar();

        m_hWndClient = m_splitter.Create(m_hWnd, rcDefault, nullptr,
                                         WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN);
        if (!m_hWndClient) {
            return -1;
        }

        if (!m_objectPane.Create(m_splitter, _T("Objects"),
                                 WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN)) {
            return -1;
        }

        m_objectPane.SetPaneContainerExtendedStyle(PANECNT_NOCLOSEBUTTON);

        if (!m_treeView.Create(m_objectPane, rcDefault, nullptr,
                               WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | TVS_HASLINES |
                               TVS_LINESATROOT | TVS_SHOWSELALWAYS, WS_EX_CLIENTEDGE)) {
            return -1;
        }

        m_objectPane.SetClient(m_treeView);

        if (!m_detailView.Create(m_splitter, rcDefault, nullptr,
                                 WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, WS_EX_CLIENTEDGE)) {
            return -1;
        }

        m_detailView.SetTitleBarWindow(m_hWnd);

        m_splitter.SetSplitterPane(0, m_objectPane);
        m_splitter.SetSplitterPane(1, m_detailView);
        m_splitter.SetSplitterPosPct(50);

        UIAddToolBar(hWndToolBar);
        UISetCheck(ID_VIEW_TOOLBAR, 1);
        UISetCheck(ID_VIEW_STATUS_BAR, 1);

        UpdateLayout();
        CenterWindow();

        // register object for message filtering and idle updates
        auto pLoop = _Module.GetMessageLoop();
        ATLASSERT(pLoop != NULL);
        pLoop->AddMessageFilter(this);
        pLoop->AddIdleHandler(this);

        return 1;
    }

    LRESULT OnFileExit(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
    {
        PostMessage(WM_CLOSE);
        return 0;
    }

    LRESULT OnFileNew(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
    {
        return 0;
    }

    LRESULT OnViewToolBar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
    {
        static auto bVisible = TRUE; // initially visible
        bVisible = !bVisible;
        CReBarCtrl rebar = m_hWndToolBar;
        auto nBandIndex = rebar.IdToIndex(ATL_IDW_BAND_FIRST + 1); // toolbar is 2nd added band
        rebar.ShowBand(nBandIndex, bVisible);
        UISetCheck(ID_VIEW_TOOLBAR, bVisible);
        UpdateLayout();
        return 0;
    }

    LRESULT OnViewStatusBar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
    {
        BOOL bVisible = !::IsWindowVisible(m_hWndStatusBar);
        ::ShowWindow(m_hWndStatusBar, bVisible ? SW_SHOWNOACTIVATE : SW_HIDE);
        UISetCheck(ID_VIEW_STATUS_BAR, bVisible);
        UpdateLayout();
        return 0;
    }

    LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
    {
        CAboutDlg dlg;
        dlg.DoModal();
        return 0;
    }

private:
    ObjectPane m_objectPane;
    ComTreeView m_treeView;
    ComDetailView m_detailView;
    CCommandBarCtrl m_cmdBar;
    CSplitterWindow m_splitter;
};
