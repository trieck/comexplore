#include "stdafx.h"
#include "mainfrm.h"

BOOL CMainFrame::PreTranslateMessage(MSG* pMsg)
{
    if (m_detailView.PreTranslateMessage(pMsg) != FALSE) {
        return TRUE;
    }

    return CFrameWindowImpl<CMainFrame>::PreTranslateMessage(pMsg);
}

BOOL CMainFrame::OnIdle()
{
    UIUpdateToolBar();
    return FALSE;
}

BOOL CMainFrame::DefCreate()
{
    enum { CX_WIDTH = 1024, CY_HEIGHT = 600 };

    RECT rect = { 0, 0, CX_WIDTH, CY_HEIGHT };
    return CreateEx(nullptr, rect) != nullptr;
}

LRESULT CMainFrame::OnTVSelChanged(LPNMHDR pnmhdr)
{
    if (pnmhdr != nullptr && pnmhdr->hwndFrom == m_treeView) {
        auto item = CTreeItem(reinterpret_cast<LPNMTREEVIEW>(pnmhdr)->itemNew.hItem, &m_treeView);
        m_detailView.SendMessage(WM_SELCHANGED, 0, item.GetData());
    }

    return 0;
}

LRESULT CMainFrame::OnCreate(LPCREATESTRUCT pcs)
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
    m_splitter.SetSplitterPosPct(30);

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

LRESULT CMainFrame::OnFileExit(WORD, WORD, HWND, BOOL&)
{
    PostMessage(WM_CLOSE);
    return 0;
}

LRESULT CMainFrame::OnFileNew(WORD, WORD, HWND, BOOL&)
{
    return 0;
}

LRESULT CMainFrame::OnViewToolBar(WORD, WORD, HWND, BOOL&)
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

LRESULT CMainFrame::OnViewStatusBar(WORD, WORD, HWND, BOOL&)
{
    BOOL bVisible = !::IsWindowVisible(m_hWndStatusBar);
    ::ShowWindow(m_hWndStatusBar, bVisible ? SW_SHOWNOACTIVATE : SW_HIDE);
    UISetCheck(ID_VIEW_STATUS_BAR, bVisible);
    UpdateLayout();
    return 0;
}

LRESULT CMainFrame::OnAppAbout(WORD, WORD, HWND, BOOL&)
{
    CAboutDlg dlg;
    dlg.DoModal();
    return 0;
}
