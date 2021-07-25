#include "stdafx.h"
#include "mainfrm.h"

BOOL CMainFrame::PreTranslateMessage(MSG* pMsg)
{
    if (m_detailView.PreTranslateMessage(pMsg) != FALSE) {
        return TRUE;
    }

    return CRibbonFrameWindowImpl<CMainFrame>::PreTranslateMessage(pMsg);
}

BOOL CMainFrame::OnIdle()
{
    UIUpdateToolBar();
    return FALSE;
}

BOOL CMainFrame::DefCreate()
{
    RECT rect = { 0, 0, 1024, 600 };
    auto hWnd = CreateEx(nullptr, rect);
    if (hWnd == nullptr) {
        ATLTRACE(_T("Unable to create main frame.\n"));
        return FALSE;
    }

    return TRUE;
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

    bool bRibbonUI = RunTimeHelper::IsRibbonUIAvailable();
    if (bRibbonUI) {
        // UI Setup and adjustments
        UIAddMenu(m_cmdBar.GetMenu(), true);

        UIRemoveUpdateElement(ID_FILE_MRU_FIRST);
        UIPersistElement(ID_GROUP_VIEW);
    }

    CreateSimpleReBar(ATL_SIMPLE_REBAR_NOBORDER_STYLE);
    AddSimpleReBarBand(hWndCmdBar);
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

    UISetCheck(ID_VIEW_STATUS_BAR, 1);
    ShowRibbonUI(TRUE);

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
