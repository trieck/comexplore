#include "stdafx.h"
#include "mainfrm.h"
#include "comerror.h"
#include "regcleanwiz.h"

BOOL CMainFrame::PreTranslateMessage(MSG* pMsg)
{
    if (m_detailView.PreTranslateMessage(pMsg) != FALSE) {
        return TRUE;
    }

    return CRibbonFrameWindowImpl<CMainFrame>::PreTranslateMessage(pMsg);
}

BOOL CMainFrame::OnIdle()
{
    UIEnable(ID_RELEASE_OBJECT, IsSelectedInstance());
    UIEnable(ID_COPY_GUID, IsGUIDSelected());
    UIEnable(ID_REGEDIT_HERE, IsGUIDSelected());

    UIUpdateMenuBar();

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
        CWaitCursor cursor;
        auto item = CTreeItem(reinterpret_cast<LPNMTREEVIEW>(pnmhdr)->itemNew.hItem, &m_treeView);
        m_detailView.SendMessage(WM_SELCHANGED, 0, item.GetData());
        m_treeView.SetFocus();
    }

    return 0;
}

LRESULT CMainFrame::OnRClick(LPNMHDR pnmh)
{
    if (pnmh->hwndFrom != m_treeView) {
        return 0;
    }

    auto item = m_treeView.GetSelectedItem();
    if (item.IsNull()) {
        return 0;
    }

    CMenu menu;
    menu.LoadMenu(IDR_CONTEXT_MENU);

    CMenuHandle popup = menu.GetSubMenu(0);

    CPoint pt;
    GetCursorPos(&pt);

    popup.TrackPopupMenu(TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_VERTICAL, pt.x, pt.y, *this);

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
        UIAddMenu(m_cmdBar.GetMenu(), true);
        UIRemoveUpdateElement(ID_FILE_MRU_FIRST);
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

LRESULT CMainFrame::OnFileExit()
{
    PostMessage(WM_CLOSE);
    return 0;
}

LRESULT CMainFrame::OnFileOpen()
{
    CFileDialog dlg(TRUE, nullptr, _T(""), OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, _T("All Files (*.*)\0*.*\0"));
    if (dlg.DoModal() != IDOK) {
        return 0;
    }

    CComPtr<IBindCtx> pBC;
    auto hr = CreateBindCtx(0, &pBC);
    if (FAILED(hr)) {
        return 0;
    }

    CLSID clsid = GUID_NULL;
    GetClassFile(dlg.m_szFileName, &clsid);

    CComPtr<IMoniker> pMoniker;
    hr = CreateFileMoniker(dlg.m_szFileName, &pMoniker);
    if (FAILED(hr)) {
        return 0;
    }

    CComPtr<IUnknown> pUnk;
    hr = pMoniker->BindToObject(pBC, nullptr, __uuidof(pUnk), IID_PPV_ARGS_Helper(&pUnk));
    if (FAILED(hr)) {
        CoMessageBox(*this, hr, pMoniker, IDR_MAINFRAME);
        return 0;
    }

    AddFileMoniker(dlg.m_szFileName, pUnk, clsid);

    return 0;
}

LRESULT CMainFrame::OnFileTypeLib()
{
    CFileDialog dlg(TRUE, nullptr, _T(""), OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
                    _T("TypeLib Files (*.tlb;*.olb;*.dll;*.ocx;*.exe)\0*.tlb;*.olb;*.dll;*.ocx;*.exe\0")
                    _T("All Files (*.*)\0*.*\0"));
    if (dlg.DoModal() != IDOK) {
        return 0;
    }

    CComPtr<ITypeLib> pTypeLib;
    auto hr = LoadTypeLib(dlg.m_szFileName, &pTypeLib);
    if (FAILED(hr)) {
        CoMessageBox(*this, hr, nullptr, IDR_MAINFRAME);
        return 0;
    }

    AddFileTypeLib(dlg.m_szFileName, pTypeLib);
    
    return 0;
}

LRESULT CMainFrame::OnViewStatusBar()
{
    BOOL bVisible = !::IsWindowVisible(m_hWndStatusBar);
    ::ShowWindow(m_hWndStatusBar, bVisible ? SW_SHOWNOACTIVATE : SW_HIDE);
    UISetCheck(ID_VIEW_STATUS_BAR, bVisible);
    UpdateLayout();
    return 0;
}

LRESULT CMainFrame::OnAppAbout()
{
    CAboutDlg dlg;
    dlg.DoModal();
    return 0;
}

LRESULT CMainFrame::OnReleaseObject()
{
    m_treeView.ReleaseSelectedObject();
    return 0;
}

LRESULT CMainFrame::OnCopyGUID()
{
    m_treeView.CopyGUIDToClipboard();
    return 0;
}

LRESULT CMainFrame::OnRegEditHere()
{
    auto item = m_treeView.GetSelectedItem();
    if (item.IsNull()) {
        return 0;
    }

    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
    if (data == nullptr || data->guid == GUID_NULL) {
        return 0;
    }

    CString strGUID;
    StringFromGUID2(data->guid, strGUID.GetBuffer(40), 40);
    strGUID.ReleaseBuffer();

    CRegKey key;
    auto lResult = key.Open(
        HKEY_CURRENT_USER, _T("Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\Regedit"),
        KEY_SET_VALUE);
    if (lResult != ERROR_SUCCESS) {
        return 0; // no access
    }

    CString strValue;
    switch (data->type) {
    case ObjectType::APPID:
        strValue.Format(_T("HKCR\\AppID\\%s"), strGUID);
        break;
    case ObjectType::CATID:
        strValue.Format(_T("HKCR\\Component Categories\\%s"), strGUID);
        break;
    case ObjectType::CLSID:
        strValue.Format(_T("HKCR\\CLSID\\%s"), strGUID);
        break;
    case ObjectType::IID:
        strValue.Format(_T("HKCR\\Interface\\%s"), strGUID);
        break;
    case ObjectType::TYPELIB:
        strValue.Format(_T("HKCR\\TypeLib\\%s"), strGUID);
        break;
    default:
        break;
    }

    if (!strValue.IsEmpty()) {
        lResult = key.SetStringValue(_T("LastKey"), strValue);
        if (lResult != ERROR_SUCCESS) {
            return 0;
        }
    }

    auto hWnd = FindWindow(_T("RegEdit_RegEdit"), nullptr);
    if (hWnd != nullptr) {
        DWORD pid = 0;
        GetWindowThreadProcessId(hWnd, &pid);
        auto hProcess = OpenProcess(PROCESS_TERMINATE, false, pid);
        TerminateProcess(hProcess, 1);
        CloseHandle(hProcess);
    }

    ShellExecute(*this, _T("open"), _T("regedit.exe"), nullptr, nullptr, SW_SHOWNORMAL);

    return 0;
}

LRESULT CMainFrame::OnRegClean()
{
    RegCleanWiz wizard;
    wizard.Execute();

    return 0;
}

BOOL CMainFrame::IsSelectedInstance() const
{
    return m_treeView.IsSelectedInstance();
}

BOOL CMainFrame::IsGUIDSelected() const
{
    return m_treeView.IsGUIDSelected();
}

void CMainFrame::AddFileMoniker(LPCTSTR pFilename, LPUNKNOWN pUnk, REFCLSID clsid)
{
    ATLASSERT(pFilename);
    ATLASSERT(pUnk);
    ATLASSERT(m_treeView);

    m_treeView.AddFileMoniker(pFilename, pUnk, clsid);
}

void CMainFrame::AddFileTypeLib(LPCTSTR pFilename, LPTYPELIB pTypeLib)
{
    ATLASSERT(pFilename);
    ATLASSERT(pTypeLib);
    ATLASSERT(m_treeView);

    m_treeView.AddFileTypeLib(pFilename, pTypeLib);
}
