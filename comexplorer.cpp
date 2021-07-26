#include "stdafx.h"
#include "comexplorer.h"
#include "MainFrm.h"

COMExplorer::~COMExplorer()
{
    if (m_hInstRich) {
        FreeLibrary(m_hInstRich);
    }

    CoUninitialize();
}

BOOL COMExplorer::Init()
{
    auto hr = CoInitialize(nullptr);
    if (FAILED(hr)) {
        ATLTRACE(_T("Cannot initialize COM libraries.\n"));
        return FALSE;
    }

    INITCOMMONCONTROLSEX iex = {
        sizeof(INITCOMMONCONTROLSEX),
        ICC_WIN95_CLASSES | ICC_COOL_CLASSES | ICC_STANDARD_CLASSES | ICC_TAB_CLASSES | ICC_USEREX_CLASSES |
        ICC_LINK_CLASS | ICC_BAR_CLASSES
    };

    if (!InitCommonControlsEx(&iex)) {
        ATLTRACE(_T("Cannot initialize common controls.\n"));
        return FALSE;
    }

    DWORD param = FE_FONTSMOOTHINGCLEARTYPE;
    SystemParametersInfo(SPI_SETFONTSMOOTHING, 1, nullptr, 0);
    SystemParametersInfo(SPI_SETFONTSMOOTHINGTYPE, 1, &param, 0);

    m_hInstRich = LoadLibrary(CRichEditCtrl::GetLibraryName());
    if (m_hInstRich == nullptr) {
        ATLTRACE(_T("Unable to load rich edit library.\n"));
        return FALSE;
    }

    return TRUE;
}

int COMExplorer::Run(HINSTANCE hInstance, LPTSTR /*lpCmdLine*/, int nCmdShow)
{
    auto hr = _Module.Init(nullptr, hInstance);
    if (FAILED(hr)) {
        ATLTRACE(_T("Unable to initialize module.\n"));
        return -1;
    }

    CMessageLoop theLoop;
    _Module.AddMessageLoop(&theLoop);

    CMainFrame wndMain;
    if (!wndMain.DefCreate()) {
        ATLTRACE(_T("Main window creation failed\n"));
        return -1;
    }

    wndMain.ShowWindow(nCmdShow);

    auto result = theLoop.Run();

    _Module.RemoveMessageLoop();

    return result;
}
