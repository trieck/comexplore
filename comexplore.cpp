#include "stdafx.h"
#include "MainFrm.h"

CAppModule _Module;

int Run(LPTSTR /*lpstrCmdLine*/  = nullptr, int nCmdShow = SW_SHOWDEFAULT)
{
    CMessageLoop theLoop;
    _Module.AddMessageLoop(&theLoop);

    auto hInstRich = LoadLibrary(CRichEditCtrl::GetLibraryName());
    if (hInstRich == nullptr) {
        ATLTRACE(_T("Unable to load rich edit library!\n"));
        return -1;
    }

    CMainFrame wndMain;

    if (!wndMain.DefCreate()) {
        ATLTRACE(_T("Main window creation failed!\n"));
        return -1;
    }

    wndMain.ShowWindow(nCmdShow);

    auto nRet = theLoop.Run();

    _Module.RemoveMessageLoop();

    FreeLibrary(hInstRich);

    return nRet;
}

int WINAPI _tWinMain(HINSTANCE hInstance, HINSTANCE /*hPrevInstance*/, LPTSTR lpstrCmdLine, int nCmdShow)
{
    auto hRes = CoInitialize(nullptr);
    ATLASSERT(SUCCEEDED(hRes));

    AtlInitCommonControls(ICC_COOL_CLASSES | ICC_BAR_CLASSES | ICC_USEREX_CLASSES);

    hRes = _Module.Init(nullptr, hInstance);
    ATLASSERT(SUCCEEDED(hRes));

    auto nRet = Run(lpstrCmdLine, nCmdShow);

    _Module.Term();

    CoUninitialize();

    return nRet;
}
