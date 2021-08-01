#include "stdafx.h"
#include "comexplorer.h"

CAppModule _Module;

int WINAPI _tWinMain(HINSTANCE hInstance, HINSTANCE /*hPrevInstance*/, LPTSTR lpCmdLine, int nShowCmd)
{
    COMExplorer explorer;
    if (!explorer.Init()) {
        ATLTRACE(_T("Cannot initialize application.\n"));
        return -1;
    }

    return explorer.Run(hInstance, lpCmdLine, nShowCmd);
}
