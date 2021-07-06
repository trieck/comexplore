#include "stdafx.h"
#include "aboutdlg.h"

LRESULT CAboutDlg::OnInitDialog(UINT, WPARAM, LPARAM, BOOL&)
{
    CenterWindow(GetParent());
    return TRUE;
}

LRESULT CAboutDlg::OnCloseCmd(WORD, WORD wID, HWND, BOOL&)
{
    EndDialog(wID);
    return 0;
}
