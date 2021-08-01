#include "stdafx.h"
#include "regcleanwelcome.h"

RegCleanWelcomePage::RegCleanWelcomePage(RegCleanInfo* pInfo, _U_STRINGorID title)
    : BasePage(title), RegCleanInfoRef(pInfo)
{
}

LRESULT RegCleanWelcomePage::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
    CFontHandle titleFont(GetExteriorPageTitleFont());
    ATLASSERT(titleFont);

    auto title = GetDlgItem(IDC_REGCLEAN_TITLE);
    ATLASSERT(title);

    title.SetFont(titleFont);

    return 0;
}

int RegCleanWelcomePage::OnSetActive()
{
    SetWizardButtons(PSWIZB_NEXT);

    return 0;
}

int RegCleanWelcomePage::OnWizardNext()
{
    return 0;
}
