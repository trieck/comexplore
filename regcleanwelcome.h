#pragma once
#include "regcleaninfo.h"
#include "resource.h"

class RegCleanWelcomePage :
    public CWizard97ExteriorPageImpl<RegCleanWelcomePage>,
    public RegCleanInfoRef
{
public:
    RegCleanWelcomePage(RegCleanInfo* pInfo, _U_STRINGorID title = nullptr);

    LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

    enum { IDD = IDD_REGCLEAN_WELCOME };

    // Overrides
    int OnSetActive();
    int OnWizardNext();

private:
    using BasePage = CWizard97ExteriorPageImpl<RegCleanWelcomePage>;

BEGIN_MSG_MAP(RegCleanWelcomePage)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        CHAIN_MSG_MAP(BasePage)
    END_MSG_MAP()
};
