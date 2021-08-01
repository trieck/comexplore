#pragma once
#include "regcleaninfo.h"
#include "resource.h"

class RegCleanResultsPage :
    public CWizard97InteriorPageImpl<RegCleanResultsPage>,
    public RegCleanInfoRef
{
public:
    RegCleanResultsPage(RegCleanInfo* pInfo, _U_STRINGorID title = nullptr);

    LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

    enum { IDD = IDD_REGCLEAN_RESULTS };

    int OnSetActive();
    int OnWizardFinish();

private:
    using BasePage = CWizard97InteriorPageImpl<RegCleanResultsPage>;

    BEGIN_MSG_MAP(RegCleanResultsPage)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        CHAIN_MSG_MAP(BasePage)
    END_MSG_MAP()

    CListViewCtrl m_results;
};
