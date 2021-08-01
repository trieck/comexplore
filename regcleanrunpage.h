#pragma once
#include "regcleaninfo.h"
#include "resource.h"

class RegCleanRunPage :
    public CWizard97InteriorPageImpl<RegCleanRunPage>,
    public RegCleanInfoRef
{
public:
    RegCleanRunPage(RegCleanInfo* pInfo, _U_STRINGorID title = nullptr);

    LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

    enum { IDD = IDD_REGCLEAN_RUN };

    int OnSetActive();
    int OnWizardBack();
    int OnWizardNext();
    BOOL OnQueryCancel();

private:
    using BasePage = CWizard97InteriorPageImpl<RegCleanRunPage>;

    BEGIN_MSG_MAP(RegCleanRun)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        MESSAGE_HANDLER(WM_PROGRESS, OnProgress)
        MESSAGE_HANDLER(WM_COMPLETE, OnComplete)
        MESSAGE_HANDLER(WM_SETRANGE, OnSetRange)
        CHAIN_MSG_MAP(BasePage)
    END_MSG_MAP()

    LRESULT OnProgress(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnComplete(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnSetRange(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

    CProgressBarCtrl m_progress;
};
