#include "stdafx.h"
#include "regcleanresults.h"

RegCleanResultsPage::RegCleanResultsPage(RegCleanInfo* pInfo, _U_STRINGorID title)
    : BasePage(title), RegCleanInfoRef(pInfo)
{
    SetHeaderTitle(_T("Results"));
    SetHeaderSubTitle(_T(""));
}

LRESULT RegCleanResultsPage::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
    m_results = ::GetDlgItem(m_hWnd, IDC_LIST_RESULTS);
    ATLASSERT(m_results);

    CRect rcList;
	m_results.GetClientRect(&rcList);
    int width = rcList.Width();
		
	auto columnWidth = MulDiv(width, 2, 5);  // 40%
	m_results.InsertColumn(0, _T("Class ID"), LVCFMT_LEFT, columnWidth, 0);

    columnWidth = width - columnWidth;  // 60%
	m_results.InsertColumn(1, _T("Filename"), LVCFMT_LEFT, columnWidth, 1);

    const auto& info = GetInfo();

    auto item = 0;
    for (const auto& p: info.clsids) {
        m_results.AddItem(item, 0, p.first);
        m_results.AddItem(item, 1, p.second);
        item++;
    }
    
    return 0;
}

int RegCleanResultsPage::OnSetActive()
{
    SetWizardButtons(PSWIZB_FINISH);


    return 0;
}

int RegCleanResultsPage::OnWizardFinish()
{
    return 0;
}

