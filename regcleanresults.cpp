#include "stdafx.h"
#include "regcleanresults.h"

#include "util.h"

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

    auto columnWidth = MulDiv(width, 2, 5); // 40%
    m_results.InsertColumn(0, _T("Class ID"), LVCFMT_LEFT, columnWidth, 0);

    columnWidth = width - columnWidth; // 60%
    m_results.InsertColumn(1, _T("Filename"), LVCFMT_LEFT, columnWidth, 1);

    const auto& info = GetInfo();

    auto item = 0;
    for (const auto& p : info.clsids) {
        m_results.AddItem(item, 0, p.first);
        m_results.AddItem(item, 1, p.second);
        item++;
    }

    if (!info.clsids.empty()) {
        m_results.SetColumnWidth(0, LVSCW_AUTOSIZE);
        m_results.SetColumnWidth(1, LVSCW_AUTOSIZE);
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
    CRegKey key;
    auto lResult = key.Open(HKEY_CLASSES_ROOT, _T("CLSID"), DELETE | KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE);
    if (lResult != ERROR_SUCCESS) {
        return 0; // no access?
    }

    const auto& info = GetInfo();

    CRegKey subKey;
    for (const auto& clsid : info.clsids) {
        lResult = RegDeleteTree(key, clsid.first);
        if (lResult != ERROR_SUCCESS) {
            WinErrorMsgBox(*this, lResult, _T("Registry Error"), MB_ICONERROR);
            continue;
        }
    }

    return 0;
}
