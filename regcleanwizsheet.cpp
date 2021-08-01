#include "stdafx.h"
#include "regcleanwizsheet.h"
#include "resource.h"

RegCleanWizSheet::RegCleanWizSheet(RegCleanInfo* pInfo) :
    BaseSheet(_T("COM Registry Cleaner"), IDB_REGCLEAN_HEADER, IDB_REGCLEAN_LARGE, 0, nullptr),
    m_welcomePage(pInfo), m_runPage(pInfo), m_resultsPage(pInfo)
{
    AddPage(m_welcomePage);
    AddPage(m_runPage);
    AddPage(m_resultsPage);
}
