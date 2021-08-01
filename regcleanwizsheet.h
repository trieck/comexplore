#pragma once
#include "regcleaninfo.h"
#include "regcleanresults.h"
#include "regcleanwelcome.h"
#include "regcleanrunpage.h"

class RegCleanWizSheet : public CWizard97SheetImpl<RegCleanWizSheet>
{
public:
    RegCleanWizSheet(RegCleanInfo* pInfo);

private:
    using BaseSheet = CWizard97SheetImpl<RegCleanWizSheet>;
    RegCleanWelcomePage m_welcomePage;
    RegCleanRunPage m_runPage;
    RegCleanResultsPage m_resultsPage;
};
