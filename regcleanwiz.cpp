#include "stdafx.h"
#include "regcleanwiz.h"
#include "regcleanwizsheet.h"

void RegCleanWiz::Execute()
{
	RegCleanWizSheet sheet(&m_info);
	sheet.DoModal();
}
