#include "stdafx.h"
#include "regcleaninfo.h"

RegCleanInfoRef::RegCleanInfoRef(RegCleanInfo* pInfo)
    : m_pInfo(pInfo)
{
}

RegCleanInfo& RegCleanInfoRef::GetInfo()
{
    ATLASSERT(m_pInfo);

    return *m_pInfo;
}
