#pragma once

/////////////////////////////////////////////////////////////////////////////
struct RegCleanInfo
{
    CStringMap<CString> clsids;

};

/////////////////////////////////////////////////////////////////////////////
class RegCleanInfoRef
{
public:
    RegCleanInfoRef(RegCleanInfo* pInfo);

    RegCleanInfo& GetInfo();

private:
    RegCleanInfo* m_pInfo;
};