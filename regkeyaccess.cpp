#include "stdafx.h"
#include "regkeyaccess.h"
#include "localbuff.h"

RegKeyAccess::~RegKeyAccess()
{
    Reset();
}

BOOL RegKeyAccess::Initialize()
{
    Reset();

    auto result = ImpersonateSelf(SecurityIdentification);
    if (!result) {
        return FALSE;
    }

    result = OpenThreadToken(GetCurrentThread(),
                             TOKEN_QUERY, TRUE, &m_token.m_h);

    RevertToSelf();

    return result;
}

BOOL RegKeyAccess::HasAccess(CRegKey& key, DWORD dwAccess)
{
    SECURITY_INFORMATION si = OWNER_SECURITY_INFORMATION |
        GROUP_SECURITY_INFORMATION | DACL_SECURITY_INFORMATION;

    DWORD dwSize;
    auto lResult = key.GetKeySecurity(si, nullptr, &dwSize);
    if (lResult != ERROR_SUCCESS) {
        return FALSE;
    }

    LocalBuffer<SECURITY_DESCRIPTOR> psd;
    if (!psd.Allocate(dwSize)) {
        return FALSE;
    }

    lResult = key.GetKeySecurity(si, psd, &dwSize);
    if (lResult != ERROR_SUCCESS) {
        return FALSE;
    }

    GENERIC_MAPPING mapping;
    mapping.GenericRead = KEY_READ;
    mapping.GenericWrite = KEY_WRITE;
    mapping.GenericExecute = KEY_EXECUTE;
    mapping.GenericAll = KEY_ALL_ACCESS;

    PRIVILEGE_SET privilegeSet;
    DWORD dwPrivSetSize = sizeof(PRIVILEGE_SET);
    DWORD dwAllowedAccess = 0;
    BOOL fAccessGranted = FALSE;

    MapGenericMask(&dwAccess, &mapping);

    AccessCheck(
        psd,
        m_token,
        dwAccess,
        &mapping,
        &privilegeSet,
        &dwPrivSetSize,
        &dwAllowedAccess,
        &fAccessGranted
    );

    if (!fAccessGranted) {
        return FALSE;
    }

    return TRUE;
}

void RegKeyAccess::Reset()
{
    m_token.Close();
    RevertToSelf();
}
