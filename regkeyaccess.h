#pragma once

/////////////////////////////////////////////////////////////////////////////
class RegKeyAccess
{
public:
    RegKeyAccess() = default;
    ~RegKeyAccess();

    BOOL Initialize();
    BOOL HasAccess(CRegKey& key, DWORD dwAccess);

private:
    void Reset();
    CHandle m_token;
};
