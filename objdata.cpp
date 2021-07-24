#include "stdafx.h"
#include "objdata.h"

/////////////////////////////////////////////////////////////////////////////
ObjectData::ObjectData() : type(ObjectType::NONE), guid(GUID_NULL),
    wMaj(0), wMin(0)
{
}

/////////////////////////////////////////////////////////////////////////////
ObjectData::ObjectData(ObjectType t, const GUID& uuid, WORD maj, WORD min)
    : type(t), guid(uuid), wMaj(maj), wMin(min)
{
}

/////////////////////////////////////////////////////////////////////////////
ObjectData::ObjectData(ObjectType t, LPCTSTR pGUID, WORD maj, WORD min)
    : type(t), guid(GUID_NULL), wMaj(maj), wMin(min)
{
    ATLASSERT(pGUID);

    CString strGUID(pGUID);
    if (pGUID[0] != _T('{') && pGUID[strGUID.GetLength() - 1] != _T('}')) {
        strGUID.Format(_T("{%s}"), pGUID);
    }

    switch (type) {
    case ObjectType::IID:
        IIDFromString(strGUID, &guid);
        break;
    case ObjectType::APPID:
    case ObjectType::CLSID:
    case ObjectType::TYPELIB:
    case ObjectType::CATID:
        CLSIDFromString(strGUID, &guid); // may work generally
        break;
    default:
        break;
    }

    if (IsEqualGUID(guid, GUID_NULL)) {
        ATLTRACE(_T("Failed to parse guid: \"%s\".\n"), static_cast<LPCTSTR>(strGUID));
    }
}
