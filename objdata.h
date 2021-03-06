#pragma once

enum class ObjectType
{
    NONE,
    APPID,
    CLSID,
    IID,
    TYPELIB,
    CATID
};

/////////////////////////////////////////////////////////////////////////////
struct ObjectData
{
    ObjectData();
    ObjectData(ObjectType t, const GUID& uuid, WORD maj = 0, WORD min = 0);
    ObjectData(ObjectType t, LPUNKNOWN pUnknown, const GUID& uuid, WORD maj = 0, WORD min = 0);
    ObjectData(ObjectType t, LPCTSTR pGUID, WORD maj = 0, WORD min = 0);
    ~ObjectData();

    // no copy or move possible
    ObjectData(const ObjectData&) = delete;
    ObjectData(ObjectData&&) = delete;
    ObjectData& operator =(const ObjectData&) = delete;
    ObjectData& operator =(ObjectData&&) = delete;

    void SafeRelease();

    ObjectType type; // type of object
    CComPtr<IUnknown> pUnknown; // object / typelib instance
    GUID guid = GUID_NULL; // UUID of object
    WORD wMaj = 0; // major version of typelib
    WORD wMin = 0; // minor version of typelib
};

using LPOBJECTDATA = ObjectData*;
