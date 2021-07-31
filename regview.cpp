#include "StdAfx.h"
#include "regview.h"
#include "resource.h"

RegistryView::RegistryView(): m_bMsgHandled(0)
{
}

LRESULT RegistryView::OnCreate(LPCREATESTRUCT pcs)
{
    auto bResult = DefWindowProc();

    m_ImageList = ImageList_Create(16, 16, ILC_MASK | ILC_COLOR32, 1, 0);

    for (auto icon : { IDI_REGNODE }) {
        auto hIcon = LoadIcon(_Module.GetResourceInstance(), MAKEINTRESOURCE(icon));
        ATLASSERT(hIcon);
        m_ImageList.AddIcon(hIcon);
    }

    SetImageList(m_ImageList, TVSIL_NORMAL);
    SetWindowLong(GWL_STYLE, GetWindowLong(GWL_STYLE)
                  | TVS_HASBUTTONS | TVS_HASLINES | TVS_FULLROWSELECT | TVS_INFOTIP
                  | TVS_LINESATROOT | TVS_SHOWSELALWAYS);

    if (!BuildView(static_cast<LPOBJECTDATA>(pcs->lpCreateParams))) {
        return -1;
    }

    SetMsgHandled(FALSE);

    return bResult;
}

BOOL RegistryView::BuildView(LPOBJECTDATA pdata)
{
    ATLASSERT(pdata && pdata->guid != GUID_NULL);
    switch (pdata->type) {
    case ObjectType::APPID:
        return BuildAppID(pdata);
    case ObjectType::CLSID:
        return BuildCLSID(pdata);
    case ObjectType::IID:
        return BuildIID(pdata);
    case ObjectType::TYPELIB:
        return BuildTypeLib(pdata);
    case ObjectType::CATID:
        return BuildCatID(pdata);
    default:
        break;
    }

    return FALSE;
}

CTreeItem RegistryView::InsertValue(HTREEITEM hParentItem, LPCTSTR keyName, LPCTSTR value, LPCTSTR data)
{
    CString strItem;

    if (value[0] == _T('\0')) {
        strItem.Format(_T("%s[<default>] = %s"), keyName, data);
    } else {
        strItem.Format(_T("%s[%s] = %s"), keyName, value, data);
    }

    return InsertItem(strItem, hParentItem, TVI_LAST);
}

CTreeItem RegistryView::InsertValue(HTREEITEM hParentItem, LPCTSTR keyName, LPCTSTR value, DWORD dwData)
{
    CString strItem;

    if (value[0] == _T('\0')) {
        strItem.Format(_T("%s[<default>] = %ld"), keyName, dwData);
    } else {
        strItem.Format(_T("%s[%s] = %ld"), keyName, value, dwData);
    }

    return InsertItem(strItem, hParentItem, TVI_LAST);
}

CTreeItem RegistryView::InsertValues(HKEY hKey, HTREEITEM hParentItem, LPCTSTR keyName)
{
    DWORD index = 0, length, type;
    TCHAR val[REG_BUFFER_SIZE];
    BYTE data[4096]{};
    DWORD dwSize = sizeof(data);
    CTreeItem item;

    for (;;) {
        length = REG_BUFFER_SIZE;
        dwSize = sizeof(data);

        auto lResult = RegEnumValue(hKey, index++, val, &length, nullptr, &type, data, &dwSize);
        if (lResult != ERROR_SUCCESS) {
            break;
        }

        if (type == REG_SZ || type == REG_MULTI_SZ || type == REG_EXPAND_SZ) {
            auto lpszData = reinterpret_cast<LPTSTR>(data);
            lpszData[dwSize] = _T('\0');
            item = InsertValue(hParentItem, keyName, val, lpszData);
        } else if (type == REG_DWORD) {
            DWORD dwData = (data[0] << 24) | (data[1] << 16) | (data[2] << 8) || data[3];
            item = InsertValue(hParentItem, keyName, val, dwData);
        }
    }

    if (!item.m_hTreeItem) {
        item = InsertItem(keyName, hParentItem, TVI_LAST);
    }

    return item;
}

void RegistryView::InsertSubkeys(CRegKey& key, HTREEITEM hParentItem)
{
    DWORD index = 0, length;
    TCHAR strSubKey[REG_BUFFER_SIZE];

    CRegKey subKey;

    for (;;) {
        length = REG_BUFFER_SIZE;
        auto lResult = key.EnumKey(index++, strSubKey, &length);
        if (lResult != ERROR_SUCCESS) {
            break;
        }

        lResult = subKey.Open(key.m_hKey, strSubKey, KEY_ENUMERATE_SUB_KEYS | KEY_READ);
        if (lResult != ERROR_SUCCESS) {
            break;
        }

        auto item = InsertValues(subKey, hParentItem, strSubKey);

        InsertSubkeys(subKey, item.m_hTreeItem);

        item.Expand();
    }
}

BOOL RegistryView::BuildCLSID(LPOBJECTDATA pdata)
{
    ATLASSERT(pdata && pdata->type == ObjectType::CLSID && pdata->guid != GUID_NULL);

    CString strGUID;
    StringFromGUID2(pdata->guid, strGUID.GetBuffer(40), 40);
    strGUID.ReleaseBuffer();

    if (!BuildCLSID(strGUID)) {
        return FALSE;
    }

    CString strPath;
    strPath.Format(_T("SOFTWARE\\Classes\\CLSID\\%s\\TypeLib"), static_cast<LPCTSTR>(strGUID));

    CRegKey key;
    auto lResult = key.Open(HKEY_LOCAL_MACHINE, strPath, KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE);
    if (lResult == ERROR_SUCCESS) {
        TCHAR typelib[REG_BUFFER_SIZE];
        DWORD length = REG_BUFFER_SIZE;
        lResult = key.QueryStringValue(nullptr, typelib, &length);
        if (lResult == ERROR_SUCCESS) {
            BuildTypeLib(typelib);
        }
    }

    strPath.Format(_T("SOFTWARE\\Classes\\CLSID\\%s\\ProgID"), static_cast<LPCTSTR>(strGUID));
    lResult = key.Open(HKEY_LOCAL_MACHINE, strPath, KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE);
    if (lResult != ERROR_SUCCESS) {
        return TRUE;
    }

    TCHAR progID[REG_BUFFER_SIZE];
    DWORD length = REG_BUFFER_SIZE;
    lResult = key.QueryStringValue(nullptr, progID, &length);
    if (lResult != ERROR_SUCCESS) {
        return TRUE;
    }

    BuildProgID(progID);

    return TRUE;
}

BOOL RegistryView::BuildCLSID(LPCTSTR pCLSID)
{
    CString strPath;
    strPath.Format(_T("SOFTWARE\\Classes\\CLSID\\%s"), pCLSID);

    CRegKey key;
    auto lResult = key.Open(HKEY_LOCAL_MACHINE, strPath, KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE);
    if (lResult != ERROR_SUCCESS) {
        return FALSE;
    }

    auto root = InsertItem(_T("CLSID="), TVI_ROOT, TVI_LAST);
    auto guid = InsertValues(key, root.m_hTreeItem, pCLSID);

    InsertSubkeys(key, guid.m_hTreeItem);

    root.Expand();
    guid.Expand();

    return TRUE;
}

BOOL RegistryView::BuildTypeLib(LPOBJECTDATA pdata)
{
    ATLASSERT(pdata && pdata->type == ObjectType::TYPELIB && pdata->guid != GUID_NULL);

    CString strGUID;
    StringFromGUID2(pdata->guid, strGUID.GetBuffer(40), 40);
    strGUID.ReleaseBuffer();

    return BuildTypeLib(strGUID);
}

BOOL RegistryView::BuildTypeLib(LPCTSTR pTypeLib)
{
    CString strPath;
    strPath.Format(_T("SOFTWARE\\Classes\\TypeLib\\%s"), pTypeLib);

    CRegKey key;
    auto lResult = key.Open(HKEY_LOCAL_MACHINE, strPath, KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE);
    if (lResult != ERROR_SUCCESS) {
        return FALSE;
    }

    auto root = InsertItem(_T("TypeLib="), TVI_ROOT, TVI_LAST);
    auto guid = InsertValues(key, root.m_hTreeItem, pTypeLib);

    InsertSubkeys(key, guid.m_hTreeItem);

    root.Expand();
    guid.Expand();

    return TRUE;
}

BOOL RegistryView::BuildProgID(LPCTSTR pProgID)
{
    CString strPath;
    strPath.Format(_T("SOFTWARE\\Classes\\%s"), pProgID);

    CRegKey key;
    auto lResult = key.Open(HKEY_LOCAL_MACHINE, strPath, KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE);
    if (lResult != ERROR_SUCCESS) {
        return FALSE;
    }

    auto root = InsertItem(_T("ProgID="), TVI_ROOT, TVI_LAST);
    auto guid = InsertValues(key, root.m_hTreeItem, pProgID);

    InsertSubkeys(key, guid.m_hTreeItem);

    root.Expand();
    guid.Expand();

    return TRUE;
}

BOOL RegistryView::BuildAppID(LPOBJECTDATA pdata)
{
    ATLASSERT(pdata && pdata->type == ObjectType::APPID && pdata->guid != GUID_NULL);

    CString strAppID;
    StringFromGUID2(pdata->guid, strAppID.GetBuffer(40), 40);
    strAppID.ReleaseBuffer();

    CString strPath;
    strPath.Format(_T("SOFTWARE\\Classes\\AppID\\%s"), static_cast<LPCTSTR>(strAppID));

    CRegKey key;
    auto lResult = key.Open(HKEY_LOCAL_MACHINE, strPath, KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE);
    if (lResult != ERROR_SUCCESS) {
        return FALSE;
    }

    auto appidRoot = InsertItem(_T("AppID="), TVI_ROOT, TVI_LAST);
    auto appid = InsertValues(key, appidRoot.m_hTreeItem, strAppID);

    InsertSubkeys(key, appid.m_hTreeItem);

    appidRoot.Expand();
    appid.Expand();

    return TRUE;
}

BOOL RegistryView::BuildIID(LPOBJECTDATA pdata)
{
    ATLASSERT(pdata && pdata->type == ObjectType::IID && pdata->guid != GUID_NULL);

    CString strIID;
    StringFromGUID2(pdata->guid, strIID.GetBuffer(40), 40);
    strIID.ReleaseBuffer();

    CString strPath;
    strPath.Format(_T("SOFTWARE\\Classes\\Interface\\%s"), static_cast<LPCTSTR>(strIID));

    CRegKey key;
    auto lResult = key.Open(HKEY_LOCAL_MACHINE, strPath, KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE);
    if (lResult != ERROR_SUCCESS) {
        return FALSE;
    }

    auto iidRoot = InsertItem(_T("IID="), TVI_ROOT, TVI_LAST);
    auto iid = InsertValues(key, iidRoot.m_hTreeItem, strIID);

    InsertSubkeys(key, iid.m_hTreeItem);

    iidRoot.Expand();
    iid.Expand();

    strPath.Format(_T("SOFTWARE\\Classes\\Interface\\%s\\ProxyStubClsid32"), static_cast<LPCTSTR>(strIID));
    lResult = key.Open(HKEY_LOCAL_MACHINE, strPath, KEY_QUERY_VALUE);
    if (lResult != ERROR_SUCCESS) {
        strPath.Format(_T("SOFTWARE\\Classes\\Interface\\%s\\ProxyStubClsid"), static_cast<LPCTSTR>(strIID));
        lResult = key.Open(HKEY_LOCAL_MACHINE, strPath, KEY_QUERY_VALUE);
    }

    if (lResult != ERROR_SUCCESS) {
        return TRUE;
    }

    TCHAR clsid[REG_BUFFER_SIZE];
    DWORD length = REG_BUFFER_SIZE;
    lResult = key.QueryStringValue(nullptr, clsid, &length);
    if (lResult != ERROR_SUCCESS) {
        return TRUE;
    }

    BuildCLSID(clsid);

    strPath.Format(_T("SOFTWARE\\Classes\\Interface\\%s\\TypeLib"), static_cast<LPCTSTR>(strIID));
    lResult = key.Open(HKEY_LOCAL_MACHINE, strPath, KEY_QUERY_VALUE);
    if (lResult != ERROR_SUCCESS) {
        return TRUE;
    }

    TCHAR typelib[REG_BUFFER_SIZE];
    length = REG_BUFFER_SIZE;
    lResult = key.QueryStringValue(nullptr, typelib, &length);
    if (lResult != ERROR_SUCCESS) {
        return TRUE;
    }

    BuildTypeLib(typelib);

    return TRUE;
}

BOOL RegistryView::BuildCatID(LPOBJECTDATA pdata)
{
    ATLASSERT(pdata && pdata->type == ObjectType::CATID && pdata->guid != GUID_NULL);

    CString strCatID;
    StringFromGUID2(pdata->guid, strCatID.GetBuffer(40), 40);
    strCatID.ReleaseBuffer();

    CString strPath;
    strPath.Format(_T("Component Categories\\%s"), static_cast<LPCTSTR>(strCatID));

    CRegKey key;
    auto lResult = key.Open(HKEY_CLASSES_ROOT, strPath, KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE);
    if (lResult != ERROR_SUCCESS) {
        return FALSE;
    }

    auto appidRoot = InsertItem(_T("Component Categories"), TVI_ROOT, TVI_LAST);
    auto appid = InsertValues(key, appidRoot.m_hTreeItem, strCatID);

    InsertSubkeys(key, appid.m_hTreeItem);

    appidRoot.Expand();
    appid.Expand();

    return TRUE;
}
