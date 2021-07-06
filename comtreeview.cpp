#include "StdAfx.h"
#include "comtreeview.h"
#include "resource.h"

#include <boost/algorithm/string.hpp>

ComTreeView::ComTreeView(): m_bMsgHandled(FALSE)
{
}

LRESULT ComTreeView::OnCreate(LPCREATESTRUCT)
{
    auto bResult = DefWindowProc();

    static constexpr auto icons = {
        IDI_NODE,
        IDI_COCLASS,
        IDI_INTERFACE,
        IDI_INTERFACE_GROUP,
        IDI_APPID,
        IDI_TYPELIB_GROUP,
        IDI_TYPELIB,
    };

    m_ImageList = ImageList_Create(16, 16, ILC_MASK | ILC_COLOR32, sizeof(icons) / sizeof(int), 0);

    for (auto icon : icons) {
        auto hIcon = LoadIcon(_Module.GetResourceInstance(), MAKEINTRESOURCE(icon));
        ATLASSERT(hIcon);
        m_ImageList.AddIcon(hIcon);
    }

    SetImageList(m_ImageList, TVSIL_NORMAL);
    CTreeViewCtrl::SetWindowLong(GWL_STYLE,
                                 CTreeViewCtrl::GetWindowLong(GWL_STYLE)
                                 | TVS_HASBUTTONS | TVS_HASLINES | TVS_FULLROWSELECT | TVS_INFOTIP
                                 | TVS_LINESATROOT | TVS_SHOWSELALWAYS);

    ConstructTree();

    SetMsgHandled(FALSE);

    return bResult;
}

void ComTreeView::OnDestroy()
{
    DeleteAllItems();
}

LRESULT ComTreeView::OnItemExpanding(LPNMHDR pnmh)
{
    const auto item = MAKE_TREEITEM(pnmh, this);

    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
    if (data == nullptr) {
        return 1;
    }

    if (data->type == ObjectType::CLSID && data->guid == GUID_NULL) {
        ExpandClasses(item);
    } else if (data->type == ObjectType::CLSID && data->guid != GUID_NULL) {
        ExpandInterfaces(item);
    } else if (data->type == ObjectType::IID && data->guid == GUID_NULL) {
        ExpandAllInterfaces(item);
    } else if (data->type == ObjectType::APPID && data->guid == GUID_NULL) {
        ExpandApps(item);
    } else if (data->type == ObjectType::TYPELIB && data->guid == GUID_NULL) {
        ExpandTypeLibs(item);
    }

    return 0;
}

LRESULT ComTreeView::OnDelete(LPNMHDR pnmh)
{
    const auto item = MAKE_OLDTREEITEM(pnmh, this);

    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());

    delete data;

    return 0;
}

void ComTreeView::ExpandClasses(const CTreeItem& item)
{
    auto result = GetItemState(item.m_hTreeItem, TVIS_EXPANDED);
    if (result) {
        return; // already expanded
    }

    auto child = item.GetChild();
    if (!child.IsNull()) {
        return; // already have children
    }

    ConstructClasses(item);

    SortChildren(item.m_hTreeItem);
}

void ComTreeView::ExpandInterfaces(const CTreeItem& item)
{
    auto result = GetItemState(item.m_hTreeItem, TVIS_EXPANDED);
    if (result) {
        return; // already expanded
    }

    auto child = item.GetChild();
    if (!child.IsNull()) {
        return; // already have children
    }

    ConstructInterfaces(item);

    child = item.GetChild();
    if (child.IsNull()) {
        // no children
        TVITEM tvitem = {};
        tvitem.hItem = item.m_hTreeItem;
        tvitem.mask = TVIF_CHILDREN;
        GetItem(&tvitem);
        tvitem.cChildren = 0;
        SetItem(&tvitem);
        return;
    }

    SortChildren(item.m_hTreeItem);
}

void ComTreeView::ExpandAllInterfaces(const CTreeItem& item)
{
    auto result = GetItemState(item.m_hTreeItem, TVIS_EXPANDED);
    if (result) {
        return; // already expanded
    }

    auto child = item.GetChild();
    if (!child.IsNull()) {
        return; // already have children
    }

    ConstructAllInterfaces(item);

    SortChildren(item.m_hTreeItem);
}

void ComTreeView::ExpandApps(const CTreeItem& item)
{
    auto result = GetItemState(item.m_hTreeItem, TVIS_EXPANDED);
    if (result) {
        return; // already expanded
    }

    auto child = item.GetChild();
    if (!child.IsNull()) {
        return; // already have children
    }

    ConstructApps(item);

    SortChildren(item.m_hTreeItem);
}

void ComTreeView::ExpandTypeLibs(const CTreeItem& item)
{
    auto result = GetItemState(item.m_hTreeItem, TVIS_EXPANDED);
    if (result) {
        return; // already expanded
    }

    auto child = item.GetChild();
    if (!child.IsNull()) {
        return; // already have children
    }

    ConstructTypeLibs(item);

    SortChildren(item.m_hTreeItem);
}

void ComTreeView::ConstructTypeLibs(const CTreeItem& item)
{
    CWaitCursor cursor;
    SetRedraw(FALSE);

    CRegKey key, subkey, verKey;
    auto lResult = key.Open(HKEY_LOCAL_MACHINE, _T("SOFTWARE\\Classes\\TypeLib"), KEY_ENUMERATE_SUB_KEYS);
    if (lResult != ERROR_SUCCESS)
        return;

    DWORD index = 0, length;

    for (;;) {
        TCHAR szTypeLibID[REG_BUFFER_SIZE]{};
        length = REG_BUFFER_SIZE;

        lResult = key.EnumKey(index++, szTypeLibID, &length);
        if (lResult != ERROR_SUCCESS) {
            break;
        }

        lResult = subkey.Open(key, szTypeLibID, KEY_READ);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }

        DWORD version = 0;

        for (;;) {
            TCHAR szVersion[REG_BUFFER_SIZE]{};
            length = REG_BUFFER_SIZE;

            lResult = subkey.EnumKey(version++, szVersion, &length);
            if (lResult != ERROR_SUCCESS) {
                break;
            }

            lResult = verKey.Open(subkey, szVersion, KEY_READ);
            if (lResult != ERROR_SUCCESS) {
                continue;
            }

            std::vector<tstring> out;
            split(out, szVersion, boost::is_any_of(_T(".")));

            WORD wMaj = 0, wMin = 0;
            if (out.size() == 2) {
                wMaj = _ttoi(out[0].c_str());
                wMin = _ttoi(out[1].c_str());
            }

            TCHAR szName[REG_BUFFER_SIZE]{};
            length = REG_BUFFER_SIZE;
            verKey.QueryStringValue(nullptr, szName, &length);

            CString strName;
            if (szName[0] == _T('\0')) {
                strName.Format(_T("%s (Ver %s)"), szTypeLibID, szVersion);
            } else {
                strName.Format(_T("%s (Ver %s)"), szName, szVersion);
            }

            TV_INSERTSTRUCT tvis;
            tvis.hParent = item.m_hTreeItem;
            tvis.hInsertAfter = TVI_LAST;
            tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
            tvis.itemex.cChildren = 0;
            tvis.itemex.pszText = const_cast<LPTSTR>(static_cast<LPCTSTR>(strName));
            tvis.itemex.iImage = 6;
            tvis.itemex.iSelectedImage = 6;

            auto pdata = std::make_unique<ObjectData>(ObjectType::TYPELIB, szTypeLibID, wMaj, wMin);
            if (pdata != nullptr && pdata->guid != GUID_NULL) {
                tvis.itemex.lParam = reinterpret_cast<LPARAM>(pdata.release());
                InsertItem(&tvis);
            }
        }
    }

    SetRedraw(TRUE);
}

void ComTreeView::ConstructApps(const CTreeItem& item)
{
    CWaitCursor cursor;
    SetRedraw(FALSE);

    CRegKey key, subkey;
    auto lResult = key.Open(HKEY_LOCAL_MACHINE, _T("SOFTWARE\\Classes\\AppID"), KEY_ENUMERATE_SUB_KEYS);
    if (lResult != ERROR_SUCCESS)
        return;

    DWORD index = 0, length = REG_BUFFER_SIZE;

    for (;;) {
        TCHAR appID[REG_BUFFER_SIZE]{};
        TCHAR name[REG_BUFFER_SIZE]{};
        length = REG_BUFFER_SIZE;

        lResult = key.EnumKey(index++, appID, &length);
        if (lResult != ERROR_SUCCESS) {
            break;
        }

        lResult = subkey.Open(key.m_hKey, appID, KEY_READ);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }

        if (appID[0] != '{') {
            // app name case
            length = REG_BUFFER_SIZE;
            _tcscpy(name, appID);
            subkey.QueryStringValue(_T("AppID"), appID, &length);
            appID[length] = _T('\0');
        }

        if (appID[0] == _T('\0')) {
            continue; // not much chance of success
        }

        if (name[0] == _T('\0')) {
            length = REG_BUFFER_SIZE;
            subkey.QueryStringValue(nullptr, name, &length);
            name[length] = _T('\0');
        }

        CString strName(name);
        strName.Trim();

        if (strName.GetLength() == 0) {
            strName = appID;
        }

        TV_INSERTSTRUCT tvis;
        tvis.hParent = item.m_hTreeItem;
        tvis.hInsertAfter = TVI_LAST;
        tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
        tvis.itemex.cChildren = 0;
        tvis.itemex.pszText = const_cast<LPTSTR>(static_cast<LPCTSTR>(strName));
        tvis.itemex.iImage = 4;
        tvis.itemex.iSelectedImage = 4;

        auto pdata = std::make_unique<ObjectData>(ObjectType::APPID, appID);
        if (pdata != nullptr && pdata->guid != GUID_NULL) {
            tvis.itemex.lParam = reinterpret_cast<LPARAM>(pdata.release());
            InsertItem(&tvis);
        }
    }

    SetRedraw(TRUE);
}

void ComTreeView::ConstructClasses(const CTreeItem& item)
{
    CWaitCursor cursor;
    SetRedraw(FALSE);

    CRegKey key, subkey;
    auto lResult = key.Open(HKEY_LOCAL_MACHINE, _T("SOFTWARE\\Classes\\CLSID"), KEY_ENUMERATE_SUB_KEYS);
    if (lResult != ERROR_SUCCESS)
        return;

    DWORD index = 0, length = REG_BUFFER_SIZE;
    TCHAR buff[REG_BUFFER_SIZE];
    TCHAR val[REG_BUFFER_SIZE];

    while ((lResult = key.EnumKey(index, buff, &length)) != ERROR_NO_MORE_ITEMS) {
        if (lResult != ERROR_SUCCESS) {
            break;
        }

        lResult = subkey.Open(key.m_hKey, buff, KEY_READ);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }

        length = REG_BUFFER_SIZE;
        subkey.QueryStringValue(nullptr, val, &length);
        val[length] = _T('\0');

        CString value(val);
        value.Trim();

        if (_tcscmp(buff, _T("CLSID")) != 0) {
            if (value.GetLength() == 0) {
                value = buff;
            }

            TV_INSERTSTRUCT tvis;
            tvis.hParent = item.m_hTreeItem;
            tvis.hInsertAfter = TVI_LAST;
            tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
            tvis.itemex.cChildren = 1;
            tvis.itemex.pszText = const_cast<LPTSTR>(static_cast<LPCTSTR>(value));
            tvis.itemex.iImage = 1;
            tvis.itemex.iSelectedImage = 1;

            auto pdata = std::make_unique<ObjectData>(ObjectType::CLSID, buff);
            if (pdata != nullptr && !IsEqualGUID(pdata->guid, GUID_NULL)) {
                tvis.itemex.lParam = reinterpret_cast<LPARAM>(pdata.release());
                InsertItem(&tvis);
            }
        }

        length = REG_BUFFER_SIZE;
        index++;
    }

    SetRedraw(TRUE);
}

void ComTreeView::ConstructInterfaces(const CTreeItem& item)
{
    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
    ATLASSERT(data != nullptr && data->type == ObjectType::CLSID && data->guid != GUID_NULL);

    CWaitCursor cursor;
    SetRedraw(FALSE);

    CComPtr<IClassFactory> pFactory;
    auto hr = CoGetClassObject(data->guid,
                               CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER,
                               nullptr, __uuidof(IClassFactory),
                               reinterpret_cast<void**>(&pFactory));
    if (FAILED(hr)) {
        SetRedraw(TRUE);
        return;
    }

    CComPtr<IUnknown> pUnk;
    hr = pFactory->CreateInstance(nullptr, IID_IUnknown, reinterpret_cast<void**>(&pUnk));
    if (FAILED(hr)) {
        SetRedraw(TRUE);
        return;
    }

    ConstructInterfaces(pUnk, item);

    SetRedraw(TRUE);
}

void ComTreeView::ConstructAllInterfaces(const CTreeItem& item)
{
    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
    ATLASSERT(data != nullptr && data->type == ObjectType::IID && data->guid == GUID_NULL);

    CWaitCursor cursor;
    SetRedraw(FALSE);

    CRegKey key, subkey;
    auto lResult = key.Open(HKEY_LOCAL_MACHINE, _T("SOFTWARE\\Classes\\Interface"), KEY_ENUMERATE_SUB_KEYS);
    if (lResult != ERROR_SUCCESS) {
        return;
    }

    DWORD index = 0;

    for (;;) {
        TCHAR szIID[REG_BUFFER_SIZE];
        TCHAR val[REG_BUFFER_SIZE];
        DWORD length = REG_BUFFER_SIZE;

        lResult = key.EnumKey(index++, szIID, &length);
        if (lResult == ERROR_NO_MORE_ITEMS) {
            break;
        }

        if (lResult != ERROR_SUCCESS) {
            break;
        }

        lResult = subkey.Open(key.m_hKey, szIID, KEY_READ);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }

        length = REG_BUFFER_SIZE;
        lResult = subkey.QueryStringValue(nullptr, val, &length);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }

        val[length] = _T('\0');

        CString value(val);
        value.Trim();

        if (value.GetLength() == 0) {
            continue;
        }

        IID iid;
        auto hr = IIDFromString(CT2OLE(szIID), &iid);
        if (FAILED(hr)) {
            continue;
        }

        TV_INSERTSTRUCT tvis{};
        tvis.hParent = item.m_hTreeItem;
        tvis.hInsertAfter = TVI_LAST;
        tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
        tvis.itemex.cChildren = 0;
        tvis.itemex.pszText = const_cast<LPTSTR>(static_cast<LPCTSTR>(value));
        tvis.itemex.iImage = 2;
        tvis.itemex.iSelectedImage = 2;
        tvis.itemex.lParam = reinterpret_cast<LPARAM>(new ObjectData(ObjectType::IID, iid));

        InsertItem(&tvis);
    }

    SetRedraw(TRUE);
}

void ComTreeView::ConstructInterfaces(CComPtr<IUnknown>& pUnk, const CTreeItem& item)
{
    CRegKey key, subkey;
    auto lResult = key.Open(HKEY_LOCAL_MACHINE, _T("SOFTWARE\\Classes\\Interface"), KEY_ENUMERATE_SUB_KEYS);
    if (lResult != ERROR_SUCCESS) {
        return;
    }

    DWORD index = 0;

    for (;;) {
        TCHAR szIID[REG_BUFFER_SIZE];
        TCHAR val[REG_BUFFER_SIZE];
        DWORD length = REG_BUFFER_SIZE;

        lResult = key.EnumKey(index++, szIID, &length);
        if (lResult == ERROR_NO_MORE_ITEMS) {
            break;
        }

        if (lResult != ERROR_SUCCESS) {
            break;
        }

        lResult = subkey.Open(key.m_hKey, szIID, KEY_READ);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }

        length = REG_BUFFER_SIZE;
        lResult = subkey.QueryStringValue(nullptr, val, &length);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }

        val[length] = _T('\0');

        CString value(val);
        value.Trim();

        IID iid;
        auto hr = IIDFromString(CT2OLE(szIID), &iid);
        if (FAILED(hr)) {
            continue;
        }

        CComPtr<IUnknown> pFoo;
        hr = pUnk->QueryInterface(iid, reinterpret_cast<void**>(&pFoo));
        if (FAILED(hr)) {
            continue;
        }

        pFoo.Release();

        TV_INSERTSTRUCT tvis{};
        tvis.hParent = item.m_hTreeItem;
        tvis.hInsertAfter = TVI_LAST;
        tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
        tvis.itemex.cChildren = 0;
        tvis.itemex.pszText = const_cast<LPTSTR>(static_cast<LPCTSTR>(value));
        tvis.itemex.iImage = 2;
        tvis.itemex.iSelectedImage = 2;
        tvis.itemex.lParam = reinterpret_cast<LPARAM>(new ObjectData(ObjectType::IID, iid));

        InsertItem(&tvis);
    }
}

void ComTreeView::ConstructTree()
{
    TV_INSERTSTRUCT tvis{};
    tvis.hParent = TVI_ROOT;
    tvis.hInsertAfter = TVI_LAST;
    tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_EXPANDEDIMAGE | TVIF_TEXT |
        TVIF_PARAM;
    tvis.itemex.cChildren = 1;
    tvis.itemex.pszText = _T("Objects");
    tvis.itemex.iImage = 0;
    tvis.itemex.iSelectedImage = 0;
    tvis.itemex.iExpandedImage = 0;
    tvis.itemex.lParam = reinterpret_cast<LPARAM>(new ObjectData(ObjectType::CLSID, GUID_NULL));

    InsertItem(&tvis);

    tvis.itemex.pszText = _T("Interfaces");
    tvis.itemex.iImage = 3;
    tvis.itemex.iSelectedImage = 3;
    tvis.itemex.iExpandedImage = 3;
    tvis.itemex.lParam = reinterpret_cast<LPARAM>(new ObjectData(ObjectType::IID, GUID_NULL));
    InsertItem(&tvis);

    tvis.itemex.pszText = _T("Applications");
    tvis.itemex.iImage = 4;
    tvis.itemex.iSelectedImage = 4;
    tvis.itemex.iExpandedImage = 4;
    tvis.itemex.lParam = reinterpret_cast<LPARAM>(new ObjectData(ObjectType::APPID, GUID_NULL));
    InsertItem(&tvis);

    tvis.itemex.pszText = _T("Type Libraries");
    tvis.itemex.iImage = 5;
    tvis.itemex.iSelectedImage = 5;
    tvis.itemex.iExpandedImage = 5;
    tvis.itemex.lParam = reinterpret_cast<LPARAM>(new ObjectData(ObjectType::TYPELIB, GUID_NULL));
    InsertItem(&tvis);

    SortChildren(TVI_ROOT);
}
