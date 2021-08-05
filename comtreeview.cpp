#include "StdAfx.h"
#include "comtreeview.h"
#include "resource.h"
#include "objdata.h"
#include "autotypeattr.h"

#include <boost/algorithm/string.hpp>

static constexpr GUID CATID_CONTROLS_GUID =
    { 0x40FC6ED4, 0x2438, 0x11CF, { 0xa3, 0xdb, 0x08, 0x00, 0x36, 0xF1, 0x25, 0x02 } };

static constexpr GUID CATID_DOCOBJECTS_GUID =
    { 0x40FC6ED8, 0x2438, 0x11CF, { 0xa3, 0xdb, 0x08, 0x00, 0x36, 0xF1, 0x25, 0x02 } };

static constexpr GUID CATID_EMBEDDABLE_GUID =
    { 0x40FC6ED3, 0x2438, 0x11CF, { 0xa3, 0xdb, 0x08, 0x00, 0x36, 0xF1, 0x25, 0x02 } };

static constexpr GUID CATID_AUTOMATION_GUID =
    { 0x40FC6ED5, 0x2438, 0x11CF, { 0xa3, 0xdb, 0x08, 0x00, 0x36, 0xF1, 0x25, 0x02 } };

static constexpr COLORREF RGBCYAN = RGB(0, 128, 128);

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
        IDI_APPID_GROUP,
        IDI_APPID,
        IDI_TYPELIB_GROUP,
        IDI_TYPELIB,
        IDI_CATID_GROUP,
        IDI_CATID
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
    } else if (data->type == ObjectType::CATID && data->guid == GUID_NULL) {
        ExpandCategories(item);
    } else if (data->type == ObjectType::CATID && data->guid != GUID_NULL) {
        ExpandCategory(item);
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

void ComTreeView::AddFileMoniker(LPCTSTR pFilename, LPUNKNOWN pUnknown, REFCLSID clsid)
{
    ATLASSERT(pFilename);
    ATLASSERT(pUnknown);

    CString strName(pFilename);
    if (clsid != GUID_NULL) {
        CString strGUID;
        StringFromGUID2(clsid, strGUID.GetBuffer(40), 40);
        strGUID.ReleaseBuffer();

        CString strPath;
        strPath.Format(_T("CLSID\\%s"), strGUID);

        CRegKey key;
        auto lResult = key.Open(HKEY_CLASSES_ROOT, strPath, KEY_READ);
        if (lResult == ERROR_SUCCESS) {
            TCHAR val[REG_BUFFER_SIZE];
            ULONG length = REG_BUFFER_SIZE;
            lResult = key.QueryStringValue(nullptr, val, &length);
            if (lResult == ERROR_SUCCESS) {
                strName.Format(_T("%s (%s)"), pFilename, val);
            }
        }
    }

    if (m_objectInstances.IsNull()) {
        TV_INSERTSTRUCT tvis{};
        tvis.hParent = TVI_ROOT;
        tvis.hInsertAfter = TVI_FIRST;
        tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_EXPANDEDIMAGE | TVIF_TEXT |
            TVIF_PARAM;
        tvis.itemex.cChildren = 1;
        tvis.itemex.pszText = _T("Object Instances");
        tvis.itemex.iImage = 0;
        tvis.itemex.iSelectedImage = 0;
        tvis.itemex.iExpandedImage = 0;
        tvis.itemex.lParam = reinterpret_cast<LPARAM>(new ObjectData(ObjectType::CLSID, GUID_NULL));

        m_objectInstances = InsertItem(&tvis);
    }

    TV_INSERTSTRUCT tvis;
    tvis.hParent = m_objectInstances;
    tvis.hInsertAfter = TVI_FIRST;
    tvis.itemex.mask = TVIF_STATE | TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
    tvis.itemex.state = TVIS_BOLD;
    tvis.itemex.stateMask = TVIS_BOLD;
    tvis.itemex.cChildren = 1;
    tvis.itemex.pszText = const_cast<LPTSTR>(static_cast<LPCTSTR>(strName));
    tvis.itemex.iImage = 1;
    tvis.itemex.iSelectedImage = 1;

    auto pdata = std::make_unique<ObjectData>(ObjectType::CLSID, pUnknown, clsid);
    tvis.itemex.lParam = reinterpret_cast<LPARAM>(pdata.release());

    auto item = InsertItem(&tvis);
    item.Expand();

    SelectItem(item);
}

void ComTreeView::AddFileTypeLib(LPCTSTR pFilename, LPTYPELIB pTypeLib)
{
    ATLASSERT(pFilename);
    ATLASSERT(pTypeLib);

    CComBSTR bstrName, bstrDoc;
    pTypeLib->GetDocumentation(MEMBERID_NIL, &bstrName, &bstrDoc, nullptr, nullptr);
    
    CString strItem;
    if (bstrDoc.Length() != 0) {
        strItem.Format(_T("%s (%s)"), bstrName, bstrDoc);
    } else {
        strItem = bstrName;
    }

    AutoTypeLibAttr attr(pTypeLib);
    attr.Get();

    if (m_fileTypeLibs.IsNull()) {
        TV_INSERTSTRUCT tvis{};
        tvis.hParent = TVI_ROOT;
        tvis.hInsertAfter = TVI_FIRST;
        tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_EXPANDEDIMAGE | TVIF_TEXT |
            TVIF_PARAM;
        tvis.itemex.cChildren = 1;
        tvis.itemex.pszText = _T("Type Library Files");
        tvis.itemex.iImage = 6;
        tvis.itemex.iSelectedImage = 6;
        tvis.itemex.iExpandedImage = 6;
        tvis.itemex.lParam = reinterpret_cast<LPARAM>(new ObjectData(ObjectType::TYPELIB, GUID_NULL));

        m_fileTypeLibs = InsertItem(&tvis);
    }

    TV_INSERTSTRUCT tvis{};
    tvis.hParent = m_fileTypeLibs;
    tvis.hInsertAfter = TVI_FIRST;
    tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
    tvis.itemex.cChildren = 0;
    tvis.itemex.pszText = const_cast<LPTSTR>(static_cast<LPCTSTR>(strItem));
    tvis.itemex.iImage = 7;
    tvis.itemex.iSelectedImage = 7;

    auto pdata = std::make_unique<ObjectData>(ObjectType::TYPELIB, pTypeLib,
                                              attr->guid, attr->wMajorVerNum,
                                              attr->wMinorVerNum);
    tvis.itemex.lParam = reinterpret_cast<LPARAM>(pdata.release());

    auto item = InsertItem(&tvis);
    SelectItem(item);
}

BOOL ComTreeView::IsSelectedInstance() const
{
    auto item = GetSelectedItem();
    if (item.IsNull()) {
        return FALSE;
    }

    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
    if (data == nullptr) {
        return FALSE;
    }

    return data->type == ObjectType::CLSID && data->pUnknown != nullptr;
}

BOOL ComTreeView::IsGUIDSelected() const
{
    auto item = GetSelectedItem();
    if (item.IsNull()) {
        return FALSE;
    }

    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
    if (data == nullptr) {
        return FALSE;
    }

    return data->guid != GUID_NULL;
}

void ComTreeView::ReleaseSelectedObject()
{
    auto item = GetSelectedItem();
    if (item.IsNull()) {
        return;
    }

    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
    if (data == nullptr) {
        return;
    }

    data->SafeRelease();

    item.Expand(TVE_COLLAPSE);
    item.SetState(0, TVIS_BOLD);
}

void ComTreeView::CopyGUIDToClipboard()
{
    auto item = GetSelectedItem();
    if (item.IsNull()) {
        return;
    }

    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
    if (data == nullptr) {
        return;
    }

    if (data->guid == GUID_NULL) {
        return;
    }

    CString strGUID;
    StringFromGUID2(data->guid, strGUID.GetBuffer(40), 40);
    strGUID.ReleaseBuffer();

    auto hGlobal = GlobalAlloc(GMEM_MOVEABLE, (strGUID.GetLength() + 1) * sizeof(TCHAR));
    if (hGlobal == nullptr) {
        return;
    }

    CT2W wszGUID(strGUID);

    auto pData = static_cast<LPWSTR>(GlobalLock(hGlobal));
    wcscpy(pData, wszGUID);

    GlobalUnlock(hGlobal);

    if (!OpenClipboard()) {
        return;
    }

    EmptyClipboard();

    SetClipboardData(CF_UNICODETEXT, hGlobal);

    CloseClipboard();
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

    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
    if (data != nullptr && data->pUnknown != nullptr) {
        if (!item.GetChild().IsNull()) {
            return; // live instance already expanded
        }
    }

    // Always remove all interfaces for new instance
    auto child = item.GetChild();
    if (!child.IsNull()) {
        auto sibling = GetNextSiblingItem(child);
        while (!sibling.IsNull()) {
            HTREEITEM hItem = sibling;
            sibling = GetNextSiblingItem(sibling);
            DeleteItem(hItem);
        }
        DeleteItem(child);
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

void ComTreeView::ExpandCategories(const CTreeItem& item)
{
    auto result = GetItemState(item.m_hTreeItem, TVIS_EXPANDED);
    if (result) {
        return; // already expanded
    }

    auto child = item.GetChild();
    if (!child.IsNull()) {
        return; // already have children
    }

    ConstructCategories(item);

    SortChildren(item.m_hTreeItem);
}

void ComTreeView::ExpandCategory(const CTreeItem& item)
{
    auto result = GetItemState(item.m_hTreeItem, TVIS_EXPANDED);
    if (result) {
        return; // already expanded
    }

    auto child = item.GetChild();
    if (!child.IsNull()) {
        return; // already have children
    }

    ConstructCategory(item);

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

void ComTreeView::ConstructTypeLibs(const CTreeItem& item)
{
    CWaitCursor cursor;
    SetRedraw(FALSE);

    CRegKey key, subkey, verKey;
    auto lResult = key.Open(HKEY_CLASSES_ROOT, _T("TypeLib"), KEY_ENUMERATE_SUB_KEYS);
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
            tvis.itemex.iImage = 7;
            tvis.itemex.iSelectedImage = 7;

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
    auto lResult = key.Open(HKEY_CLASSES_ROOT, _T("AppID"), KEY_ENUMERATE_SUB_KEYS);
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
        }

        if (appID[0] == _T('\0')) {
            continue; // not much chance of success
        }

        if (name[0] == _T('\0')) {
            length = REG_BUFFER_SIZE;
            subkey.QueryStringValue(nullptr, name, &length);
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
        tvis.itemex.iImage = 5;
        tvis.itemex.iSelectedImage = 5;

        auto pdata = std::make_unique<ObjectData>(ObjectType::APPID, appID);
        if (pdata != nullptr && pdata->guid != GUID_NULL) {
            tvis.itemex.lParam = reinterpret_cast<LPARAM>(pdata.release());
            InsertItem(&tvis);
        }
    }

    SetRedraw(TRUE);
}

int ComTreeView::LoadToolboxImage(const CRegKey& key)
{
    auto iImage = 1;

    CRegKey subkey;
    auto lResult = subkey.Open(key.m_hKey, _T("ToolboxBitmap32"), KEY_READ);
    if (lResult != ERROR_SUCCESS) {
        lResult = subkey.Open(key.m_hKey, _T("ToolboxBitmap"), KEY_READ);
        if (lResult != ERROR_SUCCESS) {
            return iImage;
        }
    }

    TCHAR szBitmap[REG_BUFFER_SIZE];
    ULONG length = REG_BUFFER_SIZE;

    lResult = subkey.QueryStringValue(nullptr, szBitmap, &length);
    if (lResult != ERROR_SUCCESS) {
        return iImage;
    }

    std::vector<tstring> out;
    split(out, szBitmap, boost::is_any_of(_T(",")));

    if (out.size() < 2) {
        return iImage;
    }

    auto fileName = out[0];
    auto id = _ttoi(out[1].c_str());

    TCHAR szFileName[REG_BUFFER_SIZE];
    ExpandEnvironmentStrings(fileName.c_str(), szFileName,
                             REG_BUFFER_SIZE);

    auto hinstBitmap = LoadLibraryEx(szFileName,
                                     nullptr,
                                     LOAD_LIBRARY_AS_DATAFILE);
    if (!hinstBitmap) {
        return 1;
    }

    auto hbmGlyph = LoadBitmap(hinstBitmap, MAKEINTRESOURCE(id));
    if (hbmGlyph) {
        iImage = m_ImageList.Add(hbmGlyph, RGBCYAN);
        DeleteObject(hbmGlyph);
    } else {
        auto hIcon = LoadIcon(hinstBitmap, MAKEINTRESOURCE(id));
        if (hIcon) {
            iImage = m_ImageList.AddIcon(hIcon);
            DestroyIcon(hIcon);
        }
    }

    FreeLibrary(hinstBitmap);

    return iImage;
}

void ComTreeView::ConstructClasses(const CTreeItem& item)
{
    CWaitCursor cursor;
    SetRedraw(FALSE);

    CRegKey key, subkey;
    auto lResult = key.Open(HKEY_CLASSES_ROOT, _T("CLSID"), KEY_ENUMERATE_SUB_KEYS);
    if (lResult != ERROR_SUCCESS)
        return;

    DWORD index = 0, length = REG_BUFFER_SIZE;
    TCHAR val[REG_BUFFER_SIZE];

    for (;;) {
        TCHAR szCLSID[REG_BUFFER_SIZE];
        length = REG_BUFFER_SIZE;

        lResult = key.EnumKey(index++, szCLSID, &length);
        if (lResult != ERROR_SUCCESS) {
            break;
        }

        if (_tcscmp(szCLSID, _T("CLSID")) == 0) {
            continue;
        }

        lResult = subkey.Open(key.m_hKey, szCLSID, KEY_READ);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }

        length = REG_BUFFER_SIZE;
        subkey.QueryStringValue(nullptr, val, &length);

        CString value(val);
        value.Trim();

        if (value.GetLength() == 0) {
            value = szCLSID;
        }

        auto iImage = LoadToolboxImage(subkey);

        TV_INSERTSTRUCT tvis;
        tvis.hParent = item.m_hTreeItem;
        tvis.hInsertAfter = TVI_LAST;
        tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
        tvis.itemex.cChildren = 1;
        tvis.itemex.pszText = const_cast<LPTSTR>(static_cast<LPCTSTR>(value));
        tvis.itemex.iImage = iImage;
        tvis.itemex.iSelectedImage = iImage;

        auto pdata = std::make_unique<ObjectData>(ObjectType::CLSID, szCLSID);
        if (pdata != nullptr && !IsEqualGUID(pdata->guid, GUID_NULL)) {
            tvis.itemex.lParam = reinterpret_cast<LPARAM>(pdata.release());
            InsertItem(&tvis);
        }
    }

    SetRedraw(TRUE);
}

void ComTreeView::ConstructInterfaces(const CTreeItem& item)
{
    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
    ATLASSERT(data != nullptr && data->type == ObjectType::CLSID && data->guid != GUID_NULL);

    CWaitCursor cursor;
    SetRedraw(FALSE);

    if (data->pUnknown == nullptr) {
        CComPtr<IClassFactory> pFactory;
        auto hr = CoGetClassObject(data->guid,
                                   CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER,
                                   nullptr, __uuidof(IClassFactory),
                                   reinterpret_cast<void**>(&pFactory));
        if (FAILED(hr)) {
            SetRedraw(TRUE);
            return;
        }

        hr = pFactory->CreateInstance(nullptr, IID_IUnknown, reinterpret_cast<void**>(&data->pUnknown));
        if (FAILED(hr)) {
            SetRedraw(TRUE);
            return;
        }

        SetItemState(item.m_hTreeItem, TVIS_BOLD, TVIS_BOLD);

        SelectItem(nullptr); // re-select a live node, if needed
        SelectItem(item.m_hTreeItem);
    }

    if (data->pUnknown) {
        ConstructInterfaces(data->pUnknown, item);
    }

    SetRedraw(TRUE);
}

void ComTreeView::ConstructAllInterfaces(const CTreeItem& item)
{
    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
    ATLASSERT(data != nullptr && data->type == ObjectType::IID && data->guid == GUID_NULL);

    CWaitCursor cursor;
    SetRedraw(FALSE);

    CRegKey key, subkey;
    auto lResult = key.Open(HKEY_CLASSES_ROOT, _T("Interface"), KEY_ENUMERATE_SUB_KEYS);
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

void ComTreeView::ConstructInterfaces(LPUNKNOWN pUnknown, const CTreeItem& item)
{
    ATLASSERT(pUnknown);

    CRegKey key, subkey;
    auto lResult = key.Open(HKEY_CLASSES_ROOT, _T("Interface"), KEY_ENUMERATE_SUB_KEYS);
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

        CString value(val);
        value.Trim();

        IID iid;
        auto hr = IIDFromString(CT2OLE(szIID), &iid);
        if (FAILED(hr)) {
            continue;
        }

        CComPtr<IUnknown> pFoo;
        hr = pUnknown->QueryInterface(iid, reinterpret_cast<void**>(&pFoo));
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

void ComTreeView::ConstructCategories(const CTreeItem& item)
{
    CRegKey key, subkey;
    auto lResult = key.Open(HKEY_CLASSES_ROOT, _T("Component Categories"), KEY_ENUMERATE_SUB_KEYS);
    if (lResult != ERROR_SUCCESS) {
        return;
    }

    DWORD index = 0;

    for (;;) {
        TCHAR szGUID[REG_BUFFER_SIZE];
        TCHAR val[REG_BUFFER_SIZE];
        DWORD length = REG_BUFFER_SIZE;

        lResult = key.EnumKey(index++, szGUID, &length);
        if (lResult != ERROR_SUCCESS) {
            break;
        }

        lResult = subkey.Open(key.m_hKey, szGUID, KEY_READ);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }

        length = REG_BUFFER_SIZE;
        lResult = subkey.QueryStringValue(_T("409"), val, &length);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }

        CString value(val);
        value.Trim();

        TV_INSERTSTRUCT tvis{};
        tvis.hParent = item.m_hTreeItem;
        tvis.hInsertAfter = TVI_LAST;
        tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
        tvis.itemex.cChildren = 1;
        tvis.itemex.pszText = const_cast<LPTSTR>(static_cast<LPCTSTR>(value));
        tvis.itemex.iImage = 9;
        tvis.itemex.iSelectedImage = 9;

        auto pdata = std::make_unique<ObjectData>(ObjectType::CATID, szGUID);
        if (pdata != nullptr && !IsEqualGUID(pdata->guid, GUID_NULL)) {
            tvis.itemex.lParam = reinterpret_cast<LPARAM>(pdata.release());
            InsertItem(&tvis);
        }
    }
}

void ComTreeView::ConstructCategory(const CTreeItem& item)
{
    auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
    ATLASSERT(data != nullptr && data->type == ObjectType::CATID && data->guid != GUID_NULL);

    CString strGUID;
    StringFromGUID2(data->guid, strGUID.GetBuffer(40), 40);
    strGUID.ReleaseBuffer();

    CWaitCursor cursor;
    SetRedraw(FALSE);

    CRegKey key, subkey;
    auto lResult = key.Open(HKEY_CLASSES_ROOT, _T("CLSID"), KEY_ENUMERATE_SUB_KEYS);
    if (lResult != ERROR_SUCCESS)
        return;

    DWORD index = 0, length = REG_BUFFER_SIZE;
    TCHAR val[REG_BUFFER_SIZE];

    for (;;) {
        TCHAR szCLSID[REG_BUFFER_SIZE];
        length = REG_BUFFER_SIZE;

        lResult = key.EnumKey(index++, szCLSID, &length);
        if (lResult != ERROR_SUCCESS) {
            break;
        }

        lResult = subkey.Open(key.m_hKey, szCLSID, KEY_READ);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }

        length = REG_BUFFER_SIZE;
        subkey.QueryStringValue(nullptr, val, &length);

        CString value(val);
        value.Trim();

        CString strPath;
        strPath.Format(_T("%s\\Implemented Categories\\%s"), szCLSID, strGUID);

        CRegKey catKey;
        lResult = catKey.Open(key.m_hKey, strPath, KEY_READ);
        auto bHasCategory = lResult == ERROR_SUCCESS;
        if (!bHasCategory) {
            if (data->guid == CATID_CONTROLS_GUID) {
                // we allow for "Control" key for controls
                strPath.Format(_T("%s\\Control"), szCLSID);
            } else if (data->guid == CATID_DOCOBJECTS_GUID) {
                // we allow for "DocObjects" key for "Document Objects"
                strPath.Format(_T("%s\\DocObject"), szCLSID);
            } else if (data->guid == CATID_EMBEDDABLE_GUID) {
                // we allow for "Insertable" key for "Embeddable Objects"
                strPath.Format(_T("%s\\Insertable"), szCLSID);
            } else if (data->guid == CATID_AUTOMATION_GUID) {
                // we allow for "Programmable" key for "Automation Objects"
                strPath.Format(_T("%s\\Programmable"), szCLSID);
            }

            lResult = catKey.Open(key.m_hKey, strPath, KEY_READ);
            bHasCategory = lResult == ERROR_SUCCESS;
        }

        if (!bHasCategory) {
            continue;
        }

        auto iImage = LoadToolboxImage(subkey);

        TV_INSERTSTRUCT tvis;
        tvis.hParent = item.m_hTreeItem;
        tvis.hInsertAfter = TVI_LAST;
        tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
        tvis.itemex.cChildren = 1;
        tvis.itemex.pszText = const_cast<LPTSTR>(static_cast<LPCTSTR>(value));
        tvis.itemex.iImage = iImage;
        tvis.itemex.iSelectedImage = iImage;

        auto pdata = std::make_unique<ObjectData>(ObjectType::CLSID, szCLSID);
        if (pdata != nullptr && !IsEqualGUID(pdata->guid, GUID_NULL)) {
            tvis.itemex.lParam = reinterpret_cast<LPARAM>(pdata.release());
            InsertItem(&tvis);
        }
    }

    SetRedraw(TRUE);
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
    tvis.itemex.iImage = 6;
    tvis.itemex.iSelectedImage = 6;
    tvis.itemex.iExpandedImage = 6;
    tvis.itemex.lParam = reinterpret_cast<LPARAM>(new ObjectData(ObjectType::TYPELIB, GUID_NULL));
    InsertItem(&tvis);

    tvis.itemex.pszText = _T("Categories");
    tvis.itemex.iImage = 8;
    tvis.itemex.iSelectedImage = 8;
    tvis.itemex.iExpandedImage = 8;
    tvis.itemex.lParam = reinterpret_cast<LPARAM>(new ObjectData(ObjectType::CATID, GUID_NULL));
    InsertItem(&tvis);

    SortChildren(TVI_ROOT);
}
