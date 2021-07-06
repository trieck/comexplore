#include "stdafx.h"
#include "resource.h"
#include "TypeLibTree.h"
#include "util.h"

TypeLibTree::TypeLibTree(): m_bMsgHandled(0)
{
}

LRESULT TypeLibTree::OnCreate(LPCREATESTRUCT pcs)
{
    auto bResult = DefWindowProc();

    static constexpr auto icons = {
        IDI_TYPELIB,
        IDI_COCLASS,
        IDI_INTERFACE,
        IDI_DISPATCH,
        IDI_RECORD,
        IDI_ENUM,
        IDI_UNION,
        IDI_MODULE,
        IDI_ALIAS,
        IDI_FUNCTION
    };

    m_ImageList = ImageList_Create(16, 16, ILC_MASK | ILC_COLOR32, sizeof(icons) / sizeof(int), 0);

    for (auto icon : icons) {
        auto hIcon = LoadIcon(_Module.GetResourceInstance(), MAKEINTRESOURCE(icon));
        ATLASSERT(hIcon);
        m_ImageList.AddIcon(hIcon);
    }

    SetImageList(m_ImageList, TVSIL_NORMAL);
    SetWindowLong(GWL_STYLE, GetWindowLong(GWL_STYLE)
                  | TVS_HASBUTTONS | TVS_HASLINES | TVS_FULLROWSELECT | TVS_INFOTIP
                  | TVS_LINESATROOT | TVS_SHOWSELALWAYS);

    if (!BuildView(static_cast<LPOBJECTDATA>(pcs->lpCreateParams))) {
        return -1; // don't create
    }

    SetMsgHandled(FALSE);

    return bResult;
}

BOOL TypeLibTree::BuildView(LPOBJECTDATA pdata)
{
    ATLASSERT(pdata && pdata->type == ObjectType::TYPELIB && pdata->guid != GUID_NULL);

    CString strGUID;
    StringFromGUID2(pdata->guid, strGUID.GetBuffer(40), 40);

    auto lcid = GetUserDefaultLCID();

    auto hr = LoadRegTypeLib(pdata->guid, pdata->wMaj, pdata->wMin, lcid, &m_pTypeLib);
    if (FAILED(hr)) {
        return FALSE;
    }

    if (m_pTypeLib != nullptr) {
        return BuildView();
    }

    return FALSE;
}

BOOL TypeLibTree::BuildView()
{
    ATLASSERT(m_pTypeLib != nullptr);

    CComBSTR bstrName, bstrDoc;
    auto hr = m_pTypeLib->GetDocumentation(MEMBERID_NIL, &bstrName, &bstrDoc, nullptr, nullptr);
    if (FAILED(hr)) {
        return FALSE;
    }

    CString strItem;
    if (bstrDoc.Length() != 0) {
        strItem.Format(_T("%s (%s)"), static_cast<LPCTSTR>(bstrName), static_cast<LPCTSTR>(bstrDoc));
    } else {
        strItem = bstrName;
    }

    auto hRoot = InsertItem(strItem, 0, 0, 1, TVI_ROOT, TVI_LAST);

    BuildTypeInfo(hRoot);
    SortChildren(hRoot);

    Expand(hRoot);

    return TRUE;
}

HTREEITEM TypeLibTree::AddTypeInfo(HTREEITEM hParent, LPTYPEINFO pTypeInfo, TYPEKIND kind)
{
    ATLASSERT(hParent != nullptr);
    ATLASSERT(pTypeInfo != nullptr);

    CComBSTR bstrName;
    auto hr = pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);
    if (FAILED(hr) || bstrName.Length() == 0) {
        return nullptr;
    }

    int nImage;
    switch (kind) {
    case TKIND_COCLASS:
        nImage = 1;
        break;
    case TKIND_INTERFACE:
        nImage = 2;
        break;
    case TKIND_DISPATCH:
        nImage = 3;
        break;
    case TKIND_RECORD:
        nImage = 4;
        break;
    case TKIND_ENUM:
        nImage = 5;
        break;
    case TKIND_UNION:
        nImage = 6;
        break;
    case TKIND_MODULE:
        nImage = 7;
        break;
    case TKIND_ALIAS:
        nImage = 8;
        break;
    default:
        return nullptr;
    }

    return InsertItem(bstrName, nImage, nImage, 1, hParent, TVI_LAST, pTypeInfo);
}

void TypeLibTree::BuildTypeInfo(HTREEITEM hParent)
{
    ATLASSERT(m_pTypeLib != nullptr);

    auto nCount = m_pTypeLib->GetTypeInfoCount();
    for (auto i = 0u; i < nCount; ++i) {
        CComPtr<ITypeInfo> pTypeInfo;
        auto hr = m_pTypeLib->GetTypeInfo(i, &pTypeInfo);
        if (FAILED(hr)) {
            continue;
        }

        TYPEKIND kind;
        hr = m_pTypeLib->GetTypeInfoType(i, &kind);
        if (FAILED(hr)) {
            continue;
        }

        HREFTYPE hRefType;
        hr = pTypeInfo->GetRefTypeOfImplType(-1, &hRefType);
        if (SUCCEEDED(hr)) {
            CComPtr<ITypeInfo> pTypeInfo2;
            hr = pTypeInfo->GetRefTypeInfo(hRefType, &pTypeInfo2);
            if (SUCCEEDED(hr)) {
                TYPEATTR* pnewAttr;
                hr = pTypeInfo2->GetTypeAttr(&pnewAttr);
                if (FAILED(hr)) {
                    continue;
                }

                auto hTreeItem = AddTypeInfo(hParent, pTypeInfo2, pnewAttr->typekind);
                pTypeInfo2->ReleaseTypeAttr(pnewAttr);

                if (hTreeItem) {
                    pTypeInfo2.Detach(); // tree owns
                }
            }
        }

        AddTypeInfo(hParent, pTypeInfo.Detach(), kind);
    }
}

void TypeLibTree::AddFunctions(const CTreeItem& item, LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    // Add functions
    for (auto i = 0u; i < pAttr->cFuncs; ++i) {
        LPFUNCDESC pFuncDesc;
        auto hr = pTypeInfo->GetFuncDesc(i, &pFuncDesc);
        if (FAILED(hr)) {
            continue;
        }

        UINT cNames;
        CComBSTR bstrName;
        hr = pTypeInfo->GetNames(pFuncDesc->memid, &bstrName, 1, &cNames);
        if (FAILED(hr)) {
            pTypeInfo->ReleaseFuncDesc(pFuncDesc);
            continue;
        }

        if (cNames == 0) {
            pTypeInfo->ReleaseFuncDesc(pFuncDesc);
            continue;
        }

        auto hItem = InsertItem(bstrName, 9, 9, 0, item.m_hTreeItem, TVI_LAST, pTypeInfo);
        if (hItem != nullptr) {
            pTypeInfo->AddRef();
        }

        pTypeInfo->ReleaseFuncDesc(pFuncDesc);
    }
}

void TypeLibTree::AddImplTypes(const CTreeItem& item, LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    // Add implemented types
    for (auto i = 0u; i < pAttr->cImplTypes; ++i) {

        HREFTYPE hRef = 0;
        auto hr = pTypeInfo->GetRefTypeOfImplType(i, &hRef);
        if (FAILED(hr)) {
            continue;
        }

        CComPtr<ITypeInfo> pTypeInfo2;
        hr = pTypeInfo->GetRefTypeInfo(hRef, &pTypeInfo2);
        if (FAILED(hr)) {
            continue;
        }

        TYPEATTR* pnewAttr;
        hr = pTypeInfo2->GetTypeAttr(&pnewAttr);
        if (FAILED(hr)) {
            continue;
        }

        auto hTreeItem = AddTypeInfo(item.m_hTreeItem, pTypeInfo2, pnewAttr->typekind);

        pTypeInfo2->ReleaseTypeAttr(pnewAttr);
        if (hTreeItem != nullptr) {
            pTypeInfo2.Detach(); // tree owns
        }
    }
}

void TypeLibTree::AddVars(const CTreeItem& item, LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    // Add variables
    for (auto i = 0u; i < pAttr->cVars; ++i) {

        LPVARDESC pVarDesc = nullptr;
        auto hr = pTypeInfo->GetVarDesc(i, &pVarDesc);
        if (FAILED(hr)) {
            continue;
        }

        UINT cNames;
        CComBSTR bstrName;
        hr = pTypeInfo->GetNames(pVarDesc->memid, &bstrName, 1, &cNames);
        if (FAILED(hr)) {
            pTypeInfo->ReleaseVarDesc(pVarDesc);
            continue;
        }

        if (cNames == 0) {
            pTypeInfo->ReleaseVarDesc(pVarDesc);
            continue;
        }

        if (pVarDesc->varkind == VAR_CONST) {
            auto desc = TYPEDESCtoString(pTypeInfo, &pVarDesc->elemdescVar.tdesc);
        }

        pTypeInfo->ReleaseVarDesc(pVarDesc);
    }
}

void TypeLibTree::ConstructChildren(const CTreeItem& item)
{
    auto pTypeInfo = reinterpret_cast<LPTYPEINFO>(item.GetData());
    if (pTypeInfo == nullptr) {
        return;
    }

    TYPEATTR* pAttr = nullptr;
    auto hr = pTypeInfo->GetTypeAttr(&pAttr);
    if (FAILED(hr)) {
        return;
    }

    AddFunctions(item, pTypeInfo, pAttr);
    AddImplTypes(item, pTypeInfo, pAttr);
    AddVars(item, pTypeInfo, pAttr);

    pTypeInfo->ReleaseTypeAttr(pAttr);
}

HTREEITEM TypeLibTree::InsertItem(LPCTSTR lpszItem, int nImage, int nSelectedImage, int nChildren, HTREEITEM hParent,
                                  HTREEITEM hInsertAfter, LPVOID lParam)
{
    TV_INSERTSTRUCT tvis;
    tvis.hParent = hParent;
    tvis.hInsertAfter = hInsertAfter;
    tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
    tvis.itemex.cChildren = nChildren;
    tvis.itemex.pszText = const_cast<LPTSTR>(lpszItem);
    tvis.itemex.iImage = nImage;
    tvis.itemex.iSelectedImage = nSelectedImage;
    tvis.itemex.lParam = reinterpret_cast<LPARAM>(lParam);

    return CTreeViewCtrl::InsertItem(&tvis);
}

void TypeLibTree::OnDestroy()
{
    DeleteAllItems();

    m_pTypeLib.Release();
}

LRESULT TypeLibTree::OnItemExpanding(LPNMHDR pnmh)
{
    CWaitCursor cursor;

    const auto item = MAKE_TREEITEM(pnmh, this);

    auto pTypeInfo = reinterpret_cast<LPTYPEINFO>(item.GetData());
    if (pTypeInfo == nullptr) {
        return 0;
    }

    auto result = GetItemState(item.m_hTreeItem, TVIS_EXPANDED);
    if (result) {
        return 0; // already expanded
    }

    auto child = item.GetChild();
    if (!child.IsNull()) {
        return 0; // already have children
    }

    ConstructChildren(item);

    SortChildren(item.m_hTreeItem);

    return 0;
}

LRESULT TypeLibTree::OnDelete(LPNMHDR pnmh)
{
    const auto item = MAKE_OLDTREEITEM(pnmh, this);

    auto* pTypeInfo = reinterpret_cast<LPTYPEINFO>(item.GetData());
    if (pTypeInfo != nullptr) {
        (void)pTypeInfo->Release();
    }

    return 0;
}
