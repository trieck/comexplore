#include "stdafx.h"


#include "autodesc.h"
#include "autotypeattr.h"
#include "resource.h"
#include "typeinfonode.h"
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
        IDI_FUNCTION,
        IDI_VARIABLE
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

HTREEITEM TypeLibTree::AddTypeInfo(HTREEITEM hParent, LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr)
{
    ATLASSERT(hParent != nullptr);
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    CComBSTR bstrName;
    auto hr = pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);
    if (FAILED(hr) || bstrName.Length() == 0) {
        return nullptr;
    }

    CString strName(bstrName), strTitle;

    int nImage;
    switch (pAttr->typekind) {
    case TKIND_COCLASS:
        strTitle.Format(_T("coclass %s"), static_cast<LPCTSTR>(strName));
        nImage = 1;
        break;
    case TKIND_INTERFACE:
        strTitle.Format(_T("interface %s"), static_cast<LPCTSTR>(strName));
        nImage = 2;
        break;
    case TKIND_DISPATCH:
        strTitle.Format(_T("dispinterface %s"), static_cast<LPCTSTR>(strName));
        nImage = 3;
        break;
    case TKIND_RECORD:
        strTitle.Format(_T("typedef struct %s"), static_cast<LPCTSTR>(strName));
        nImage = 4;
        break;
    case TKIND_ENUM:
        strTitle.Format(_T("typedef enum %s"), static_cast<LPCTSTR>(strName));
        nImage = 5;
        break;
    case TKIND_UNION:
        strTitle.Format(_T("typedef union %s"), static_cast<LPCTSTR>(strName));
        nImage = 6;
        break;
    case TKIND_MODULE:
        strTitle.Format(_T("module %s"), static_cast<LPCTSTR>(strName));
        nImage = 7;
        break;
    case TKIND_ALIAS:
        strTitle.Format(_T("typedef %s %s"),
                        static_cast<LPCTSTR>(TYPEDESCtoString(pTypeInfo, &pAttr->tdescAlias)),
                        static_cast<LPCTSTR>(strName));
        nImage = 8;
        break;
    default:
        return nullptr;
    }

    return InsertItem(strTitle, nImage, 1, hParent, pTypeInfo, MEMBERID_NIL);
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

        AutoTypeAttr attr(pTypeInfo);
        hr = attr.Get();
        if (FAILED(hr)) {
            continue;
        }

        HREFTYPE hRefType;
        hr = pTypeInfo->GetRefTypeOfImplType(-1, &hRefType);
        if (SUCCEEDED(hr)) {
            CComPtr<ITypeInfo> pTypeInfo2;
            hr = pTypeInfo->GetRefTypeInfo(hRefType, &pTypeInfo2);
            if (SUCCEEDED(hr)) {
                AutoTypeAttr attr2(pTypeInfo2);
                hr = attr2.Get();
                if (FAILED(hr)) {
                    continue;
                }

                AddTypeInfo(hParent, pTypeInfo2, static_cast<LPTYPEATTR>(attr2));
            }
        }

        AddTypeInfo(hParent, pTypeInfo, static_cast<LPTYPEATTR>(attr));
    }
}

void TypeLibTree::AddFunctions(const CTreeItem& item, LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    AutoDesc<FUNCDESC> funcdesc(pTypeInfo);

    // Add functions
    for (auto i = 0u; i < pAttr->cFuncs; ++i) {
        auto hr = funcdesc.Get(i);
        if (FAILED(hr)) {
            continue;
        }

        UINT cNames;
        CComBSTR bstrName;
        hr = pTypeInfo->GetNames(funcdesc->memid, &bstrName, 1, &cNames);
        if (FAILED(hr)) {
            continue;
        }

        if (cNames == 0) {
            continue;
        }

        InsertItem(bstrName, 9, 0, item.m_hTreeItem, pTypeInfo, funcdesc->memid);
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

        AutoTypeAttr newAttr(pTypeInfo2);
        hr = newAttr.Get();
        if (FAILED(hr)) {
            continue;
        }

        AddTypeInfo(item.m_hTreeItem, pTypeInfo2, static_cast<LPTYPEATTR>(newAttr));
    }
}

void TypeLibTree::AddVars(const CTreeItem& item, LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    AutoDesc<VARDESC> vardesc(pTypeInfo);

    // Add variables
    for (auto i = 0u; i < pAttr->cVars; ++i) {
        auto hr = vardesc.Get(i);
        if (FAILED(hr)) {
            continue;
        }

        UINT cNames;
        CComBSTR bstrName;
        hr = pTypeInfo->GetNames(vardesc->memid, &bstrName, 1, &cNames);
        if (FAILED(hr)) {
            continue;
        }

        if (cNames == 0) {
            continue;
        }

        if (vardesc->varkind == VAR_CONST) {
            // what about properties?
            auto desc = TYPEDESCtoString(pTypeInfo, &vardesc->elemdescVar.tdesc);

            CComVariant vtValue;
            hr = vtValue.ChangeType(VT_BSTR, vardesc->lpvarValue);
            if (FAILED(hr)) {
                if (vardesc->lpvarValue->vt == VT_ERROR || vardesc->lpvarValue->vt == VT_HRESULT) {
                    vtValue.bstrVal = CComBSTR(GetScodeString(vardesc->lpvarValue->scode));
                } else {
                    continue;
                }
            }

            auto strEscaped = Escape(CString(vtValue));

            CString strName(bstrName);
            CString strValue;
            if (V_VT(vardesc->lpvarValue) == VT_BSTR) {
                strValue.Format(_T("const %s %s = \"%s\""),
                                static_cast<LPCTSTR>(desc),
                                static_cast<LPCTSTR>(strName),
                                static_cast<LPCTSTR>(strEscaped));
            } else {
                strValue.Format(_T("const %s %s = %s"),
                                static_cast<LPCTSTR>(desc),
                                static_cast<LPCTSTR>(strName),
                                static_cast<LPCTSTR>(strEscaped));
            }

            InsertItem(strValue, 10, 0, item.m_hTreeItem, pTypeInfo, vardesc->memid);
        }
    }
}

void TypeLibTree::ConstructChildren(const CTreeItem& item)
{
    auto pNode = reinterpret_cast<LPTYPEINFONODE>(item.GetData());
    if (pNode == nullptr || pNode->pTypeInfo == nullptr) {
        return;
    }

    CComPtr<ITypeInfo> pTypeInfo(pNode->pTypeInfo);

    AutoTypeAttr attr(pTypeInfo);
    auto hr = attr.Get();
    if (FAILED(hr)) {
        return;
    }

    if (pNode->memberID == MEMBERID_NIL) {
        AddFunctions(item, pTypeInfo, static_cast<LPTYPEATTR>(attr));
        AddImplTypes(item, pTypeInfo, static_cast<LPTYPEATTR>(attr));
        AddVars(item, pTypeInfo, static_cast<LPTYPEATTR>(attr));
    }
}

HTREEITEM TypeLibTree::InsertItem(LPCTSTR lpszName, int nImage, int nChildren, HTREEITEM hParent, LPTYPEINFO pTypeInfo,
                                  MEMBERID memberID)
{
    auto* pNode = new TypeInfoNode(pTypeInfo, memberID);
    return InsertItem(lpszName, nImage, nImage, nChildren, hParent, TVI_LAST, pNode);
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

    auto pNode = reinterpret_cast<LPTYPEINFONODE>(item.GetData());
    if (pNode == nullptr) {
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

    auto* pNode = reinterpret_cast<LPTYPEINFONODE>(item.GetData());

    delete pNode;

    return 0;
}
