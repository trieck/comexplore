#pragma once

#include "typeinfonode.h"
#include "objdata.h"

class TypeLibTree : public CWindowImpl<TypeLibTree, CTreeViewCtrlEx>
{
public:
    DECLARE_WND_SUPERCLASS(NULL, CTreeViewCtrl::GetWndClassName())

BEGIN_MSG_MAP_EX(TypeLibTree)
        MSG_WM_CREATE(OnCreate)
        MSG_WM_DESTROY(OnDestroy)
        REFLECTED_NOTIFY_CODE_HANDLER_EX(TVN_ITEMEXPANDING, OnItemExpanding)
        REFLECTED_NOTIFY_CODE_HANDLER_EX(TVN_DELETEITEM, OnDelete)
        DEFAULT_REFLECTION_HANDLER()
    END_MSG_MAP()

    TypeLibTree();
    LRESULT OnCreate(LPCREATESTRUCT /*pcs*/);
    void OnDestroy();
    LRESULT OnItemExpanding(LPNMHDR pnmh);
    LRESULT OnDelete(LPNMHDR pnmh);
    HRESULT GetTypeLib(ITypeLib** pTypeLib);
private:
    BOOL BuildView();
    BOOL BuildView(LPOBJECTDATA pdata);
    HTREEITEM InsertItem(LPCTSTR lpszItem, int nImage, int nSelectedImage,
                         int nChildren, HTREEITEM hParent, HTREEITEM hInsertAfter,
                         LPVOID = nullptr);
    HTREEITEM InsertItem(LPCTSTR lpszName, int nImage, int nChildren, HTREEITEM hParent,
                         TypeInfoType type, LPTYPEINFO pTypeInfo, MEMBERID memberID = MEMBERID_NIL);

    void BuildTypeInfo(HTREEITEM hParent);
    HTREEITEM AddTypeInfo(HTREEITEM hParent, LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr);
    void AddFunctions(const CTreeItem& item, LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr);
    void AddImplTypes(const CTreeItem& item, LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr);
    void AddVars(const CTreeItem& item, LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr);
    void AddAliases(const CTreeItem& item, LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr);
    void ConstructChildren(const CTreeItem& item);

    CImageList m_ImageList;
    CComPtr<ITypeLib> m_pTypeLib;
};
