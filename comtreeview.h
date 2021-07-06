#pragma once

#include "common.h"

class ComTreeView : public CWindowImpl<ComTreeView, CTreeViewCtrlEx>
{
public:
    DECLARE_WND_SUPERCLASS(NULL, CTreeViewCtrl::GetWndClassName())

BEGIN_MSG_MAP_EX(ComTreeView)
        MSG_WM_CREATE(OnCreate)
        MSG_WM_DESTROY(OnDestroy)
        REFLECTED_NOTIFY_CODE_HANDLER_EX(TVN_ITEMEXPANDING, OnItemExpanding)
        REFLECTED_NOTIFY_CODE_HANDLER_EX(TVN_DELETEITEM, OnDelete)
        DEFAULT_REFLECTION_HANDLER()
    END_MSG_MAP()

    ComTreeView();
    LRESULT OnCreate(LPCREATESTRUCT /*pcs*/);
    void OnDestroy();
    LRESULT OnItemExpanding(LPNMHDR pnmh);
    LRESULT OnDelete(LPNMHDR pnmh);

private:
    void ExpandClasses(const CTreeItem& item);
    void ExpandInterfaces(const CTreeItem& item);
    void ExpandAllInterfaces(const CTreeItem& item);
    void ExpandApps(const CTreeItem& item);
    void ExpandTypeLibs(const CTreeItem& item);
    void ConstructTypeLibs(const CTreeItem& item);
    void ConstructApps(const CTreeItem& item);
    void ConstructClasses(const CTreeItem& item);
    void ConstructInterfaces(const CTreeItem& item);
    void ConstructAllInterfaces(const CTreeItem& item);
    void ConstructInterfaces(CComPtr<IUnknown>& pUnk, const CTreeItem& item);
    void ConstructTree();

    CImageList m_ImageList;
};

/////////////////////////////////////////////////////////////////////////////
