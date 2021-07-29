#pragma once

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

    void AddFileMoniker(LPCTSTR pFilename, LPUNKNOWN pUnknown, REFCLSID clsid);
    BOOL IsSelectedInstance() const;
    BOOL IsGUIDSelected() const;

    void ReleaseSelectedObject();
    void CopyGUIDToClipboard();

private:
    void ExpandClasses(const CTreeItem& item);
    void ExpandInterfaces(const CTreeItem& item);
    void ExpandAllInterfaces(const CTreeItem& item);
    void ExpandApps(const CTreeItem& item);
    void ExpandTypeLibs(const CTreeItem& item);
    void ExpandCategories(const CTreeItem& item);
    void ExpandCategory(const CTreeItem& item);
    void ConstructTypeLibs(const CTreeItem& item);
    void ConstructApps(const CTreeItem& item);
    int LoadToolboxImage(const CRegKey& key);
    void ConstructClasses(const CTreeItem& item);
    void ConstructInterfaces(const CTreeItem& item);
    void ConstructAllInterfaces(const CTreeItem& item);
    void ConstructInterfaces(LPUNKNOWN pUnknown, const CTreeItem& item);
    void ConstructCategories(const CTreeItem& item);
    void ConstructCategory(const CTreeItem& item);
    void ConstructTree();

    CImageList m_ImageList;
    CTreeItem m_objectInstances;
};

/////////////////////////////////////////////////////////////////////////////
