#pragma once

#include "objdata.h"

class RegistryView : public CWindowImpl<RegistryView, CTreeViewCtrlEx>
{
public:
BEGIN_MSG_MAP_EX(RegistryView)
        MSG_WM_CREATE(OnCreate)
    END_MSG_MAP()

    RegistryView();
    LRESULT OnCreate(LPCREATESTRUCT pcs);
private:
    BOOL BuildView(LPOBJECTDATA pdata);

    CTreeItem InsertValue(HTREEITEM hParentItem, LPCTSTR keyName, LPCTSTR value, LPCTSTR data);
    CTreeItem InsertValue(HTREEITEM hParentItem, LPCTSTR keyName, LPCTSTR value, DWORD dwData);
    CTreeItem InsertValues(HKEY hKey, HTREEITEM hParentItem, LPCTSTR keyName);
    void InsertSubkeys(CRegKey& key, HTREEITEM hParentItem);
    BOOL BuildCLSID(LPOBJECTDATA pdata);
    BOOL BuildCLSID(LPCTSTR pCLSID);
    BOOL BuildTypeLib(LPOBJECTDATA pdata);
    BOOL BuildTypeLib(LPCTSTR pTypeLib);
    BOOL BuildProgID(LPCTSTR pProgID);
    BOOL BuildAppID(LPOBJECTDATA pdata);
    BOOL BuildIID(LPOBJECTDATA pdata);
    BOOL BuildCatID(LPOBJECTDATA pdata);

    CImageList m_ImageList;
};

