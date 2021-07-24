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
    void BuildView(LPOBJECTDATA pdata);

    CTreeItem InsertValue(HTREEITEM hParentItem, LPCTSTR keyName, LPCTSTR value, LPCTSTR data);
    CTreeItem InsertValue(HTREEITEM hParentItem, LPCTSTR keyName, LPCTSTR value, DWORD dwData);
    CTreeItem InsertValues(HKEY hKey, HTREEITEM hParentItem, LPCTSTR keyName);
    void InsertSubkeys(CRegKey& key, HTREEITEM hParentItem);
    void BuildCLSID(LPOBJECTDATA pdata);
    void BuildCLSID(LPCTSTR pCLSID);
    void BuildTypeLib(LPOBJECTDATA pdata);
    void BuildTypeLib(LPCTSTR pTypeLib);
    void BuildProgID(LPCTSTR pProgID);
    void BuildAppID(LPOBJECTDATA pdata);
    void BuildIID(LPOBJECTDATA pdata);
    void BuildCatID(LPOBJECTDATA pdata);

    CImageList m_ImageList;
};

