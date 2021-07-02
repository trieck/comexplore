// comexploreView.h : interface of the CComexploreView class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_COMEXPLOREVIEW_H__A78413B2_6F5F_4064_A07F_F14647350754__INCLUDED_)
#define AFX_COMEXPLOREVIEW_H__A78413B2_6F5F_4064_A07F_F14647350754__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#define MAKE_TREEITEM(n, t)	\
	CTreeItem(((LPNMTREEVIEW)n)->itemNew.hItem, t)
#define MAKE_OLDTREEITEM(n, t)	\
	CTreeItem(((LPNMTREEVIEW)n)->itemOld.hItem, t)

#define CLSID_NODE _T("Class IDs")

class CComexploreView : public CWindowImpl<CComexploreView, CTreeViewCtrlEx>
{
public:
    DECLARE_WND_SUPERCLASS(NULL, CTreeViewCtrl::GetWndClassName())

    BOOL PreTranslateMessage(MSG* pMsg)
    {
        return FALSE;
    }

BEGIN_MSG_MAP_EX(CComexploreView)
        MSG_WM_CREATE(OnCreate)
        MSG_OCM_NOTIFY(OnNotify)
        DEFAULT_REFLECTION_HANDLER()
    END_MSG_MAP()

    LRESULT OnCreate(LPCREATESTRUCT pcs)
    {
        LRESULT bResult = DefWindowProc();

        // Create a masked image list large enough to hold the icons.
        m_ImageList = ImageList_Create(16, 16, ILC_MASK, 1, 0);

        // Load the icon resources, and add the icons to the image list.
        HICON hIcon = LoadIcon(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDI_NODE));
        m_ImageList.AddIcon(hIcon);

        SetImageList(m_ImageList, TVSIL_NORMAL);
        CTreeViewCtrl::SetWindowLong(GWL_STYLE,
                                     CTreeViewCtrl::GetWindowLong(GWL_STYLE)
                                     | TVS_HASBUTTONS | TVS_HASLINES | TVS_FULLROWSELECT | TVS_INFOTIP
                                     | TVS_LINESATROOT | TVS_SHOWSELALWAYS);

        ConstructTree();

        SetMsgHandled(FALSE);

        return bResult;
    }

    LRESULT OnNotify(int id, LPNMHDR pnmh)
    {
        switch (pnmh->code) {
        case TVN_ITEMEXPANDING:
            OnExpanding(MAKE_TREEITEM(pnmh, this));
            break;
        case TVN_DELETEITEM:
            OnDelete(MAKE_OLDTREEITEM(pnmh, this));
            break;
        default:
            SetMsgHandled(FALSE);
        }
        return 0;
    }

    void OnExpanding(const CTreeItem& item)
    {
        auto data = reinterpret_cast<TString*>(item.GetData());
        if (data && data->Compare(CLSID_NODE) == 0) {
            ExpandClasses(item);
        }
    }

    void OnDelete(const CTreeItem& item)
    {
        auto data = reinterpret_cast<TString*>(item.GetData());
        delete data;
    }

    void ExpandClasses(const CTreeItem& item)
    {
        CTreeItem child = item.GetChild();
        if (!child.IsNull())
            return; // already expanded

        ConstructClasses(item);
        SortChildren(item.m_hTreeItem);
    }

    void ConstructClasses(const CTreeItem& item)
    {
        SetCursor(LoadCursor(nullptr, IDC_WAIT));

        HTREEITEM hItem = item.m_hTreeItem;

        CRegKey key, subkey;
        LONG lResult = key.Open(HKEY_CLASSES_ROOT, _T("CLSID"), KEY_ENUMERATE_SUB_KEYS);
        if (lResult != ERROR_SUCCESS)
            goto exit; // no access

        DWORD index = 0, length = MAX_PATH + 1;
        TCHAR buff[MAX_PATH + 1];
        TCHAR val[MAX_PATH + 1];

        while ((lResult = key.EnumKey(index, buff, &length)) != ERROR_NO_MORE_ITEMS) {
            if (lResult != ERROR_SUCCESS)
                break;

            buff[length] = '\0';

            lResult = subkey.Open(key.m_hKey, buff, KEY_READ);
            if (lResult != ERROR_SUCCESS)
                break;

            length = MAX_PATH + 1;
            subkey.QueryStringValue(nullptr, val, &length);
            val[length] = '\0';

            TString value(val);
            value.Trim();

            if (value.GetLength() > 0) {
                TV_INSERTSTRUCT tvis;
                tvis.hParent = item.m_hTreeItem;
                tvis.hInsertAfter = hItem;
                tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_TEXT | TVIF_PARAM;
                tvis.itemex.cChildren = 1;
                tvis.itemex.pszText = (LPTSTR)static_cast<LPCTSTR>(value);
                tvis.itemex.iImage = 0;
                tvis.itemex.iSelectedImage = 0;
                tvis.itemex.lParam = (LPARAM)new TString(buff);
                hItem = CTreeViewCtrl::InsertItem(&tvis);
            }

            length = MAX_PATH + 1;
            index++;
        }

    exit:
        SetCursor(LoadCursor(nullptr, IDC_ARROW));
    }

private:
    void ConstructTree()
    {
        TV_INSERTSTRUCT tvis;
        tvis.hParent = TVI_ROOT;
        tvis.hInsertAfter = TVI_ROOT;
        tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_TEXT | TVIF_PARAM;
        tvis.itemex.cChildren = 1;
        tvis.itemex.pszText = _T("Class IDs");
        tvis.itemex.iImage = 0;
        tvis.itemex.iSelectedImage = 0;
        tvis.itemex.lParam = (LPARAM)new TString(CLSID_NODE);

        CTreeViewCtrl::InsertItem(&tvis);
    }

    CImageList m_ImageList;
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_COMEXPLOREVIEW_H__A78413B2_6F5F_4064_A07F_F14647350754__INCLUDED_)
