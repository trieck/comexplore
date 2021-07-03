#pragma once

#include "common.h"
#include "resource.h"

#define MAKE_TREEITEM(n, t) \
    CTreeItem(((LPNMTREEVIEW)(n))->itemNew.hItem, t)
#define MAKE_OLDTREEITEM(n, t) \
    CTreeItem(((LPNMTREEVIEW)(n))->itemOld.hItem, t)

#define CLSID_NODE _T("Class IDs")

class ComTreeView : public CWindowImpl<ComTreeView, CTreeViewCtrlEx>
{
public:
    DECLARE_WND_SUPERCLASS(NULL, CTreeViewCtrl::GetWndClassName())

    ComTreeView(): m_bMsgHandled(FALSE)
    {
    }

BEGIN_MSG_MAP_EX(ComTreeView)
        MSG_WM_CREATE(OnCreate)
        MSG_WM_DESTROY(OnDestroy)
        REFLECTED_NOTIFY_CODE_HANDLER_EX(TVN_ITEMEXPANDING, OnItemExpanding)
        REFLECTED_NOTIFY_CODE_HANDLER_EX(TVN_DELETEITEM, OnDelete)
        DEFAULT_REFLECTION_HANDLER()
    END_MSG_MAP()

    LRESULT OnCreate(LPCREATESTRUCT /*pcs*/)
    {
        auto bResult = DefWindowProc();

        // Create a masked image list large enough to hold the icons.
        m_ImageList = ImageList_Create(16, 16, ILC_MASK, 1, 0);

        for (auto icon : { IDI_NODE, IDI_COCLASS }) {
            // Load the icon resources, and add the icons to the image list.
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

    void OnDestroy()
    {
        DeleteAllItems();
    }

    LRESULT OnItemExpanding(LPNMHDR pnmh)
    {
        const auto item = MAKE_TREEITEM(pnmh, this);

        auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());
        if (data && data->type == ObjectType::CLSID && IsEqualGUID(data->guid, GUID_NULL)) {
            ExpandClasses(item);
        }

        return 0;
    }

    LRESULT OnDelete(LPNMHDR pnmh)
    {
        const auto item = MAKE_OLDTREEITEM(pnmh, this);

        auto data = reinterpret_cast<LPOBJECTDATA>(item.GetData());

        delete data;

        return 0;
    }

    void ExpandClasses(const CTreeItem& item)
    {
        auto child = item.GetChild();
        if (!child.IsNull())
            return; // already expanded

        ConstructClasses(item);
        SortChildren(item.m_hTreeItem);
    }

    void ConstructClasses(const CTreeItem& item)
    {
        auto hItem = item.m_hTreeItem;

        CWaitCursor cursor;

        CRegKey key, subkey;
        auto lResult = key.Open(HKEY_LOCAL_MACHINE, _T("SOFTWARE\\Classes\\CLSID"), KEY_ENUMERATE_SUB_KEYS);
        if (lResult != ERROR_SUCCESS)
            return;

        DWORD index = 0, length = MAX_PATH + 1;
        TCHAR buff[MAX_PATH + 1];
        TCHAR val[MAX_PATH + 1];

        while ((lResult = key.EnumKey(index, buff, &length)) != ERROR_NO_MORE_ITEMS) {
            if (lResult != ERROR_SUCCESS)
                break;

            lResult = subkey.Open(key.m_hKey, buff, KEY_READ);
            if (lResult != ERROR_SUCCESS)
                break;

            length = MAX_PATH + 1;
            subkey.QueryStringValue(nullptr, val, &length);
            val[length] = '\0';

            TString value(val);
            value.Trim();

            if (_tcscmp(buff, _T("CLSID")) != 0) {
                if (value.GetLength() == 0) {
                    value = buff;
                }

                TV_INSERTSTRUCT tvis;
                tvis.hParent = item.m_hTreeItem;
                tvis.hInsertAfter = hItem;
                tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_TEXT | TVIF_PARAM;
                tvis.itemex.cChildren = 1;
                tvis.itemex.pszText = const_cast<LPTSTR>(static_cast<LPCTSTR>(value));
                tvis.itemex.iImage = 1;
                tvis.itemex.iSelectedImage = 1;

                auto pdata = std::make_unique<ObjectData>(ObjectType::CLSID, buff);
                if (pdata != nullptr && !IsEqualGUID(pdata->guid, GUID_NULL)) {
                    tvis.itemex.lParam = reinterpret_cast<LPARAM>(pdata.release());
                    hItem = CTreeViewCtrl::InsertItem(&tvis);
                }
            }

            length = MAX_PATH + 1;
            index++;
        }
    }

private:
    void ConstructTree()
    {
        TV_INSERTSTRUCT tvis;
        tvis.hParent = TVI_ROOT;
        tvis.hInsertAfter = TVI_ROOT;
        tvis.itemex.mask = TVIF_CHILDREN | TVIF_IMAGE | TVIF_TEXT | TVIF_PARAM;
        tvis.itemex.cChildren = 1;
        tvis.itemex.pszText = CLSID_NODE;
        tvis.itemex.iImage = 0;
        tvis.itemex.iSelectedImage = 0;
        tvis.itemex.lParam = reinterpret_cast<LPARAM>(new ObjectData{ ObjectType::CLSID, nullptr });

        CTreeViewCtrl::InsertItem(&tvis);
    }

    CImageList m_ImageList;
};


/////////////////////////////////////////////////////////////////////////////
