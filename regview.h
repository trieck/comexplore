#pragma once

class RegistryView : public CWindowImpl<ComTreeView, CTreeViewCtrlEx>
{
public:
BEGIN_MSG_MAP_EX(RegistryView)
        MSG_WM_CREATE(OnCreate)
        MESSAGE_HANDLER(WM_SELCHANGED, OnSelChanged)
    END_MSG_MAP()

    DECLARE_WND_SUPERCLASS(_T("RegistryView"), CTreeViewCtrlEx::GetWndClassName())

    RegistryView(): m_bMsgHandled(0)
    {
    }

    LRESULT OnCreate(LPCREATESTRUCT /*pcs*/)
    {
        auto bResult = DefWindowProc();

        m_ImageList = ImageList_Create(16, 16, ILC_MASK | ILC_COLOR32, 1, 0);

        for (auto icon : { IDI_REGNODE }) {
            auto hIcon = LoadIcon(_Module.GetResourceInstance(), MAKEINTRESOURCE(icon));
            ATLASSERT(hIcon);
            m_ImageList.AddIcon(hIcon);
        }

        SetImageList(m_ImageList, TVSIL_NORMAL);
        CTreeViewCtrl::SetWindowLong(GWL_STYLE,
                                     CTreeViewCtrl::GetWindowLong(GWL_STYLE)
                                     | TVS_HASBUTTONS | TVS_HASLINES | TVS_FULLROWSELECT | TVS_INFOTIP
                                     | TVS_LINESATROOT | TVS_SHOWSELALWAYS);

        SetMsgHandled(FALSE);

        return bResult;
    }

    LRESULT OnSelChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
    {
        auto pdata = reinterpret_cast<LPOBJECTDATA>(lParam);
        if (pdata != nullptr && pdata->guid != GUID_NULL) {
            BuildTree(pdata);
        }

        return 0;
    }

private:
    CTreeItem InsertValue(HTREEITEM hParentItem, LPCTSTR keyName, LPCTSTR value, LPCTSTR data)
    {
        TString strItem;

        if (value[0] == _T('\0')) {
            strItem.Format(_T("%s[<default>] = %s"), keyName, data);
        } else {
            strItem.Format(_T("%s[%s] = %s"), keyName, value, data);
        }

        return InsertItem(strItem, hParentItem, TVI_LAST);
    }

    CTreeItem InsertValue(HTREEITEM hParentItem, LPCTSTR keyName, LPCTSTR value, DWORD dwData)
    {
        TString strItem;

        if (value[0] == _T('\0')) {
            strItem.Format(_T("%s[<default>] = %ld"), keyName, dwData);
        } else {
            strItem.Format(_T("%s[%s] = %ld"), keyName, value, dwData);
        }

        return InsertItem(strItem, hParentItem, TVI_LAST);
    }

    CTreeItem InsertValues(HKEY hKey, HTREEITEM hParentItem, LPCTSTR keyName)
    {
        DWORD index = 0, length = MAX_PATH + 1, type;
        TCHAR val[MAX_PATH + 1];
        BYTE data[4096]{};
        DWORD dwSize = sizeof(data);
        CTreeItem item;

        while (RegEnumValue(hKey, index++, val, &length, nullptr,
                            &type, data, &dwSize) == ERROR_SUCCESS) {
            val[length] = _T('\0');

            if (type == REG_SZ || type == REG_MULTI_SZ || type == REG_EXPAND_SZ) {
                auto lpszData = reinterpret_cast<LPTSTR>(data);
                lpszData[dwSize] = _T('\0');
                item = InsertValue(hParentItem, keyName, val, lpszData);
            } else if (type == REG_DWORD) {
                DWORD dwData = (data[0] << 24) | (data[1] << 16) | (data[2] << 8) || data[3];
                item = InsertValue(hParentItem, keyName, val, dwData);
            }

            length = MAX_PATH + 1;
            dwSize = sizeof(data);
        }

        if (!item.m_hTreeItem) {
            item = InsertItem(keyName, hParentItem, TVI_LAST);
        }

        return item;
    }

    void InsertSubkeys(CRegKey& key, HTREEITEM hParentItem)
    {
        DWORD index = 0, length = MAX_PATH + 1;
        TCHAR strSubKey[MAX_PATH + 1];

        CRegKey subKey;
        LONG lResult;
        while ((lResult = key.EnumKey(index++, strSubKey, &length)) != ERROR_NO_MORE_ITEMS) {
            if (lResult != ERROR_SUCCESS) {
                break;
            }

            lResult = subKey.Open(key.m_hKey, strSubKey, KEY_READ);
            if (lResult != ERROR_SUCCESS) {
                break;
            }

            auto item = InsertValues(subKey, hParentItem, strSubKey);

            InsertSubkeys(subKey, item.m_hTreeItem);

            item.Expand();

            length = MAX_PATH + 1;
        }
    }

    void BuildTree(LPOBJECTDATA pdata)
    {
        ATLASSERT(pdata);

        this->DeleteAllItems();

        TString strGUID;
        StringFromGUID2(pdata->guid, strGUID.GetBuffer(40), 40);

        TString strPath;
        strPath.Format(_T("SOFTWARE\\Classes\\CLSID\\%s"), static_cast<LPCTSTR>(strGUID));

        CWaitCursor cursor;
        CRegKey key, subkey;
        auto lResult = key.Open(HKEY_LOCAL_MACHINE, strPath, KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE);
        if (lResult != ERROR_SUCCESS) {
            return;
        }

        auto root = InsertItem(_T("CLSID="), TVI_ROOT, TVI_LAST);
        auto guid = InsertValues(key, root.m_hTreeItem, strGUID);

        InsertSubkeys(key, guid.m_hTreeItem);

        root.Expand();
        guid.Expand();
    }

    CImageList m_ImageList;
};
