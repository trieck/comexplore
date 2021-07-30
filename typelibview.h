#pragma once
#include "TypeLibTree.h"
#include "IDLView.h"

class TypeLibView : public CSplitterWindow
{
public:
BEGIN_MSG_MAP_EX(TypeLibView)
        MESSAGE_HANDLER(WM_CREATE, OnCreate)
        NOTIFY_CODE_HANDLER_EX(TVN_SELCHANGED, OnTVSelChanged)
        CHAIN_MSG_MAP(CSplitterWindow)
    END_MSG_MAP()

    TypeLibView();
    LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
    LRESULT OnTVSelChanged(LPNMHDR pnmhdr);

private:
    void SelectItem(LPVOID pv);

    TypeLibTree m_tree;
    IDLView m_idlView;
};
