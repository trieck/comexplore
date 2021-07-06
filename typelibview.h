#pragma once
#include "TypeLibTree.h"


class TypeLibView : public CSplitterWindow
{
public:
BEGIN_MSG_MAP_EX(TypeLibView)
        MESSAGE_HANDLER(WM_CREATE, OnCreate)
        CHAIN_MSG_MAP(CSplitterWindow)
    END_MSG_MAP()

    TypeLibView();
    LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

private:
    TypeLibTree m_tree;
};
