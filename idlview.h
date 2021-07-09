#pragma once
#include "typeinfonode.h"

class IDLView : public CWindowImpl<IDLView, CStatic>
{
public:
BEGIN_MSG_MAP(IDLView)
        MSG_WM_CREATE(OnCreate)
        MESSAGE_HANDLER(WM_SELCHANGED, OnSelChanged)
    END_MSG_MAP()

    LRESULT OnCreate(LPCREATESTRUCT pcs);
    LRESULT OnSelChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);

private:
    void Decompile(LPTYPEINFONODE pNode);
    void DecompileCoClass(LPTYPEINFONODE pNode, LPTYPEATTR pAttr);
    void Update(LPTYPEINFONODE pNode);
    BOOL Write(LPCTSTR format, ...);
    BOOL FlushStream();

    CComPtr<IStream> m_pStream;
    CFont m_font;
};
