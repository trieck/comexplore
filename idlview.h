#pragma once
#include "typeinfonode.h"

class IDLView : public CScrollWindowImpl<IDLView>
{
public:
BEGIN_MSG_MAP(IDLView)
        MSG_WM_CREATE(OnCreate)
        MESSAGE_HANDLER(WM_SELCHANGED, OnSelChanged)
        CHAIN_MSG_MAP(CScrollWindowImpl<IDLView>)
    END_MSG_MAP()

    void DoPaint(CDCHandle dc);
    LRESULT OnCreate(LPCREATESTRUCT pcs);
    LRESULT OnSelChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);

private:
    void Decompile(LPTYPEINFONODE pNode);
    void DecompileCoClass(LPTYPEINFONODE pNode, LPTYPEATTR pAttr);
    void Update(LPTYPEINFONODE pNode);
    BOOL Write(LPCTSTR format, ...);
    BOOL WriteStream();

    CComPtr<IStream> m_pStream;
    CFont m_font;
    CDC m_memDC;
    CBitmap m_bitmap;
};
