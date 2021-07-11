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
    void DecompileFunc(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, MEMBERID memID, int level = 0);
    void DecompileConst(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, LPVARDESC pVarDesc, int level = 0);
    void DecompileVar(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, MEMBERID memID, int level = 0);
    void DecompileRecord(LPTYPEINFONODE pNode, LPTYPEATTR pAttr);
    void DecompileInterface(LPTYPEINFONODE pNode, LPTYPEATTR pAttr);
    void DecompileDispatch(LPTYPEINFONODE pNode, LPTYPEATTR pAttr);
    void Update(LPTYPEINFONODE pNode);
    void WriteAttributes(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr);
    BOOL Write(LPCTSTR format, ...);
    BOOL WriteV(LPCTSTR format, va_list args);
    BOOL WriteAttr(BOOL& hasAttributes, BOOL bNewLine, LPCTSTR format, ...);
    BOOL WriteIndent(int level = 1);
    BOOL WriteStream();

    CComPtr<IStream> m_pStream;
    CFont m_font;
    CDC m_memDC;
    CBitmap m_bitmap;
};


