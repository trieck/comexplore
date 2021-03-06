#pragma once
#include "textstream.h"
#include "typeinfonode.h"

class IDLView : public CWindowImpl<IDLView, CRichEditCtrl>
{
public:
BEGIN_MSG_MAP(IDLView)
        MSG_WM_CREATE(OnCreate)
        MESSAGE_HANDLER(WM_SELCHANGED, OnSelChanged)
    END_MSG_MAP()

    DECLARE_WND_SUPERCLASS(NULL, CRichEditCtrl::GetWndClassName())

    LRESULT OnCreate(LPCREATESTRUCT pcs);
    LRESULT OnSelChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);

private:
    void Decompile(LPTYPELIB pTypeLib);
    void Decompile(LPTYPEINFONODE pNode, int level);
    void Decompile(LPTYPEINFO pTypeInfo, int level);
    void DecompileAlias(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level);
    void DecompileCoClass(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level);
    void DecompileDispatch(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level);
    void DecompileEnum(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level);
    void DecompileInterface(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level);
    void DecompileModule(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level);
    void DecompileRecord(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level);
    void DecompileUnion(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level);
    void DecompileFunc(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, MEMBERID memID, int level);
    void DecompileConst(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, MEMBERID memID, int level);
    void DecompileVar(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, MEMBERID memID, int level);

    // Helpers
    void Update(LPTYPELIB pTypeLib, LPTYPEINFONODE pNode);
    void WriteCustAttr(BOOL& hasAttributes, int level, LPCUSTDATA pCustData);
    void WriteCustAttr(BOOL& hasAttributes, int level, LPTYPELIB pTypeLib);
    void WriteCustAttr(BOOL& hasAttributes, int level, LPTYPEINFO pTypeInfo);

    void WriteAttributes(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, BOOL fNewLine,
                         MEMBERID memID, int level);
    BOOL Write(LPCTSTR format, ...);
    BOOL WriteLevel(int level, LPCTSTR format, ...);
    BOOL WriteAttr(BOOL& hasAttributes, BOOL fNewLine, int level, LPCTSTR format, ...);
    BOOL WriteIndent(int level);
    BOOL FlushStream();

    CComObjectStack<TextStream> m_stream;
};
