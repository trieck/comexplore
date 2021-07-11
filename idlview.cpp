#include "stdafx.h"
#include "idlview.h"

#include "autobstrarray.h"
#include "autodesc.h"
#include "autotypeattr.h"
#include "typeinfonode.h"
#include "util.h"

static constexpr auto MAX_NAMES = 64;

void IDLView::DoPaint(CDCHandle dc)
{
    if (m_memDC && m_bitmap) {
        auto hOldBitmap = m_memDC.SelectBitmap(m_bitmap);

        CRect rc;
        dc.GetClipBox(&rc);

        dc.BitBlt(rc.left, rc.top,
                  rc.Width(), rc.Height(), m_memDC,
                  rc.left, rc.top, SRCCOPY);

        m_memDC.SelectBitmap(hOldBitmap);
    }
}

LRESULT IDLView::OnCreate(LPCREATESTRUCT /*pcs*/)
{
    auto lRet = DefWindowProc();

    if (m_font) {
        m_font.DeleteObject();
    }

    m_pStream.Release();
    m_memDC.DeleteDC();
    if (m_bitmap) {
        m_bitmap.DeleteObject();
    }

    if (!m_font.CreatePointFont(100, _T("Cascadia Mono"))) {
        return -1;
    }

    SetScrollOffset(0, 0, FALSE);
    SetScrollSize({ 1, 1 });

    return lRet;
}

LRESULT IDLView::OnSelChanged(UINT, WPARAM, LPARAM lParam, BOOL& bHandled)
{
    auto pNode = reinterpret_cast<LPTYPEINFONODE>(lParam);

    Update(pNode);

    bHandled = TRUE;

    return 0;
}

void IDLView::Update(LPTYPEINFONODE pNode)
{
    m_pStream.Release();
    m_memDC.DeleteDC();
    if (m_bitmap) {
        m_bitmap.DeleteObject();
    }

    SetScrollOffset(0, 0, FALSE);
    SetScrollSize({ 1, 1 });

    if (pNode != nullptr && pNode->pTypeInfo != nullptr) {
        m_pStream = SHCreateMemStream(nullptr, 0);
        Decompile(pNode);
        WriteStream();
    }

    Invalidate();
}

void IDLView::WriteAttributes(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    auto fAttributes = FALSE;

    if (pAttr->guid != GUID_NULL) {
        CString strGUID;
        StringFromGUID2(pAttr->guid, strGUID.GetBuffer(40), 40);
        WriteAttr(fAttributes, TRUE, _T("  uuid(%.36s)"), static_cast<LPCTSTR>(strGUID) + 1);
    }

    if (pAttr->wMajorVerNum || pAttr->wMinorVerNum) {
        WriteAttr(fAttributes, TRUE, _T("  version(%d.%d)"), pAttr->wMajorVerNum, pAttr->wMinorVerNum);
    }

    CComBSTR bstrDoc;
    DWORD dwHelpID = 0;

    pTypeInfo->GetDocumentation(MEMBERID_NIL, nullptr, &bstrDoc, &dwHelpID, nullptr);
    if (bstrDoc.Length()) {
        WriteAttr(fAttributes, TRUE, _T("  helpstring(\"%s\")"), CString(bstrDoc));
    }

    if (dwHelpID > 0) {
        WriteAttr(fAttributes, TRUE, _T("  helpcontext(%#08.8x)"), dwHelpID);
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FAPPOBJECT) {
        WriteAttr(fAttributes, TRUE, _T("  appobject"));
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FCONTROL) {
        WriteAttr(fAttributes, TRUE, _T("  control"));
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FDUAL) {
        WriteAttr(fAttributes, TRUE, _T("  dual"));
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FHIDDEN) {
        WriteAttr(fAttributes, TRUE, _T("  hidden"));
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FLICENSED) {
        WriteAttr(fAttributes, TRUE, _T("  licensed"));
    }

    if (pAttr->typekind == TKIND_COCLASS && !(pAttr->wTypeFlags & TYPEFLAG_FCANCREATE)) {
        WriteAttr(fAttributes, TRUE, _T("  noncreatable"));
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FNONEXTENSIBLE) {
        WriteAttr(fAttributes, TRUE, _T("  nonextensible"));
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FOLEAUTOMATION) {
        WriteAttr(fAttributes, TRUE, _T("  oleautomation"));
    }

    if (fAttributes) {
        Write(_T("\n]\n"));
    }
}

BOOL IDLView::WriteStream()
{
    ATLASSERT(m_pStream != nullptr);

    STATSTG statstg;
    auto hr = m_pStream->Stat(&statstg, STATFLAG_NONAME);
    if (FAILED(hr)) {
        return FALSE;
    }

    if (statstg.cbSize.LowPart == 0) {
        return FALSE;
    }

    LARGE_INTEGER li{};
    hr = m_pStream->Seek(li, STREAM_SEEK_SET, nullptr);
    if (FAILED(hr)) {
        return FALSE;
    }

    CString strText;
    auto* buffer = strText.GetBuffer(int(statstg.cbSize.LowPart + 1));

    ULONG uRead;
    hr = m_pStream->Read(buffer, statstg.cbSize.LowPart, &uRead);
    if (FAILED(hr)) {
        return FALSE;
    }

    buffer[uRead / sizeof(TCHAR)] = _T('\0');
    strText.ReleaseBuffer();

    m_pStream.Release();

    CClientDC dc(*this);
    if (!m_memDC.CreateCompatibleDC(dc)) {
        return FALSE;
    }

    m_memDC.SetMapMode(dc.GetMapMode());
    auto hOldFont = m_memDC.SelectFont(m_font);

    CRect rc;
    m_memDC.DrawText(strText, -1, &rc, DT_CALCRECT);

    if (!m_bitmap.CreateCompatibleBitmap(dc, rc.Width(), rc.Height())) {
        return FALSE;
    }

    auto hOldBitmap = m_memDC.SelectBitmap(m_bitmap);

    auto clrWindow = GetSysColor(COLOR_WINDOW);
    auto clrText = RGB(0, 80, 128);

    m_memDC.FillSolidRect(rc, clrWindow);
    m_memDC.SetBkColor(clrWindow);
    m_memDC.SetTextColor(clrText);
    m_memDC.DrawText(strText, -1, &rc, 0);
    m_memDC.SelectBitmap(hOldBitmap);
    m_memDC.SelectFont(hOldFont);

    SetScrollOffset(0, 0, FALSE);
    SetScrollSize(CSize(rc.Width(), rc.Height()));

    return TRUE;
}

BOOL IDLView::WriteV(LPCTSTR format, va_list args)
{
    ATLASSERT(m_pStream != nullptr && format != nullptr);

    CString strValue;
    strValue.FormatV(format, args);

    auto hr = m_pStream->Write(strValue, strValue.GetLength() * sizeof(TCHAR), nullptr);

    return SUCCEEDED(hr);
}

BOOL IDLView::Write(LPCTSTR format, ...)
{
    va_list argList;
    va_start(argList, format);
    auto result = WriteV(format, argList);
    va_end(argList);

    return result;
}

BOOL IDLView::WriteAttr(BOOL& hasAttributes, BOOL bNewLine, LPCTSTR format, ...)
{
    auto result = TRUE;
    if (!hasAttributes) {
        result &= Write(_T("["));
        if (bNewLine) {
            result &= Write(_T("\n"));
        }
        hasAttributes = TRUE;
    } else {
        result &= Write(_T(","));

        if (bNewLine) {
            result &= Write(_T("\n"));
        } else {
            result &= Write(_T(" "));
        }
    }

    va_list argList;
    va_start(argList, format);
    result &= WriteV(format, argList);
    va_end(argList);

    return result;
}

BOOL IDLView::WriteIndent(int level)
{
    auto result = TRUE;

    while (level-- > 0) {
        result &= Write(_T("    "));
    }

    return result;
}

void IDLView::Decompile(LPTYPEINFONODE pNode)
{
    ATLASSERT(pNode != nullptr && pNode->pTypeInfo != nullptr);

    AutoTypeAttr attr(pNode->pTypeInfo);
    auto hr = attr.Get();
    if (FAILED(hr)) {
        return;
    }

    switch (attr->typekind) {
    case TKIND_ALIAS:
        break;
    case TKIND_COCLASS:
        DecompileCoClass(pNode, static_cast<LPTYPEATTR>(attr));
        break;
    case TKIND_DISPATCH:
        DecompileDispatch(pNode, static_cast<LPTYPEATTR>(attr));
        break;
    case TKIND_ENUM:
        break;
    case TKIND_INTERFACE:
        DecompileInterface(pNode, static_cast<LPTYPEATTR>(attr));
        break;
    case TKIND_RECORD:
        DecompileRecord(pNode, static_cast<LPTYPEATTR>(attr));
        break;
    case TKIND_UNION:
        break;
    default:
        break;
    }
}

void IDLView::DecompileRecord(LPTYPEINFONODE pNode, LPTYPEATTR pAttr)
{
    ATLASSERT(pNode != nullptr && pNode->pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);
    ATLASSERT(pAttr->typekind == TKIND_RECORD);

    if (pNode->memberID != MEMBERID_NIL) {
        DecompileVar(pNode->pTypeInfo, pAttr, pNode->memberID);
        return;
    }

    auto pTypeInfo(pNode->pTypeInfo);
    WriteAttributes(pTypeInfo, pAttr);

    CComBSTR bstrName;
    pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);

    Write(_T("typedef struct tag%s {\n"), bstrName);

    for (auto i = 0u; i < pAttr->cVars; ++i) {
        DecompileVar(pNode->pTypeInfo, pAttr, i, 1);
    }

    Write(_T("} %s;"), bstrName);
}

void IDLView::DecompileCoClass(LPTYPEINFONODE pNode, LPTYPEATTR pAttr)
{
    ATLASSERT(pNode != nullptr && pNode->pTypeInfo != nullptr);
    ATLASSERT(pNode->memberID == MEMBERID_NIL);
    ATLASSERT(pAttr != nullptr);

    auto pTypeInfo(pNode->pTypeInfo);
    WriteAttributes(pTypeInfo, pAttr);

    CComBSTR bstrName;
    auto hr = pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);
    if (SUCCEEDED(hr)) {
        Write(_T("coclass %s {\n"), CString(bstrName));
    }

    for (auto i = 0u; i < pAttr->cImplTypes; ++i) {
        auto impltype = 0;
        hr = pTypeInfo->GetImplTypeFlags(i, &impltype);
        if (FAILED(hr)) {
            continue;
        }

        HREFTYPE hRef;
        hr = pTypeInfo->GetRefTypeOfImplType(i, &hRef);
        if (FAILED(hr)) {
            continue;
        }

        CComPtr<ITypeInfo> pTypeInfoImpl;
        hr = pTypeInfo->GetRefTypeInfo(hRef, &pTypeInfoImpl);
        if (FAILED(hr)) {
            continue;
        }

        CComBSTR bstrDoc, bstrHelp;
        DWORD dwHelpID = 0;
        hr = pTypeInfoImpl->GetDocumentation(MEMBERID_NIL, &bstrName, &bstrDoc, &dwHelpID, &bstrHelp);
        if (FAILED(hr)) {
            continue;
        }

        AutoTypeAttr attrImpl(pTypeInfoImpl);
        hr = attrImpl.Get();
        if (FAILED(hr)) {
            continue;
        }

        WriteIndent();

        auto fAttributes = FALSE;
        if (impltype) {
            if (impltype & IMPLTYPEFLAG_FDEFAULT) {
                WriteAttr(fAttributes, FALSE, _T("default"));
            }

            if (impltype & IMPLTYPEFLAG_FSOURCE) {
                WriteAttr(fAttributes, FALSE, _T("source"));
            }

            if (impltype & IMPLTYPEFLAG_FRESTRICTED) {
                WriteAttr(fAttributes, FALSE, _T("restricted"));
            }

            if (fAttributes) {
                Write(_T("] "));
            }
        }

        if (attrImpl->typekind == TKIND_INTERFACE) {
            Write(_T("interface "));
        }

        if (attrImpl->typekind == TKIND_DISPATCH) {
            Write(_T("dispinterface "));
        }

        Write(_T("%s;\n"), bstrName);
    }

    Write(_T("};\n"));
}

void IDLView::DecompileFunc(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, MEMBERID memID, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);
    ATLASSERT(memID != MEMBERID_NIL);

    AutoDesc<FUNCDESC> funcdesc(pTypeInfo);

    auto hr = funcdesc.Get(memID);
    if (FAILED(hr)) {
        return;
    }

    if (funcdesc->wFuncFlags & (FUNCFLAG_FHIDDEN | FUNCFLAG_FNONBROWSABLE | FUNCFLAG_FRESTRICTED)) {
        if (pAttr->guid != IID_IUnknown && pAttr->guid != IID_IDispatch) {
            return;
        }
    }

    WriteIndent(level);

    auto fAttributes = FALSE;

    if (pAttr->typekind == TKIND_DISPATCH || pAttr->wTypeFlags & TYPEFLAG_FDUAL) {
        WriteAttr(fAttributes, FALSE, _T("id(0x%.8x)"), funcdesc->memid);
    } else if (pAttr->typekind == TKIND_MODULE) {
        WriteAttr(fAttributes, FALSE, _T("entry(%d)"), memID);
    }

    if (funcdesc->invkind & INVOKE_PROPERTYGET) {
        WriteAttr(fAttributes, FALSE, _T("propget"));
    }

    if (funcdesc->invkind & INVOKE_PROPERTYPUT) {
        WriteAttr(fAttributes, FALSE, _T("propput"));
    }

    if (funcdesc->invkind & INVOKE_PROPERTYPUTREF) {
        WriteAttr(fAttributes, FALSE, _T("propputref"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FRESTRICTED) {
        WriteAttr(fAttributes, FALSE, _T("restricted"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FSOURCE) {
        WriteAttr(fAttributes, FALSE, _T("source"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FBINDABLE) {
        WriteAttr(fAttributes, FALSE, _T("bindable"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FREQUESTEDIT) {
        WriteAttr(fAttributes, FALSE, _T("requestedit"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FDISPLAYBIND) {
        WriteAttr(fAttributes, FALSE, _T("displaybind"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FDEFAULTBIND) {
        WriteAttr(fAttributes, FALSE, _T("defaultbind"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FHIDDEN) {
        WriteAttr(fAttributes, FALSE, _T("hidden"));
    }

    if (funcdesc->cParamsOpt == -1) {
        WriteAttr(fAttributes, FALSE, _T("vararg")); // optional params
    }

    CComBSTR bstrDoc;
    DWORD dwHelpID = 0;
    pTypeInfo->GetDocumentation(funcdesc->memid, nullptr, &bstrDoc, &dwHelpID, nullptr);

    if (bstrDoc.Length()) {
        WriteAttr(fAttributes, FALSE, _T("helpstring(\"%s\")"), CString(bstrDoc));
        if (dwHelpID > 0) {
            WriteAttr(fAttributes, FALSE, _T(", helpcontext(%#08.8x)"), dwHelpID);
        }
    } else if (dwHelpID > 0) {
        WriteAttr(fAttributes, FALSE, _T("helpcontext(%#08.8x)"), dwHelpID);
    }

    if (fAttributes) {
        Write(_T("]\n"));
        WriteIndent(level);
    }

    // Write return type
    Write(_T("%s "), TYPEDESCtoString(pTypeInfo, &funcdesc->elemdescFunc.tdesc));

    // Calling convention
    if (!(pAttr->wTypeFlags & TYPEFLAG_FDUAL)) {
        switch (funcdesc->callconv) {
        case CC_CDECL:
            Write(_T("_cdecl "));
            break;
        case CC_PASCAL:
            Write(_T("_pascal "));
            break;
        case CC_STDCALL:
            Write(_T("_stdcall "));
            break;
        default:
            break;
        }
    }

    AutoBstrArray<MAX_NAMES> bstrNames;

    UINT cNames = 0;
    hr = pTypeInfo->GetNames(funcdesc->memid, static_cast<LPBSTR>(bstrNames), bstrNames.size(), &cNames);
    if (FAILED(hr)) {
        return;
    }

    // For property put and put reference functions, the right side of the assignment is unnamed.
    // If cMaxNames is less than is required to return all of the names of the parameters of a function,
    // then only the names of the first cMaxNames - 1 parameters are returned.
    // The names of the parameters are returned in the array in the same order that they appear elsewhere
    // in the interface (for example, the same order in the parameter array associated with the FUNCDESC enumeration).
    if (static_cast<int>(cNames) < funcdesc->cParams + 1) {
        bstrNames[cNames++] = SysAllocString(OLESTR("rhs"));
    }

    ATLASSERT(static_cast<int>(cNames) == funcdesc->cParams + 1);

    Write(_T("%s("), bstrNames[0]);

    for (auto i = 0; i < funcdesc->cParams; ++i) {
        fAttributes = FALSE;

        if (funcdesc->lprgelemdescParam[i].idldesc.wIDLFlags & IDLFLAG_FIN) {
            WriteAttr(fAttributes, FALSE, _T("in"));
        }

        if (funcdesc->lprgelemdescParam[i].idldesc.wIDLFlags & IDLFLAG_FOUT) {
            WriteAttr(fAttributes, FALSE, _T("out"));
        }

        if (funcdesc->lprgelemdescParam[i].idldesc.wIDLFlags & IDLFLAG_FLCID) {
            WriteAttr(fAttributes, FALSE, _T("lcid"));
        }

        if (funcdesc->lprgelemdescParam[i].idldesc.wIDLFlags & IDLFLAG_FRETVAL) {
            WriteAttr(fAttributes, FALSE, _T("retval"));
        }

        // If we have an optional last parameter and we're on the last paramter
        // or we are into the optional parameters...
        if ((funcdesc->cParamsOpt == -1 && i == funcdesc->cParams - 1) ||
            (i > (funcdesc->cParams - funcdesc->cParamsOpt))) {
            WriteAttr(fAttributes, FALSE, _T("optional"));
            fAttributes = TRUE;
        }

        if (fAttributes) {
            Write(_T("] "));
        }

        if ((funcdesc->lprgelemdescParam[i].tdesc.vt & 0x0FFF) == VT_CARRAY) {
            Write(_T("%s %s"),
                  TYPEDESCtoString(pTypeInfo, &funcdesc->lprgelemdescParam[i].tdesc.lpadesc->tdescElem)
                  , bstrNames[i + 1]);

            for (auto j = 0u; j < funcdesc->lprgelemdescParam[i].tdesc.lpadesc->cDims; ++j) {
                Write(_T("[%d]"), funcdesc->lprgelemdescParam[i].tdesc.lpadesc->rgbounds[j].cElements);
            }
        } else {
            Write(_T("%s %s"), TYPEDESCtoString(pTypeInfo, &funcdesc->lprgelemdescParam[i].tdesc)
                  , bstrNames[i + 1]);
        }

        if (i < funcdesc->cParams - 1) {
            Write(_T(",\n"));
            WriteIndent(level + 1);
        }
    }

    Write(_T(");\n"));
}

void IDLView::DecompileConst(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, LPVARDESC pVarDesc, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);
    ATLASSERT(pVarDesc != nullptr);
    ATLASSERT(pVarDesc->varkind == VAR_CONST);

    WriteAttributes(pTypeInfo, pAttr);
}

void IDLView::DecompileVar(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, MEMBERID memID, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);
    ATLASSERT(memID != MEMBERID_NIL);

    AutoDesc<VARDESC> vardesc(pTypeInfo);
    auto hr = vardesc.Get(memID);
    if (FAILED(hr)) {
        return;
    }

    if (vardesc->varkind == VAR_CONST) {
        DecompileConst(pTypeInfo, pAttr, static_cast<LPVARDESC>(vardesc), level);
        return;
    }

    WriteIndent(level);

    auto fAttributes = FALSE;

    if (pAttr->typekind == TKIND_DISPATCH || pAttr->wTypeFlags & TYPEFLAG_FDUAL) {
        WriteAttr(fAttributes, FALSE, _T("id(0x%.8x)"), vardesc->memid);

        if (vardesc->wVarFlags & VARFLAG_FREADONLY) {
            WriteAttr(fAttributes, FALSE, _T("readonly"));
        }
        if (vardesc->wVarFlags & VARFLAG_FSOURCE) {
            WriteAttr(fAttributes, FALSE, _T("source"));
        }
        if (vardesc->wVarFlags & VARFLAG_FBINDABLE) {
            WriteAttr(fAttributes, FALSE, _T("bindable"));
        }
        if (vardesc->wVarFlags & VARFLAG_FREQUESTEDIT) {
            WriteAttr(fAttributes, FALSE, _T("requestedit"));
        }
        if (vardesc->wVarFlags & VARFLAG_FDISPLAYBIND) {
            WriteAttr(fAttributes, FALSE, _T("displaybind"));
        }
        if (vardesc->wVarFlags & VARFLAG_FDEFAULTBIND) {
            WriteAttr(fAttributes, FALSE, _T("defaultbind"));
        }
        if (vardesc->wVarFlags & VARFLAG_FHIDDEN) {
            WriteAttr(fAttributes, FALSE, _T("hidden"));
        }
    }

    CComBSTR bstrName, bstrDoc;
    DWORD dwHelpID = 0;

    pTypeInfo->GetDocumentation(vardesc->memid, &bstrName, &bstrDoc, &dwHelpID, nullptr);
    if (!bstrName.Length()) {
        bstrName = _T("(nameless)");
    }

    if (bstrDoc.Length()) {
        WriteAttr(fAttributes, FALSE, _T("helpstring(\"%s\")"), CString(bstrDoc));
    }

    if (dwHelpID > 0) {
        WriteAttr(fAttributes, TRUE, _T("helpcontext(%#08.8x)"), dwHelpID);
    }

    if (fAttributes) {
        Write(_T("] "));
    }
        
    if ((vardesc->elemdescVar.tdesc.vt & 0x0FFF) == VT_CARRAY) {
        Write(_T("%s "), TYPEDESCtoString(pTypeInfo, &vardesc->elemdescVar.tdesc.lpadesc->tdescElem));
        Write(bstrName);

        for (auto i = 0u; i < vardesc->elemdescVar.tdesc.lpadesc->cDims; ++i) {
            Write(_T("[%d]"), vardesc->elemdescVar.tdesc.lpadesc->rgbounds[i].cElements);
        }
    } else {
        Write(_T("%s "), TYPEDESCtoString(pTypeInfo, &vardesc->elemdescVar.tdesc));
        Write(bstrName);
    }

    Write(_T(";\n"));
}

void IDLView::DecompileDispatch(LPTYPEINFONODE pNode, LPTYPEATTR pAttr)
{
    ATLASSERT(pNode != nullptr && pNode->pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    if (pNode->memberID != MEMBERID_NIL) {
        DecompileFunc(pNode->pTypeInfo, pAttr, pNode->memberID);
        return;
    }

    auto pTypeInfo(pNode->pTypeInfo);
    WriteAttributes(pTypeInfo, pAttr);

    CComBSTR bstrName;
    pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);

    Write(_T("dispinterface %s {\n"), CString(bstrName));

    Write(_T("    properties:\n"));
    for (auto i = 0u; i < pAttr->cVars; ++i) {
        DecompileVar(pTypeInfo, pAttr, i, 2);
    }

    Write(_T("    methods:\n"));
    for (auto i = 0u; i < pAttr->cFuncs; ++i) {
        DecompileFunc(pTypeInfo, pAttr, i, 2);
    }

    Write(_T("};\n"));
}

void IDLView::DecompileInterface(LPTYPEINFONODE pNode, LPTYPEATTR pAttr)
{
    ATLASSERT(pNode != nullptr && pNode->pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    if (pNode->memberID != MEMBERID_NIL) {
        DecompileFunc(pNode->pTypeInfo, pAttr, pNode->memberID);
        return;
    }

    auto pTypeInfo(pNode->pTypeInfo);
    WriteAttributes(pTypeInfo, pAttr);

    CComBSTR bstrName;
    auto hr = pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);
    if (SUCCEEDED(hr)) {
        Write(_T("interface %s"), CString(bstrName));
    }

    auto bases = 0;
    for (auto i = 0u; i < pAttr->cImplTypes; ++i) {
        HREFTYPE hRef = 0;
        hr = pTypeInfo->GetRefTypeOfImplType(i, &hRef);
        if (FAILED(hr)) {
            continue;
        }

        CComPtr<ITypeInfo> pTypeInfoImpl;
        hr = pTypeInfo->GetRefTypeInfo(hRef, &pTypeInfoImpl);
        if (FAILED(hr)) {
            continue;
        }

        hr = pTypeInfoImpl->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);
        if (FAILED(hr)) {
            continue;
        }

        if (bases++ == 0) {
            Write(_T(" : %s"), bstrName);
        } else {
            Write(_T(", %s"), bstrName);
        }
    }

    Write(_T(" {\n"));

    for (auto i = 0u; i < pAttr->cFuncs; ++i) {
        DecompileFunc(pTypeInfo, pAttr, i, 1);
    }

    Write(_T("};"));
}
