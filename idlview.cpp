#include "stdafx.h"
#include "idlview.h"

#include "autobstrarray.h"
#include "autodesc.h"
#include "autotypeattr.h"
#include "idlstreamformatter.h"
#include "typeinfonode.h"
#include "util.h"

static constexpr auto MAX_NAMES = 64;

LRESULT IDLView::OnCreate(LPCREATESTRUCT /*pcs*/)
{
    auto lRet = DefWindowProc();

    SetModify(FALSE);

    return lRet;
}

LRESULT IDLView::OnSelChanged(UINT, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
    CComPtr<ITypeLib> pTypeLib(reinterpret_cast<ITypeLib*>(wParam));

    auto pNode = reinterpret_cast<LPTYPEINFONODE>(lParam);

    Update(pTypeLib, pNode);

    bHandled = TRUE;

    return 0;
}

void IDLView::Update(LPTYPELIB pTypeLib, LPTYPEINFONODE pNode)
{
    ATLASSERT(pTypeLib != nullptr);

    CWaitCursor cursor;

    m_stream.Reset();

    if (pNode != nullptr && pNode->pTypeInfo != nullptr) {
        Decompile(pNode, 0);
    } else {
        Decompile(pTypeLib);
    }

    FlushStream();
}

void IDLView::WriteAttributes(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, BOOL fNewLine, MEMBERID memID, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    auto fAttributes = FALSE;

    if (pAttr->guid != GUID_NULL) {
        CString strGUID;
        StringFromGUID2(pAttr->guid, strGUID.GetBuffer(40), 40);
        WriteAttr(fAttributes, fNewLine, level, _T("uuid(%.36s)"), static_cast<LPCTSTR>(strGUID) + 1);
    }

    if (pAttr->wMajorVerNum || pAttr->wMinorVerNum) {
        WriteAttr(fAttributes, fNewLine, level, _T("version(%d.%d)"), pAttr->wMajorVerNum, pAttr->wMinorVerNum);
    }

    CComBSTR bstrDoc;
    DWORD dwHelpID = 0;

    pTypeInfo->GetDocumentation(memID, nullptr, &bstrDoc, &dwHelpID, nullptr);
    if (bstrDoc.Length()) {
        WriteAttr(fAttributes, fNewLine, level, _T("helpstring(\"%s\")"), CString(bstrDoc));
    }

    if (dwHelpID > 0) {
        WriteAttr(fAttributes, fNewLine, level, _T("helpcontext(%#08.8x)"), dwHelpID);
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FAPPOBJECT) {
        WriteAttr(fAttributes, fNewLine, level, _T("appobject"));
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FCONTROL) {
        WriteAttr(fAttributes, fNewLine, level, _T("control"));
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FDUAL) {
        WriteAttr(fAttributes, fNewLine, level, _T("dual"));
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FHIDDEN) {
        WriteAttr(fAttributes, fNewLine, level, _T("hidden"));
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FLICENSED) {
        WriteAttr(fAttributes, fNewLine, level, _T("licensed"));
    }

    if (pAttr->typekind == TKIND_COCLASS && !(pAttr->wTypeFlags & TYPEFLAG_FCANCREATE)) {
        WriteAttr(fAttributes, fNewLine, level, _T("noncreatable"));
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FNONEXTENSIBLE) {
        WriteAttr(fAttributes, fNewLine, level, _T("nonextensible"));
    }

    if (pAttr->wTypeFlags & TYPEFLAG_FOLEAUTOMATION) {
        WriteAttr(fAttributes, fNewLine, level, _T("oleautomation"));
    }

    if (pAttr->typekind == TKIND_ALIAS) {
        WriteAttr(fAttributes, fNewLine, level, _T("public"));
    }

    if (pAttr->typekind == TKIND_MODULE) {
        CComBSTR bstrName;
        auto hr = pTypeInfo->GetDllEntry(MEMBERID_NIL, INVOKE_FUNC, &bstrName, nullptr, nullptr);
        if (FAILED(hr) || bstrName.Length() == 0) {
            bstrName = _T("<no entry points>");
        }

        WriteAttr(fAttributes, fNewLine, level, _T("dllname(\"%s\")"), bstrName);
    }

    if (fAttributes) {
        if (fNewLine) {
            Write(_T("\n"));
            WriteLevel(level, _T("]\n"));
        } else {
            Write(_T("] "));
        }
    }
}

static DWORD CALLBACK StreamCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG* pcb)
{
    auto** ptext = reinterpret_cast<LPCSTR*>(dwCookie);
    size_t nLength = strlen(*ptext);

    auto sz = std::min<size_t>(nLength, cb);

    memcpy(pbBuff, *ptext, sz);

    *ptext += sz;
    *pcb = static_cast<ULONG>(sz);

    return 0;
}

/////////////////////////////////////////////////////////////////////////////
BOOL IDLView::FlushStream()
{
    auto str = IDLStreamFormatter::Format(m_stream);
    LPCSTR pstr(str);

    EDITSTREAM stream{};
    stream.dwCookie = reinterpret_cast<DWORD_PTR>(&pstr);
    stream.pfnCallback = StreamCallback;

    StreamIn(SF_RTF, stream);

    if (stream.dwError != NOERROR) {
        return FALSE;
    }

    CPoint origin;
    SetScrollPos(&origin);

    m_stream.Reset();

    return TRUE;
}

BOOL IDLView::Write(LPCTSTR format, ...)
{
    va_list argList;
    va_start(argList, format);
    auto result = m_stream.WriteV(format, argList);
    va_end(argList);

    return result;
}

BOOL IDLView::WriteLevel(int level, LPCTSTR format, ...)
{
    auto result = WriteIndent(level);

    va_list argList;
    va_start(argList, format);
    result &= m_stream.WriteV(format, argList);
    va_end(argList);

    return result;
}

BOOL IDLView::WriteAttr(BOOL& hasAttributes, BOOL fNewLine, int level, LPCTSTR format, ...)
{
    auto result = TRUE;
    if (!hasAttributes) {
        if (fNewLine) {
            result &= WriteLevel(level, _T("[\n"));
        } else {
            result &= WriteLevel(level, _T("["));
        }
        hasAttributes = TRUE;
    } else {
        result &= Write(_T(","));

        if (fNewLine) {
            result &= Write(_T("\n"));
        } else {
            result &= Write(_T(" "));
        }
    }

    if (fNewLine) {
        result &= WriteIndent(level + 1);
    }

    va_list argList;
    va_start(argList, format);
    result &= m_stream.WriteV(format, argList);
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

void IDLView::Decompile(LPTYPELIB pTypeLib)
{
    ATLASSERT(pTypeLib != nullptr);

    AutoTypeLibAttr attr(pTypeLib);
    auto hr = attr.Get();
    if (FAILED(hr)) {
        return;
    }

    Write(_T("// Generated .IDL file (by COM Explorer)\n"));
    Write(_T("//\n"));
    Write(_T("// typelib filename: "));

    CComBSTR bstrFileName;
    hr = QueryPathOfRegTypeLib(attr->guid, attr->wMajorVerNum, attr->wMinorVerNum, attr->lcid,
                               &bstrFileName);
    if (FAILED(hr)) {
        bstrFileName = _T("<Unable to determine filename>");
    }

    Write(_T("%s\n"), bstrFileName);

    auto fAttributes = FALSE;

    if (attr->guid != GUID_NULL) {
        CString strGUID;
        StringFromGUID2(attr->guid, strGUID.GetBuffer(40), 40);
        WriteAttr(fAttributes, TRUE, 0, _T("uuid(%.36s)"), static_cast<LPCTSTR>(strGUID) + 1);
    }

    if (attr->wMajorVerNum || attr->wMinorVerNum) {
        WriteAttr(fAttributes, TRUE, 0, _T("version(%d.%d)"), attr->wMajorVerNum, attr->wMinorVerNum);
    }

    CComBSTR bstrName, bstrDoc, bstrHelp;
    DWORD dwHelpID = 0;

    pTypeLib->GetDocumentation(MEMBERID_NIL, &bstrName, &bstrDoc, &dwHelpID, &bstrHelp);
    if (bstrDoc.Length()) {
        WriteAttr(fAttributes, TRUE, 0, _T("helpstring(\"%s\")"), CString(bstrDoc));
    }

    if (bstrHelp.Length()) {
        WriteAttr(fAttributes, TRUE, 0,_T("helpfile(\"%s\")"), bstrHelp);
        WriteAttr(fAttributes, TRUE, 0,_T("helpcontext(%#08.8x)"), dwHelpID);
    }

    if (fAttributes) {
        Write(_T("\n]\n"));
    }

    Write(_T("library %s\n{\n"), bstrName);

    for (auto i = 0u; i < pTypeLib->GetTypeInfoCount(); ++i) {
        CComPtr<ITypeInfo> pTypeInfo;
        hr = pTypeLib->GetTypeInfo(i, &pTypeInfo);
        if (FAILED(hr)) {
            continue;
        }

        Decompile(pTypeInfo, 1);
    }

    Write(_T("};\n"));
}

void IDLView::Decompile(LPTYPEINFONODE pNode, int level)
{
    ATLASSERT(pNode != nullptr);
    ATLASSERT(pNode->pTypeInfo != nullptr);

    auto pTypeInfo(pNode->pTypeInfo);
    AutoTypeAttr attr(pTypeInfo);
    auto hr = attr.Get();
    if (FAILED(hr)) {
        return;
    }

    switch (pNode->type) {
    case TypeInfoType::T_ENUM:
        DecompileEnum(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TypeInfoType::T_RECORD:
        DecompileRecord(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TypeInfoType::T_MODULE:
        DecompileModule(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TypeInfoType::T_INTERFACE:
        DecompileInterface(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TypeInfoType::T_DISPATCH:
        DecompileDispatch(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TypeInfoType::T_COCLASS:
        DecompileCoClass(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TypeInfoType::T_ALIAS:
        DecompileAlias(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TypeInfoType::T_UNION:
        DecompileUnion(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TypeInfoType::T_FUNCTION:
        DecompileFunc(pTypeInfo, static_cast<LPTYPEATTR>(attr), pNode->memberID, level);
        break;
    case TypeInfoType::T_VAR:
        DecompileVar(pTypeInfo, static_cast<LPTYPEATTR>(attr), pNode->memberID, level);
        break;
    case TypeInfoType::T_CONST:
        DecompileConst(pTypeInfo, static_cast<LPTYPEATTR>(attr), pNode->memberID, level);
        break;
    default:
        break;
    }
}

void IDLView::Decompile(LPTYPEINFO pTypeInfo, int level)
{
    ATLASSERT(pTypeInfo != nullptr);

    AutoTypeAttr attr(pTypeInfo);
    auto hr = attr.Get();
    if (FAILED(hr)) {
        return;
    }

    switch (attr->typekind) {
    case TKIND_ALIAS:
        DecompileAlias(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TKIND_COCLASS:
        DecompileCoClass(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TKIND_DISPATCH:
        DecompileDispatch(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TKIND_ENUM:
        DecompileEnum(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TKIND_INTERFACE:
        DecompileInterface(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TKIND_MODULE:
        DecompileModule(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TKIND_RECORD:
        DecompileRecord(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    case TKIND_UNION:
        DecompileUnion(pTypeInfo, static_cast<LPTYPEATTR>(attr), level);
        break;
    default:
        break;
    }
}

void IDLView::DecompileAlias(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);
    ATLASSERT(pAttr->typekind == TKIND_ALIAS);

    WriteLevel(level, _T("typedef "));
    WriteAttributes(pTypeInfo, pAttr, FALSE, MEMBERID_NIL, 0);

    CComBSTR bstrName;
    pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);

    auto typeDesc = TYPEDESCtoString(pTypeInfo, &pAttr->tdescAlias);

    Write(_T("%s %s;\n"), typeDesc, bstrName);
}

void IDLView::DecompileRecord(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);
    ATLASSERT(pAttr->typekind == TKIND_RECORD);

    CComBSTR bstrName;
    pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);

    WriteLevel(level, _T("typedef "));
    WriteAttributes(pTypeInfo, pAttr, FALSE, MEMBERID_NIL, 0);

    Write(_T("struct tag%s {\n"), bstrName);

    for (auto i = 0u; i < pAttr->cVars; ++i) {
        DecompileVar(pTypeInfo, pAttr, i, level + 1);
    }

    WriteLevel(level, _T("} %s;\n"), bstrName);
}

void IDLView::DecompileUnion(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);
    ATLASSERT(pAttr->typekind == TKIND_UNION);

    CComBSTR bstrName;
    pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);

    WriteLevel(level, _T("typedef "));
    WriteAttributes(pTypeInfo, pAttr, FALSE, MEMBERID_NIL, 0);

    Write(_T("union tag%s {\n"), bstrName);

    for (auto i = 0u; i < pAttr->cVars; ++i) {
        DecompileVar(pTypeInfo, pAttr, i, level + 1);
    }

    WriteLevel(level, _T("} %s;\n"), bstrName);
}

void IDLView::DecompileCoClass(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    WriteAttributes(pTypeInfo, pAttr, TRUE, MEMBERID_NIL, level);

    CComBSTR bstrName;
    auto hr = pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);
    if (SUCCEEDED(hr)) {
        WriteLevel(level, _T("coclass %s {\n"), CString(bstrName));
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

        WriteIndent(level + 1);

        auto fAttributes = FALSE;
        if (impltype) {
            if (impltype & IMPLTYPEFLAG_FDEFAULT) {
                WriteAttr(fAttributes, FALSE, 0, _T("default"));
            }

            if (impltype & IMPLTYPEFLAG_FSOURCE) {
                WriteAttr(fAttributes, FALSE, 0, _T("source"));
            }

            if (impltype & IMPLTYPEFLAG_FRESTRICTED) {
                WriteAttr(fAttributes, FALSE, 0, _T("restricted"));
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

    WriteLevel(level, _T("};\n"));
}

void IDLView::DecompileEnum(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);
    ATLASSERT(pAttr->typekind == TKIND_ENUM);

    WriteLevel(level, _T("typedef "));
    WriteAttributes(pTypeInfo, pAttr, FALSE, MEMBERID_NIL, 0);

    CComBSTR bstrName;
    pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);

    Write(_T("enum %s {\n"), bstrName);

    for (auto i = 0u; i < pAttr->cVars; ++i) {
        DecompileConst(pTypeInfo, pAttr, i, level + 1);
    }

    WriteLevel(level, _T("};\n"));
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
    
    auto fAttributes = FALSE;

    if (pAttr->typekind == TKIND_DISPATCH || pAttr->wTypeFlags & TYPEFLAG_FDUAL) {
        WriteAttr(fAttributes, FALSE, level, _T("id(0x%.8x)"), funcdesc->memid);
    } else if (pAttr->typekind == TKIND_MODULE) {
        WriteAttr(fAttributes, FALSE, level, _T("entry(%d)"), memID);
    }

    if (funcdesc->invkind & INVOKE_PROPERTYGET) {
        WriteAttr(fAttributes, FALSE, level, _T("propget"));
    }

    if (funcdesc->invkind & INVOKE_PROPERTYPUT) {
        WriteAttr(fAttributes, FALSE, level, _T("propput"));
    }

    if (funcdesc->invkind & INVOKE_PROPERTYPUTREF) {
        WriteAttr(fAttributes, FALSE, level, _T("propputref"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FRESTRICTED) {
        WriteAttr(fAttributes, FALSE, level, _T("restricted"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FSOURCE) {
        WriteAttr(fAttributes, FALSE, level, _T("source"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FBINDABLE) {
        WriteAttr(fAttributes, FALSE, level, _T("bindable"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FREQUESTEDIT) {
        WriteAttr(fAttributes, FALSE, level, _T("requestedit"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FDISPLAYBIND) {
        WriteAttr(fAttributes, FALSE, level, _T("displaybind"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FDEFAULTBIND) {
        WriteAttr(fAttributes, FALSE, level, _T("defaultbind"));
    }

    if (funcdesc->wFuncFlags & FUNCFLAG_FHIDDEN) {
        WriteAttr(fAttributes, FALSE, level, _T("hidden"));
    }

    if (funcdesc->cParamsOpt == -1) {
        WriteAttr(fAttributes, FALSE, level, _T("vararg")); // optional params
    }

    CComBSTR bstrDoc;
    DWORD dwHelpID = 0;
    pTypeInfo->GetDocumentation(funcdesc->memid, nullptr, &bstrDoc, &dwHelpID, nullptr);

    if (bstrDoc.Length()) {
        WriteAttr(fAttributes, FALSE, level, _T("helpstring(\"%s\")"), CString(bstrDoc));
    }

    if (dwHelpID > 0) {
        WriteAttr(fAttributes, FALSE, level, _T("helpcontext(%#08.8x)"), dwHelpID);
    }

    if (fAttributes) {
        Write(_T("]\n"));
    }

    // Write return type
    WriteLevel(level, _T("%s "), TYPEDESCtoString(pTypeInfo, &funcdesc->elemdescFunc.tdesc));

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
            WriteAttr(fAttributes, FALSE, 0, _T("in"));
        }

        if (funcdesc->lprgelemdescParam[i].idldesc.wIDLFlags & IDLFLAG_FOUT) {
            WriteAttr(fAttributes, FALSE, 0, _T("out"));
        }

        if (funcdesc->lprgelemdescParam[i].idldesc.wIDLFlags & IDLFLAG_FLCID) {
            WriteAttr(fAttributes, FALSE, 0, _T("lcid"));
        }

        if (funcdesc->lprgelemdescParam[i].idldesc.wIDLFlags & IDLFLAG_FRETVAL) {
            WriteAttr(fAttributes, FALSE, 0, _T("retval"));
        }

        // If we have an optional last parameter and we're on the last paramter
        // or we are into the optional parameters...
        if ((funcdesc->cParamsOpt == -1 && i == funcdesc->cParams - 1) ||
            (i > (funcdesc->cParams - funcdesc->cParamsOpt))) {
            WriteAttr(fAttributes, FALSE, 0, _T("optional"));
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

void IDLView::DecompileConst(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, MEMBERID memID, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    AutoDesc<VARDESC> vardesc(pTypeInfo);
    auto hr = vardesc.Get(memID);
    if (FAILED(hr)) {
        return;
    }

    ATLASSERT(vardesc->varkind == VAR_CONST);

    WriteIndent(level);

    auto fAttributes = FALSE;

    if (pAttr->typekind == TKIND_MODULE) {
        WriteAttr(fAttributes, FALSE, 0, _T("entry(%d)"), vardesc->memid);
    }

    CComBSTR bstrName, bstrDoc;
    DWORD dwHelpID = 0;
    pTypeInfo->GetDocumentation(vardesc->memid, &bstrName, &bstrDoc, &dwHelpID, nullptr);

    if (bstrDoc.Length()) {
        WriteAttr(fAttributes, FALSE, 0, _T("helpstring(\"%s\")"), CString(bstrDoc));
    }

    if (dwHelpID > 0) {
        WriteAttr(fAttributes, FALSE, 0, _T("helpcontext(%#08.8x)"), dwHelpID);
    }

    if (fAttributes) {
        Write(_T("] "));
    }

    auto typeDesc = TYPEDESCtoString(pTypeInfo, &vardesc->elemdescVar.tdesc);

    CComVariant vtValue;
    hr = vtValue.ChangeType(VT_BSTR, vardesc->lpvarValue);
    if (FAILED(hr)) {
        if (vardesc->lpvarValue->vt == VT_ERROR || vardesc->lpvarValue->vt == VT_HRESULT) {
            vtValue.bstrVal = CComBSTR(GetScodeString(vardesc->lpvarValue->scode));
        }
    }

    auto strEscaped = Escape(CString(vtValue));

    CString strValue;
    CString strName(bstrName);
    if (V_VT(vardesc->lpvarValue) == VT_BSTR) {
        strValue.Format(_T("const %s %s = \"%s\";\n"),
                        static_cast<LPCTSTR>(typeDesc),
                        static_cast<LPCTSTR>(strName),
                        static_cast<LPCTSTR>(strEscaped));
    } else {
        strValue.Format(_T("const %s %s = %s;\n"),
                        static_cast<LPCTSTR>(typeDesc),
                        static_cast<LPCTSTR>(strName),
                        static_cast<LPCTSTR>(strEscaped));
    }

    Write(strValue);
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

    ATLASSERT(vardesc->varkind != VAR_CONST);

    WriteIndent(level);

    auto fAttributes = FALSE;

    if (pAttr->typekind == TKIND_DISPATCH || pAttr->wTypeFlags & TYPEFLAG_FDUAL) {
        WriteAttr(fAttributes, FALSE, 0, _T("id(0x%.8x)"), vardesc->memid);

        if (vardesc->wVarFlags & VARFLAG_FREADONLY) {
            WriteAttr(fAttributes, FALSE, 0, _T("readonly"));
        }
        if (vardesc->wVarFlags & VARFLAG_FSOURCE) {
            WriteAttr(fAttributes, FALSE, 0, _T("source"));
        }
        if (vardesc->wVarFlags & VARFLAG_FBINDABLE) {
            WriteAttr(fAttributes, FALSE, 0, _T("bindable"));
        }
        if (vardesc->wVarFlags & VARFLAG_FREQUESTEDIT) {
            WriteAttr(fAttributes, FALSE, 0, _T("requestedit"));
        }
        if (vardesc->wVarFlags & VARFLAG_FDISPLAYBIND) {
            WriteAttr(fAttributes, FALSE, 0, _T("displaybind"));
        }
        if (vardesc->wVarFlags & VARFLAG_FDEFAULTBIND) {
            WriteAttr(fAttributes, FALSE, 0, _T("defaultbind"));
        }
        if (vardesc->wVarFlags & VARFLAG_FHIDDEN) {
            WriteAttr(fAttributes, FALSE, 0, _T("hidden"));
        }
    }

    CComBSTR bstrName, bstrDoc;
    DWORD dwHelpID = 0;

    pTypeInfo->GetDocumentation(vardesc->memid, &bstrName, &bstrDoc, &dwHelpID, nullptr);
    if (!bstrName.Length()) {
        bstrName = _T("(nameless)");
    }

    if (bstrDoc.Length()) {
        WriteAttr(fAttributes, FALSE, 0, _T("helpstring(\"%s\")"), CString(bstrDoc));
    }

    if (dwHelpID > 0) {
        WriteAttr(fAttributes, FALSE, 0, _T("helpcontext(%#08.8x)"), dwHelpID);
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

void IDLView::DecompileDispatch(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    WriteAttributes(pTypeInfo, pAttr, TRUE, MEMBERID_NIL, level);

    CComBSTR bstrName;
    pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);

    WriteLevel(level, _T("dispinterface %s {\n"), CString(bstrName));

    if (pAttr->cVars > 0) {
        WriteLevel(level + 1, _T("properties:\n"));
    }
    for (auto i = 0u; i < pAttr->cVars; ++i) {
        DecompileVar(pTypeInfo, pAttr, i, level + 2);
    }

    if (pAttr->cFuncs > 0) {
        WriteLevel(level + 1, _T("methods:\n"));
    }
    for (auto i = 0u; i < pAttr->cFuncs; ++i) {
        DecompileFunc(pTypeInfo, pAttr, i, level + 2);
    }

    WriteLevel(level, _T("};\n"));
}

void IDLView::DecompileInterface(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    WriteAttributes(pTypeInfo, pAttr, TRUE, MEMBERID_NIL, level);

    CComBSTR bstrName;
    pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);

    WriteLevel(level, _T("interface %s"), bstrName);

    auto bases = 0;
    for (auto i = 0u; i < pAttr->cImplTypes; ++i) {
        HREFTYPE hRef = 0;
        auto hr = pTypeInfo->GetRefTypeOfImplType(i, &hRef);
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
        DecompileFunc(pTypeInfo, pAttr, i, level + 1);
    }

    WriteLevel(level, _T("};\n"));
}

void IDLView::DecompileModule(LPTYPEINFO pTypeInfo, LPTYPEATTR pAttr, int level)
{
    ATLASSERT(pTypeInfo != nullptr);
    ATLASSERT(pAttr != nullptr);

    WriteAttributes(pTypeInfo, pAttr, TRUE, MEMBERID_NIL, level);

    CComBSTR bstrName;
    pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);

    WriteLevel(level, _T("module %s {\n"), bstrName);

    for (auto i = 0u; i < pAttr->cFuncs; ++i) {
        DecompileFunc(pTypeInfo, pAttr, i, level + 1);
    }

    for (auto i = 0u; i < pAttr->cVars; ++i) {
        DecompileConst(pTypeInfo, pAttr, i, level + 1);
    }

    WriteLevel(level, _T("};\n"));
}
