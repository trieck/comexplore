#include "StdAfx.h"
#include "util.h"

static LPCTSTR g_rgszVT[] =
{
    _T("void"), //VT_EMPTY           = 0,   /* [V]   [P]  nothing                     */
    _T("null"), //VT_NULL            = 1,   /* [V]        SQL style Null              */
    _T("short"), //VT_I2              = 2,   /* [V][T][P]  2 byte signed int           */
    _T("long"), //VT_I4              = 3,   /* [V][T][P]  4 byte signed int           */
    _T("float"), //VT_R4              = 4,   /* [V][T][P]  4 byte real                 */
    _T("double"), //VT_R8              = 5,   /* [V][T][P]  8 byte real                 */
    _T("CURRENCY"), //VT_CY              = 6,   /* [V][T][P]  currency                    */
    _T("DATE"), //VT_DATE            = 7,   /* [V][T][P]  date                        */
    _T("BSTR"), //VT_BSTR            = 8,   /* [V][T][P]  binary string               */
    _T("IDispatch*"), //VT_DISPATCH        = 9,   /* [V][T]     IDispatch FAR*              */
    _T("SCODE"), //VT_ERROR           = 10,  /* [V][T]     SCODE                       */
    _T("boolean"), //VT_BOOL            = 11,  /* [V][T][P]  True=-1, False=0            */
    _T("VARIANT"), //VT_VARIANT         = 12,  /* [V][T][P]  VARIANT FAR*                */
    _T("IUnknown*"), //VT_UNKNOWN         = 13,  /* [V][T]     IUnknown FAR*               */
    _T("wchar_t"), //VT_WBSTR           = 14,  /* [V][T]     wide binary string          */
    _T(""), //                   = 15,
    _T("char"), //VT_I1              = 16,  /*    [T]     signed char                 */
    _T("unsigned char"), //VT_UI1             = 17,  /*    [T]     unsigned char               */
    _T("unsigned short"), //VT_UI2             = 18,  /*    [T]     unsigned short              */
    _T("unsigned long"), //VT_UI4             = 19,  /*    [T]     unsigned short              */
    _T("int64"), //VT_I8              = 20,  /*    [T][P]  signed 64-bit int           */
    _T("uint64"), //VT_UI8             = 21,  /*    [T]     unsigned 64-bit int         */
    _T("int"), //VT_INT             = 22,  /*    [T]     signed machine int          */
    _T("unsigned int"), //VT_UINT            = 23,  /*    [T]     unsigned machine int        */
    _T("void"), //VT_VOID            = 24,  /*    [T]     C style void                */
    _T("HRESULT"), //VT_HRESULT         = 25,  /*    [T]                                 */
    _T("PTR"), //VT_PTR             = 26,  /*    [T]     pointer type                */
    _T("SAFEARRAY"), //VT_SAFEARRAY       = 27,  /*    [T]     (use VT_ARRAY in VARIANT)   */
    _T("CARRAY"), //VT_CARRAY          = 28,  /*    [T]     C style array               */
    _T("USERDEFINED"), //VT_USERDEFINED     = 29,  /*    [T]     user defined type         */
    _T("LPSTR"), //VT_LPSTR           = 30,  /*    [T][P]  null terminated string      */
    _T("LPWSTR"), //VT_LPWSTR          = 31,  /*    [T][P]  wide null terminated string */
    _T(""), //                   = 32,
    _T(""), //                   = 33,
    _T(""), //                   = 34,
    _T(""), //                   = 35,
    _T(""), //                   = 36,
    _T(""), //                   = 37,
    _T(""), //                   = 38,
    _T(""), //                   = 39,
    _T(""), //                   = 40,
    _T(""), //                   = 41,
    _T(""), //                   = 42,
    _T(""), //                   = 43,
    _T(""), //                   = 44,
    _T(""), //                   = 45,
    _T(""), //                   = 46,
    _T(""), //                   = 47,
    _T(""), //                   = 48,
    _T(""), //                   = 49,
    _T(""), //                   = 50,
    _T(""), //                   = 51,
    _T(""), //                   = 52,
    _T(""), //                   = 53,
    _T(""), //                   = 54,
    _T(""), //                   = 55,
    _T(""), //                   = 56,
    _T(""), //                   = 57,
    _T(""), //                   = 58,
    _T(""), //                   = 59,
    _T(""), //                   = 60,
    _T(""), //                   = 61,
    _T(""), //                   = 62,
    _T(""), //                   = 63,
    _T("FILETIME"), //VT_FILETIME        = 64,  /*       [P]  FILETIME                    */
    _T("BLOB"), //VT_BLOB            = 65,  /*       [P]  Length prefixed bytes       */
    _T("STREAM"), //VT_STREAM          = 66,  /*       [P]  Name of the stream follows  */
    _T("STORAGE"), //VT_STORAGE         = 67,  /*       [P]  Name of the storage follows */
    _T("STREAMED_OBJECT"), //VT_STREAMED_OBJECT = 68,  /*       [P]  Stream contains an object   */
    _T("STORED_OBJECT"), //VT_STORED_OBJECT   = 69,  /*       [P]  Storage contains an object  */
    _T("BLOB_OBJECT"), //VT_BLOB_OBJECT     = 70,  /*       [P]  Blob contains an object     */
    _T("CF"), //VT_CF              = 71,  /*       [P]  Clipboard format            */
    _T("CLSID"), //VT_CLSID           = 72   /*       [P]  A Class ID                  */
};

/////////////////////////////////////////////////////////////////////////////
CString VTtoString(VARTYPE vt)
{
    CString str;

    vt &= ~0xF000;

    if (vt <= VT_CLSID) {
        str = g_rgszVT[vt];
    } else {
        str = _T("Unknown");
    }

    return str;
}

/////////////////////////////////////////////////////////////////////////////
CString ErrorMessage(DWORD dwError)
{
    CString output;

    LPTSTR pmsg = nullptr;

    FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
                  nullptr, dwError, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                  reinterpret_cast<LPTSTR>(&pmsg), 0, nullptr);

    if (pmsg != nullptr) {
        size_t N = _tcslen(pmsg);
        if (N > 1 && pmsg[N - 1] == '\n')
            pmsg[N - 1] = '\0';

        if (N > 1 && pmsg[N - 2] == '\r')
            pmsg[N - 2] = '\0';

        output = pmsg;

        LocalFree(pmsg);
    }

    return output;
}

/////////////////////////////////////////////////////////////////////////////
int WinErrorMsgBox(HWND hWndOwner, DWORD dwError, ATL::_U_STRINGorID title,
                   UINT uType)
{
    auto message = ErrorMessage(dwError);
    if (message.IsEmpty()) {
        message = _T("Unknown Error.");
    }

    return AtlMessageBox(hWndOwner, static_cast<LPCTSTR>(message), title,
                         uType);
}

/////////////////////////////////////////////////////////////////////////////
CString TYPEDESCtoString(LPTYPEINFO pTypeInfo, TYPEDESC* pTypeDesc)
{
    ATLASSERT(pTypeInfo != nullptr && pTypeDesc != nullptr);

    CString str(_T("Unknown"));

    if (pTypeDesc->vt == VT_PTR) {
        str.Format(_T("%s*"), static_cast<LPCTSTR>(TYPEDESCtoString(pTypeInfo, pTypeDesc->lptdesc)));
    } else if ((pTypeDesc->vt & 0xFFF) == VT_CARRAY) {
        str = TYPEDESCtoString(pTypeInfo, &pTypeDesc->lpadesc->tdescElem);

        CString strTemp;
        for (auto n = 0u; n < pTypeDesc->lpadesc->cDims; ++n) {
            strTemp.Format(_T("[%d]"), pTypeDesc->lpadesc->rgbounds[n].cElements);
            str += strTemp;
        }
    } else if ((pTypeDesc->vt & 0xFFF) == VT_SAFEARRAY) {
        str.Format(_T("SAFEARRAY(%s)"), static_cast<LPCTSTR>(TYPEDESCtoString(pTypeInfo, pTypeDesc->lptdesc)));
    } else if (pTypeDesc->vt == VT_USERDEFINED) {
        CComPtr<ITypeInfo> pTypeInfoRef;
        auto hr = pTypeInfo->GetRefTypeInfo(pTypeDesc->hreftype, &pTypeInfoRef);
        if (FAILED(hr)) {
            return str;
        }

        CComBSTR bstrName;
        hr = pTypeInfoRef->GetDocumentation(MEMBERID_NIL, &bstrName, nullptr, nullptr, nullptr);
        if (FAILED(hr)) {
            return str;
        }

        str = bstrName;
    } else {
        str = VTtoString(pTypeDesc->vt);
    }

    return str;
}

/////////////////////////////////////////////////////////////////////////////
CString Escape(LPCTSTR lpszInput)
{
    CString strOutput;

    int nLen = static_cast<int>(_tcslen(lpszInput));

    auto lpszOutput = strOutput.GetBuffer(nLen * 2 + 1);

    for (auto i = 0; i < nLen; ++i) {
        if (!_istprint(*lpszInput) || *lpszInput == _T('\"') ||
            *lpszInput == _T('\\')) {
            *lpszOutput++ = _T('\\');

            switch (*lpszInput) {
            case '\"':
                *lpszOutput++ = '\"';
                break;
            case '\\':
                *lpszOutput++ = '\\';
                break;
            case '\a':
                *lpszOutput++ = 'a';
                break;
            case '\b':
                *lpszOutput++ = 'b';
                break;
            case '\f':
                *lpszOutput++ = 'f';
                break;
            case '\n':
                *lpszOutput++ = 'n';
                break;
            case '\r':
                *lpszOutput++ = 'r';
                break;
            case '\t':
                *lpszOutput++ = 't';
                break;
            case '\v':
                *lpszOutput++ = 'v';
                break;
            case '\0':
                *lpszOutput++ = '0';
                break;
            default:
                lpszOutput += _stprintf_s(lpszOutput, nLen * 2 + 1, _T("x%02X"),
                                          static_cast<UINT>(*lpszInput));
                break;
            }

            lpszInput++;
        } else {
            *lpszOutput++ = *lpszInput++;
        }
    }

    *lpszOutput = _T('\0');

    strOutput.ReleaseBuffer();

    return strOutput;
}

/////////////////////////////////////////////////////////////////////////////
CString GetScodeString(SCODE sc)
{
    struct SCODE_ENTRY
    {
        SCODE sc;
        LPCTSTR lpszName;
    };
#define MAKE_SCODE_ENTRY(sc)    { sc, _T(#sc) }

    static const SCODE_ENTRY scNameTable[] =
    {
        MAKE_SCODE_ENTRY(S_OK),
        MAKE_SCODE_ENTRY(S_FALSE),

        MAKE_SCODE_ENTRY(CACHE_S_FORMATETC_NOTSUPPORTED),
        MAKE_SCODE_ENTRY(CACHE_S_SAMECACHE),
        MAKE_SCODE_ENTRY(CACHE_S_SOMECACHES_NOTUPDATED),
        MAKE_SCODE_ENTRY(CONVERT10_S_NO_PRESENTATION),
        MAKE_SCODE_ENTRY(DATA_S_SAMEFORMATETC),
        MAKE_SCODE_ENTRY(DRAGDROP_S_CANCEL),
        MAKE_SCODE_ENTRY(DRAGDROP_S_DROP),
        MAKE_SCODE_ENTRY(DRAGDROP_S_USEDEFAULTCURSORS),
        MAKE_SCODE_ENTRY(INPLACE_S_TRUNCATED),
        MAKE_SCODE_ENTRY(MK_S_HIM),
        MAKE_SCODE_ENTRY(MK_S_ME),
        MAKE_SCODE_ENTRY(MK_S_MONIKERALREADYREGISTERED),
        MAKE_SCODE_ENTRY(MK_S_REDUCED_TO_SELF),
        MAKE_SCODE_ENTRY(MK_S_US),
        MAKE_SCODE_ENTRY(OLE_S_MAC_CLIPFORMAT),
        MAKE_SCODE_ENTRY(OLE_S_STATIC),
        MAKE_SCODE_ENTRY(OLE_S_USEREG),
        MAKE_SCODE_ENTRY(OLEOBJ_S_CANNOT_DOVERB_NOW),
        MAKE_SCODE_ENTRY(OLEOBJ_S_INVALIDHWND),
        MAKE_SCODE_ENTRY(OLEOBJ_S_INVALIDVERB),
        MAKE_SCODE_ENTRY(OLEOBJ_S_LAST),
        MAKE_SCODE_ENTRY(STG_S_CONVERTED),
        MAKE_SCODE_ENTRY(VIEW_S_ALREADY_FROZEN),

        MAKE_SCODE_ENTRY(E_UNEXPECTED),
        MAKE_SCODE_ENTRY(E_NOTIMPL),
        MAKE_SCODE_ENTRY(E_OUTOFMEMORY),
        MAKE_SCODE_ENTRY(E_INVALIDARG),
        MAKE_SCODE_ENTRY(E_NOINTERFACE),
        MAKE_SCODE_ENTRY(E_POINTER),
        MAKE_SCODE_ENTRY(E_HANDLE),
        MAKE_SCODE_ENTRY(E_ABORT),
        MAKE_SCODE_ENTRY(E_FAIL),
        MAKE_SCODE_ENTRY(E_ACCESSDENIED),

        MAKE_SCODE_ENTRY(CACHE_E_NOCACHE_UPDATED),
        MAKE_SCODE_ENTRY(CLASS_E_CLASSNOTAVAILABLE),
        MAKE_SCODE_ENTRY(CLASS_E_NOAGGREGATION),
        MAKE_SCODE_ENTRY(CLIPBRD_E_BAD_DATA),
        MAKE_SCODE_ENTRY(CLIPBRD_E_CANT_CLOSE),
        MAKE_SCODE_ENTRY(CLIPBRD_E_CANT_EMPTY),
        MAKE_SCODE_ENTRY(CLIPBRD_E_CANT_OPEN),
        MAKE_SCODE_ENTRY(CLIPBRD_E_CANT_SET),
        MAKE_SCODE_ENTRY(CO_E_ALREADYINITIALIZED),
        MAKE_SCODE_ENTRY(CO_E_APPDIDNTREG),
        MAKE_SCODE_ENTRY(CO_E_APPNOTFOUND),
        MAKE_SCODE_ENTRY(CO_E_APPSINGLEUSE),
        MAKE_SCODE_ENTRY(CO_E_BAD_PATH),
        MAKE_SCODE_ENTRY(CO_E_CANTDETERMINECLASS),
        MAKE_SCODE_ENTRY(CO_E_CLASS_CREATE_FAILED),
        MAKE_SCODE_ENTRY(CO_E_CLASSSTRING),
        MAKE_SCODE_ENTRY(CO_E_DLLNOTFOUND),
        MAKE_SCODE_ENTRY(CO_E_ERRORINAPP),
        MAKE_SCODE_ENTRY(CO_E_ERRORINDLL),
        MAKE_SCODE_ENTRY(CO_E_IIDSTRING),
        MAKE_SCODE_ENTRY(CO_E_NOTINITIALIZED),
        MAKE_SCODE_ENTRY(CO_E_OBJISREG),
        MAKE_SCODE_ENTRY(CO_E_OBJNOTCONNECTED),
        MAKE_SCODE_ENTRY(CO_E_OBJNOTREG),
        MAKE_SCODE_ENTRY(CO_E_OBJSRV_RPC_FAILURE),
        MAKE_SCODE_ENTRY(CO_E_SCM_ERROR),
        MAKE_SCODE_ENTRY(CO_E_SCM_RPC_FAILURE),
        MAKE_SCODE_ENTRY(CO_E_SERVER_EXEC_FAILURE),
        MAKE_SCODE_ENTRY(CO_E_SERVER_STOPPING),
        MAKE_SCODE_ENTRY(CO_E_WRONGOSFORAPP),
        MAKE_SCODE_ENTRY(CONVERT10_E_OLESTREAM_BITMAP_TO_DIB),
        MAKE_SCODE_ENTRY(CONVERT10_E_OLESTREAM_FMT),
        MAKE_SCODE_ENTRY(CONVERT10_E_OLESTREAM_GET),
        MAKE_SCODE_ENTRY(CONVERT10_E_OLESTREAM_PUT),
        MAKE_SCODE_ENTRY(CONVERT10_E_STG_DIB_TO_BITMAP),
        MAKE_SCODE_ENTRY(CONVERT10_E_STG_FMT),
        MAKE_SCODE_ENTRY(CONVERT10_E_STG_NO_STD_STREAM),
        MAKE_SCODE_ENTRY(DISP_E_ARRAYISLOCKED),
        MAKE_SCODE_ENTRY(DISP_E_BADCALLEE),
        MAKE_SCODE_ENTRY(DISP_E_BADINDEX),
        MAKE_SCODE_ENTRY(DISP_E_BADPARAMCOUNT),
        MAKE_SCODE_ENTRY(DISP_E_BADVARTYPE),
        MAKE_SCODE_ENTRY(DISP_E_EXCEPTION),
        MAKE_SCODE_ENTRY(DISP_E_MEMBERNOTFOUND),
        MAKE_SCODE_ENTRY(DISP_E_NONAMEDARGS),
        MAKE_SCODE_ENTRY(DISP_E_NOTACOLLECTION),
        MAKE_SCODE_ENTRY(DISP_E_OVERFLOW),
        MAKE_SCODE_ENTRY(DISP_E_PARAMNOTFOUND),
        MAKE_SCODE_ENTRY(DISP_E_PARAMNOTOPTIONAL),
        MAKE_SCODE_ENTRY(DISP_E_TYPEMISMATCH),
        MAKE_SCODE_ENTRY(DISP_E_UNKNOWNINTERFACE),
        MAKE_SCODE_ENTRY(DISP_E_UNKNOWNLCID),
        MAKE_SCODE_ENTRY(DISP_E_UNKNOWNNAME),
        MAKE_SCODE_ENTRY(DRAGDROP_E_ALREADYREGISTERED),
        MAKE_SCODE_ENTRY(DRAGDROP_E_INVALIDHWND),
        MAKE_SCODE_ENTRY(DRAGDROP_E_NOTREGISTERED),
        MAKE_SCODE_ENTRY(DV_E_CLIPFORMAT),
        MAKE_SCODE_ENTRY(DV_E_DVASPECT),
        MAKE_SCODE_ENTRY(DV_E_DVTARGETDEVICE),
        MAKE_SCODE_ENTRY(DV_E_DVTARGETDEVICE_SIZE),
        MAKE_SCODE_ENTRY(DV_E_FORMATETC),
        MAKE_SCODE_ENTRY(DV_E_LINDEX),
        MAKE_SCODE_ENTRY(DV_E_NOIVIEWOBJECT),
        MAKE_SCODE_ENTRY(DV_E_STATDATA),
        MAKE_SCODE_ENTRY(DV_E_STGMEDIUM),
        MAKE_SCODE_ENTRY(DV_E_TYMED),
        MAKE_SCODE_ENTRY(INPLACE_E_NOTOOLSPACE),
        MAKE_SCODE_ENTRY(INPLACE_E_NOTUNDOABLE),
        MAKE_SCODE_ENTRY(MEM_E_INVALID_LINK),
        MAKE_SCODE_ENTRY(MEM_E_INVALID_ROOT),
        MAKE_SCODE_ENTRY(MEM_E_INVALID_SIZE),
        MAKE_SCODE_ENTRY(MK_E_CANTOPENFILE),
        MAKE_SCODE_ENTRY(MK_E_CONNECTMANUALLY),
        MAKE_SCODE_ENTRY(MK_E_ENUMERATION_FAILED),
        MAKE_SCODE_ENTRY(MK_E_EXCEEDEDDEADLINE),
        MAKE_SCODE_ENTRY(MK_E_INTERMEDIATEINTERFACENOTSUPPORTED),
        MAKE_SCODE_ENTRY(MK_E_INVALIDEXTENSION),
        MAKE_SCODE_ENTRY(MK_E_MUSTBOTHERUSER),
        MAKE_SCODE_ENTRY(MK_E_NEEDGENERIC),
        MAKE_SCODE_ENTRY(MK_E_NO_NORMALIZED),
        MAKE_SCODE_ENTRY(MK_E_NOINVERSE),
        MAKE_SCODE_ENTRY(MK_E_NOOBJECT),
        MAKE_SCODE_ENTRY(MK_E_NOPREFIX),
        MAKE_SCODE_ENTRY(MK_E_NOSTORAGE),
        MAKE_SCODE_ENTRY(MK_E_NOTBINDABLE),
        MAKE_SCODE_ENTRY(MK_E_NOTBOUND),
        MAKE_SCODE_ENTRY(MK_E_SYNTAX),
        MAKE_SCODE_ENTRY(MK_E_UNAVAILABLE),
        MAKE_SCODE_ENTRY(OLE_E_ADVF),
        MAKE_SCODE_ENTRY(OLE_E_ADVISENOTSUPPORTED),
        MAKE_SCODE_ENTRY(OLE_E_BLANK),
        MAKE_SCODE_ENTRY(OLE_E_CANT_BINDTOSOURCE),
        MAKE_SCODE_ENTRY(OLE_E_CANT_GETMONIKER),
        MAKE_SCODE_ENTRY(OLE_E_CANTCONVERT),
        MAKE_SCODE_ENTRY(OLE_E_CLASSDIFF),
        MAKE_SCODE_ENTRY(OLE_E_ENUM_NOMORE),
        MAKE_SCODE_ENTRY(OLE_E_INVALIDHWND),
        MAKE_SCODE_ENTRY(OLE_E_INVALIDRECT),
        MAKE_SCODE_ENTRY(OLE_E_NOCACHE),
        MAKE_SCODE_ENTRY(OLE_E_NOCONNECTION),
        MAKE_SCODE_ENTRY(OLE_E_NOSTORAGE),
        MAKE_SCODE_ENTRY(OLE_E_NOT_INPLACEACTIVE),
        MAKE_SCODE_ENTRY(OLE_E_NOTRUNNING),
        MAKE_SCODE_ENTRY(OLE_E_OLEVERB),
        MAKE_SCODE_ENTRY(OLE_E_PROMPTSAVECANCELLED),
        MAKE_SCODE_ENTRY(OLE_E_STATIC),
        MAKE_SCODE_ENTRY(OLE_E_WRONGCOMPOBJ),
        MAKE_SCODE_ENTRY(OLEOBJ_E_INVALIDVERB),
        MAKE_SCODE_ENTRY(OLEOBJ_E_NOVERBS),
        MAKE_SCODE_ENTRY(REGDB_E_CLASSNOTREG),
        MAKE_SCODE_ENTRY(REGDB_E_IIDNOTREG),
        MAKE_SCODE_ENTRY(REGDB_E_INVALIDVALUE),
        MAKE_SCODE_ENTRY(REGDB_E_KEYMISSING),
        MAKE_SCODE_ENTRY(REGDB_E_READREGDB),
        MAKE_SCODE_ENTRY(REGDB_E_WRITEREGDB),
        MAKE_SCODE_ENTRY(RPC_E_ATTEMPTED_MULTITHREAD),
        MAKE_SCODE_ENTRY(RPC_E_CALL_CANCELED),
        MAKE_SCODE_ENTRY(RPC_E_CALL_REJECTED),
        MAKE_SCODE_ENTRY(RPC_E_CANTCALLOUT_AGAIN),
        MAKE_SCODE_ENTRY(RPC_E_CANTCALLOUT_INASYNCCALL),
        MAKE_SCODE_ENTRY(RPC_E_CANTCALLOUT_INEXTERNALCALL),
        MAKE_SCODE_ENTRY(RPC_E_CANTCALLOUT_ININPUTSYNCCALL),
        MAKE_SCODE_ENTRY(RPC_E_CANTPOST_INSENDCALL),
        MAKE_SCODE_ENTRY(RPC_E_CANTTRANSMIT_CALL),
        MAKE_SCODE_ENTRY(RPC_E_CHANGED_MODE),
        MAKE_SCODE_ENTRY(RPC_E_CLIENT_CANTMARSHAL_DATA),
        MAKE_SCODE_ENTRY(RPC_E_CLIENT_CANTUNMARSHAL_DATA),
        MAKE_SCODE_ENTRY(RPC_E_CLIENT_DIED),
        MAKE_SCODE_ENTRY(RPC_E_CONNECTION_TERMINATED),
        MAKE_SCODE_ENTRY(RPC_E_DISCONNECTED),
        MAKE_SCODE_ENTRY(RPC_E_FAULT),
        MAKE_SCODE_ENTRY(RPC_E_INVALID_CALLDATA),
        MAKE_SCODE_ENTRY(RPC_E_INVALID_DATA),
        MAKE_SCODE_ENTRY(RPC_E_INVALID_DATAPACKET),
        MAKE_SCODE_ENTRY(RPC_E_INVALID_PARAMETER),
        MAKE_SCODE_ENTRY(RPC_E_INVALIDMETHOD),
        MAKE_SCODE_ENTRY(RPC_E_NOT_REGISTERED),
        MAKE_SCODE_ENTRY(RPC_E_OUT_OF_RESOURCES),
        MAKE_SCODE_ENTRY(RPC_E_RETRY),
        MAKE_SCODE_ENTRY(RPC_E_SERVER_CANTMARSHAL_DATA),
        MAKE_SCODE_ENTRY(RPC_E_SERVER_CANTUNMARSHAL_DATA),
        MAKE_SCODE_ENTRY(RPC_E_SERVER_DIED),
        MAKE_SCODE_ENTRY(RPC_E_SERVER_DIED_DNE),
        MAKE_SCODE_ENTRY(RPC_E_SERVERCALL_REJECTED),
        MAKE_SCODE_ENTRY(RPC_E_SERVERCALL_RETRYLATER),
        MAKE_SCODE_ENTRY(RPC_E_SERVERFAULT),
        MAKE_SCODE_ENTRY(RPC_E_SYS_CALL_FAILED),
        MAKE_SCODE_ENTRY(RPC_E_THREAD_NOT_INIT),
        MAKE_SCODE_ENTRY(RPC_E_UNEXPECTED),
        MAKE_SCODE_ENTRY(RPC_E_WRONG_THREAD),
        MAKE_SCODE_ENTRY(STG_E_ABNORMALAPIEXIT),
        MAKE_SCODE_ENTRY(STG_E_ACCESSDENIED),
        MAKE_SCODE_ENTRY(STG_E_CANTSAVE),
        MAKE_SCODE_ENTRY(STG_E_DISKISWRITEPROTECTED),
        MAKE_SCODE_ENTRY(STG_E_EXTANTMARSHALLINGS),
        MAKE_SCODE_ENTRY(STG_E_FILEALREADYEXISTS),
        MAKE_SCODE_ENTRY(STG_E_FILENOTFOUND),
        MAKE_SCODE_ENTRY(STG_E_INSUFFICIENTMEMORY),
        MAKE_SCODE_ENTRY(STG_E_INUSE),
        MAKE_SCODE_ENTRY(STG_E_INVALIDFLAG),
        MAKE_SCODE_ENTRY(STG_E_INVALIDFUNCTION),
        MAKE_SCODE_ENTRY(STG_E_INVALIDHANDLE),
        MAKE_SCODE_ENTRY(STG_E_INVALIDHEADER),
        MAKE_SCODE_ENTRY(STG_E_INVALIDNAME),
        MAKE_SCODE_ENTRY(STG_E_INVALIDPARAMETER),
        MAKE_SCODE_ENTRY(STG_E_INVALIDPOINTER),
        MAKE_SCODE_ENTRY(STG_E_LOCKVIOLATION),
        MAKE_SCODE_ENTRY(STG_E_MEDIUMFULL),
        MAKE_SCODE_ENTRY(STG_E_NOMOREFILES),
        MAKE_SCODE_ENTRY(STG_E_NOTCURRENT),
        MAKE_SCODE_ENTRY(STG_E_NOTFILEBASEDSTORAGE),
        MAKE_SCODE_ENTRY(STG_E_OLDDLL),
        MAKE_SCODE_ENTRY(STG_E_OLDFORMAT),
        MAKE_SCODE_ENTRY(STG_E_PATHNOTFOUND),
        MAKE_SCODE_ENTRY(STG_E_READFAULT),
        MAKE_SCODE_ENTRY(STG_E_REVERTED),
        MAKE_SCODE_ENTRY(STG_E_SEEKERROR),
        MAKE_SCODE_ENTRY(STG_E_SHAREREQUIRED),
        MAKE_SCODE_ENTRY(STG_E_SHAREVIOLATION),
        MAKE_SCODE_ENTRY(STG_E_TOOMANYOPENFILES),
        MAKE_SCODE_ENTRY(STG_E_UNIMPLEMENTEDFUNCTION),
        MAKE_SCODE_ENTRY(STG_E_UNKNOWN),
        MAKE_SCODE_ENTRY(STG_E_WRITEFAULT),
        MAKE_SCODE_ENTRY(TYPE_E_AMBIGUOUSNAME),
        MAKE_SCODE_ENTRY(TYPE_E_BADMODULEKIND),
        MAKE_SCODE_ENTRY(TYPE_E_BUFFERTOOSMALL),
        MAKE_SCODE_ENTRY(TYPE_E_CANTCREATETMPFILE),
        MAKE_SCODE_ENTRY(TYPE_E_CANTLOADLIBRARY),
        MAKE_SCODE_ENTRY(TYPE_E_CIRCULARTYPE),
        MAKE_SCODE_ENTRY(TYPE_E_DLLFUNCTIONNOTFOUND),
        MAKE_SCODE_ENTRY(TYPE_E_DUPLICATEID),
        MAKE_SCODE_ENTRY(TYPE_E_ELEMENTNOTFOUND),
        MAKE_SCODE_ENTRY(TYPE_E_INCONSISTENTPROPFUNCS),
        MAKE_SCODE_ENTRY(TYPE_E_INVALIDSTATE),
        MAKE_SCODE_ENTRY(TYPE_E_INVDATAREAD),
        MAKE_SCODE_ENTRY(TYPE_E_IOERROR),
        MAKE_SCODE_ENTRY(TYPE_E_LIBNOTREGISTERED),
        MAKE_SCODE_ENTRY(TYPE_E_NAMECONFLICT),
        MAKE_SCODE_ENTRY(TYPE_E_OUTOFBOUNDS),
        MAKE_SCODE_ENTRY(TYPE_E_QUALIFIEDNAMEDISALLOWED),
        MAKE_SCODE_ENTRY(TYPE_E_REGISTRYACCESS),
        MAKE_SCODE_ENTRY(TYPE_E_SIZETOOBIG),
        MAKE_SCODE_ENTRY(TYPE_E_TYPEMISMATCH),
        MAKE_SCODE_ENTRY(TYPE_E_UNDEFINEDTYPE),
        MAKE_SCODE_ENTRY(TYPE_E_UNKNOWNLCID),
        MAKE_SCODE_ENTRY(TYPE_E_UNSUPFORMAT),
        MAKE_SCODE_ENTRY(TYPE_E_WRONGTYPEKIND),
        MAKE_SCODE_ENTRY(VIEW_E_DRAW),
    };
#undef MAKE_SCODE_ENTRY

    // look for it in the table
    for (auto i : scNameTable) {
        if (sc == i.sc) {
            return i.lpszName;
        }
    }

    TCHAR szError[32];
    _stprintf_s(szError, 32, _T("%#8.8x"), sc);

    return szError; // not found
}
