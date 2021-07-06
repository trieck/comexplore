#include "StdAfx.h"
#include "util.h"

static LPCTSTR g_rgszVT[] =
{
    _T("void"), //VT_EMPTY           = 0,   /* [V]   [P]  nothing                     */
    _T("null"), //VT_NULL            = 1,   /* [V]        SQL style Null              */
    _T("short"), //VT_I2              = 2,   /* [V][T][P]  2 byte signed int           */
    _T("long"), //VT_I4              = 3,   /* [V][T][P]  4 byte signed int           */
    _T("single"), //VT_R4              = 4,   /* [V][T][P]  4 byte real                 */
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
