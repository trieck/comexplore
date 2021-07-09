#include "stdafx.h"
#include "idlview.h"


#include "autotypeattr.h"
#include "typeinfonode.h"

LRESULT IDLView::OnCreate(LPCREATESTRUCT pcs)
{
    auto lRet = DefWindowProc();

    if (m_font.m_hFont != nullptr) {
        m_font.DeleteObject();
    }

    if (!m_font.CreatePointFont(100, _T("Cascadia Mono"))) {
        return -1;
    }

    SetFont(m_font);

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

    SetWindowText(_T(""));

    if (pNode == nullptr || pNode->pTypeInfo == nullptr) {
        return;
    }

    m_pStream = SHCreateMemStream(nullptr, 0);
    if (m_pStream == nullptr) {
        return;
    }

    Decompile(pNode);

    FlushStream();
}

BOOL IDLView::FlushStream()
{
    ATLASSERT(m_pStream != nullptr);

    STATSTG statstg;
    auto hr = m_pStream->Stat(&statstg, STATFLAG_NONAME);
    if (FAILED(hr)) {
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

    SetWindowText(strText);

    m_pStream.Release();

    return TRUE;
}

BOOL IDLView::Write(LPCTSTR format, ...)
{
    ATLASSERT(m_pStream != nullptr && format != nullptr);

    CString strValue;

    va_list argList;
    va_start(argList, format);
    strValue.FormatV(format, argList);
    va_end(argList);

    auto hr = m_pStream->Write(strValue, strValue.GetLength() * sizeof(TCHAR), nullptr);

    return SUCCEEDED(hr);
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
    case TKIND_COCLASS:
        DecompileCoClass(pNode, static_cast<LPTYPEATTR>(attr));
        break;
    default:
        break;
    }
}

void IDLView::DecompileCoClass(LPTYPEINFONODE pNode, LPTYPEATTR pAttr)
{
    ATLASSERT(pNode != nullptr && pNode->pTypeInfo != nullptr);
    ATLASSERT(pNode->memberID == MEMBERID_NIL);
    ATLASSERT(pAttr != nullptr);

    CString strGUID;
    StringFromGUID2(pAttr->guid, strGUID.GetBuffer(40), 40);

    Write(_T("[\nuuid(%s)"), strGUID);

    if (pAttr->wMajorVerNum || pAttr->wMinorVerNum) {
        Write(_T(",\n  version(%d.%d)"), pAttr->wMajorVerNum, pAttr->wMinorVerNum);
    }

    CComBSTR bstrName, bstrDoc, bstrHelp;
    DWORD dwHelpID = 0;

    auto pTypeInfo(pNode->pTypeInfo);
    pTypeInfo->GetDocumentation(MEMBERID_NIL, &bstrName, &bstrDoc, &dwHelpID, &bstrHelp);
    if (bstrDoc.Length()) {
        Write(_T(",\n  helpstring(\"%s\")"), CString(bstrDoc));
    }

    if (dwHelpID > 0) {
        Write(_T(",\n  helpcontext(%#08.8x)"), dwHelpID);
    }

    if (pAttr->wTypeFlags == TYPEFLAG_FAPPOBJECT) {
        Write(_T(",\n  appobject"));
    }
    if (pAttr->wTypeFlags == TYPEFLAG_FHIDDEN) {
        Write(_T(",  hidden"));
    }
    if (pAttr->wTypeFlags == TYPEFLAG_FLICENSED) {
        Write(_T(",\n  licensed"));
    }
    if (pAttr->wTypeFlags == TYPEFLAG_FCONTROL) {
        Write(_T(",\n  control"));
    }

    Write(_T("\n]\ncoclass %s"), CString(bstrName));
}
