#include "stdafx.h"
#include "idlview.h"
#include "autotypeattr.h"
#include "typeinfonode.h"

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

LRESULT IDLView::OnCreate(LPCREATESTRUCT pcs)
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
    auto clrText = RGB(0, 128, 0);

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
