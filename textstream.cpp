#include "stdafx.h"
#include "textstream.h"

BOOL TextStream::Create()
{
    m_pStream.Release();

    m_pStream = SHCreateMemStream(nullptr, 0);
    if (m_pStream == nullptr) {
        return FALSE;
    }

    m_cxChar = m_cyChar = 0;
    m_strText.Empty();
    m_lines.clear();

    if (m_font) {
        m_font.DeleteObject();
    }

    if (!m_font.CreatePointFont(100, _T("Cascadia Mono"))) {
        return FALSE;
    }

    return TRUE;
}

uint64_t TextStream::Size() const
{
    ATLASSERT(m_pStream != nullptr);

    STATSTG statstg;
    auto hr = m_pStream->Stat(&statstg, STATFLAG_NONAME);
    if (FAILED(hr)) {
        return 0;
    }

    return statstg.cbSize.QuadPart;
}

void TextStream::Prepare(CWindow& window)
{
    ATLASSERT(m_pStream != nullptr);

    // Convert stream to series of lines for rendering
    m_strText = Read();

    ResetStream();

    CClientDC dc(window);

    auto hOldFont = dc.SelectFont(m_font);
    dc.DrawText(m_strText, -1, &m_rc, DT_CALCRECT);

    auto* buffer = m_strText.LockBuffer();

    m_lines.clear();
    m_lines.push_back(buffer);

    for (auto* p = buffer + 1; *p != _T('\0'); ++p) {
        if (*p == '\n') {
            if (p[1] != _T('\0')) {
                *p = _T('\0');
                m_lines.push_back(&p[1]);
            }
        }
    }

    m_strText.UnlockBuffer();

    TEXTMETRIC tm;
    dc.GetTextMetrics(&tm);

    dc.GetCharWidth32('X', 'X', &m_cxChar);
    m_cyChar = tm.tmHeight + tm.tmExternalLeading;

    dc.SelectFont(hOldFont);
}

void TextStream::Render(CDCHandle dc)
{
    if (!m_lines.empty()) {
        CRect rc;
        dc.GetClipBox(&rc);

        auto nstart = rc.top / m_cyChar;
        auto nend = std::min<int>(static_cast<int>(m_lines.size()) - 1,
                                  (rc.bottom + m_cyChar - 1) / m_cyChar);

        auto hOldFont = dc.SelectFont(m_font);

        for (auto i = nstart; i <= nend; ++i) {
            auto y = m_cyChar * i;
            auto* pLine = m_lines[i];
            auto nLen = int(_tcslen(pLine));
            dc.ExtTextOut(0, y, ETO_CLIPPED, &rc, pLine, nLen);
        }

        dc.SelectFont(hOldFont);
    }
}

CSize TextStream::GetCharSize() const
{
    return { m_cxChar, m_cyChar };
}

CSize TextStream::GetDocSize() const
{
    return { m_rc.Width(), m_rc.Height() };
}

BOOL TextStream::ResetStream()
{
    ATLASSERT(m_pStream != nullptr);

    ULARGE_INTEGER uli{};

    auto hr = m_pStream->SetSize(uli);
    if (FAILED(hr)) {
        return hr;
    }

    LARGE_INTEGER li{};
    hr = m_pStream->Seek(li, STREAM_SEEK_SET, nullptr);
    if (FAILED(hr)) {
        return FALSE;
    }

    return TRUE;
}

CString TextStream::Read() const
{
    ATLASSERT(m_pStream != nullptr);

    auto size = Size();

    CString strText;
    auto* buffer = strText.GetBuffer(int(size + 1));

    LARGE_INTEGER li{};
    auto hr = m_pStream->Seek(li, STREAM_SEEK_SET, nullptr);
    if (FAILED(hr)) {
        return _T("");
    }

    ULONG uRead;
    hr = m_pStream->Read(buffer, ULONG(size), &uRead);
    if (FAILED(hr)) {
        return _T("");
    }

    if (size != uRead) {
        return _T("");
    }

    buffer[uRead / sizeof(TCHAR)] = _T('\0');
    strText.ReleaseBuffer();

    return strText;
}

BOOL TextStream::WriteV(LPCTSTR format, va_list args)
{
    ATLASSERT(m_pStream != nullptr && format != nullptr);

    CString strValue;
    strValue.FormatV(format, args);

    ULONG cb = strValue.GetLength() * sizeof(TCHAR);
    ULONG written;

    auto hr = m_pStream->Write(strValue, cb, &written);
    if (FAILED(hr)) {
        return FALSE;
    }

    if (written != cb) {
        return FALSE;
    }

    return TRUE;
}

BOOL TextStream::Write(LPCTSTR format, ...)
{
    va_list argList;
    va_start(argList, format);
    auto result = WriteV(format, argList);
    va_end(argList);

    return result;
}
