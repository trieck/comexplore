#pragma once

/////////////////////////////////////////////////////////
class IDLStream
{
public:
    IDLStream() = default;
    ~IDLStream() = default;

    BOOL Create();
    BOOL Write(LPCTSTR format, ...);
    BOOL WriteV(LPCTSTR format, va_list args);
    CString Read() const;
    uint64_t Size() const;
    void Parse();
    void Render(CDCHandle dc);
    CSize GetCharSize() const;
    CSize GetDocSize() const;
    BOOL ResetStream();

private:
    int m_cxChar = 0;
    int m_cyChar = 0;
    CRect m_rc;
    CFont m_font;
    CDC m_memDC;
    CBitmap m_bitmap;
    CComPtr<IStream> m_pStream;
};
