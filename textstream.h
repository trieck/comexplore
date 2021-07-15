#pragma once

/////////////////////////////////////////////////////////
class TextStream
{
public:
    TextStream() = default;
    virtual ~TextStream() = default;

    BOOL Create();
    BOOL Write(LPCTSTR format, ...);
    BOOL WriteV(LPCTSTR format, va_list args);
    CString Read() const;
    uint64_t Size() const;
    void Prepare(CWindow& window);
    void Render(CDCHandle dc);
    CSize GetCharSize() const;
    CSize GetDocSize() const;

    BOOL ResetStream();
private:
    CFont m_font;
    int m_cxChar = 0;
    int m_cyChar = 0;
    CComPtr<IStream> m_pStream;
    CString m_strText;
    CRect m_rc;
    std::vector<LPCTSTR> m_lines;
};
