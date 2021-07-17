#include "stdafx.h"
#include "idlstreamrenderer.h"

static const string_set KEYWORDS({
    _T("coclass"),
    _T("const"),
    _T("default"),
    _T("dispinterface"),
    _T("enum"),
    _T("import"),
    _T("importlib"),
    _T("include"),
    _T("interface"),
    _T("library"),
    _T("methods"),
    _T("module"),
    _T("properties"),
    _T("struct"),
    _T("typedef"),
    _T("union"),
});

static const string_set PROP_KEYWORDS({
    _T("appobject"),
    _T("bindable"),
    _T("control"),
    _T("custom"),
    _T("dllname"),
    _T("dual"),
    _T("entry"),
    _T("helpcontext"),
    _T("helpcontext"),
    _T("helpfile"),
    _T("helpstring"),
    _T("hidden"),
    _T("id"),
    _T("in"),
    _T("licensed"),
    _T("nonbrowsable"),
    _T("noncreatable"),
    _T("nonextensible"),
    _T("object"),
    _T("oleautomation"),
    _T("optional"),
    _T("out"),
    _T("propget"),
    _T("propput"),
    _T("propputref"),
    _T("readonly"),
    _T("restricted"),
    _T("retval"),
    _T("source"),
    _T("unique"),
    _T("uuid"),
    _T("version")
});

static const COLORREF KEYWORD_COLOR = RGB(0, 0, 200);
static const COLORREF PROP_KEYWORD_COLOR = RGB(200, 0, 0);
static const COLORREF COMMENT_COLOR = RGB(0, 128, 0);
static const COLORREF LITERAL_COLOR = RGB(200, 0, 200);

struct Token
{
    enum class Type
    {
        EMPTY = 0,
        ID,
        LITERAL,
        COMMENT,
        WHITESPACE,
        NEWLINE
    };

    Token() : type(Type::EMPTY)
    {
    }

    Type type;
    CString value;
};

static Token GetToken(LPCTSTR* ppin);
static BOOL ParseGUID(LPCTSTR* ppin);

IDLStreamRenderer::IDLStreamRenderer()
{
    m_font.CreatePointFont(100, _T("Cascadia Mono"));
}

void IDLStreamRenderer::Parse(TextStream& stream)
{
    auto strText = stream.ReadString();

    LPCTSTR pText = strText;

    CClientDC dc(nullptr);
    if (m_memDC) {
        m_memDC.DeleteDC();
    }
    m_memDC.CreateCompatibleDC(dc);

    m_memDC.SetMapMode(dc.GetMapMode());
    auto hOldFont = m_memDC.SelectFont(m_font);
    m_memDC.DrawText(strText, -1, &m_rc, DT_CALCRECT);

    TEXTMETRIC tm;
    m_memDC.GetTextMetrics(&tm);
    m_memDC.GetCharWidth32('X', 'X', &m_cxChar);
    m_cyChar = tm.tmHeight + tm.tmExternalLeading;

    if (m_bitmap) {
        m_bitmap.DeleteObject();
    }

    BITMAPINFOHEADER bmih{};
    bmih.biSize = sizeof(BITMAPINFOHEADER);
    bmih.biWidth = m_rc.Width();
    bmih.biHeight = m_rc.Height();
    bmih.biPlanes = 1;
    bmih.biBitCount = 32;
    bmih.biCompression = BI_RGB;

    BITMAPINFO bmi{};
    bmi.bmiHeader = bmih;
    if (!m_bitmap.CreateDIBSection(dc, &bmi, DIB_RGB_COLORS, nullptr, nullptr, 0)) {
        return;
    }

    auto hOldBitmap = m_memDC.SelectBitmap(m_bitmap);
    auto clrWindow = GetSysColor(COLOR_WINDOW);
    auto clrText = RGB(0, 0, 0);

    m_memDC.FillSolidRect(m_rc, clrWindow);
    m_memDC.SetBkColor(clrWindow);
    m_memDC.SetBkMode(OPAQUE);
    m_memDC.SetTextColor(clrText);

    auto x = 0, y = 0;
    string_set::const_iterator it;

    auto bDone = FALSE;
    while (!bDone) {
        auto token = GetToken(&pText);
        LPCTSTR value = token.value;

        switch (token.type) {
        case Token::Type::EMPTY:
            ATLASSERT(*pText == '\0');
            bDone = TRUE;
            break;
        case Token::Type::NEWLINE:
            x = 0;
            y += m_cyChar;
            break;
        case Token::Type::ID:
            it = KEYWORDS.find(value);
            if (it != KEYWORDS.end()) {
                m_memDC.SetTextColor(KEYWORD_COLOR);
            }

            it = PROP_KEYWORDS.find(value);
            if (it != PROP_KEYWORDS.end()) {
                m_memDC.SetTextColor(PROP_KEYWORD_COLOR);
            }
            break;
        case Token::Type::COMMENT:
            m_memDC.SetTextColor(COMMENT_COLOR);
            break;
        case Token::Type::LITERAL:
            m_memDC.SetTextColor(LITERAL_COLOR);
            break;
        default:
            break;
        }

        CSize sz;
        m_memDC.GetTextExtent(value, -1, &sz);
        m_memDC.TextOut(x, y, value);
        x += sz.cx;

        m_memDC.SetTextColor(clrText);
    }

    m_memDC.SelectFont(hOldFont);
    m_memDC.SelectBitmap(hOldBitmap);
}

void IDLStreamRenderer::Render(CDCHandle dc)
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

CSize IDLStreamRenderer::GetCharSize() const
{
    return { m_cxChar, m_cyChar };
}

CSize IDLStreamRenderer::GetDocSize() const
{
    return { m_rc.Width(), m_rc.Height() };
}

// Helper functions
Token GetToken(LPCTSTR* ppin)
{
    Token token;

    for (;;) {
        switch (**ppin) {
        case '\0':
            token.type = Token::Type::EMPTY;
            return token;
        case '\r':
        case '\t':
        case ' ':
            while (_istspace(**ppin)) {
                token.value += *(*ppin)++;
            }
            token.type = Token::Type::WHITESPACE;
            return token;
        case '/':
            if ((*ppin)[1] == '/') {
                while (**ppin != '\0' && **ppin != '\n') {
                    token.value += *(*ppin)++;
                }
                token.type = Token::Type::COMMENT;
            } else if ((*ppin)[1] == '*') {
                while (**ppin != '\0') {
                    token.value += **ppin;
                    if (**ppin == '/' && (*ppin)[-1] == '*') {
                        break;
                    }
                    (*ppin)++;
                }
                token.type = Token::Type::COMMENT;
            } else {
                token.value = *(*ppin)++;
                token.type = Token::Type::ID;
            }
            return token;
        case '"':
            token.value = *(*ppin)++;
            while (**ppin != '\0' && **ppin != '\n') {
                token.value += *(*ppin)++;
                if ((*ppin)[-1] == '"') {
                    break;
                }
            }
            token.type = Token::Type::LITERAL;
            return token;
        case '\n':
            token.value = *(*ppin)++;
            token.type = Token::Type::NEWLINE;
            return token;
        default:
            if (isdigit(**ppin) || isxdigit(**ppin)) {
                if (ParseGUID(ppin)) {
                    token.value = CString(*ppin, 36);
                    *ppin += 36;
                    token.type = Token::Type::LITERAL;
                    return token;
                }

                if ((*ppin)[0] == '0' && ((*ppin)[1] == 'x' || (*ppin)[1] == 'X')) {
                    token.value += *(*ppin)++; // hex number
                    token.value += *(*ppin)++;
                    while (isxdigit(**ppin)) {
                        token.value += *(*ppin)++;
                    }
                    token.type = Token::Type::LITERAL;
                    return token;
                }
                if (isdigit(**ppin)) {
                    while (isdigit(**ppin) || **ppin == '.') {
                        token.value += *(*ppin)++;
                    }
                    token.type = Token::Type::LITERAL;
                    return token;
                }
                // fall through
            }
            if (iscsym(**ppin)) {
                while (iscsym(**ppin)) {
                    token.value += *(*ppin)++;
                }
                token.type = Token::Type::ID;
                return token;
            }

            if (ispunct(**ppin)) {
                token.value = *(*ppin)++;
                token.type = Token::Type::ID;
                return token;
            }

            (*ppin)++; // eat
            break;
        }
    }
}

BOOL ParseGUID(LPCTSTR* ppin)
{
    GUID guid;
    auto result = _stscanf(*ppin, _T("%08lX-%04hX-%04hX-%02hhX%02hhX-%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX"),
                           &guid.Data1, &guid.Data2, &guid.Data3,
                           &guid.Data4[0], &guid.Data4[1], &guid.Data4[2],
                           &guid.Data4[3], &guid.Data4[4], &guid.Data4[5],
                           &guid.Data4[6], &guid.Data4[7]);
    if (result != 11) {
        return FALSE;
    }

    return TRUE;
}
