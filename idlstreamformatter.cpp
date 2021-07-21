#include "stdafx.h"
#include "idlstreamformatter.h"

static const string_set KEYWORDS({
    _T("_cdecl"),
    _T("_pascal"),
    _T("_stdcall"),
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
    _T("public"),
    _T("readonly"),
    _T("restricted"),
    _T("retval"),
    _T("source"),
    _T("unique"),
    _T("uuid"),
    _T("version")
});

static const string_set TYPES({
    _T("BLOB"),
    _T("BLOB_OBJECT"),
    _T("boolean"),
    _T("BSTR"),
    _T("BSTR*"),
    _T("CARRAY"),
    _T("CF"),
    _T("char"),
    _T("char*"),
    _T("CLSID")
    _T("CURRENCY"),
    _T("DATE"),
    _T("DATE*"),
    _T("double"),
    _T("double*"),
    _T("FILETIME"),
    _T("FILETIME*"),
    _T("GUID"),
    _T("GUID*"),
    _T("float"),
    _T("float*"),
    _T("HRESULT"),
    _T("IDispatch*"),
    _T("int"),
    _T("int*"),
    _T("int64"),
    _T("int64*"),
    _T("IUnknown*"),
    _T("long"),
    _T("long*"),
    _T("LPSTR"),
    _T("LPSTR*"),
    _T("LPWSTR"),
    _T("LPWSTR*"),
    _T("null"),
    _T("PTR"),
    _T("SAFEARRAY"),
    _T("SAFEARRAY*"),
    _T("SCODE"),
    _T("short"),
    _T("short*"),
    _T("STORAGE"),
    _T("STORED_OBJECT"),
    _T("STREAM"),
    _T("STREAMED_OBJECT"),
    _T("uint64"),
    _T("uint64*"),
    _T("unsigned"),
    _T("USERDEFINED"),
    _T("VARIANT"),
    _T("VARIANT*"),
    _T("void"),
    _T("void*"),
    _T("wchar_t")
    _T("wchar_t*")
});

struct Token
{
    enum class Type
    {
        EMPTY = 0,
        ID,
        LITERAL,
        COMMENT,
        WHITESPACE,
        NEWLINE,
        OTHER
    };

    Token() : type(Type::EMPTY)
    {
    }

    Type type;
    CString value;
};

static Token GetToken(LPCTSTR* ppin);
static BOOL ParseGUID(LPCTSTR pin);

enum COLORS
{
    COLOR_TEXT= 1,
    COLOR_KEYWORD,
    COLOR_PROPERTY,
    COLOR_COMMENT,
    COLOR_LITERAL,
    COLOR_TYPE,
};

CStringA IDLStreamFormatter::Format(TextStream& stream)
{
    CString output(_T("{\\rtf\\ansi\\deff0{\\fonttbl{\\f0 Cascadia Mono;\\f1 Courier New;}}"
        "{\\colortbl;"
        "\\red0\\green0\\blue0;" // COLOR_TEXT
        "\\red0\\green0\\blue200;" // COLOR_KEYWORD
        "\\red200\\green0\\blue0;" // COLOR_PROPERTY
        "\\red0\\green128\\blue0;" // COLOR_COMMENT
        "\\red200\\green0\\blue200;" // COLOR_LITERAL
        "\\red128\\green0\\blue0;" // COLOR_TYPE
        ";}\n"
        "\\cf1\n"));

    auto strText = stream.ReadString();
    auto textColor = COLOR_TEXT;
    COLORREF lastColor = textColor;

    LPCTSTR pText = strText;

    string_set::const_iterator it;

    auto bDone = FALSE;
    while (!bDone) {
        textColor = COLOR_TEXT;

        auto token = GetToken(&pText);

        switch (token.type) {
        case Token::Type::EMPTY:
            ATLASSERT(*pText == '\0');
            bDone = TRUE;
            break;
        case Token::Type::ID:
            it = KEYWORDS.find(token.value);
            if (it != KEYWORDS.end()) {
                textColor = COLOR_KEYWORD;
                break;
            }
            it = PROP_KEYWORDS.find(token.value);
            if (it != PROP_KEYWORDS.end()) {
                textColor = COLOR_PROPERTY;
                break;
            }
            it = TYPES.find(token.value);
            if (it != TYPES.end()) {
                textColor = COLOR_TYPE;
                break;
            }
            break;
        case Token::Type::COMMENT:
            textColor = COLOR_COMMENT;
            break;
        case Token::Type::LITERAL:
            textColor = COLOR_LITERAL;
            break;
        default:
            break;
        }

        if (textColor != lastColor) {
            CString strColor;
            strColor.Format(_T("\\cf%d\n"), textColor);
            output += strColor;
        }

        CString value(token.value);
        if (value[0] == '\n') {
            value = _T("\\line\n");
        } else {
            value.Replace(_T("\\"), _T("\\\\"));
            value.Replace(_T("{"), _T("\\{"));
            value.Replace(_T("}"), _T("\\}"));
        }

        output += value;

        if (textColor != lastColor) {
            output += _T("\\cf1\n");
        }

        lastColor = textColor;
    }

    output += '}';

    CT2CA utf8Str(output, CP_UTF8);

    return CStringA(utf8Str);
}

// Helper functions
Token GetToken(LPCTSTR* ppin)
{
    ATLASSERT(ppin);

    for (Token token;;) {
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
                if (ParseGUID(*ppin)) {
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

                // pointer identifier
                while (**ppin == '*') {
                    token.value += *(*ppin)++;
                }

                token.type = Token::Type::ID;
                return token;
            }

            if (ispunct(**ppin)) {
                token.value = *(*ppin)++;
                token.type = Token::Type::OTHER;
                return token;
            }

            (*ppin)++; // eat
            break;
        }
    }
}

BOOL ParseGUID(LPCTSTR pin)
{
    ATLASSERT(pin);

    for (auto i = 0; i < 8; ++i) {
        if (!isxdigit(*pin++)) {
            return FALSE;
        }
    }

    if (*pin++ != '-') {
        return FALSE;
    }

    for (auto i = 0; i < 4; ++i) {
        if (!isxdigit(*pin++)) {
            return FALSE;
        }
    }

    if (*pin++ != '-') {
        return FALSE;
    }

    for (auto i = 0; i < 4; ++i) {
        if (!isxdigit(*pin++)) {
            return FALSE;
        }
    }

    if (*pin++ != '-') {
        return FALSE;
    }

    for (auto i = 0; i < 4; ++i) {
        if (!isxdigit(*pin++)) {
            return FALSE;
        }
    }

    if (*pin++ != '-') {
        return FALSE;
    }

    for (auto i = 0; i < 12; ++i) {
        if (!isxdigit(*pin++)) {
            return FALSE;
        }
    }

    return TRUE;
}
