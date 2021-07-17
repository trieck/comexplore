#pragma once

#ifdef _ATL_MIN_CRT
#undef _ATL_MIN_CRT
#endif

#define _WTL_FORWARD_DECLARE_CSTRING

#define WINVER          _WIN32_WINNT_WIN10
#define _WIN32_WINNT    _WIN32_WINNT_WIN10
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#define WM_SELCHANGED (WM_APP + 1)

#include <atlbase.h>
#include <atlapp.h>
#include <atlstr.h>
#include <atlcrack.h>
#include <atlctrls.h>
#include <atldlgs.h>
#include <atlwin.h>
#include <atlframe.h>
#include <atlctrls.h>
#include <atldlgs.h>
#include <atlctrlw.h>
#include <atlctrlx.h>
#include <atlsplit.h>
#include <atltrace.h>
#include <atltypes.h>
#include <atlscrl.h>
#include <atlstr.h>

#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <string>

#ifdef _UNICODE
using tstring = std::wstring;
#else
using tstring = std::string;
#endif

#define MAKE_TREEITEM(n, t) \
    CTreeItem(((LPNMTREEVIEW)(n))->itemNew.hItem, t)
#define MAKE_OLDTREEITEM(n, t) \
    CTreeItem(((LPNMTREEVIEW)(n))->itemOld.hItem, t)

#define MSG_WM_PAINT2(func) \
	if (uMsg == WM_PAINT) \
	{ \
		this->SetMsgHandled(TRUE); \
		func(CPaintDC(*this)); \
		lResult = 0; \
		if(this->IsMsgHandled()) \
			return TRUE; \
	}

constexpr auto REG_BUFFER_SIZE = 1024;

struct equal_to_string
{
    bool operator()(const LPCTSTR& lhs, const LPCTSTR& rhs) const
    {
        ATLASSERT(lhs != nullptr);
        ATLASSERT(rhs != nullptr);

        return _tcscmp(lhs, rhs) == 0;
    }
};

struct string_hash
{
#if defined(_WIN64)
    static const size_t FNV_PRIME = 1099511628211ULL;
#else // defined(_WIN64)
    static const size_t FNV_PRIME = 16777619U;
#endif // defined(_WIN64)

    size_t operator()(LPCTSTR p) const
    {
        ATLASSERT(p != nullptr);

        size_t result = 0;

        while (*p != '\0') {
            result ^= *p++;
            result *= FNV_PRIME;
        }

        return result;
    }
};

template <typename Value>
class string_key_map
{
    string_key_map() = default;
public:
    using type = std::unordered_map<LPCTSTR, Value, string_hash, equal_to_string>;
};

using string_set = std::unordered_set<LPCTSTR, string_hash, equal_to_string>;

extern CAppModule _Module;

#if defined _M_IX86
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
