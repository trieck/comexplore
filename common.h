/////////////////////////////////////////////////////////////////////////////
//
// COMMON.H : common include file
//

#pragma once

#include <functional>
#include <list>
#include <map>
#include <memory>

#include <string>
#include <utility>
#include <vector>

#ifdef _UNICODE
using tstring = std::wstring;
#else
using tstring = std::string;
#endif

#define WINVER          _WIN32_WINNT_WIN10
#define _WIN32_WINNT    _WIN32_WINNT_WIN10
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#define WM_SELCHANGED (WM_APP + 1)

#include <atltrace.h>
#include <atlstr.h>
#include <windows.h>
#include <objbase.h>
#include <tchar.h>
#include <ctime>
#include <comutil.h>

#include "coallocator.h"
#include "coobject.h"

/////////////////////////////////////////////////////////////////////////////
// Iterator interface
template <typename T>
DECLARE_INTERFACE(IIterator)
{
    virtual ~IIterator() = default;
    virtual T GetNext() PURE;
    virtual bool HasNext() const PURE;
    virtual void Reset() PURE;
};

/////////////////////////////////////////////////////////////////////////////
// COM initializer
class coinit
{
public:
    coinit()
    {
        CoInitialize(nullptr);
    }

    ~coinit()
    {
        CoUninitialize();
    }
};

/////////////////////////////////////////////////////////////////////////////
// COM allocated string construction
using costring = std::basic_string<TCHAR, std::char_traits<TCHAR>,
                                   coallocator<TCHAR>>;

using costringptr = coobject<costring, coallocator<costring>>;

/////////////////////////////////////////////////////////////////////////////
// case insensitive costring comparison
struct stringless : std::binary_function<costring, costring, bool>
{
    bool operator ()(const costring& x, const costring& y) const
    {
        return (_tcsicmp(x.c_str(), y.c_str()) > 0);
    }
};

/////////////////////////////////////////////////////////////////////////////
// COM allocated vector construction
using costringvec = std::vector<costring, coallocator<costring>>;
using costringvecptr = coobject<costringvec, coallocator<costringvec>>;

/////////////////////////////////////////////////////////////////////////////
// COM allocated map construction
using costringpair = std::pair<costring, costring>;
using costringmap = std::map<costring, costring, stringless,
                             coallocator<costringpair>>;
using costringmapptr = coobject<costringmap, coallocator<costringmap>>;

/////////////////////////////////////////////////////////////////////////////
// COM allocated list construction
using costringlist = std::list<costring, coallocator<costring>>;
using costringlistptr = coobject<costringlist, coallocator<costringlist>>;

/////////////////////////////////////////////////////////////////////////////
class StringVecIterator : public IIterator<costring>
{
public:
    StringVecIterator(costringvec vec) : v(std::move(vec))
    {
        Reset();
    }

    StringVecIterator(const StringVecIterator& rhs)
    {
        *this = rhs;
    }

    StringVecIterator& operator =(const StringVecIterator& rhs)
    {
        if (this != &rhs) {
            v = rhs.v;
            Reset();
        }
        return *this;
    }

    costring GetNext() override
    {
        return *it++;
    }

    bool HasNext() const override
    {
        return it != v.end();
    }

    void Reset() override
    {
        it = v.begin();
    }

private:
    costringvec v;
    costringvec::const_iterator it;
};

/////////////////////////////////////////////////////////////////////////////
class StringListIterator : public IIterator<costring>
{
public:
    StringListIterator(costringlist list) : l(std::move(list))
    {
        Reset();
    }

    StringListIterator(const StringListIterator& rhs)
    {
        *this = rhs;
    }

    StringListIterator& operator =(const StringListIterator& rhs)
    {
        if (this != &rhs) {
            l = rhs.l;
            Reset();
        }
        return *this;
    }

    costring GetNext() override
    {
        return *it++;
    }

    bool HasNext() const override
    {
        return it != l.end();
    }

    void Reset() override
    {
        it = l.begin();
    }

private:
    costringlist l;
    costringlist::const_iterator it;
};

/////////////////////////////////////////////////////////////////////////////
class StringMapIterator : public IIterator<costringpair>
{
public:
    StringMapIterator(costringmap map) : m(std::move(map))
    {
        Reset();
    }

    StringMapIterator(const StringMapIterator& rhs)
    {
        *this = rhs;
    }

    StringMapIterator& operator =(const StringMapIterator& rhs)
    {
        if (this != &rhs) {
            m = rhs.m;
            Reset();
        }
        return *this;
    }

    costringpair GetNext() override
    {
        return *it++;
    }

    bool HasNext() const override
    {
        return it != m.end();
    }

    void Reset() override
    {
        it = m.begin();
    }

private:
    costringmap m;
    costringmap::const_iterator it;
};

/////////////////////////////////////////////////////////////////////////////
class cofiletime
{
    cofiletime() = default;

public:
    using filetime = coobject<FILETIME>;
    using systemtime = coobject<SYSTEMTIME>;

    filetime get()
    {
        FILETIME ft;
        auto hr = CoFileTimeNow(&ft);
        _com_util::CheckError(hr);
        return ft;
    }

    systemtime getsystime()
    {
        auto f = get();
        SYSTEMTIME st;
        FileTimeToSystemTime(&*f, &st);
        return st;
    }

    costring fmttime(const costring& fmt)
    {
        auto s = getsystime();

        struct tm atm;
        atm.tm_sec = (*s).wSecond;
        atm.tm_min = (*s).wMinute;
        atm.tm_hour = (*s).wHour;
        atm.tm_mday = (*s).wDay;
        atm.tm_mon = (*s).wMonth - 1;
        atm.tm_year = (*s).wYear - 1900;
        auto t = _mktime64(&atm);

        TCHAR buf[128];
        auto tmp = _localtime64(&t);
        _tcsftime(buf, 128, fmt.c_str(), tmp);

        return buf;
    }
};

#include "objdata.h"

constexpr auto REG_BUFFER_SIZE = 1024;

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
