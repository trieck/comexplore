/////////////////////////////////////////////////////////////////////////////
//
// COMMON.H : common include file
//

#pragma once

#ifndef __COMMON_H__
#define __COMMON_H__

#include <memory>
#include <functional>
#include <string>
#include <vector>
#include <map>
#include <list>

#include "coallocator.h"
#include "coobject.h"

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <objbase.h>
#include <tchar.h>
#include <time.h>
#include <comutil.h>

/////////////////////////////////////////////////////////////////////////////
// Iterator interface
template <typename T>
DECLARE_INTERFACE(IIterator)
{
public:
    virtual T GetNext() PURE;
    virtual bool HasNext() const PURE;
    virtual void Reset() PURE;
};

/////////////////////////////////////////////////////////////////////////////
// COM initializer
class coinit {
public:
    coinit () {
        CoInitialize(NULL);
    }
    ~coinit () {
        CoUninitialize();
    }
};

/////////////////////////////////////////////////////////////////////////////
// COM allocated string construction
typedef std::basic_string<TCHAR, std::char_traits<TCHAR>,
        coallocator<TCHAR> > costring;

typedef coobject<costring, coallocator<costring> > costringptr;

/////////////////////////////////////////////////////////////////////////////
// case insensitive costring comparison
struct stringless : std::binary_function <costring, costring, bool> {
    bool operator () (const costring & x, const costring & y) const {
        return (stricmp(x.c_str(), y.c_str()) > 0);
    }
};

/////////////////////////////////////////////////////////////////////////////
// COM allocated vector construction
typedef std::vector<costring, coallocator<costring> > costringvec;
typedef coobject<costringvec, coallocator<costringvec> > costringvecptr;

/////////////////////////////////////////////////////////////////////////////
// COM allocated map construction
typedef std::pair<costring, costring> costringpair;
typedef std::map<costring, costring, stringless,
        coallocator <costringpair> > costringmap;
typedef coobject<costringmap, coallocator<costringmap> > costringmapptr;

/////////////////////////////////////////////////////////////////////////////
// COM allocated list construction
typedef std::list<costring, coallocator<costring> > costringlist;
typedef coobject<costringlist, coallocator<costringlist> > costringlistptr;

/////////////////////////////////////////////////////////////////////////////
class StringVecIterator : public IIterator<costring> {
public:
    StringVecIterator(const costringvec &vec) : v(vec) {
        Reset();
    }
    StringVecIterator(const StringVecIterator &rhs) {
        *this = rhs;
    }
    StringVecIterator &operator = (const StringVecIterator &rhs) {
        if (this != &rhs) {
            v = rhs.v;
            Reset();
        }
        return *this;
    }
    costring GetNext() {
        return *it++;
    }
    bool HasNext() const {
        return it != v.end();
    }
    void Reset() {
        it = v.begin();
    }
private:
    costringvec v;
    costringvec::const_iterator it;
};

/////////////////////////////////////////////////////////////////////////////
class StringListIterator : public IIterator<costring> {
public:
    StringListIterator(const costringlist &list) : l(list) {
        Reset();
    }
    StringListIterator(const StringListIterator &rhs) {
        *this = rhs;
    }
    StringListIterator &operator = (const StringListIterator &rhs) {
        if (this != &rhs) {
            l = rhs.l;
            Reset();
        }
        return *this;
    }
    costring GetNext() {
        return *it++;
    }
    bool HasNext() const {
        return it != l.end();
    }
    void Reset() {
        it = l.begin();
    }
private:
    costringlist l;
    costringlist::const_iterator it;
};

/////////////////////////////////////////////////////////////////////////////
class StringMapIterator : public IIterator<costringpair> {
public:
    StringMapIterator(const costringmap &map) : m(map) {
        Reset();
    }
    StringMapIterator(const StringMapIterator &rhs)  {
        *this = rhs;
    }
    StringMapIterator &operator = (const StringMapIterator &rhs) {
        if (this != &rhs) {
            m = rhs.m;
            Reset();
        }
        return *this;
    }
    costringpair GetNext() {
        return *it++;
    }
    bool HasNext() const {
        return it != m.end();
    }
    void Reset() {
        it = m.begin();
    }
private:
    costringmap m;
    costringmap::const_iterator it;
};

/////////////////////////////////////////////////////////////////////////////
class cofiletime {
private:
    cofiletime() {}
public:
    typedef coobject<FILETIME> filetime;
    typedef coobject<SYSTEMTIME> systemtime;

    static filetime get() {
        FILETIME ft;
        HRESULT hr = CoFileTimeNow(&ft);
        _com_util::CheckError(hr);
        return ft;
    }

    static systemtime getsystime() {
        filetime f = get();
        SYSTEMTIME st;
        FileTimeToSystemTime(&*f, &st);
        return st;
    }

    static costring fmttime(const costring &fmt) {
        systemtime s = getsystime();

        struct tm atm;
        atm.tm_sec = (*s).wSecond;
        atm.tm_min = (*s).wMinute;
        atm.tm_hour = (*s).wHour;
        atm.tm_mday = (*s).wDay;
        atm.tm_mon = (*s).wMonth - 1;
        atm.tm_year = (*s).wYear - 1900;
        __time64_t t = _mktime64(&atm);

        TCHAR buf[128];
        struct tm* tmp = _localtime64(&t);
        _tcsftime(buf, 128, fmt.c_str(), tmp);

        return buf;
    }
};

extern StringMapIterator GetClassGuids();

#endif // __COMMON_H__
