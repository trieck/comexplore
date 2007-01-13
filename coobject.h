/////////////////////////////////////////////////////////////////////////////
//
//	COOBJECT.H : dynamically allocated object wrapper 
//
//	Copyright(c) 2003 KnowX.com, All Rights Reserved
//

#pragma once

#ifndef __COOBJECT_H__
#define __COOBJECT_H__

/////////////////////////////////////////////////////////////////////////////
template <typename T, typename Alloc = coallocator<T> >
class coobject 
{
public:
	typedef T value_type;
	typedef typename Alloc::pointer pointer;
	typedef typename Alloc::reference reference;

	coobject(const T& obj = T(), const Alloc &a = coallocator<T>())
		: alloc(a), ptr(0)
	{
		T *p = alloc.allocate(sizeof(T));
		try {
			alloc.construct(p, obj);
		} catch (...) {
			alloc.deallocate(p, 1);
			throw;
		}
		ptr = p;
	}

	coobject(coobject<typename T, typename Alloc> &rhs) 
		: ptr(rhs.release())
	{
	}

	coobject<typename T, typename Alloc>& operator = (const coobject &rhs) {
		reset(rhs.release());
		return *this;
	}

	~coobject()
	{
		destroy();
	}

	typename Alloc::reference operator *() { return *ptr; }
	value_type operator*() const { return *ptr; }

private:
	typename Alloc::pointer release() {
		typename Alloc::pointer tmp = ptr;
		ptr = 0;	// give up ownership
		return tmp;
	}

	void reset(typename Alloc::pointer p = 0) {
		if (p != ptr) 
			destroy();
		ptr = p;	// take ownership
	}

	void destroy() {
		if (ptr != 0) {
			alloc.destroy(ptr);
			alloc.deallocate(ptr, 1);
		}
		ptr = 0;
	}

	Alloc alloc;
	typename Alloc::pointer ptr;
};
/////////////////////////////////////////////////////////////////////////////
#endif // __COOBJECT_H__