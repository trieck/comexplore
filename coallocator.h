/////////////////////////////////////////////////////////////////////////////
//
// COALLOCATOR.H : STL Custom COM Allocator
//

#pragma once

#ifndef __COALLOCATOR_H__
#define __COALLOCATOR_H__

#include <limits>
#include "comalloc.h"

/////////////////////////////////////////////////////////////////////////////
template <typename T>
class coallocator
{
public:
    typedef std::size_t size_type;
    typedef std::ptrdiff_t difference_type;
    typedef T* pointer;
    typedef const T* const_pointer;
    typedef T& reference;
    typedef const T& const_reference;
    typedef T value_type;

    template <class U>
    struct rebind
    {
        using other = coallocator<U>;
    };

    coallocator() noexcept = default;

    coallocator(const coallocator&) throw() = default;

    template <class U>
    coallocator(const coallocator<U>&) noexcept
    {
    }

    ~coallocator() noexcept = default;

    pointer address(reference x) const
    {
        return &x;
    }

    const_pointer address(const_reference x) const
    {
        return &x;
    }

    pointer allocate(size_type n, void* /*hint*/ = nullptr)
    {
        auto pv = static_cast<pointer>(comalloc(n * sizeof(T)));
        if (pv == nullptr)
            throw std::bad_alloc();
        return pv;
    }

    void deallocate(pointer p, size_type n)
    {
        cofree(static_cast<void*>(p));
    }

    size_type max_size() const noexcept
    {
        return std::numeric_limits<size_type>::max() / sizeof(T);
    }

    void construct(pointer p, const T& val)
    {
        new(static_cast<void*>(p)) T(val);
    }

    void destroy(pointer p)
    {
        p->~T();
    }
};

/////////////////////////////////////////////////////////////////////////////
template <typename T>
bool operator ==(const coallocator<T>&, const coallocator<T>&)
{
    return true;
}

/////////////////////////////////////////////////////////////////////////////
template <typename T>
bool operator !=(const coallocator<T>&, const coallocator<T>&)
{
    return false;
}

#endif // __COALLOCATOR_H__
