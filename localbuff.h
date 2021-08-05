#pragma once

/////////////////////////////////////////////////////////////////////////////
template <typename T>
class LocalBuffer
{
public:
    LocalBuffer();
    ~LocalBuffer();

    BOOL Allocate(size_t sz);
    operator T*();

private:
    void Free();

    T* m_p;
};

template <typename T>
LocalBuffer<T>::LocalBuffer() : m_p(nullptr)
{
}

template <typename T>
LocalBuffer<T>::~LocalBuffer()
{
    Free();
}

template <typename T>
BOOL LocalBuffer<T>::Allocate(size_t sz)
{
    Free();

    m_p = static_cast<T*>(LocalAlloc(LMEM_FIXED, sz));
    if (m_p == nullptr) {
        return FALSE;
    }

    return true;
}

template <typename T>
LocalBuffer<T>::operator T*()
{
    return m_p;
}

template <typename T>
void LocalBuffer<T>::Free()
{
    if (m_p != nullptr) {
        LocalFree(m_p);
        m_p = nullptr;
    }
}
