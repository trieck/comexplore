#pragma once

template <typename T>
class AutoDesc
{
public:
    AutoDesc(LPTYPEINFO pTypeInfo);
    ~AutoDesc();

    HRESULT Get(UINT index);
    void Release();

    T* operator->();

private:
    CComPtr<ITypeInfo> m_pTypeInfo;
    T* m_pDesc;
};

template <typename T>
AutoDesc<T>::AutoDesc(LPTYPEINFO pTypeInfo)
    : m_pTypeInfo(pTypeInfo), m_pDesc(nullptr)
{
    ATLASSERT(m_pTypeInfo != nullptr);
}

template <typename T>
AutoDesc<T>::~AutoDesc()
{
    Release();
    m_pTypeInfo.Release();
}

template <typename T>
T* AutoDesc<T>::operator->()
{
    ATLASSERT(m_pDesc != nullptr);

    return m_pDesc;
}

template <>
inline HRESULT AutoDesc<VARDESC>::Get(UINT index)
{
    Release();

    auto hr = m_pTypeInfo->GetVarDesc(index, &m_pDesc);

    return hr;
}

template <>
inline void AutoDesc<VARDESC>::Release()
{
    if (m_pDesc != nullptr) {
        m_pTypeInfo->ReleaseVarDesc(m_pDesc);
        m_pDesc = nullptr;
    }
}

template <>
inline HRESULT AutoDesc<FUNCDESC>::Get(UINT index)
{
    Release();

    auto hr = m_pTypeInfo->GetFuncDesc(index, &m_pDesc);

    return hr;
}

template <>
inline void AutoDesc<FUNCDESC>::Release()
{
    m_pTypeInfo->ReleaseFuncDesc(m_pDesc);
    m_pDesc = nullptr;
}
