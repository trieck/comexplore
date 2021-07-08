#include "stdafx.h"
#include "autotypeattr.h"

AutoTypeAttr::AutoTypeAttr(LPTYPEINFO pTypeInfo)
    : m_pTypeInfo(pTypeInfo), m_pTypeAttr(nullptr)
{
}

AutoTypeAttr::~AutoTypeAttr()
{
    Release();
    m_pTypeInfo.Release();
}

HRESULT AutoTypeAttr::Get()
{
    Release();

    auto hr = m_pTypeInfo->GetTypeAttr(&m_pTypeAttr);

    return hr;
}

void AutoTypeAttr::Release()
{
    if (m_pTypeAttr != nullptr) {
        m_pTypeInfo->ReleaseTypeAttr(m_pTypeAttr);
        m_pTypeAttr = nullptr;
    }
}

LPTYPEATTR AutoTypeAttr::operator->()
{
    ATLASSERT(m_pTypeAttr != nullptr);

    return m_pTypeAttr;
}

AutoTypeAttr::operator LPTYPEATTR()
{
    ATLASSERT(m_pTypeAttr != nullptr);

    return m_pTypeAttr;
}
