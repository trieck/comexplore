#pragma once

class AutoTypeAttr
{
public:
    AutoTypeAttr(LPTYPEINFO pTypeInfo);
    ~AutoTypeAttr();

    HRESULT Get();
    void Release();

    LPTYPEATTR operator->();
    explicit operator LPTYPEATTR();

private:
    CComPtr<ITypeInfo> m_pTypeInfo;
    LPTYPEATTR m_pTypeAttr;
};
