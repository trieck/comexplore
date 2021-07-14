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

class AutoTypeLibAttr
{
public:
    AutoTypeLibAttr(LPTYPELIB pTypeLib);
    ~AutoTypeLibAttr();

    HRESULT Get();
    void Release();

    LPTLIBATTR operator->();
    explicit operator LPTLIBATTR();

private:
    CComPtr<ITypeLib> m_pTypeLib;
    LPTLIBATTR m_pTypeLibAttr;
};
