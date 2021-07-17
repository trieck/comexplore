#pragma once

/////////////////////////////////////////////////////////////////////////////
class ATL_NO_VTABLE TextStream :
    public CComObjectRoot,
    public IStream
{
public:
BEGIN_COM_MAP(TextStream)
            COM_INTERFACE_ENTRY(IStream)
    END_COM_MAP()

    HRESULT FinalConstruct();
    void FinalRelease();

    HRESULT Write(LPCTSTR format, ...);
    HRESULT WriteV(LPCTSTR format, va_list args);
    HRESULT Reset();
    CString ReadString();

    // IStream members
    HRESULT Read(void* pv, ULONG cb, ULONG* pcbRead) override;
    HRESULT Write(const void* pv, ULONG cb, ULONG* pcbWritten) override;
    HRESULT Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition) override;
    HRESULT SetSize(ULARGE_INTEGER libNewSize) override;
    HRESULT CopyTo(IStream* pstm, ULARGE_INTEGER cb, ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWritten) override;
    HRESULT Commit(DWORD grfCommitFlags) override;
    HRESULT Revert() override;
    HRESULT LockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType) override;
    HRESULT UnlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType) override;
    HRESULT Stat(STATSTG* pstatstg, DWORD grfStatFlag) override;
    HRESULT Clone(IStream** ppstm) override;

private:
    CComPtr<IStream> m_pImpl;
};
