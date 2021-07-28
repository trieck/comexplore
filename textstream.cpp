#include "stdafx.h"
#include "textstream.h"

HRESULT TextStream::Reset()
{
    ATLASSERT(m_pImpl != nullptr);

    ULARGE_INTEGER uli{};

    auto hr = m_pImpl->SetSize(uli);
    if (FAILED(hr)) {
        return hr;
    }

    LARGE_INTEGER li{};
    hr = m_pImpl->Seek(li, STREAM_SEEK_SET, nullptr);
    if (FAILED(hr)) {
        return hr;
    }

    return S_OK;
}

CString TextStream::ReadString()
{
    ATLASSERT(m_pImpl != nullptr);

    STATSTG statstg;
    auto hr = Stat(&statstg, STATFLAG_NONAME);
    if (FAILED(hr)) {
        return _T("");
    }

    auto size = statstg.cbSize.QuadPart;

    CString strText;
    auto* buffer = strText.GetBuffer(static_cast<int>(size + 1));

    LARGE_INTEGER li{};
    hr = Seek(li, STREAM_SEEK_SET, nullptr);
    if (FAILED(hr)) {
        return _T("");
    }

    ULONG uRead;
    hr = Read(buffer, static_cast<ULONG>(size), &uRead);
    if (FAILED(hr)) {
        return _T("");
    }

    if (size != uRead) {
        return _T("");
    }

    buffer[uRead / sizeof(TCHAR)] = _T('\0');
    strText.ReleaseBuffer();

    return strText;
}

HRESULT TextStream::WriteV(LPCTSTR format, va_list args)
{
    ATLASSERT(m_pImpl != nullptr && format != nullptr);

    CString strValue;
    strValue.FormatV(format, args);

    ULONG cb = strValue.GetLength() * sizeof(TCHAR);
    ULONG written;

    auto hr = m_pImpl->Write(strValue, cb, &written);
    if (FAILED(hr)) {
        return hr;
    }

    return S_OK;
}

HRESULT TextStream::Write(LPCTSTR format, ...)
{
    va_list argList;
    va_start(argList, format);
    auto result = WriteV(format, argList);
    va_end(argList);

    return result;
}

TextStream::~TextStream()
{
}

HRESULT TextStream::FinalConstruct()
{
    m_pImpl = SHCreateMemStream(nullptr, 0);
    if (!m_pImpl) {
        return E_FAIL;
    }

    return S_OK;
}

void TextStream::FinalRelease()
{
    m_pImpl.Release();
}

HRESULT TextStream::Read(void* pv, ULONG cb, ULONG* pcbRead)
{
    ATLASSERT(m_pImpl);
    return m_pImpl->Read(pv, cb, pcbRead);
}

HRESULT TextStream::Write(const void* pv, ULONG cb, ULONG* pcbWritten)
{
    ATLASSERT(m_pImpl);
    return m_pImpl->Write(pv, cb, pcbWritten);
}

HRESULT TextStream::Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition)
{
    ATLASSERT(m_pImpl);
    return m_pImpl->Seek(dlibMove, dwOrigin, plibNewPosition);
}

HRESULT TextStream::SetSize(ULARGE_INTEGER libNewSize)
{
    ATLASSERT(m_pImpl);
    return m_pImpl->SetSize(libNewSize);
}

HRESULT TextStream::CopyTo(IStream* pstm, ULARGE_INTEGER cb, ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWritten)
{
    ATLASSERT(m_pImpl);
    return m_pImpl->CopyTo(pstm, cb, pcbRead, pcbWritten);
}

HRESULT TextStream::Commit(DWORD grfCommitFlags)
{
    ATLASSERT(m_pImpl);
    return m_pImpl->Commit(grfCommitFlags);
}

HRESULT TextStream::Revert()
{
    ATLASSERT(m_pImpl);
    return m_pImpl->Revert();
}

HRESULT TextStream::LockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
{
    ATLASSERT(m_pImpl);
    return m_pImpl->LockRegion(libOffset, cb, dwLockType);
}

HRESULT TextStream::UnlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
{
    ATLASSERT(m_pImpl);
    return m_pImpl->UnlockRegion(libOffset, cb, dwLockType);
}

HRESULT TextStream::Stat(STATSTG* pstatstg, DWORD grfStatFlag)
{
    ATLASSERT(m_pImpl);
    return m_pImpl->Stat(pstatstg, grfStatFlag);
}

HRESULT TextStream::Clone(IStream** ppstm)
{
    ATLASSERT(m_pImpl);
    return m_pImpl->Clone(ppstm);
}
