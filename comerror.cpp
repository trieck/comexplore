#include "stdafx.h"
#include "comerror.h"

COMError::COMError(HRESULT hr, LPUNKNOWN pUnknown, REFIID iid)
{
    SetErrorMessage(hr, pUnknown, iid);
}

COMError::operator CString() const
{
    return m_message;
}

void COMError::SetErrorMessage(HRESULT hr, LPUNKNOWN pUnknown, const IID& iid)
{
    _com_error error(hr);
    m_message = error.ErrorMessage();

    if (!pUnknown) {
        return;
    }

    ISupportErrorInfoPtr pSupportErrorInfo;
    hr = pUnknown->QueryInterface(IID_ISupportErrorInfo,
                                  IID_PPV_ARGS_Helper(&pSupportErrorInfo));
    if (FAILED(hr)) {
        return;
    }

    hr = pSupportErrorInfo->InterfaceSupportsErrorInfo(iid);
    if (hr != S_OK) {
        return;
    }

    IErrorInfoPtr pErrorInfo;
    hr = GetErrorInfo(0, &pErrorInfo);
    if (hr != S_OK) {
        return;
    }

    CComBSTR bstrDescription;
    hr = pErrorInfo->GetDescription(&bstrDescription);
    if (FAILED(hr)) {
        return;
    }

    m_message = bstrDescription;
}

/////////////////////////////////////////////////////////////////////////////
int CoMessageBox(HWND hWndOwner, HRESULT hr, LPUNKNOWN pUnknown,
                 ATL::_U_STRINGorID title, UINT uType)
{
    auto message = COM_ERROR_DESC(hr, pUnknown);

    return AtlMessageBox(hWndOwner, static_cast<LPCTSTR>(message), title,
                         uType);
}
