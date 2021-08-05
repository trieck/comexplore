#pragma once

/////////////////////////////////////////////////////////////////////////////
CString Escape(LPCTSTR lpszInput);
CString GetScodeString(SCODE sc);
CString TYPEDESCtoString(LPTYPEINFO pTypeInfo, TYPEDESC* pTypeDesc);
CString VTtoString(VARTYPE vt);
CString ErrorMessage(DWORD dwError = GetLastError());

int WinErrorMsgBox(HWND hWndOwner,
    DWORD dwError = GetLastError(),
    ATL::_U_STRINGorID title = nullptr, 
    UINT uType = MB_OK | MB_ICONINFORMATION);