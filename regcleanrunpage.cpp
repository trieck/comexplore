#include "stdafx.h"
#include "regcleanrunpage.h"
#include "util.h"

static DWORD WINAPI CleanProc(LPVOID pv);
static BOOL ReadFile(LPCTSTR pFilename);
static CString GetServerPath(LPCTSTR pFilename, BOOL inProc);
static CString ParseLocalServer(LPCTSTR pFilename);
static std::vector<CString> SplitPath(LPCTSTR pFilename);

RegCleanRunPage::RegCleanRunPage(RegCleanInfo* pInfo, _U_STRINGorID title)
    : BasePage(title), RegCleanInfoRef(pInfo)
{
    SetHeaderTitle(_T("Looking for invalid COM entries..."));
    SetHeaderSubTitle(_T(""));
}

LRESULT RegCleanRunPage::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    m_progress = ::GetDlgItem(m_hWnd, IDC_REGCLEAN_PROGRESS);

    return 0;
}

int RegCleanRunPage::OnSetActive()
{
    SetWizardButtons(0);

    m_progress.SetRange(0, 100);
    m_progress.SetPos(0);
    m_progress.SetStep(1);

    auto hThread = CreateThread(nullptr, 0, CleanProc, this, 0, nullptr);
    if (hThread == nullptr) {
        ATLTRACE(_T("Unable to create worker thread.\n"));
        return 0;
    }

    CloseHandle(hThread);

    return 0;
}

BOOL RegCleanRunPage::OnQueryCancel()
{
    return FALSE;
}

LRESULT RegCleanRunPage::OnProgress(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    ATLASSERT(m_progress.IsWindow());

    m_progress.SetPos(static_cast<int>(wParam));

    return 0;
}

LRESULT RegCleanRunPage::OnComplete(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
    if (wParam == 0) {
        // SUCCESS
        Sleep(500);
        SetWizardButtons(PSWIZB_BACK | PSWIZB_NEXT);
    } else {
        // ERROR condition
        auto lower = 0, upper = 0;
        m_progress.GetRange(lower, upper);
        m_progress.SetPos(upper);
        m_progress.SetState(PBST_ERROR);

        WinErrorMsgBox(*this, static_cast<DWORD>(lParam), _T("RegClean Error"), MB_ICONERROR);

    }

    return 0;
}

LRESULT RegCleanRunPage::OnSetRange(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
{
    m_progress.SetRange(static_cast<int>(wParam), static_cast<int>(lParam));

    return 0;
}

int RegCleanRunPage::OnWizardBack()
{
    return 0;
}

int RegCleanRunPage::OnWizardNext()
{
    return 0;
}

DWORD WINAPI CleanProc(LPVOID pv)
{
    auto* pThis = static_cast<RegCleanRunPage*>(pv);
    ATLASSERT(pThis);

    CRegKey key, subkey;
    auto lResult = key.Open(HKEY_CLASSES_ROOT, _T("CLSID"), KEY_READ);
    if (lResult != ERROR_SUCCESS) {
        SendMessage(*pThis, WM_COMPLETE, lResult, lResult);
        return -1;
    }

    DWORD dwSubKeys = 0;
    lResult = RegQueryInfoKey(key, nullptr, nullptr, nullptr, &dwSubKeys,
                              nullptr, nullptr, nullptr, nullptr, nullptr,
                              nullptr, nullptr);
    if (lResult != ERROR_SUCCESS) {
        SendMessage(*pThis, WM_COMPLETE, lResult, lResult);
        return -1;
    }

    SendMessage(*pThis, WM_SETRANGE, 0, dwSubKeys);

    DWORD index = 0, length = REG_BUFFER_SIZE;
    TCHAR filename[MAX_PATH];
    auto& info = pThis->GetInfo();

    for (;;) {
        TCHAR szCLSID[REG_BUFFER_SIZE];
        length = REG_BUFFER_SIZE;

        SendMessage(pThis->m_hWnd, WM_PROGRESS, index, 0);

        lResult = key.EnumKey(index++, szCLSID, &length);
        if (lResult != ERROR_SUCCESS) {
            break;
        }

        lResult = subkey.Open(key.m_hKey, szCLSID, 
            DELETE | KEY_ENUMERATE_SUB_KEYS | KEY_READ | KEY_QUERY_VALUE);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }

        BOOL inProc = TRUE;
        CString strPath;
        strPath.Format(_T("%s\\InprocServer32"), szCLSID);

        lResult = subkey.Open(key.m_hKey, strPath, KEY_READ);
        if (lResult != ERROR_SUCCESS) {
            // check for LocalServer
            strPath.Format(_T("%s\\LocalServer32"), szCLSID);
            lResult = subkey.Open(key.m_hKey, strPath, KEY_READ);
            if (lResult != ERROR_SUCCESS) {
                continue;
            }
            inProc = FALSE;
        }

        length = REG_BUFFER_SIZE;
        lResult = subkey.QueryStringValue(nullptr, filename, &length);
        if (lResult != ERROR_SUCCESS) {
            continue;
        }
        CString path = GetServerPath(filename, inProc);
        if (!ReadFile(path)) {
            info.clsids[szCLSID] = filename; // cannot read
        }
    }

    SendMessage(*pThis, WM_COMPLETE, 0, 0);

    return 0;
}

BOOL ReadFile(LPCTSTR pFilename)
{
    ATLASSERT(pFilename);

    auto attr = GetFileAttributes(pFilename);
    if (attr == INVALID_FILE_ATTRIBUTES) {
        return FALSE;
    }

    HANDLE hFile = CreateFile(pFilename,
                              GENERIC_READ,
                              FILE_SHARE_READ,
                              nullptr,
                              OPEN_EXISTING,
                              FILE_ATTRIBUTE_NORMAL,
                              nullptr);
    if (hFile == INVALID_HANDLE_VALUE) {
        return FALSE;
    }

    CloseHandle(hFile);

    return TRUE;
}

std::vector<CString> SplitPath(LPCTSTR pFilename)
{
    std::vector<CString> output;
    ATLASSERT(pFilename);

    CString arg;
    BOOL inQuotes = FALSE;

    for (;;) {
        switch (*pFilename) {
        case '\0':
            if (!arg.IsEmpty()) {
                output.push_back(arg);
            }
            return output;
        case '\"':
            if (inQuotes) {
                output.push_back(arg);
                inQuotes = FALSE;
            } else {
                inQuotes = TRUE;
            }
            arg.Empty();
            break;
        case ' ':
            if (inQuotes) {
                arg += ' ';
            } else {
                // for the first path argument, treat all spaces
                // as legal file path characters and not as delimiters.
                if (output.empty() && _tcsncmp(&pFilename[-4], _T(".exe"), 4) != 0) {
                    arg += ' ';
                } else if (!arg.IsEmpty()) {
                    arg.Trim();
                    output.push_back(arg);
                    arg.Empty();
                }
            }
            break;
        case ',':
        case '-':
        case '/':
            if (!arg.IsEmpty() && pFilename[-1] == ' ') {
                arg.Trim();
                output.push_back(arg);
                arg.Empty();
            }
            arg += *pFilename;
            break;
        default:
            arg += *pFilename;
            break;
        }
        ++pFilename;
    }
}

CString ParseLocalServer(LPCTSTR pFilename)
{
    ATLASSERT(pFilename);
    auto args = SplitPath(pFilename);
    if (!args.empty()) {
        return args[0];
    }

    return _T("");
}

CString GetServerPath(LPCTSTR pFilename, BOOL inProc)
{
    ATLASSERT(pFilename);

    CString filename;
    if (inProc) {
        filename = pFilename;
        filename.Remove('\"');
    } else {
        filename = ParseLocalServer(pFilename);
    }

    TCHAR fileBuffer[MAX_PATH];
    ExpandEnvironmentStrings(filename, fileBuffer, REG_BUFFER_SIZE);
    pFilename = fileBuffer;

    if (ReadFile(pFilename)) {
        return pFilename;
    }

    // might be a system object with no path -- one last shot
    if (_tcschr(pFilename, '\\') != nullptr) {
        // already contains path information
        return pFilename;
    }

    TCHAR systemDir[MAX_PATH];
    TCHAR combinedFile[MAX_PATH];

    GetSystemDirectory(systemDir, MAX_PATH);
    PathCombine(combinedFile, systemDir, pFilename);

    return combinedFile;
}
