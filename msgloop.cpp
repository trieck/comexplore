#include "stdafx.h"
#include "msgloop.h"

int MessageLoopEx::Run()
{
    BOOL bDoIdle = TRUE;
    int nIdleCount = 0;

    for (;;) {
        while (bDoIdle && !::PeekMessage(&m_msg, nullptr, 0, 0, PM_NOREMOVE)) {
            if (!OnIdle(nIdleCount++))
                bDoIdle = FALSE;
        }

        auto bRet = ::GetMessage(&m_msg, nullptr, 0, 0);

        if (bRet == -1) {
            ATLTRACE(_T("::GetMessage returned -1 (error)\n"));
            continue; // error, don't process
        }

        if (!bRet) {
            ATLTRACE(_T("CMessageLoop::Run - exiting\n"));
            break; // WM_QUIT, exit message loop
        }
        try {
            if (!PreTranslateMessage(&m_msg)) {
                ::TranslateMessage(&m_msg);
                ::DispatchMessage(&m_msg);
            }
        } catch (...) {
            ATLTRACE(_T("Unexpected error.\n"));
            continue;
        }

        if (IsIdleMessage(&m_msg)) {
            bDoIdle = TRUE;
            nIdleCount = 0;
        }
    }

    return static_cast<int>(m_msg.wParam);
}
