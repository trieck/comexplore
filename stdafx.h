// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__C948BDDA_086B_4731_AE11_F06D2809D133__INCLUDED_)
#define AFX_STDAFX_H__C948BDDA_086B_4731_AE11_F06D2809D133__INCLUDED_

// Change these values to use different versions
#define WINVER		0x0400
//#define _WIN32_WINNT	0x0400
#define _WIN32_IE	0x0400
#define _RICHEDIT_VER	0x0100

#include "common.h"

#include <atlbase.h>
#include <atlapp.h>
#include <atlstr.h>
#include <atlcrack.h>		// WTL enhanced message map macros

typedef CStringT<TCHAR, StrTraitATL<TCHAR> > TString;

extern CAppModule _Module;

#include <atlwin.h>

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__C948BDDA_086B_4731_AE11_F06D2809D133__INCLUDED_)
