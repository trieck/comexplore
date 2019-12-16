// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__C948BDDA_086B_4731_AE11_F06D2809D133__INCLUDED_)
#define AFX_STDAFX_H__C948BDDA_086B_4731_AE11_F06D2809D133__INCLUDED_

#pragma warning(disable:4996)	// disable deprecation warnings

#include "common.h"

// cms070922: new for wtl8.0 to get atldlgs.h to compile.
#define _WTL_FORWARD_DECLARE_CSTRING

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
