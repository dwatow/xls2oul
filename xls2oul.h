// xls2oul.h : main header file for the XLS2OUL application
//

#if !defined(AFX_XLS2OUL_H__8334AA6C_287D_4AF4_BFCB_3903FBEB0C62__INCLUDED_)
#define AFX_XLS2OUL_H__8334AA6C_287D_4AF4_BFCB_3903FBEB0C62__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CXls2oulApp:
// See xls2oul.cpp for the implementation of this class
//

class CXls2oulApp : public CWinApp
{
public:
	CXls2oulApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CXls2oulApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CXls2oulApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_XLS2OUL_H__8334AA6C_287D_4AF4_BFCB_3903FBEB0C62__INCLUDED_)
