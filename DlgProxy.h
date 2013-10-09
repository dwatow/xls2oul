// DlgProxy.h : header file
//

#if !defined(AFX_DLGPROXY_H__4663DFE0_8461_4C1B_92C4_44A1C362C3A4__INCLUDED_)
#define AFX_DLGPROXY_H__4663DFE0_8461_4C1B_92C4_44A1C362C3A4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CXls2oulDlg;

/////////////////////////////////////////////////////////////////////////////
// CXls2oulDlgAutoProxy command target

class CXls2oulDlgAutoProxy : public CCmdTarget
{
	DECLARE_DYNCREATE(CXls2oulDlgAutoProxy)

	CXls2oulDlgAutoProxy();           // protected constructor used by dynamic creation

// Attributes
public:
	CXls2oulDlg* m_pDialog;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CXls2oulDlgAutoProxy)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual ~CXls2oulDlgAutoProxy();

	// Generated message map functions
	//{{AFX_MSG(CXls2oulDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(CXls2oulDlgAutoProxy)

	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CXls2oulDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGPROXY_H__4663DFE0_8461_4C1B_92C4_44A1C362C3A4__INCLUDED_)
