// DlgProxy.cpp : implementation file
//

#include "stdafx.h"
#include "xls2oul.h"
#include "DlgProxy.h"
#include "xls2oulDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CXls2oulDlgAutoProxy

IMPLEMENT_DYNCREATE(CXls2oulDlgAutoProxy, CCmdTarget)

CXls2oulDlgAutoProxy::CXls2oulDlgAutoProxy()
{
	EnableAutomation();
	
	// To keep the application running as long as an automation 
	//	object is active, the constructor calls AfxOleLockApp.
	AfxOleLockApp();

	// Get access to the dialog through the application's
	//  main window pointer.  Set the proxy's internal pointer
	//  to point to the dialog, and set the dialog's back pointer to
	//  this proxy.
	ASSERT (AfxGetApp()->m_pMainWnd != NULL);
	ASSERT_VALID (AfxGetApp()->m_pMainWnd);
	ASSERT_KINDOF(CXls2oulDlg, AfxGetApp()->m_pMainWnd);
	m_pDialog = (CXls2oulDlg*) AfxGetApp()->m_pMainWnd;
	m_pDialog->m_pAutoProxy = this;
}

CXls2oulDlgAutoProxy::~CXls2oulDlgAutoProxy()
{
	// To terminate the application when all objects created with
	// 	with automation, the destructor calls AfxOleUnlockApp.
	//  Among other things, this will destroy the main dialog
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CXls2oulDlgAutoProxy::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CXls2oulDlgAutoProxy, CCmdTarget)
	//{{AFX_MSG_MAP(CXls2oulDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CXls2oulDlgAutoProxy, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CXls2oulDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IXls2oul to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {F203F1D3-2F2A-4C9E-A169-EED0C5F825D9}
static const IID IID_IXls2oul =
{ 0xf203f1d3, 0x2f2a, 0x4c9e, { 0xa1, 0x69, 0xee, 0xd0, 0xc5, 0xf8, 0x25, 0xd9 } };

BEGIN_INTERFACE_MAP(CXls2oulDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CXls2oulDlgAutoProxy, IID_IXls2oul, Dispatch)
END_INTERFACE_MAP()

// The IMPLEMENT_OLECREATE2 macro is defined in StdAfx.h of this project
// {6770C287-5B05-438D-B852-377249FE2D35}
IMPLEMENT_OLECREATE2(CXls2oulDlgAutoProxy, "Xls2oul.Application", 0x6770c287, 0x5b05, 0x438d, 0xb8, 0x52, 0x37, 0x72, 0x49, 0xfe, 0x2d, 0x35)

/////////////////////////////////////////////////////////////////////////////
// CXls2oulDlgAutoProxy message handlers
