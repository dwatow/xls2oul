// xls2oulDlg.h : header file
//

#if !defined(AFX_XLS2OULDLG_H__E330877C_F09A_481C_86EF_860FA8D5E23D__INCLUDED_)
#define AFX_XLS2OULDLG_H__E330877C_F09A_481C_86EF_860FA8D5E23D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
#include "meeting.h"

class CXls2oulDlgAutoProxy;

/////////////////////////////////////////////////////////////////////////////
// CXls2oulDlg dialog

class CXls2oulDlg : public CDialog
{
	std::vector<meeting> tranMeeting;

	DECLARE_DYNAMIC(CXls2oulDlg);
	friend class CXls2oulDlgAutoProxy;

// Construction
public:
	CXls2oulDlg(CWnd* pParent = NULL);	// standard constructor
	virtual ~CXls2oulDlg();

	void ReadExl(const CString& path);
	void WriteOul();

// Dialog Data
	//{{AFX_DATA(CXls2oulDlg)
	enum { IDD = IDD_XLS2OUL_DIALOG };
	CProgressCtrl	m_bar;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CXls2oulDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CXls2oulDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// Generated message map functions
	//{{AFX_MSG(CXls2oulDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	afx_msg void OnTran();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_XLS2OULDLG_H__E330877C_F09A_481C_86EF_860FA8D5E23D__INCLUDED_)
