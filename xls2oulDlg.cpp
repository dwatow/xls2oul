// xls2oulDlg.cpp : implementation file
//

#include "stdafx.h"
#include "xls2oul.h"
#include "xls2oulDlg.h"
#include "DlgProxy.h"
#include "xlef.h"
#include "Oft.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CXls2oulDlg dialog

IMPLEMENT_DYNAMIC(CXls2oulDlg, CDialog);

CXls2oulDlg::CXls2oulDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CXls2oulDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CXls2oulDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
}

CXls2oulDlg::~CXls2oulDlg()
{
	// If there is an automation proxy for this dialog, set
	//  its back pointer to this dialog to NULL, so it knows
	//  the dialog has been deleted.
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;
}

void CXls2oulDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CXls2oulDlg)
	DDX_Control(pDX, IDC_BAR, m_bar);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CXls2oulDlg, CDialog)
	//{{AFX_MSG_MAP(CXls2oulDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_CLOSE()
	ON_BN_CLICKED(ID_TRAN, OnTran)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CXls2oulDlg message handlers

BOOL CXls2oulDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CXls2oulDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CXls2oulDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CXls2oulDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

// Automation servers should not exit when a user closes the UI
//  if a controller still holds on to one of its objects.  These
//  message handlers make sure that if the proxy is still in use,
//  then the UI is hidden but the dialog remains around if it
//  is dismissed.

void CXls2oulDlg::OnClose() 
{
	if (CanExit())
		CDialog::OnClose();
}

void CXls2oulDlg::OnOK() 
{
	if (CanExit())
		CDialog::OnOK();
}

void CXls2oulDlg::OnCancel() 
{
	if (CanExit())
		CDialog::OnCancel();
}

BOOL CXls2oulDlg::CanExit()
{
	// If the proxy object is still around, then the automation
	//  controller is still holding on to this application.  Leave
	//  the dialog around, but hide its UI.
	if (m_pAutoProxy != NULL)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}

void CXls2oulDlg::ReadExl(const CString& path)
{
	xlsFile ReadFile;
	
	ReadFile.Open(path);
	ReadFile.SelectSheet(1);
		
	int indexXls = 3;
	int totalXlsData = ReadFile.GetVrticlTotalCell() - indexXls+1;

	m_bar.SetRange(0, totalXlsData*2);
	m_bar.OffsetPos(0);
	m_bar.SetStep(1);

	int indexSafe = 0;
	do 
	{
		meeting currentMeeting;
		
		currentMeeting.SetSubject( ReadFile.SelectCell('A', indexXls)->GetCell2CStr());
		currentMeeting.SetDate(    ReadFile.SelectCell('B', indexXls)->GetCell2Int(), 
			ReadFile.SelectCell('C', indexXls)->GetCell2Int(), 
			ReadFile.SelectCell('D', indexXls)->GetCell2Int() );
		
		currentMeeting.SetTime(    ReadFile.SelectCell('E', indexXls)->GetCell2Int(), 
			ReadFile.SelectCell('F', indexXls)->GetCell2Int(), 
			ReadFile.SelectCell('G', indexXls)->GetCell2Int() );
		currentMeeting.SetPlace(   ReadFile.SelectCell('H', indexXls)->GetCell2CStr());
		currentMeeting.SetHowLongBeforeStart(ReadFile.SelectCell('I', indexXls)->GetCell2Int());
		currentMeeting.SetHowlongActive(ReadFile.SelectCell('J', indexXls)->GetCell2Int());
		tranMeeting.push_back(currentMeeting);
		indexXls++;
		indexSafe++;
		m_bar.SetPos(indexSafe);
		
		//避免當機的安全值
		if(indexSafe == 100)
			break;
	}
	while(!ReadFile.SelectCell('A', indexXls)->GetCell2CStr().IsEmpty());
}

void CXls2oulDlg::WriteOul()
{
	olkFile WriteFile;

	int indexCurrentBar = m_bar.GetPos();
	for (std::vector<meeting>::iterator itor = tranMeeting.begin(); itor != tranMeeting.end(); ++itor)
	{
		WriteFile.AddMeeting(itor->GetSubject(), 
			           (DATE)itor->GetDateAndTime(), 
					         itor->GetPlace(), 
							 itor->GetMinutesBeforeStart(), 
							 itor->GetMinutesActive(), 
							 itor->GetDescription());
		++indexCurrentBar;
		m_bar.SetPos(indexCurrentBar);
	}

	tranMeeting.clear();
}

void CXls2oulDlg::OnTran() 
{
	// TODO: Add your control notification handler code here
	CFileDialog dlgFile(TRUE);
	if (dlgFile.DoModal() == IDOK)
	{
		ReadExl(dlgFile.GetPathName());
		WriteOul();
	}
}
