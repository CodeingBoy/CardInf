// CardInfDlg.cpp : 实现文件
//

#define _CRT_SECURE_NO_WARNINGS

#include "stdafx.h"
#include "CardInf.h"
#include "CardInfDlg.h"
#include "afxdialogex.h"
#include "ExcelUtils.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define STU_INF_EXCEL "学生卡信息.xls" // 物理卡和学号对应的表格名称
#define ID_INF_EXCEL "10-261.xls"	// 学生信息表格名称

// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

	// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CCardInfDlg 对话框



CCardInfDlg::CCardInfDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_CARDINF_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CCardInfDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CCardInfDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_Search_ID, &CCardInfDlg::OnBnClickedSearchId)
	ON_BN_CLICKED(IDC_getInf, &CCardInfDlg::OnBnClickedgetinf)
	ON_BN_CLICKED(IDC_read_card, &CCardInfDlg::OnBnClickedreadcard)
	ON_BN_CLICKED(IDC_read_card_cycle, &CCardInfDlg::OnBnClickedreadcardcycle)
	ON_WM_TIMER()
END_MESSAGE_MAP()


// CCardInfDlg 消息处理程序

BOOL CCardInfDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	ShowWindow(SW_MINIMIZE);

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CCardInfDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CCardInfDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CCardInfDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CCardInfDlg::OnBnClickedSearchId()
{
	GetDlgItem(IDC_Search_ID)->SetWindowTextW(_T("查询中……"));
	GetDlgItem(IDC_Search_ID)->EnableWindow(false);

	CEdit* edit = (CEdit*)GetDlgItem(IDC_phy_number);
	CString phy_num_str;
	edit->GetWindowTextW(phy_num_str);

	CString ExcelLocation = GetProgramCurrentPath();
	ExcelLocation += STU_INF_EXCEL;

	CExcelUtils phy_inf;
	phy_inf.InitExcel();
	phy_inf.OpenExcelFile(ExcelLocation);
	phy_inf.LoadSheet(1, true);
	wchar_t *phy_num_wchar = new wchar_t[256];
	long r;
	wcscpy_s(phy_num_wchar, 256, phy_num_str);
	if (phy_inf.FindRow(phy_num_wchar, &r, 1)) {
		CEdit* edit = (CEdit*)GetDlgItem(IDC_ID);
		edit->SetWindowTextW(phy_inf.GetCellString(r, 2));
	}
	phy_inf.CloseExcelFile(false);
	phy_inf.ReleaseExcel();

	GetDlgItem(IDC_Search_ID)->SetWindowTextW(_T("查询完毕"));
	GetDlgItem(IDC_Search_ID)->EnableWindow(true);
}


void CCardInfDlg::OnBnClickedgetinf()
{
	// 改变按钮文字和状态
	GetDlgItem(IDC_getInf)->SetWindowTextW(_T("查询中……"));
	GetDlgItem(IDC_getInf)->EnableWindow(false);

	// 获取输入框对象
	CEdit* edit = (CEdit*)GetDlgItem(IDC_ID);
	CString ID_str;
	edit->GetWindowTextW(ID_str);

	CString ExcelLocation = GetProgramCurrentPath();
	ExcelLocation += ID_INF_EXCEL;

	CExcelUtils Stu_inf;
	Stu_inf.InitExcel();
	Stu_inf.OpenExcelFile(ExcelLocation);
	Stu_inf.LoadSheet(1, true);
	wchar_t *student_ID = new wchar_t[256];
	long r;
	wcscpy_s(student_ID, 256, ID_str); // Cstring->wchar_t
	if (Stu_inf.FindRow(student_ID, &r, 1)) {
		CEdit* Name = (CEdit*)GetDlgItem(IDC_Name);
		CEdit* Sex = (CEdit*)GetDlgItem(IDC_Sex);
		CEdit* Collage = (CEdit*)GetDlgItem(IDC_Collage);
		CEdit* Professionals = (CEdit*)GetDlgItem(IDC_Professionals);
		CEdit* Class = (CEdit*)GetDlgItem(IDC_Class);

		Name->SetWindowTextW(Stu_inf.GetCellString(r, 2));
		Sex->SetWindowTextW(Stu_inf.GetCellString(r, 3));
		Collage->SetWindowTextW(Stu_inf.GetCellString(r, 4));
		Professionals->SetWindowTextW(Stu_inf.GetCellString(r, 5));
		Class->SetWindowTextW(Stu_inf.GetCellString(r, 6));
	}
	Stu_inf.CloseExcelFile(false);
	Stu_inf.ReleaseExcel();

	GetDlgItem(IDC_getInf)->SetWindowTextW(_T("查询完毕"));
	GetDlgItem(IDC_getInf)->EnableWindow(true);
}

void CCardInfDlg::OnBnClickedreadcard()
{
	//卡序列号缓冲
	unsigned char myserial[4];
	unsigned char status;
	/*if (!FileExists(FileName))
	{//如果文件不存在
		ShowMessageb("无法在应用程序的文件夹找到IC卡读写卡器动态库");
		return; //返回
	}*/

	typedef unsigned char(__stdcall *ppiccrequest)(unsigned char* serial);
	HINSTANCE hDLL = LoadLibrary(_T("OUR_MIFARE.dll"));// 加载DLL

	ppiccrequest piccrequest;
	piccrequest = (ppiccrequest)GetProcAddress(hDLL, "piccrequest");

	//提取动态库
	//piccrequest = (unsigned char(__stdcall *piccrequest)(unsigned char *serial))GetProcAddress(hDll, "piccread");


	status = piccrequest(myserial);
	//返回值处理
	//调用读卡函数，如果没有寻到卡返回1，拿卡太快返回2，没注册发卡机返回4，没有驱动程序返回3
	switch (status)
	{
	case '0':
	{
		// 获取输入框对象
		CEdit* edit = (CEdit*)GetDlgItem(IDC_phy_number);
		CString ID_str;
		ID_str.Format(_T("%s"), myserial);
		edit->SetWindowTextW(ID_str);
		break;
	}
	case '1':
		MessageBox(_T("没有寻找到卡！"));
		break;
	case '2':
		MessageBox(_T("拿卡太快！"));
		break;
	case '3':
		MessageBox(_T("没有驱动程序！"));
		break;
	case '4':
		MessageBox(_T("没注册发卡机！"));
		break;
	}

}

CString CCardInfDlg::GetProgramCurrentPath(void)
{
	TCHAR path_buffer[_MAX_PATH];
	TCHAR drive[_MAX_DRIVE];
	TCHAR dir[_MAX_DIR];
	TCHAR fname[_MAX_FNAME];
	TCHAR ext[_MAX_EXT];

	GetModuleFileName(NULL, path_buffer, _MAX_PATH);
	_wsplitpath_s(path_buffer, drive, dir, fname, ext);
	CString strDir;
	strDir += drive;
	strDir += dir;
	return strDir;
}

void CCardInfDlg::OnBnClickedreadcardcycle()
{
	CString btntext;
	GetDlgItem(IDC_read_card_cycle)->GetWindowTextW(btntext);
	if (btntext == _T("停止循环读卡"))
	{
		KillTimer(1);
		GetDlgItem(IDC_read_card_cycle)->SetWindowTextW(_T("循环读卡"));
	}else{
		SetTimer(1, 500, NULL);
		GetDlgItem(IDC_read_card_cycle)->SetWindowTextW(_T("停止循环读卡"));
	}
}



void CCardInfDlg::OnTimer(UINT_PTR nIDEvent)
{
	switch (nIDEvent)
	{
	case 1:
	{
		//卡序列号缓冲
		unsigned char myserial[4];
		unsigned char status;
		/*if (!FileExists(FileName))
		{//如果文件不存在
		ShowMessageb("无法在应用程序的文件夹找到IC卡读写卡器动态库");
		return; //返回
		}*/

		typedef unsigned char(__stdcall *ppiccrequest)(unsigned char* serial);
		HINSTANCE hDLL = LoadLibrary(_T("OUR_MIFARE.dll"));// 加载DLL

		ppiccrequest piccrequest;
		piccrequest = (ppiccrequest)GetProcAddress(hDLL, "piccrequest");

		//提取动态库
		//piccrequest = (unsigned char(__stdcall *piccrequest)(unsigned char *serial))GetProcAddress(hDll, "piccread");

		while (true)
		{
			status = piccrequest(myserial);
			//返回值处理
			//调用读卡函数，如果没有寻到卡返回1，拿卡太快返回2，没注册发卡机返回4，没有驱动程序返回3
			switch (status)
			{
			case '0':
			{
				// 获取输入框对象
				CEdit* edit = (CEdit*)GetDlgItem(IDC_phy_number);
				CString ID_str;
				ID_str.Format(_T("%s"), myserial);
				edit->SetWindowTextW(ID_str);
				break;
			}
			case '1':
				//MessageBox(_T("没有寻找到卡！"));
				break;
			case '2':
				//MessageBox(_T("拿卡太快！"));
				break;
			case '3':
				//MessageBox(_T("没有驱动程序！"));
				break;
			case '4':
				//MessageBox(_T("没注册发卡机！"));
				break;
			}
		}
	}
	default:
		break;
	}

	CDialogEx::OnTimer(nIDEvent);
}
