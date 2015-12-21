// CardInfDlg.cpp : 实现文件
//

#define _CRT_SECURE_NO_WARNINGS

#include "stdafx.h"
#include "CardInf.h"
#include "CardInfDlg.h"
#include "afxdialogex.h"
#include "ExcelUtils.h"
#include "Sheet3.h"
#include "PrintPreView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

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
	//students = (Student*)malloc(sizeof(Student) * 50);
	//Student_num = 0;
}

void CCardInfDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST, m_List);
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
	ON_WM_CLOSE()
	ON_BN_CLICKED(IDC_CLEAR_LIST, &CCardInfDlg::OnBnClickedClearList)
	ON_BN_CLICKED(IDC_NewFile, &CCardInfDlg::OnBnClickedNewfile)
	ON_BN_CLICKED(IDC_Print_Preview, &CCardInfDlg::OnBnClickedPrintPreview)
	ON_BN_CLICKED(IDC_BUTTON2, &CCardInfDlg::OnBnClickedButton2)
	ON_WM_CHAR()
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

	ShowWindow(SW_SHOW);
	hDLL = LoadLibrary(_T("OUR_MIFARE.dll"));// 加载DLL

	if (!hDLL)
	{
		MessageBox(_T("DLL 没有加载成功！"), _T("加载失败"), MB_ICONERROR);
		exit(0);
	}

	piccrequest = (ppiccrequest)GetProcAddress(hDLL, "piccrequest");
	pcdbeep = (ppcdbeep)GetProcAddress(hDLL, "pcdbeep");

	m_List.InsertColumn(0, _T("学号"), LVCFMT_CENTER, 100);
	m_List.InsertColumn(1, _T("姓名"), LVCFMT_CENTER, 50);
	m_List.InsertColumn(2, _T("性别"), LVCFMT_CENTER, 40);
	m_List.InsertColumn(3, _T("学院"), LVCFMT_CENTER, 100);
	m_List.InsertColumn(4, _T("专业"), LVCFMT_CENTER, 100);
	m_List.InsertColumn(5, _T("班级"), LVCFMT_CENTER, 100);

	OnBnClickedNewfile();

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

	CString SQL_statement;
	SQL_statement.Format(_T("SELECT * FROM Sheet3 WHERE 物理卡号=\'%s\'"), phy_num_str);

	CSheet3 sheet;
	sheet.Open(AFX_DB_USE_DEFAULT_TYPE, SQL_statement, CRecordset::readOnly);

	if (sheet.GetRecordCount() < 1) {
		pcdbeep(500);
		MessageBox(_T("没有查询到结果！"), _T("产生错误"), MB_ICONERROR);
		CString out;
		out.Format(_T("没有查询到结果！物理卡号为：%s"), phy_num_str);
		Log(out);
	}
	else if (sheet.GetRecordCount() > 1) {
		pcdbeep(500);
		MessageBox(_T("查询结果不唯一！"), _T("产生错误"), MB_ICONERROR);
		CString out;
		out.Format(_T("查询结果不唯一！物理卡号为：%s"), phy_num_str);
		Log(out);
	}
	else {
		pcdbeep(50);

		CString ID_str;

		sheet.GetFieldValue((short)0, ID_str);

		CEdit* edit1 = (CEdit*)GetDlgItem(IDC_ID);
		edit1->SetWindowTextW(ID_str);

		if (txt.m_hFile != CFile::hFileNull) {
			CString out;
			out.Format(_T("查询成功！卡号为：%s"), ID_str);
			txt.Write(out, out.GetLength()*sizeof(wchar_t));
			txt.Flush();
		}

		OnBnClickedgetinf();
	}
	sheet.Close();

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

	CString SQL_statement;
	SQL_statement.Format(_T("SELECT * FROM Sheet3 WHERE 学号=\'%s\'"), ID_str);
	CSheet3 sheet;
	sheet.Open(AFX_DB_USE_DEFAULT_TYPE, SQL_statement, CRecordset::readOnly);

	if (sheet.GetRecordCount() < 1) {
		MessageBox(_T("没有查询到结果！"), _T("产生错误"), MB_ICONERROR);
		CString out;
		out.Format(_T("没有查询到结果！学号为：%s"), ID_str);
		Log(out);
	}
	else if (sheet.GetRecordCount() > 1) {
		MessageBox(_T("查询结果不唯一！"), _T("产生错误"), MB_ICONERROR);
		CString out;
		out.Format(_T("查询结果不唯一！学号为：%s"), ID_str);
		Log(out);
	}
	else {

		CString Name_str, Sex_str, Collage_str, Professionals_str, Class_str;

		sheet.GetFieldValue((short)1, Name_str);
		sheet.GetFieldValue((short)2, Sex_str);
		sheet.GetFieldValue((short)3, Collage_str);
		sheet.GetFieldValue((short)4, Professionals_str);
		sheet.GetFieldValue((short)5, Class_str);

		CEdit* Name = (CEdit*)GetDlgItem(IDC_Name);
		CEdit* Sex = (CEdit*)GetDlgItem(IDC_Sex);
		CEdit* Collage = (CEdit*)GetDlgItem(IDC_Collage);
		CEdit* Professionals = (CEdit*)GetDlgItem(IDC_Professionals);
		CEdit* Class = (CEdit*)GetDlgItem(IDC_Class);

		Name->SetWindowTextW(Name_str);
		Sex->SetWindowTextW(Sex_str);
		Collage->SetWindowTextW(Collage_str);
		Professionals->SetWindowTextW(Professionals_str);
		Class->SetWindowTextW(Class_str);

		if (txt.m_hFile != CFile::hFileNull){
			CString temp,temp2;
			temp.Format(_T("查询成功。学号：%s 姓名：%s 性别：%s 学院：%s 专业：%s 班级：%s"), ID_str, Name_str, Sex_str, Collage_str, Professionals_str, Class_str);
			temp2.Format(_T("%s,%s,%s,%s,%s,%s"), ID_str, Name_str, Sex_str, Collage_str, Professionals_str, Class_str);
			Log(temp);
			WriteCSV(temp2);
		}

		int Count = m_List.GetItemCount();
		m_List.InsertItem(Count, ID_str);
		m_List.SetItemText(Count, 1, Name_str);
		m_List.SetItemText(Count, 2, Sex_str);
		m_List.SetItemText(Count, 3, Collage_str);
		m_List.SetItemText(Count, 4, Professionals_str);
		m_List.SetItemText(Count, 5, Class_str);

		Student temp;
		temp.Name = Name_str;
		temp.Sex = Sex_str;
		temp.Collage = Collage_str;
		temp.Profession = Professionals_str;
		temp.Class = Class_str;

		Student_list.push_back(temp);

// 		students[Student_num].Name = Name_str;
// 		students[Student_num].Sex = Sex_str;
// 		students[Student_num].Collage = Collage_str;
// 		students[Student_num].Profession = Professionals_str;
// 		students[Student_num].Class = Class_str;
// 		Student_num++;


		if (Sex_str == _T("男"))
		{
			StuNum_inf.Male++;
			wchar_t num[5];
			_itow_s(StuNum_inf.Male, num, 5, 10);
			GetDlgItem(IDC_Num_Male)->SetWindowTextW(num);
		}
		else {
			StuNum_inf.Female++;
			wchar_t num[5];
			_itow_s(StuNum_inf.Female, num, 5, 10);
			GetDlgItem(IDC_Num_Female)->SetWindowTextW(num);
		}
		StuNum_inf.total++;
		wchar_t num[5];
		_itow_s(StuNum_inf.total, num, 5, 10);
		GetDlgItem(IDC_Num_Total)->SetWindowTextW(num);
	}
	sheet.Close();

	GetDlgItem(IDC_getInf)->SetWindowTextW(_T("查询完毕"));
	GetDlgItem(IDC_getInf)->EnableWindow(true);
}

void CCardInfDlg::OnBnClickedreadcard()
{
	//卡序列号缓冲
	unsigned char myserial[4];
	unsigned char status;

	status = piccrequest(myserial);

	CString ID_str_last, ID_str;
	GetDlgItem(IDC_phy_number)->GetWindowTextW(ID_str_last);

	//返回值处理
	//调用读卡函数，如果没有寻到卡返回1，拿卡太快返回2，没注册发卡机返回4，没有驱动程序返回3
	switch (status)
	{
	case 0:
	{
		double cardNumBer;
		cardNumBer = myserial[3];
		cardNumBer = cardNumBer * 256;
		cardNumBer = cardNumBer + myserial[2];
		cardNumBer = cardNumBer * 256;
		cardNumBer = cardNumBer + myserial[1];
		cardNumBer = cardNumBer * 256;
		cardNumBer = cardNumBer + myserial[0];

		// 获取输入框对象
		ID_str.Format(_T("%d"), (int)cardNumBer);
		GetDlgItem(IDC_phy_number)->SetWindowTextW(ID_str);
		break;
	}
	case ERR_REQUEST:
		MessageBox(_T("没有寻找到卡！"), _T("找不到卡片"), MB_ICONERROR);
		break;
	case ERR_NONEDLL:
		MessageBox(_T("找不到动态库！"), _T("找不到动态库"), MB_ICONERROR);
		break;
	case ERR_DRIVERNULL:
		MessageBox(_T("驱动程序没找到或安装错误！"), _T("找不到驱动"), MB_ICONERROR);
		break;
	default:
		MessageBox(_T("出现了未知异常！"), _T("找不到驱动"), MB_ICONERROR);
		break;
	}

	if (ID_str_last != ID_str && ID_str != "") {
		pcdbeep(50);
		OnBnClickedSearchId();
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
	}
	else {
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

		status = piccrequest(myserial);

		CString ID_str_last, ID_str;
		GetDlgItem(IDC_phy_number)->GetWindowTextW(ID_str_last);

		if (status == 0)
		{
			double cardNumBer;
			cardNumBer = myserial[3];
			cardNumBer = cardNumBer * 256;
			cardNumBer = cardNumBer + myserial[2];
			cardNumBer = cardNumBer * 256;
			cardNumBer = cardNumBer + myserial[1];
			cardNumBer = cardNumBer * 256;
			cardNumBer = cardNumBer + myserial[0];

			// 获取输入框对象
			ID_str.Format(_T("%010.0lf"), cardNumBer);
			GetDlgItem(IDC_phy_number)->SetWindowTextW(ID_str);
		}

		if (ID_str_last != ID_str && ID_str != "") {
			OnBnClickedSearchId();
		}

	}
	default:
		break;
	}

	CDialogEx::OnTimer(nIDEvent);
}


void CCardInfDlg::OnClose()
{
	FreeLibrary(hDLL);

	CDialogEx::OnClose();
}

void CCardInfDlg::OnBnClickedClearList()
{
	for (int i = m_List.GetItemCount(); i >= 0; i--)
		m_List.DeleteItem(i);

	Student_list.clear();
	StuNum_inf = { 0,0,0 };
	GetDlgItem(IDC_Num_Male)->SetWindowTextW(_T("0"));
	GetDlgItem(IDC_Num_Female)->SetWindowTextW(_T("0"));
	GetDlgItem(IDC_Num_Total)->SetWindowTextW(_T("0"));
}

CString CCardInfDlg::CreateNewFile() {
	SYSTEMTIME sys;
	GetLocalTime(&sys);

	CString path;
	path.Format(_T("%s%d%d%d-%02d%02d%02d.txt"), GetProgramCurrentPath(), sys.wYear, sys.wMonth, sys.wDay, sys.wHour, sys.wMinute, sys.wSecond);
	if (!txt.Open(path, CFile::modeCreate | CFile::modeReadWrite))
		MessageBox(_T("文件打开失败！"), _T("文件打开失败"), MB_ICONERROR);
	else
		txt.Write("\xff\xfe", 2);

	CString path2;
	path2.Format(_T("%s%d%d%d-%02d%02d%02d.csv"), GetProgramCurrentPath(), sys.wYear, sys.wMonth, sys.wDay, sys.wHour, sys.wMinute, sys.wSecond);
	if (!csv.Open(path2, CFile::modeCreate | CFile::modeReadWrite))
		MessageBox(_T("文件打开失败！"), _T("文件打开失败"), MB_ICONERROR);
	else
		csv.Write("\xff\xfe", 2);

	return path;
}


void CCardInfDlg::OnBnClickedNewfile()
{
	if (txt.m_hFile != CFile::hFileNull) {
		txt.Close();
	}
	if (csv.m_hFile != CFile::hFileNull) {
		csv.Close();
	}
	GetDlgItem(IDC_NewFile)->SetWindowText(CreateNewFile());
}

void CCardInfDlg::Log(CString Content)
{
	SYSTEMTIME sys;
	GetLocalTime(&sys);
	CString temp;
	temp.Format(_T("[%d-%d-%d,%02d:%02d:%02d]"), sys.wYear, sys.wMonth, sys.wDay, sys.wHour, sys.wMinute, sys.wSecond);
	temp += Content;
	txt.Write(temp, temp.GetLength()*sizeof(wchar_t));
	txt.Write(_T("\r\n"), 2*sizeof(wchar_t));
	txt.Flush();
}

void CCardInfDlg::WriteCSV(CString Content)
{
	csv.Write(Content, Content.GetLength()*sizeof(wchar_t));
	csv.Write(_T("\r\n"), 2 * sizeof(wchar_t));
	csv.Flush();
}


void CCardInfDlg::OnBnClickedPrintPreview()
{
	CString CompetitionName, CompanyName;
	GetDlgItem(IDC_CompetitionName)->GetWindowTextW(CompetitionName);
	GetDlgItem(IDC_CompanyName)->GetWindowTextW(CompanyName);
	CPrintPreView PrintPreviewDlg(&Student_list, Student_num, CompetitionName.GetBuffer(), CompanyName.GetBuffer(), StuNum_inf);
	PrintPreviewDlg.DoModal();
}


void CCardInfDlg::OnBnClickedButton2()
{
#ifdef _DEBUG
	for (int i = 0; i < 100; i++)
	{
		// 获取输入框对象
		CEdit* edit = (CEdit*)GetDlgItem(IDC_ID);
		CString ID_str;
		double temp = 201510098016;
		ID_str.Format(_T("%.0lf"), temp + i);
		edit->SetWindowTextW(ID_str);
		OnBnClickedgetinf();
	}
#endif

}


BOOL CCardInfDlg::PreTranslateMessage(MSG* pMsg)
{
	if (WM_KEYDOWN == pMsg->message)
	{
		CEdit* pEdit = (CEdit*)GetDlgItem(IDC_ID);
		ASSERT(pEdit);
		if (pMsg->hwnd == pEdit->GetSafeHwnd() && VK_RETURN == pMsg->wParam)
		{
			OnBnClickedgetinf();
			return TRUE;
		}
	}

	return CDialogEx::PreTranslateMessage(pMsg);
}
