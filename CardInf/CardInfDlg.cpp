// CardInfDlg.cpp : ʵ���ļ�
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

// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

	// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CCardInfDlg �Ի���



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


// CCardInfDlg ��Ϣ�������

BOOL CCardInfDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	ShowWindow(SW_SHOW);
	hDLL = LoadLibrary(_T("OUR_MIFARE.dll"));// ����DLL

	if (!hDLL)
	{
		MessageBox(_T("DLL û�м��سɹ���"), _T("����ʧ��"), MB_ICONERROR);
		exit(0);
	}

	piccrequest = (ppiccrequest)GetProcAddress(hDLL, "piccrequest");
	pcdbeep = (ppcdbeep)GetProcAddress(hDLL, "pcdbeep");

	m_List.InsertColumn(0, _T("ѧ��"), LVCFMT_CENTER, 100);
	m_List.InsertColumn(1, _T("����"), LVCFMT_CENTER, 50);
	m_List.InsertColumn(2, _T("�Ա�"), LVCFMT_CENTER, 40);
	m_List.InsertColumn(3, _T("ѧԺ"), LVCFMT_CENTER, 100);
	m_List.InsertColumn(4, _T("רҵ"), LVCFMT_CENTER, 100);
	m_List.InsertColumn(5, _T("�༶"), LVCFMT_CENTER, 100);

	OnBnClickedNewfile();

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CCardInfDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CCardInfDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CCardInfDlg::OnBnClickedSearchId()
{
	GetDlgItem(IDC_Search_ID)->SetWindowTextW(_T("��ѯ�С���"));
	GetDlgItem(IDC_Search_ID)->EnableWindow(false);

	CEdit* edit = (CEdit*)GetDlgItem(IDC_phy_number);
	CString phy_num_str;
	edit->GetWindowTextW(phy_num_str);

	CString SQL_statement;
	SQL_statement.Format(_T("SELECT * FROM Sheet3 WHERE ������=\'%s\'"), phy_num_str);

	CSheet3 sheet;
	sheet.Open(AFX_DB_USE_DEFAULT_TYPE, SQL_statement, CRecordset::readOnly);

	if (sheet.GetRecordCount() < 1) {
		pcdbeep(500);
		MessageBox(_T("û�в�ѯ�������"), _T("��������"), MB_ICONERROR);
		CString out;
		out.Format(_T("û�в�ѯ�������������Ϊ��%s"), phy_num_str);
		Log(out);
	}
	else if (sheet.GetRecordCount() > 1) {
		pcdbeep(500);
		MessageBox(_T("��ѯ�����Ψһ��"), _T("��������"), MB_ICONERROR);
		CString out;
		out.Format(_T("��ѯ�����Ψһ��������Ϊ��%s"), phy_num_str);
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
			out.Format(_T("��ѯ�ɹ�������Ϊ��%s"), ID_str);
			txt.Write(out, out.GetLength()*sizeof(wchar_t));
			txt.Flush();
		}

		OnBnClickedgetinf();
	}
	sheet.Close();

	GetDlgItem(IDC_Search_ID)->SetWindowTextW(_T("��ѯ���"));
	GetDlgItem(IDC_Search_ID)->EnableWindow(true);
}


void CCardInfDlg::OnBnClickedgetinf()
{
	// �ı䰴ť���ֺ�״̬
	GetDlgItem(IDC_getInf)->SetWindowTextW(_T("��ѯ�С���"));
	GetDlgItem(IDC_getInf)->EnableWindow(false);

	// ��ȡ��������
	CEdit* edit = (CEdit*)GetDlgItem(IDC_ID);
	CString ID_str;
	edit->GetWindowTextW(ID_str);

	CString SQL_statement;
	SQL_statement.Format(_T("SELECT * FROM Sheet3 WHERE ѧ��=\'%s\'"), ID_str);
	CSheet3 sheet;
	sheet.Open(AFX_DB_USE_DEFAULT_TYPE, SQL_statement, CRecordset::readOnly);

	if (sheet.GetRecordCount() < 1) {
		MessageBox(_T("û�в�ѯ�������"), _T("��������"), MB_ICONERROR);
		CString out;
		out.Format(_T("û�в�ѯ�������ѧ��Ϊ��%s"), ID_str);
		Log(out);
	}
	else if (sheet.GetRecordCount() > 1) {
		MessageBox(_T("��ѯ�����Ψһ��"), _T("��������"), MB_ICONERROR);
		CString out;
		out.Format(_T("��ѯ�����Ψһ��ѧ��Ϊ��%s"), ID_str);
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
			temp.Format(_T("��ѯ�ɹ���ѧ�ţ�%s ������%s �Ա�%s ѧԺ��%s רҵ��%s �༶��%s"), ID_str, Name_str, Sex_str, Collage_str, Professionals_str, Class_str);
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


		if (Sex_str == _T("��"))
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

	GetDlgItem(IDC_getInf)->SetWindowTextW(_T("��ѯ���"));
	GetDlgItem(IDC_getInf)->EnableWindow(true);
}

void CCardInfDlg::OnBnClickedreadcard()
{
	//�����кŻ���
	unsigned char myserial[4];
	unsigned char status;

	status = piccrequest(myserial);

	CString ID_str_last, ID_str;
	GetDlgItem(IDC_phy_number)->GetWindowTextW(ID_str_last);

	//����ֵ����
	//���ö������������û��Ѱ��������1���ÿ�̫�췵��2��ûע�ᷢ��������4��û���������򷵻�3
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

		// ��ȡ��������
		ID_str.Format(_T("%d"), (int)cardNumBer);
		GetDlgItem(IDC_phy_number)->SetWindowTextW(ID_str);
		break;
	}
	case ERR_REQUEST:
		MessageBox(_T("û��Ѱ�ҵ�����"), _T("�Ҳ�����Ƭ"), MB_ICONERROR);
		break;
	case ERR_NONEDLL:
		MessageBox(_T("�Ҳ�����̬�⣡"), _T("�Ҳ�����̬��"), MB_ICONERROR);
		break;
	case ERR_DRIVERNULL:
		MessageBox(_T("��������û�ҵ���װ����"), _T("�Ҳ�������"), MB_ICONERROR);
		break;
	default:
		MessageBox(_T("������δ֪�쳣��"), _T("�Ҳ�������"), MB_ICONERROR);
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
	if (btntext == _T("ֹͣѭ������"))
	{
		KillTimer(1);
		GetDlgItem(IDC_read_card_cycle)->SetWindowTextW(_T("ѭ������"));
	}
	else {
		SetTimer(1, 500, NULL);
		GetDlgItem(IDC_read_card_cycle)->SetWindowTextW(_T("ֹͣѭ������"));
	}
}



void CCardInfDlg::OnTimer(UINT_PTR nIDEvent)
{
	switch (nIDEvent)
	{
	case 1:
	{
		//�����кŻ���
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

			// ��ȡ��������
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
		MessageBox(_T("�ļ���ʧ�ܣ�"), _T("�ļ���ʧ��"), MB_ICONERROR);
	else
		txt.Write("\xff\xfe", 2);

	CString path2;
	path2.Format(_T("%s%d%d%d-%02d%02d%02d.csv"), GetProgramCurrentPath(), sys.wYear, sys.wMonth, sys.wDay, sys.wHour, sys.wMinute, sys.wSecond);
	if (!csv.Open(path2, CFile::modeCreate | CFile::modeReadWrite))
		MessageBox(_T("�ļ���ʧ�ܣ�"), _T("�ļ���ʧ��"), MB_ICONERROR);
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
		// ��ȡ��������
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
