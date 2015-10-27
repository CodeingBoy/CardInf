// CardInfDlg.cpp : ʵ���ļ�
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

#define STU_INF_EXCEL "ѧ������Ϣ.xls" // ������ѧ�Ŷ�Ӧ�ı������
#define ID_INF_EXCEL "10-261.xls"	// ѧ����Ϣ�������

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

	ShowWindow(SW_MINIMIZE);

	// TODO: �ڴ���Ӷ���ĳ�ʼ������

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

	GetDlgItem(IDC_getInf)->SetWindowTextW(_T("��ѯ���"));
	GetDlgItem(IDC_getInf)->EnableWindow(true);
}

void CCardInfDlg::OnBnClickedreadcard()
{
	//�����кŻ���
	unsigned char myserial[4];
	unsigned char status;
	/*if (!FileExists(FileName))
	{//����ļ�������
		ShowMessageb("�޷���Ӧ�ó�����ļ����ҵ�IC����д������̬��");
		return; //����
	}*/

	typedef unsigned char(__stdcall *ppiccrequest)(unsigned char* serial);
	HINSTANCE hDLL = LoadLibrary(_T("OUR_MIFARE.dll"));// ����DLL

	ppiccrequest piccrequest;
	piccrequest = (ppiccrequest)GetProcAddress(hDLL, "piccrequest");

	//��ȡ��̬��
	//piccrequest = (unsigned char(__stdcall *piccrequest)(unsigned char *serial))GetProcAddress(hDll, "piccread");


	status = piccrequest(myserial);
	//����ֵ����
	//���ö������������û��Ѱ��������1���ÿ�̫�췵��2��ûע�ᷢ��������4��û���������򷵻�3
	switch (status)
	{
	case '0':
	{
		// ��ȡ��������
		CEdit* edit = (CEdit*)GetDlgItem(IDC_phy_number);
		CString ID_str;
		ID_str.Format(_T("%s"), myserial);
		edit->SetWindowTextW(ID_str);
		break;
	}
	case '1':
		MessageBox(_T("û��Ѱ�ҵ�����"));
		break;
	case '2':
		MessageBox(_T("�ÿ�̫�죡"));
		break;
	case '3':
		MessageBox(_T("û����������"));
		break;
	case '4':
		MessageBox(_T("ûע�ᷢ������"));
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
	if (btntext == _T("ֹͣѭ������"))
	{
		KillTimer(1);
		GetDlgItem(IDC_read_card_cycle)->SetWindowTextW(_T("ѭ������"));
	}else{
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
		/*if (!FileExists(FileName))
		{//����ļ�������
		ShowMessageb("�޷���Ӧ�ó�����ļ����ҵ�IC����д������̬��");
		return; //����
		}*/

		typedef unsigned char(__stdcall *ppiccrequest)(unsigned char* serial);
		HINSTANCE hDLL = LoadLibrary(_T("OUR_MIFARE.dll"));// ����DLL

		ppiccrequest piccrequest;
		piccrequest = (ppiccrequest)GetProcAddress(hDLL, "piccrequest");

		//��ȡ��̬��
		//piccrequest = (unsigned char(__stdcall *piccrequest)(unsigned char *serial))GetProcAddress(hDll, "piccread");

		while (true)
		{
			status = piccrequest(myserial);
			//����ֵ����
			//���ö������������û��Ѱ��������1���ÿ�̫�췵��2��ûע�ᷢ��������4��û���������򷵻�3
			switch (status)
			{
			case '0':
			{
				// ��ȡ��������
				CEdit* edit = (CEdit*)GetDlgItem(IDC_phy_number);
				CString ID_str;
				ID_str.Format(_T("%s"), myserial);
				edit->SetWindowTextW(ID_str);
				break;
			}
			case '1':
				//MessageBox(_T("û��Ѱ�ҵ�����"));
				break;
			case '2':
				//MessageBox(_T("�ÿ�̫�죡"));
				break;
			case '3':
				//MessageBox(_T("û����������"));
				break;
			case '4':
				//MessageBox(_T("ûע�ᷢ������"));
				break;
			}
		}
	}
	default:
		break;
	}

	CDialogEx::OnTimer(nIDEvent);
}
