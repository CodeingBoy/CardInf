// PrintPreView.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "CardInf.h"
#include "PrintPreView.h"
#include "afxdialogex.h"
#include "CardInfDlg.h"
#include <list>

// CPrintPreView �Ի���

IMPLEMENT_DYNAMIC(CPrintPreView, CDialogEx)

CPrintPreView::CPrintPreView(std::list<Student> *StudentsList, int Number, wchar_t* Competition_Name, wchar_t* Comapany_Name, Student_Number_Inf num_inf, CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_DIALOG1, pParent)
{
	AfxInitRichEdit2();
	CPrintPreView::StudentsList = StudentsList;
	CPrintPreView::Number = Number;
	CPrintPreView::Comapany_Name = Comapany_Name;
	CPrintPreView::Competition_Name = Competition_Name;
	CPrintPreView::num_inf = num_inf;
}

CPrintPreView::~CPrintPreView()
{

}

void CPrintPreView::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	//  DDX_Control(pDX, IDC_RichEdit, Text);
	DDX_Control(pDX, IDC_RichEdit, m_Text);
}


BEGIN_MESSAGE_MAP(CPrintPreView, CDialogEx)
	ON_BN_CLICKED(IDC_Export, &CPrintPreView::OnBnClickedExport)
END_MESSAGE_MAP()


// CPrintPreView ��Ϣ�������


BOOL CPrintPreView::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	CRichEditCtrl *edit = (CRichEditCtrl*)GetDlgItem(IDC_RichEdit);

	BuildText();

	return TRUE;  // return TRUE unless you set the focus to a control
				  // �쳣: OCX ����ҳӦ���� FALSE
}

void CPrintPreView::BuildText() {

	CString temp;

	//temp.Format(_T("��������ѧ����ѧԺ\r\n2015��У�˻� %s��Ŀ\r\n%s��������\r\n\n"), Competition_Name, Comapany_Name);
	temp.Format(_T("��������ѧ����ѧԺ\r\n2015����������������������\r\n%s�μ�����\r\n\n"), Comapany_Name);
	Outputf(CF_TITLE, PF_CENTER, temp);

	time_t nowtime;
	nowtime = time(NULL); //��ȡ����ʱ��  
	struct tm local;
	localtime_s(&local, &nowtime);  //��ȡ��ǰϵͳʱ��  
	wchar_t buf[80];
	wcsftime(buf, 80, _T("%Y��%m��%d�� %H:%M:%S"), &local);

	temp.Format(_T("��¼ʱ�䣺%s ���%s\r\n----------------------------------------------------------------------\r\n"),
		buf, Comapany_Name);
	Outputf(CF_NORMAL, PF_LEFT, temp);

	temp.Format(_T("%2s %s %s %5s %10s %10s\r\n"), _T("���"), _T("����"), _T("�Ա�"), _T("ѧԺ"), _T("רҵ"), _T("�༶"));
	Outputf(CF_STUINF, PF_LEFT, temp);

	std::list<Student>::iterator StudentsListIterator;

	int i = 0;
	for (StudentsListIterator = StudentsList->begin();
	StudentsListIterator != StudentsList->end();
		++StudentsListIterator, i++)
	{
		temp.Format(_T("%02d %4s %1s %5s %5s %10s\r\n"), i + 1, StudentsListIterator->Name, StudentsListIterator->Sex,
			StudentsListIterator->Collage, StudentsListIterator->Profession, StudentsListIterator->Class);
		Outputf(CF_STUINF, PF_LEFT, temp);
	}

	temp.Format(_T("\r\n��%d�ˡ���������%d�ˣ�Ů��%d�ˡ�\r\n"), num_inf.total, num_inf.Male, num_inf.Female);
	Outputf(CF_STUINF, PF_RIGHT, temp);

	temp.Format(_T("----------------------------------------------------------------------\r\n"));
	Outputf(CF_NORMAL, PF_RIGHT, temp);
}

void CPrintPreView::Outputf(int style, int alignment, CString text) {

	CHARFORMAT2 cf;
	memset(&cf, 0, sizeof(CHARFORMAT2));
	cf.cbSize = sizeof(CHARFORMAT2);
	cf.dwMask = CFM_BOLD | CFM_COLOR | CFM_FACE | CFM_ITALIC | CFM_UNDERLINE | CFM_SIZE | CFM_STRIKEOUT;
	cf.dwEffects = CFE_BOLD;

	switch (style)
	{
	case CF_TITLE: {
		cf.yHeight = 22 * 1440 / 96;
		cf.yOffset = 0;
		cf.crTextColor = RGB(0, 0, 0);
		cf.bCharSet = 0;
		cf.bPitchAndFamily = 0;
		wcscpy_s(cf.szFaceName, _T("����"));

		break;
	}
	case CF_STUINF: {
		cf.yHeight = 13 * 1440 / 96;
		cf.yOffset = 0;
		cf.crTextColor = RGB(0, 0, 0);
		cf.bCharSet = 0;
		cf.bPitchAndFamily = 0;
		wcscpy_s(cf.szFaceName, _T("����"));

		break;
	}
	case CF_NORMAL:
	default: {
		cf.yHeight = 14 * 1440 / 96;
		cf.yOffset = 0;
		cf.crTextColor = RGB(0, 0, 0);
		cf.bCharSet = 0;
		cf.bPitchAndFamily = 0;
		wcscpy_s(cf.szFaceName, _T("����"));
	}

	}

	PARAFORMAT pf;
	pf.dwMask = PFM_ALIGNMENT;
	switch (alignment)
	{
	case PF_CENTER:
		pf.wAlignment = PFA_CENTER;
		break;
	case PF_RIGHT:
		pf.wAlignment = PFA_RIGHT;
		break;
	case PF_LEFT:
	default:
		pf.wAlignment = PFA_LEFT;

	}

	m_Text.SetWordCharFormat(cf);
	m_Text.SetParaFormat(pf);
	m_Text.ReplaceSel(text);
}

void CPrintPreView::OnBnClickedExport()
{
	EDITSTREAM es;
	CString name;
	//name.Format(_T("%s-%s"), Competition_Name, Comapany_Name);
	name.Format(_T("������-%s"), Comapany_Name);
	CFile rtf;
	CreateNewFile(rtf, name);
	es.dwCookie = (DWORD)&rtf;
	es.pfnCallback = (EDITSTREAMCALLBACK)MyStreamOutCallback;
	m_Text.StreamOut(SF_RTF, es);
	rtf.Close();

	CString temp;
	temp.Format(_T("�ѱ��浽%s.rtf��"), name);
	MessageBox(temp, NULL, MB_ICONINFORMATION);
}

static DWORD CALLBACK MyStreamOutCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb)
{
	CFile* pFile = (CFile*)dwCookie;

	pFile->Write(pbBuff, cb);
	*pcb = cb;

	return 0;
}

BOOL CPrintPreView::CreateNewFile(CFile &file, CString Name) {
	SYSTEMTIME sys;
	GetLocalTime(&sys);

	CString path;
	path.Format(_T("%s%s.rtf"), GetProgramCurrentPath(), Name);
	if (!(file.Open(path, CFile::modeCreate | CFile::modeReadWrite))) {
		MessageBox(_T("rtf�ļ�����ʧ�ܣ�"), _T("rtf�ļ���ʧ��"), MB_ICONERROR);
		return FALSE;
	}
	return TRUE;
}

CString CPrintPreView::GetProgramCurrentPath(void)
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