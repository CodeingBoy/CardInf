// PrintPreView.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "CardInf.h"
#include "PrintPreView.h"
#include "afxdialogex.h"
#include "CardInfDlg.h"

// CPrintPreView �Ի���

IMPLEMENT_DYNAMIC(CPrintPreView, CDialogEx)

CPrintPreView::CPrintPreView(Student* Students, int Number, wchar_t* Competition_Name, wchar_t* Comapany_Name, Student_Number_Inf num_inf, CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_DIALOG1, pParent)
{
	AfxInitRichEdit2();
	CPrintPreView::Students = Students;
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

	
// 	edit->ReplaceSel(_T("��������ѧ����ѧԺ\
// 		----------------------------------------------------------------------\
// 		������Ա������\
// 		���  	����  �Ա� 		ѧԺ 		�༶\
// 		1   	�Ʒɺ�  ��	���������ѧԺ	�������4��\
// 		2\
// \
// \
// \
// \
// 		��XX�ˣ����У�����X�ˣ�Ů��X�ˣ�����____��Ŀ���������涨��\
// 		-------------------------------------------------------------------------- -\
// 		����ɼ���_____\
// 		"));

	BuildText();

	

	

	return TRUE;  // return TRUE unless you set the focus to a control
				  // �쳣: OCX ����ҳӦ���� FALSE
}

void CPrintPreView::BuildText() {
	CHARFORMAT2 title;
	memset(&title, 0, sizeof(CHARFORMAT2));

	title.cbSize = sizeof(CHARFORMAT2);
	title.dwMask = CFM_BOLD | CFM_COLOR | CFM_FACE | CFM_ITALIC | CFM_UNDERLINE | CFM_SIZE | CFM_STRIKEOUT;
	title.dwEffects = CFE_BOLD;
	title.yHeight = 22 * 1440 / 96;
	title.yOffset = 0;
	title.crTextColor = RGB(0, 0, 0);
	title.bCharSet = 0;
	title.bPitchAndFamily = 0;
	wcscpy_s(title.szFaceName, _T("����"));

	CHARFORMAT2 nornal;
	memset(&nornal, 0, sizeof(CHARFORMAT2));

	nornal.cbSize = sizeof(CHARFORMAT2);
	nornal.dwMask = CFM_BOLD | CFM_COLOR | CFM_FACE | CFM_ITALIC | CFM_UNDERLINE | CFM_SIZE | CFM_STRIKEOUT;
	nornal.dwEffects = CFE_BOLD;
	nornal.yHeight = 14 * 1440 / 96;
	nornal.yOffset = 0;
	nornal.crTextColor = RGB(0, 0, 0);
	nornal.bCharSet = 0;
	nornal.bPitchAndFamily = 0;
	wcscpy_s(nornal.szFaceName, _T("����"));

	CHARFORMAT2 students_list;
	memset(&students_list, 0, sizeof(CHARFORMAT2));

	students_list.cbSize = sizeof(CHARFORMAT2);
	students_list.dwMask = CFM_BOLD | CFM_COLOR | CFM_FACE | CFM_ITALIC | CFM_UNDERLINE | CFM_SIZE | CFM_STRIKEOUT;
	students_list.dwEffects = CFE_BOLD;
	students_list.yHeight = 13 * 1440 / 96;
	students_list.yOffset = 0;
	students_list.crTextColor = RGB(0, 0, 0);
	students_list.bCharSet = 0;
	students_list.bPitchAndFamily = 0;
	wcscpy_s(students_list.szFaceName, _T("����"));
	
	CString temp;

	PARAFORMAT pf;
	m_Text.GetParaFormat(pf);
	pf.dwMask = PFM_ALIGNMENT;
	pf.wAlignment = PFA_CENTER;
	m_Text.SetParaFormat(pf);
	//temp.Format(_T("��������ѧ����ѧԺ\r\n2015��У�˻� %s��Ŀ\r\n%s��������\r\n\n"), Competition_Name, Comapany_Name);
	temp.Format(_T("��������ѧ����ѧԺ\r\n2015����������������������\r\n%s�μ�����\r\n\n"), Comapany_Name);
	
	m_Text.SetWordCharFormat(title);
	m_Text.ReplaceSel(temp);

	m_Text.GetParaFormat(pf);
	pf.dwMask = PFM_ALIGNMENT;
	pf.wAlignment = PFA_LEFT;
	m_Text.SetParaFormat(pf);

	m_Text.SetWordCharFormat(nornal);
	temp.Format(_T("��¼ʱ�䣺%s ���%s\r\n��ӡʱ�䣺%s\n----------------------------------------------------------------------\r\n"),
				_T("2015-12-05 13:31:04"), Comapany_Name, _T("2015-12-05 13:31:16"));
	m_Text.ReplaceSel(temp);

	m_Text.SetWordCharFormat(students_list);
	temp.Format(_T("%2s %s %s %5s %10s %10s\r\n"), _T("���"),_T("����"), _T("�Ա�"), _T("ѧԺ"), _T("רҵ"), _T("�༶"));
	m_Text.ReplaceSel(temp);

	for (int i = 0; i < Number;i++)
	{
		temp.Format(_T("%02d %4s %1s %5s %5s %10s\r\n"), i+1,Students[i].Name,Students[i].Sex,Students[i].Collage,Students[i].Profession,Students[i].Class);
		m_Text.ReplaceSel(temp);
	}
	
	m_Text.GetParaFormat(pf);
	pf.dwMask = PFM_ALIGNMENT;
	pf.wAlignment = PFA_RIGHT;
	m_Text.SetParaFormat(pf);
	temp.Format(_T("\r\n��%d�ˡ���������%d�ˣ�Ů��%d�ˡ�\r\n"), num_inf.total, num_inf.Male, num_inf.Female);
	m_Text.ReplaceSel(temp);
	
	m_Text.SetWordCharFormat(nornal);
	temp.Format(_T("----------------------------------------------------------------------\r\n"));
	m_Text.ReplaceSel(temp);
}

void CPrintPreView::OnBnClickedExport()
{
	EDITSTREAM es;
	CString name;
	name.Format(_T("%s-%s"), Competition_Name, Comapany_Name);
	CFile rtf;
	CreateNewFile(rtf, name);
	es.dwCookie = (DWORD)&rtf;
	es.pfnCallback = (EDITSTREAMCALLBACK)MyStreamOutCallback;
	m_Text.StreamOut(SF_RTF, es);
	rtf.Close();
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