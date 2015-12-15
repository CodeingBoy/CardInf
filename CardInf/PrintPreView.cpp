// PrintPreView.cpp : 实现文件
//

#include "stdafx.h"
#include "CardInf.h"
#include "PrintPreView.h"
#include "afxdialogex.h"
#include "CardInfDlg.h"

// CPrintPreView 对话框

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


// CPrintPreView 消息处理程序


BOOL CPrintPreView::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	CRichEditCtrl *edit = (CRichEditCtrl*)GetDlgItem(IDC_RichEdit);

	
// 	edit->ReplaceSel(_T("华南理工大学广州学院\
// 		----------------------------------------------------------------------\
// 		参赛人员名单：\
// 		序号  	姓名  性别 		学院 		班级\
// 		1   	黄飞豪  男	计算机工程学院	软件工程4班\
// 		2\
// \
// \
// \
// \
// 		共XX人，其中，男性X人，女性X人，符合____项目比赛人数规定。\
// 		-------------------------------------------------------------------------- -\
// 		该组成绩：_____\
// 		"));

	BuildText();

	

	

	return TRUE;  // return TRUE unless you set the focus to a control
				  // 异常: OCX 属性页应返回 FALSE
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
	wcscpy_s(title.szFaceName, _T("黑体"));

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
	wcscpy_s(nornal.szFaceName, _T("宋体"));

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
	wcscpy_s(students_list.szFaceName, _T("宋体"));
	
	CString temp;

	PARAFORMAT pf;
	m_Text.GetParaFormat(pf);
	pf.dwMask = PFM_ALIGNMENT;
	pf.wAlignment = PFA_CENTER;
	m_Text.SetParaFormat(pf);
	//temp.Format(_T("华南理工大学广州学院\r\n2015届校运会 %s项目\r\n%s参赛名单\r\n\n"), Competition_Name, Comapany_Name);
	temp.Format(_T("华南理工大学广州学院\r\n2015年阳光体育冬季环湖长跑\r\n%s参加名单\r\n\n"), Comapany_Name);
	
	m_Text.SetWordCharFormat(title);
	m_Text.ReplaceSel(temp);

	m_Text.GetParaFormat(pf);
	pf.dwMask = PFM_ALIGNMENT;
	pf.wAlignment = PFA_LEFT;
	m_Text.SetParaFormat(pf);

	m_Text.SetWordCharFormat(nornal);
	temp.Format(_T("检录时间：%s 组别：%s\r\n打印时间：%s\n----------------------------------------------------------------------\r\n"),
				_T("2015-12-05 13:31:04"), Comapany_Name, _T("2015-12-05 13:31:16"));
	m_Text.ReplaceSel(temp);

	m_Text.SetWordCharFormat(students_list);
	temp.Format(_T("%2s %s %s %5s %10s %10s\r\n"), _T("序号"),_T("姓名"), _T("性别"), _T("学院"), _T("专业"), _T("班级"));
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
	temp.Format(_T("\r\n共%d人。其中男性%d人，女性%d人。\r\n"), num_inf.total, num_inf.Male, num_inf.Female);
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
		MessageBox(_T("rtf文件创建失败！"), _T("rtf文件打开失败"), MB_ICONERROR);
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