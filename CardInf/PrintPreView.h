#pragma once
#include "afxcmn.h"
#include <afx.h>
#include <list>
#include <algorithm>

#define CF_NORMAL	0
#define CF_TITLE	1
#define CF_STUINF	2

#define PF_LEFT		0
#define PF_CENTER	1
#define PF_RIGHT	2

// CPrintPreView 对话框

class CPrintPreView : public CDialogEx
{
	DECLARE_DYNAMIC(CPrintPreView)

public:
	//CPrintPreView(CWnd* pParent = NULL);
	//CPrintPreView(Student * Students, int Number, wchar_t* Competition_Name, wchar_t* Comapany_Name, CWnd * pParent = NULL);
	//CPrintPreView(Student * Students, int Number, wchar_t * Competition_Name, wchar_t * Comapany_Name, Student_Number_Inf num_inf, CWnd * pParent = NULL);
	CPrintPreView(std::list<Student> *StudentsList, int Number, wchar_t* Competition_Name, wchar_t* Comapany_Name, Student_Number_Inf num_inf, CWnd* pParent = NULL);
	// 标准构造函数
	virtual ~CPrintPreView();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG1 };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
	virtual BOOL OnInitDialog();
	void BuildText();
	void OutputItem(Student & student);
	void Outputf(int style, int alignment, CString text);
	CRichEditCtrl m_Text;
	std::list<Student> *StudentsList;
	int Number;
	wchar_t* Competition_Name;
	wchar_t* Comapany_Name;
	Student_Number_Inf num_inf;
	afx_msg void OnBnClickedExport();
	BOOL CreateNewFile(CFile & file, CString Name);
	CString GetProgramCurrentPath(void);
};

static DWORD CALLBACK MyStreamOutCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG * pcb);
