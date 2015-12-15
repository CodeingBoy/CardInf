#pragma once
#include "afxcmn.h"
#include <afx.h>

// CPrintPreView �Ի���

class CPrintPreView : public CDialogEx
{
	DECLARE_DYNAMIC(CPrintPreView)

public:
	//CPrintPreView(CWnd* pParent = NULL);
	//CPrintPreView(Student * Students, int Number, wchar_t* Competition_Name, wchar_t* Comapany_Name, CWnd * pParent = NULL);
	CPrintPreView(Student * Students, int Number, wchar_t * Competition_Name, wchar_t * Comapany_Name, Student_Number_Inf num_inf, CWnd * pParent = NULL);
	// ��׼���캯��
	virtual ~CPrintPreView();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG1 };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
//	CRichEditCtrl Text;
	virtual BOOL OnInitDialog();
	void BuildText();
	CRichEditCtrl m_Text;
public:
	Student* Students;
	int Number;
	wchar_t* Competition_Name;
	wchar_t* Comapany_Name;
	Student_Number_Inf num_inf;
	afx_msg void OnBnClickedExport();
	BOOL CreateNewFile(CFile & file, CString Name);
	CString GetProgramCurrentPath(void);
};

static DWORD CALLBACK MyStreamOutCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG * pcb);
