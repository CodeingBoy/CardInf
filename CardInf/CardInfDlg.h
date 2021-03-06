
// CardInfDlg.h : 头文件
//

#pragma once
#include "ExcelUtils.h"

// CCardInfDlg 对话框
class CCardInfDlg : public CDialogEx
{
// 构造
public:
	CCardInfDlg(CWnd* pParent = NULL);	// 标准构造函数
	Search_param param, param_ID;

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_CARDINF_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

	CWinThread* pThread, *pThread_ID;
public:
	afx_msg void OnBnClickedSearchId();
	afx_msg void OnBnClickedgetinf();
	afx_msg void OnBnClickedreadcard();
	static CString GetProgramCurrentPath(void);
	afx_msg void OnBnClickedreadcardcycle();
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	long r = 0, indicator = 0;
	long r_ID = 0, indicator_ID = 0;
protected:
	afx_msg LRESULT OnUpdateId(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnUpdateInf(WPARAM wParam, LPARAM lParam);
};
