
// CardInfDlg.h : ͷ�ļ�
//

#pragma once
#include "ExcelUtils.h"

// CCardInfDlg �Ի���
class CCardInfDlg : public CDialogEx
{
// ����
public:
	CCardInfDlg(CWnd* pParent = NULL);	// ��׼���캯��
	Search_param param, param_ID;

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_CARDINF_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
