
// CardInfDlg.h : ͷ�ļ�
//

#pragma once


// CCardInfDlg �Ի���
class CCardInfDlg : public CDialogEx
{
// ����
public:
	CCardInfDlg(CWnd* pParent = NULL);	// ��׼���캯��

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
public:
	afx_msg void OnBnClickedSearchId();
	afx_msg void OnBnClickedgetinf();
	afx_msg void OnBnClickedreadcard();
	CString GetProgramCurrentPath(void);
	afx_msg void OnBnClickedreadcardcycle();
	afx_msg void OnTimer(UINT_PTR nIDEvent);
};