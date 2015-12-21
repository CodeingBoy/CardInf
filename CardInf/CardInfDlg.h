
// CardInfDlg.h : ͷ�ļ�
//

#pragma once
#include "afxcmn.h"
#include <list>
#include <algorithm>

#define ERR_REQUEST 8//Ѱ������
#define ERR_READSERIAL 9//�����������
#define ERR_SELECTCARD 10//ѡ������
#define ERR_LOADKEY 11//װ���������
#define ERR_AUTHKEY 12//������֤����
#define ERR_READ 13//��������
#define ERR_WRITE 14//д������
#define ERR_NONEDLL 21//û�ж�̬��
#define ERR_DRIVERORDLL 22//��̬������������쳣
#define ERR_DRIVERNULL 23//��������������δ��װ
#define ERR_TIMEOUT 24//������ʱ��һ���Ƕ�̬��û�з�ӳ
#define ERR_TXSIZE 25//������������
#define ERR_TXCRC 26//���͵�CRC��
#define ERR_RXSIZE 27//���յ���������
#define ERR_RXCRC 28//���յ�CRC��

typedef unsigned char(__stdcall *ppiccrequest)(unsigned char* serial);
typedef unsigned char(__stdcall *ppcdbeep)(unsigned long xms);

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
	CListCtrl m_List;
	HINSTANCE hDLL;
	ppiccrequest piccrequest; 
	ppcdbeep pcdbeep;
	afx_msg void OnClose();
	afx_msg void OnBnClickedClearList();
	CString CreateNewFile();
	CFile txt, csv;
	afx_msg void OnBnClickedNewfile();
	void Log(CString Content);
	void WriteCSV(CString Content);
	afx_msg void OnBnClickedPrintPreview();
	//Student students[100];
	int Student_num = 0;
	Student_Number_Inf StuNum_inf = { 0,0,0 };
	std::list<Student> Student_list;
	afx_msg void OnBnClickedButton2();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
};
