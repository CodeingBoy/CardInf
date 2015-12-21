
// CardInfDlg.h : 头文件
//

#pragma once
#include "afxcmn.h"
#include <list>
#include <algorithm>

#define ERR_REQUEST 8//寻卡错误
#define ERR_READSERIAL 9//读序列吗错误
#define ERR_SELECTCARD 10//选卡错误
#define ERR_LOADKEY 11//装载密码错误
#define ERR_AUTHKEY 12//密码认证错误
#define ERR_READ 13//读卡错误
#define ERR_WRITE 14//写卡错误
#define ERR_NONEDLL 21//没有动态库
#define ERR_DRIVERORDLL 22//动态库或驱动程序异常
#define ERR_DRIVERNULL 23//驱动程序错误或尚未安装
#define ERR_TIMEOUT 24//操作超时，一般是动态库没有反映
#define ERR_TXSIZE 25//发送字数不够
#define ERR_TXCRC 26//发送的CRC错
#define ERR_RXSIZE 27//接收的字数不够
#define ERR_RXCRC 28//接收的CRC错

typedef unsigned char(__stdcall *ppiccrequest)(unsigned char* serial);
typedef unsigned char(__stdcall *ppcdbeep)(unsigned long xms);

// CCardInfDlg 对话框
class CCardInfDlg : public CDialogEx
{
// 构造
public:
	CCardInfDlg(CWnd* pParent = NULL);	// 标准构造函数

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
