// ExcelUtils_Thread.cpp : 实现文件
//

#include "stdafx.h"
#include "CardInf.h"
#include "ExcelUtils_Thread.h"

// CExcelUtils_Thread

IMPLEMENT_DYNCREATE(CExcelUtils_Thread, CWinThread)

CExcelUtils_Thread::CExcelUtils_Thread()
{
}

CExcelUtils_Thread::~CExcelUtils_Thread()
{
}

BOOL CExcelUtils_Thread::InitInstance()
{
	CString ExcelLocation = CCardInfDlg::GetProgramCurrentPath();
	ExcelLocation += STU_INF_EXCEL;

	Excel.InitExcel();
	Excel.OpenExcelFile(ExcelLocation);
	Excel.LoadSheet(1, true);
	/*wchar_t *phy_num_wchar = new wchar_t[256];
	wcscpy_s(phy_num_wchar, 256, CardNumber);
	if (phy_inf.FindRow(phy_num_wchar, row, 1))
		rtnval = true;
	else
		rtnval = false;
	phy_inf.CloseExcelFile(false);
	phy_inf.ReleaseExcel();*/

	return TRUE;
}

int CExcelUtils_Thread::ExitInstance()
{
	// TODO:    在此执行任意逐线程清理
	return CWinThread::ExitInstance();
}

BEGIN_MESSAGE_MAP(CExcelUtils_Thread, CWinThread)
END_MESSAGE_MAP()


// CExcelUtils_Thread 消息处理程序
