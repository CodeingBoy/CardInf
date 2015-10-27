#pragma once

#include "ExcelUtils.h"
#include "CardInfDlg.h"

CExcelUtils Excel;

// CExcelUtils_Thread

class CExcelUtils_Thread : public CWinThread
{
	DECLARE_DYNCREATE(CExcelUtils_Thread)

protected:
	CExcelUtils_Thread();           // 动态创建所使用的受保护的构造函数
	virtual ~CExcelUtils_Thread();

public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();

protected:
	DECLARE_MESSAGE_MAP()
};


