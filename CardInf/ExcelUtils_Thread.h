#pragma once

#include "ExcelUtils.h"
#include "CardInfDlg.h"

CExcelUtils Excel;

// CExcelUtils_Thread

class CExcelUtils_Thread : public CWinThread
{
	DECLARE_DYNCREATE(CExcelUtils_Thread)

protected:
	CExcelUtils_Thread();           // ��̬������ʹ�õ��ܱ����Ĺ��캯��
	virtual ~CExcelUtils_Thread();

public:
	virtual BOOL InitInstance();
	virtual int ExitInstance();

protected:
	DECLARE_MESSAGE_MAP()
};


