#include "stdafx.h"
#include "ExcelUtils.h"
#include "CardInfDlg.h"

CExcelUtils::CExcelUtils()
{
}


CExcelUtils::~CExcelUtils()
{
}

//************************************
// 函数名：		Find
// 函数说明：		遍历各个单元格查找对应内容
// 参数：			wchar_t * KeyWord	关键字
// 参数：			long * row			填入用于接收的 long 类型行变量
// 参数：			long * column		填入用于接收的 long 类型列变量
// 返回值：		bool，true 表示找到内容，false 表示找不到
//************************************
bool CExcelUtils::Find(wchar_t* KeyWord, long* row, long* column) {
	for (int i = 1; i <= GetRowCount(); i++) {
		for (int j = 1; j <= GetColumnCount(); j++) {
			if (wcscmp(KeyWord, GetCellString(i, j)) == 0)
			{
				*row = i;
				*column = j;
				return true;
			}
		}
	}

	return false;
}

//************************************
// 函数名：		FindRow
// 函数说明：		逐行查找对应内容
// 参数：			wchar_t * KeyWord	关键字
// 参数：			long * row			填入用于接收的 long 类型行变量
// 参数：			long column			固定查找的列
// 返回值：		bool，true 表示找到内容，false 表示找不到
//************************************
bool CExcelUtils::FindRow(wchar_t* KeyWord, long* row, long column) {
	for (int i = 1; i <= GetRowCount(); i++) {
		if (wcscmp(KeyWord, GetCellString(i, column)) == 0)
		{
			*row = i;
			return true;
		}
	}

	return false;
}

bool CExcelUtils::FindRow(wchar_t* KeyWord, long* row, long column, long* indicator) {
	for (int i = 1; i <= GetRowCount(); i++) {
		*indicator = i;
		if (wcscmp(KeyWord, GetCellString(i, column)) == 0)
		{
			*row = i;
			return true;
		}
	}

	return false;
}

UINT CExcelUtils::SearchCardNumber(LPVOID lpParam) {
	UINT rtnval;

	CoInitialize(NULL);
	//CoInitializeEx(NULL, COINIT_MULTITHREADED);

	CString ExcelLocation = CCardInfDlg::GetProgramCurrentPath();
	ExcelLocation += STU_INF_EXCEL;

	CExcelUtils phy_inf;
	phy_inf.InitExcel();
	phy_inf.OpenExcelFile(ExcelLocation);
	phy_inf.LoadSheet(1, true);
	wchar_t *phy_num_wchar = new wchar_t[256];
	wcscpy_s(phy_num_wchar, 256, (*((Search_param*)lpParam)).CardNumber);
	if (phy_inf.FindRow(phy_num_wchar, (*((Search_param*)lpParam)).row, 1, (*((Search_param*)lpParam)).progresspointer))
		rtnval = 0;
	else
		rtnval = 1;
	phy_inf.CloseExcelFile(false);
	phy_inf.ReleaseExcel();

	CoUninitialize();

	return 0;
}

UINT CExcelUtils::SearchInfByID(LPVOID lpParam) {
	UINT rtnval;

	CoInitialize(NULL);
	//CoInitializeEx(NULL, COINIT_MULTITHREADED);

	CString ExcelLocation = CCardInfDlg::GetProgramCurrentPath();
	ExcelLocation += ID_INF_EXCEL;

	CExcelUtils phy_inf;
	phy_inf.InitExcel();
	phy_inf.OpenExcelFile(ExcelLocation);
	phy_inf.LoadSheet(1, true);
	wchar_t *phy_num_wchar = new wchar_t[256];
	wcscpy_s(phy_num_wchar, 256, (*((Search_param*)lpParam)).CardNumber);
	if (phy_inf.FindRow(phy_num_wchar, (*((Search_param*)lpParam)).row, 1, (*((Search_param*)lpParam)).progresspointer))
		rtnval = 0;
	else
		rtnval = 1;
	phy_inf.CloseExcelFile(false);
	phy_inf.ReleaseExcel();

	CoUninitialize();

	return 0;
}