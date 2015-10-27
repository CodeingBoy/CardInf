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
// ��������		Find
// ����˵����		����������Ԫ����Ҷ�Ӧ����
// ������			wchar_t * KeyWord	�ؼ���
// ������			long * row			�������ڽ��յ� long �����б���
// ������			long * column		�������ڽ��յ� long �����б���
// ����ֵ��		bool��true ��ʾ�ҵ����ݣ�false ��ʾ�Ҳ���
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
// ��������		FindRow
// ����˵����		���в��Ҷ�Ӧ����
// ������			wchar_t * KeyWord	�ؼ���
// ������			long * row			�������ڽ��յ� long �����б���
// ������			long column			�̶����ҵ���
// ����ֵ��		bool��true ��ʾ�ҵ����ݣ�false ��ʾ�Ҳ���
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