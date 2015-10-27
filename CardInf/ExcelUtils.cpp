#include "stdafx.h"
#include "ExcelUtils.h"


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