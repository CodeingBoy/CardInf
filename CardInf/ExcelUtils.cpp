#include "stdafx.h"
#include "ExcelUtils.h"


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