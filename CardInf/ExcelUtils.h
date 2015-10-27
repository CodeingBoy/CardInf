#pragma once
#include "IllusionExcelFile.h"

#define STU_INF_EXCEL "学生卡信息.xls" // 物理卡和学号对应的表格名称
#define ID_INF_EXCEL "10-261.xls"	// 学生信息表格名称

struct Search_param
{
	CString CardNumber;
	long* row;
	long* progresspointer = NULL;
};

class CExcelUtils :
	public IllusionExcelFile
{
public:
	CExcelUtils();
	~CExcelUtils();
	bool Find(wchar_t* KeyWord, long * row, long * column);
	bool FindRow(wchar_t * KeyWord, long * row, long column);
	bool FindRow(wchar_t * KeyWord, long * row, long column, long * indicator);
	static UINT SearchCardNumber(LPVOID lpParam);
	static UINT SearchInfByID(LPVOID lpParam);
};

