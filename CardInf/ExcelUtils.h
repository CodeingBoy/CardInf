#pragma once
#include "IllusionExcelFile.h"

#define STU_INF_EXCEL "ѧ������Ϣ.xls" // ������ѧ�Ŷ�Ӧ�ı������
#define ID_INF_EXCEL "10-261.xls"	// ѧ����Ϣ�������

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

