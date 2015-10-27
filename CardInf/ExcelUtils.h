#pragma once
#include "IllusionExcelFile.h"
class CExcelUtils :
	public IllusionExcelFile
{
public:
	CExcelUtils();
	~CExcelUtils();
	bool Find(wchar_t* KeyWord, long * row, long * column);
	bool FindRow(wchar_t * KeyWord, long * row, long column);
};

