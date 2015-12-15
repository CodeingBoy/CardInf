
// CardInf.h : PROJECT_NAME 应用程序的主头文件
//

#pragma once

#ifndef __AFXWIN_H__
	#error "在包含此文件之前包含“stdafx.h”以生成 PCH 文件"
#endif

#include "resource.h"		// 主符号


// CCardInfApp: 
// 有关此类的实现，请参阅 CardInf.cpp
//

class CCardInfApp : public CWinApp
{
public:
	CCardInfApp();

// 重写
public:
	virtual BOOL InitInstance();

// 实现

	DECLARE_MESSAGE_MAP()
};

extern CCardInfApp theApp;

struct Student
{
	CString Name;
	CString	Sex; // true = male
	CString Collage;
	CString Profession;
	CString Class;
};

struct Student_Number_Inf
{
	int total;
	int Male;
	int Female;
};