
// CardInf.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CCardInfApp: 
// �йش����ʵ�֣������ CardInf.cpp
//

class CCardInfApp : public CWinApp
{
public:
	CCardInfApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CCardInfApp theApp;