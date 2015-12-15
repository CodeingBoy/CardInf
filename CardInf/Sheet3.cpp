// Sheet3.h : CSheet3 类的实现



// CSheet3 实现

// 代码生成在 2015年10月28日, 20:58

#include "stdafx.h"
#include "Sheet3.h"
#include "CardInf.h"

IMPLEMENT_DYNAMIC(CSheet3, CRecordset)

CSheet3::CSheet3(CDatabase* pdb)
	: CRecordset(pdb)
{
	column1 = L"";
	column2 = L"";
	column3 = L"";
	column4 = L"";
	column5 = L"";
	column6 = L"";
	column7 = 0.0;
	column8 = L"";
	column9 = L"";
	column10 = L"";
	m_nFields = 10;
	m_nDefaultType = dynaset;
}

CString CSheet3::GetDefaultConnect()
{
	CString path = GetProgramCurrentPath();
	return _T("DBQ=") + path + _T("学生数据库.mdb;DefaultDir=") + path + _T(";Driver={Driver do Microsoft Access (*.mdb)};DriverId=25;FIL=MS Access;FILEDSN=")
		+ path + _T("学生数据库.mdb.dsn;MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;");
}

CString CSheet3::GetDefaultSQL()
{
	return _T("[Sheet3]");
}

CString CSheet3::GetProgramCurrentPath(void)
{
	TCHAR path_buffer[_MAX_PATH];
	TCHAR drive[_MAX_DRIVE];
	TCHAR dir[_MAX_DIR];
	TCHAR fname[_MAX_FNAME];
	TCHAR ext[_MAX_EXT];

	GetModuleFileName(NULL, path_buffer, _MAX_PATH);
	_wsplitpath_s(path_buffer, drive, dir, fname, ext);
	CString strDir;
	strDir += drive;
	strDir += dir;
	return strDir;
}

void CSheet3::DoFieldExchange(CFieldExchange* pFX)
{
	pFX->SetFieldType(CFieldExchange::outputColumn);
// RFX_Text() 和 RFX_Int() 这类宏依赖的是
// 成员变量的类型，而不是数据库字段的类型。
// ODBC 尝试自动将列值转换为所请求的类型
	RFX_Text(pFX, _T("[学号]"), column1);
	RFX_Text(pFX, _T("[姓名]"), column2);
	RFX_Text(pFX, _T("[性别]"), column3);
	RFX_Text(pFX, _T("[学院]"), column4);
	RFX_Text(pFX, _T("[专业名称]"), column5);
	RFX_Text(pFX, _T("[行政班]"), column6);
	RFX_Double(pFX, _T("[年级]"), column7);
	RFX_Text(pFX, _T("[是否在校]"), column8);
	RFX_Text(pFX, _T("[是否注册]"), column9);
	RFX_Text(pFX, _T("[物理卡号]"), column10);

}
/////////////////////////////////////////////////////////////////////////////
// CSheet3 诊断

#ifdef _DEBUG
void CSheet3::AssertValid() const
{
	CRecordset::AssertValid();
}

void CSheet3::Dump(CDumpContext& dc) const
{
	CRecordset::Dump(dc);
}
#endif //_DEBUG


