// Sheet3.h : CSheet3 ���ʵ��



// CSheet3 ʵ��

// ���������� 2015��10��28��, 20:58

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
	return _T("DBQ=") + path + _T("ѧ�����ݿ�.mdb;DefaultDir=") + path + _T(";Driver={Driver do Microsoft Access (*.mdb)};DriverId=25;FIL=MS Access;FILEDSN=")
		+ path + _T("ѧ�����ݿ�.mdb.dsn;MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;");
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
// RFX_Text() �� RFX_Int() �������������
// ��Ա���������ͣ����������ݿ��ֶε����͡�
// ODBC �����Զ�����ֵת��Ϊ�����������
	RFX_Text(pFX, _T("[ѧ��]"), column1);
	RFX_Text(pFX, _T("[����]"), column2);
	RFX_Text(pFX, _T("[�Ա�]"), column3);
	RFX_Text(pFX, _T("[ѧԺ]"), column4);
	RFX_Text(pFX, _T("[רҵ����]"), column5);
	RFX_Text(pFX, _T("[������]"), column6);
	RFX_Double(pFX, _T("[�꼶]"), column7);
	RFX_Text(pFX, _T("[�Ƿ���У]"), column8);
	RFX_Text(pFX, _T("[�Ƿ�ע��]"), column9);
	RFX_Text(pFX, _T("[������]"), column10);

}
/////////////////////////////////////////////////////////////////////////////
// CSheet3 ���

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


