/******************************************************************************************
Copyright           : 2000-2004, Appache  2.0
FileName            : illusion_excel_file.cpp
Author              : Sail
Version             :
Date Of Creation    : 2009��4��3��
Description         :

Others              :
Function List       :
1.  ......
Modification History:
1.Date  :
Author  :
Modification  :

������Ǵ��������صģ���������ɣ���лԭ�������ߣ���ֻ��������������һ��������
��������һЩ������ʹ�ò�������bool ��ΪBOOL��,���ڶ����ϵ���Ҹ���һ���֣��о�ԭ�������߶���OO��˼·���ֲ��Ǻ������
�������ණ��OLE������ȫ���˽⣬�ñ��˷�װ�Ķ����о����Ƿ����˺ܶ࣬C++��ΰ���C++
http://blog.csdn.net/gyssoft/archive/2007/04/29/1592104.aspx

OLE��дEXCEL���Ƚ���������Ӧ�þ�������OLE�Ĵ���
���ڶ�ȡ�����н��������������һ��Ԥ���صķ�ʽ���������һ�μ������еĶ�ȡ����,����ٶȾͷɿ��ˡ�
��˵д������û��ʲô�����ӿ��
http://topic.csdn.net/t/20030626/21/1962211.html

������һЩд�뷽ʽ�Ĵ��룬��֤����д��EXCEL�����������Ƕ��ڱ��棬�ҷ����������CLOSE���ұ���ķ�ʽ��
�ٶȷǳ������Ҳ����Ϊʲô��
�����Ұ�EXCEL���ˣ�������к�������,


******************************************************************************************/




//-----------------------excelfile.cpp----------------

#include "StdAfx.h"
#include "illusionexcelfile.h"



COleVariant
covTrue((short)TRUE),
covFalse((short)FALSE),
covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

//
CApplication IllusionExcelFile::excel_application_;


IllusionExcelFile::IllusionExcelFile() :
	already_preload_(FALSE)
{
}

IllusionExcelFile::~IllusionExcelFile()
{
	//
	CloseExcelFile();
}


//��ʼ��EXCEL�ļ���
BOOL IllusionExcelFile::InitExcel()
{

	//����Excel 2000������(����Excel) 
	if (!excel_application_.CreateDispatch(_T("Excel.Application"), NULL))
	{
		AfxMessageBox(_T("����Excel����ʧ��,�����û�а�װEXCEL������!"));
		return FALSE;
	}

	excel_application_.put_DisplayAlerts(FALSE);
	return TRUE;
}

//
void IllusionExcelFile::ReleaseExcel()
{
	excel_application_.Quit();
	excel_application_.ReleaseDispatch();
	excel_application_ = NULL;
}

//��excel�ļ�
BOOL IllusionExcelFile::OpenExcelFile(const wchar_t *file_name)
{
	//�ȹر�
	CloseExcelFile();

	//����ģ���ļ��������ĵ� 
	excel_books_.AttachDispatch(excel_application_.get_Workbooks(), true);

	LPDISPATCH lpDis = NULL;
	lpDis = excel_books_.Add(COleVariant(file_name));
	if (lpDis)
	{
		excel_work_book_.AttachDispatch(lpDis);
		//�õ�Worksheets 
		excel_sheets_.AttachDispatch(excel_work_book_.get_Worksheets(), true);

		//��¼�򿪵��ļ�����
		open_excel_file_ = file_name;

		return TRUE;
	}

	return FALSE;
}

//�رմ򿪵�Excel �ļ�,Ĭ������������ļ�
void IllusionExcelFile::CloseExcelFile(BOOL if_save)
{
	//����Ѿ��򿪣��ر��ļ�
	if (open_excel_file_.IsEmpty() == FALSE)
	{
		//�������,�����û�����,���û��Լ��棬����Լ�SAVE�������Ī���ĵȴ�
		if (if_save)
		{
			ShowInExcel(TRUE);
		}
		else
		{
			//
			excel_work_book_.Close(COleVariant(short(FALSE)), COleVariant(open_excel_file_), covOptional);
			excel_books_.Close();
		}

		//���ļ����������
		open_excel_file_.Empty();
	}



	excel_sheets_.ReleaseDispatch();
	excel_work_sheet_.ReleaseDispatch();
	excel_current_range_.ReleaseDispatch();
	excel_work_book_.ReleaseDispatch();
	excel_books_.ReleaseDispatch();
}

void IllusionExcelFile::SaveasXSLFile(const CString &xls_file)
{
	excel_work_book_.SaveAs(COleVariant(xls_file),
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		0,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional);
	return;
}


int IllusionExcelFile::GetSheetCount()
{
	return excel_sheets_.get_Count();
}


CString IllusionExcelFile::GetSheetName(long table_index)
{
	CWorksheet sheet;
	sheet.AttachDispatch(excel_sheets_.get_Item(COleVariant((long)table_index)), true);
	CString name = sheet.get_Name();
	sheet.ReleaseDispatch();
	return name;
}

//������ż���Sheet���,������ǰ�������еı���ڲ�����
BOOL IllusionExcelFile::LoadSheet(long table_index, BOOL pre_load)
{
	LPDISPATCH lpDis = NULL;
	excel_current_range_.ReleaseDispatch();
	excel_work_sheet_.ReleaseDispatch();
	lpDis = excel_sheets_.get_Item(COleVariant((long)table_index));
	if (lpDis)
	{
		excel_work_sheet_.AttachDispatch(lpDis, true);
		excel_current_range_.AttachDispatch(excel_work_sheet_.get_Cells(), true);
	}
	else
	{
		return FALSE;
	}

	already_preload_ = FALSE;
	//�������Ԥ�ȼ���
	if (pre_load)
	{
		PreLoadSheet();
		already_preload_ = TRUE;
	}

	return TRUE;
}

//�������Ƽ���Sheet���,������ǰ�������еı���ڲ�����
BOOL IllusionExcelFile::LoadSheet(const wchar_t* sheet, BOOL pre_load)
{
	LPDISPATCH lpDis = NULL;
	excel_current_range_.ReleaseDispatch();
	excel_work_sheet_.ReleaseDispatch();
	lpDis = excel_sheets_.get_Item(COleVariant(sheet));
	if (lpDis)
	{
		excel_work_sheet_.AttachDispatch(lpDis, true);
		excel_current_range_.AttachDispatch(excel_work_sheet_.get_Cells(), true);

	}
	else
	{
		return FALSE;
	}
	//
	already_preload_ = FALSE;
	//�������Ԥ�ȼ���
	if (pre_load)
	{
		already_preload_ = TRUE;
		PreLoadSheet();
	}

	return TRUE;
}

//�õ��е�����
int IllusionExcelFile::GetColumnCount()
{
	CRange range;
	CRange usedRange;
	usedRange.AttachDispatch(excel_work_sheet_.get_UsedRange(), true);
	range.AttachDispatch(usedRange.get_Columns(), true);
	int count = range.get_Count();
	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();
	return count;
}

//�õ��е�����
int IllusionExcelFile::GetRowCount()
{
	CRange range;
	CRange usedRange;
	usedRange.AttachDispatch(excel_work_sheet_.get_UsedRange(), true);
	range.AttachDispatch(usedRange.get_Rows(), true);
	int count = range.get_Count();
	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();
	return count;
}

//���һ��CELL�Ƿ����ַ���
BOOL IllusionExcelFile::IsCellString(long irow, long icolumn)
{
	CRange range;
	range.AttachDispatch(excel_current_range_.get_Item(COleVariant((long)irow), COleVariant((long)icolumn)).pdispVal, true);
	COleVariant vResult = range.get_Value2();
	//VT_BSTR��ʾ�ַ���
	if (vResult.vt == VT_BSTR)
	{
		return TRUE;
	}
	return FALSE;
}

//���һ��CELL�Ƿ�����ֵ
BOOL IllusionExcelFile::IsCellInt(long irow, long icolumn)
{
	CRange range;
	range.AttachDispatch(excel_current_range_.get_Item(COleVariant((long)irow), COleVariant((long)icolumn)).pdispVal, true);
	COleVariant vResult = range.get_Value2();
	//����һ�㶼��VT_R8
	if (vResult.vt == VT_INT || vResult.vt == VT_R8)
	{
		return TRUE;
	}
	return FALSE;
}

//
CString IllusionExcelFile::GetCellString(long irow, long icolumn)
{

	COleVariant vResult;
	CString str;
	//�ַ���
	if (already_preload_ == FALSE)
	{
		CRange range;
		range.AttachDispatch(excel_current_range_.get_Item(COleVariant((long)irow), COleVariant((long)icolumn)).pdispVal, true);
		vResult = range.get_Value2();
		range.ReleaseDispatch();
	}
	//�����������Ԥ�ȼ�����
	else
	{
		long read_address[2];
		VARIANT val;
		read_address[0] = irow;
		read_address[1] = icolumn;
		ole_safe_array_.GetElement(read_address, &val);
		vResult = val;
	}

	if (vResult.vt == VT_BSTR)
	{
		str = vResult.bstrVal;
	}
	//����
	else if (vResult.vt == VT_INT)
	{
		str.Format(_T("%d"), vResult.pintVal);
	}
	//8�ֽڵ����� 
	else if (vResult.vt == VT_R8)
	{
		str.Format(_T("%0.0f"), vResult.dblVal);
	}
	//ʱ���ʽ
	else if (vResult.vt == VT_DATE)
	{
		SYSTEMTIME st;
		VariantTimeToSystemTime(vResult.date, &st);
		CTime tm(st);
		str = tm.Format("%Y-%m-%d");

	}
	//��Ԫ��յ�
	else if (vResult.vt == VT_EMPTY)
	{
		str = "";
	}

	return str;
}

double IllusionExcelFile::GetCellDouble(long irow, long icolumn)
{
	double rtn_value = 0;
	COleVariant vresult;
	//�ַ���
	if (already_preload_ == FALSE)
	{
		CRange range;
		range.AttachDispatch(excel_current_range_.get_Item(COleVariant((long)irow), COleVariant((long)icolumn)).pdispVal, true);
		vresult = range.get_Value2();
		range.ReleaseDispatch();
	}
	//�����������Ԥ�ȼ�����
	else
	{
		long read_address[2];
		VARIANT val;
		read_address[0] = irow;
		read_address[1] = icolumn;
		ole_safe_array_.GetElement(read_address, &val);
		vresult = val;
	}

	if (vresult.vt == VT_R8)
	{
		rtn_value = vresult.dblVal;
	}

	return rtn_value;
}

//VT_R8
int IllusionExcelFile::GetCellInt(long irow, long icolumn)
{
	int num;
	COleVariant vresult;

	if (already_preload_ == FALSE)
	{
		CRange range;
		range.AttachDispatch(excel_current_range_.get_Item(COleVariant((long)irow), COleVariant((long)icolumn)).pdispVal, true);
		vresult = range.get_Value2();
		range.ReleaseDispatch();
	}
	else
	{
		long read_address[2];
		VARIANT val;
		read_address[0] = irow;
		read_address[1] = icolumn;
		ole_safe_array_.GetElement(read_address, &val);
		vresult = val;
	}
	//ԭ���룺
	//num = static_cast<int>(vresult.dblVal);
	CString temp(vresult);
	num = _ttoi(temp);

	return num;
}

void IllusionExcelFile::SetCellString(long irow, long icolumn, CString new_string)
{
	COleVariant new_value(new_string);
	CRange start_range = excel_work_sheet_.get_Range(COleVariant(_T("A1")), covOptional);
	CRange write_range = start_range.get_Offset(COleVariant((long)irow - 1), COleVariant((long)icolumn - 1));
	write_range.put_Value2(new_value);
	start_range.ReleaseDispatch();
	write_range.ReleaseDispatch();

}

void IllusionExcelFile::SetCellInt(long irow, long icolumn, int new_int)
{
	COleVariant new_value((long)new_int);

	CRange start_range = excel_work_sheet_.get_Range(COleVariant(_T("A1")), covOptional);
	CRange write_range = start_range.get_Offset(COleVariant((long)irow - 1), COleVariant((long)icolumn - 1));
	write_range.put_Value2(new_value);
	start_range.ReleaseDispatch();
	write_range.ReleaseDispatch();
}


//
void IllusionExcelFile::ShowInExcel(BOOL bShow)
{
	excel_application_.put_Visible(bShow);
	excel_application_.put_UserControl(bShow);
}

//���ش򿪵�EXCEL�ļ�����
CString IllusionExcelFile::GetOpenFileName()
{
	return open_excel_file_;
}

//ȡ�ô�sheet������
CString IllusionExcelFile::GetLoadSheetName()
{
	return excel_work_sheet_.get_Name();
}

//ȡ���е����ƣ�����27->AA
char *IllusionExcelFile::GetColumnName(long icolumn)
{
	static char column_name[64];
	size_t str_len = 0;

	while (icolumn > 0)
	{
		int num_data = icolumn % 26;
		icolumn /= 26;
		if (num_data == 0)
		{
			num_data = 26;
			icolumn--;
		}
		column_name[str_len] = (char)((num_data - 1) + 'A');
		str_len++;
	}
	column_name[str_len] = '\0';
	//��ת
	_strrev(column_name);

	return column_name;
}

//Ԥ�ȼ���
void IllusionExcelFile::PreLoadSheet()
{

	CRange used_range;

	used_range = excel_work_sheet_.get_UsedRange();


	VARIANT ret_ary = used_range.get_Value2();
	if (!(ret_ary.vt & VT_ARRAY))
	{
		return;
	}
	//
	ole_safe_array_.Clear();
	ole_safe_array_.Attach(ret_ary);
}