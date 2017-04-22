#include "stdafx.h"
#include "ExcelProcess.h"
#include "..\Header\CommonHeader.h"

//为避免在代码中多处对ExcelProcess进行new，将ExcelProcess设计为单例模式
//通过getInstance和destroyInstance获取和销毁对象
ExcelProcess* ExcelProcess::excel = nullptr;

ExcelProcess::ExcelProcess()
{
	lpDisp = NULL;
	excelVer = 0;
	//init();
}

//获取对象
ExcelProcess* ExcelProcess::getInstance()
{
	if(excel == NULL)
	{
		excel = new ExcelProcess();
	}
	return excel;
}
//销毁对象
void ExcelProcess::destroyInstance()
{
	if(excel != NULL)
	{
		delete excel;
		excel = NULL;
	}
}

ExcelProcess::~ExcelProcess()
{
	try
	{
		sheet.ReleaseDispatch();
		sheets.ReleaseDispatch();
		book.ReleaseDispatch();
		books.ReleaseDispatch();
		ExcelApp.ReleaseDispatch();
		ExcelApp.Quit();
		//退出伪装的app
		if(!ExcelApp_fake.get_ActiveSheet())
		{
			books_fake.ReleaseDispatch();
			ExcelApp_fake.ReleaseDispatch();
			ExcelApp_fake.Quit();
		}
	}
	catch (COleDispatchException*)
	{
		AfxMessageBox(Notice_get_by_id(IDS_POW_OFF_EXCEL_FAIL));
		//AfxMessageBox(_T("关闭Excel服务出错。"));
	}
}


/************************************************************************/
/* 初始化                                                               */
/************************************************************************/
BOOL ExcelProcess::init()
{
	CString strOfficeVer[5] = {_T("office 2003"),_T("office 2007"),_T("office 2010"),_T("office 2013"), _T("office 2016")};
	BOOL result = FALSE;
	for(int i = 4; i >= 0; i--)
	{
		if(!createServer(strOfficeVer[i]))
			continue;
		else
		{
			result = TRUE;
		}
	}
	if(excelVer == 0)
	{
		result = FALSE;
	}
	return result;
}

//获取office版本
int ExcelProcess::getOfficeVer()
{
	return excelVer;
}

/************************************************************************/
/* 启动伪Excel服务
/* 因为有些office版本(如2007)的机制是一个EXE进程对应多个Excel文件
/* 如果不启动伪Excel服务，当用户打开其他excel文件时，会对原来的excel文件产生干扰
/* 该伪Excel服务就是提供给用户操作其他excel文件(可能发生)
/************************************************************************/
BOOL ExcelProcess::createFakeServer()
{
	CString strOfficeVer[5] = {_T("office 2003"),_T("office 2007"),_T("office 2010"),_T("office 2013"), _T("office 2016")};
	for(int i = 4; i >= 0; i--)
	{
		CString officeVer = strOfficeVer[i];
		//去除前后空格
		officeVer.Trim();
		//获取版本号字符
		CString verNum = officeVer.Right(4);
		int ver = _ttoi(verNum);
		switch(ver)
		{
		case 2003:
			if(judgeExcelVer(11))
			{
				if(ExcelApp_fake.CreateDispatch(_T("Excel.Application.11"), NULL))
				{
					ExcelApp_fake.put_DisplayAlerts(FALSE);
					books_fake.AttachDispatch(ExcelApp_fake.get_Workbooks());
					return TRUE;
				}
				else
				{
					return FALSE;
				}
			}
			break;
		case 2007:
			if(judgeExcelVer(12))
			{
				if(ExcelApp_fake.CreateDispatch(_T("Excel.Application.12"), NULL))
				{
					ExcelApp_fake.put_DisplayAlerts(FALSE);
					books_fake.AttachDispatch(ExcelApp_fake.get_Workbooks());
					return TRUE;
				}
				else
				{
					return FALSE;
				}
			}
			break;
		case 2010:
			if(judgeExcelVer(14))
			{
				if(ExcelApp_fake.CreateDispatch(_T("Excel.Application.14"), NULL))
				{
					ExcelApp_fake.put_DisplayAlerts(FALSE);
					books_fake.AttachDispatch(ExcelApp_fake.get_Workbooks());
					return TRUE;
				}
				else
				{
					return FALSE;
				}
			}
			break;
		case 2013:
			if( judgeExcelVer(15))
			{
				if(ExcelApp_fake.CreateDispatch(_T("Excel.Application.15"), NULL))
				{
					ExcelApp_fake.put_DisplayAlerts(FALSE);
					books_fake.AttachDispatch(ExcelApp_fake.get_Workbooks());
					return TRUE;
				}
				else
				{
					return FALSE;
				}
			}
			break;
		case 2016:
			if( judgeExcelVer(16))
			{
				if(ExcelApp_fake.CreateDispatch(_T("Excel.Application.16"), NULL))
				{
					ExcelApp_fake.put_DisplayAlerts(FALSE);
					books_fake.AttachDispatch(ExcelApp_fake.get_Workbooks());
					return TRUE;
				}
				else
				{
					return FALSE;
				}
			}
			break;
		}
	}
	return FALSE;
}

/************************************************************************/
/* 判断Excel版本号                                                      */
/************************************************************************/
BOOL ExcelProcess::judgeExcelVer(int Ver)
{
	HKEY hkey;
	int ret;
	CString str;	
	LONG len;
	str.Format(_T("Excel.Application.%d"),Ver);
	str += _T("\\CLSID");
	ret = RegCreateKey(HKEY_CLASSES_ROOT, str, &hkey);
	if(ret == ERROR_SUCCESS)
	{
		RegQueryValue(HKEY_CLASSES_ROOT, str, NULL, &len);
		//如果注册表中 HKEY_CLASSES_ROOT\Excel.Application.x\CPLSID中的值为空，则读取到'\0'，长度为2
		return len == 2 ? FALSE : TRUE;
	}
	else
	{
		return FALSE;
	}
}

/************************************************************************/
/* 启动Excel服务，传入的字符串参数格式为 office ****                    */
/************************************************************************/
BOOL ExcelProcess::createServer( CString officeVer )
{
	//去除前后空格
	officeVer.Trim();
	//获取版本号字符
	CString verNum = officeVer.Right(4);
	int ver = _ttoi(verNum);
	switch(ver)
	{
	case 2003:
		if(judgeExcelVer(11))
		{
			if(ExcelApp.CreateDispatch(_T("Excel.Application.11"), NULL))
			{
				excelVer = 2003;
			}
			else
			{
				return FALSE;
			}
		}
		else
		{
			return FALSE;
		}
		break;
	case 2007:
		if(judgeExcelVer(12))
		{
			if(ExcelApp.CreateDispatch(_T("Excel.Application.12"), NULL))
			{
				excelVer = 2007;
			}
			else
			{
				return FALSE;
			}
		}
		else
		{
			return FALSE;
		}
		break;
	case 2010:
		if(judgeExcelVer(14))
		{
			if(ExcelApp.CreateDispatch(_T("Excel.Application.14"), NULL))
			{
				excelVer = 2010;
			}
			else
			{
				return FALSE;
			}
		}
		else
		{
			return FALSE;
		}
		break;
	case 2013:
		if( judgeExcelVer(15))
		{
			if(ExcelApp.CreateDispatch(_T("Excel.Application.15"), NULL))
			{
				excelVer = 2013;
			}
			else
			{
				return FALSE;
			}
		}
		else
		{
			return FALSE;
		}
		break;
	case 2016:
		if( judgeExcelVer(16))
		{
			if(ExcelApp.CreateDispatch(_T("Excel.Application.16"), NULL))
			{
				excelVer = 2016;
			}
			else
			{
				return FALSE;
			}
		}
		else
		{
			return FALSE;
		}
		break;
	}
	return TRUE;
}

BOOL ExcelProcess::setExcelExportSheets( CString filePath )
{
// 	ExcelApp.put_Visible(TRUE);
// 	ExcelApp.put_UserControl(FALSE);
	//获取工作薄集合 
	books.AttachDispatch(ExcelApp.get_Workbooks());
	//打开一个工作簿
	try
	{
		lpDisp = books.Open(filePath, 
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
			vtMissing, vtMissing, vtMissing, vtMissing);
		book.AttachDispatch(lpDisp);
	}
	catch(...)
	{
		//新建
		lpDisp = books.Add(vtMissing);
		book.AttachDispatch(lpDisp);
		saveExcelAs(filePath);
	}

	sheets.AttachDispatch(book.get_Sheets());
	return TRUE;
}

//显示excel文件，在配置批处理导出结束后使用
void ExcelProcess::setView()
{
	ExcelApp.put_Visible(TRUE);
	ExcelApp.put_UserControl(FALSE);
}

BOOL ExcelProcess::openExcelFile( CString filePath )
{
	return setExcelImportSheets(filePath);
}

BOOL ExcelProcess::setExcelImportSheets( CString filePath )
{
	if (filePath.GetLength() == 0)
		return FALSE;
	// 	ExcelApp.put_Visible(TRUE);
	// 	ExcelApp.put_UserControl(FALSE);
	//获取工作薄集合 
	books.AttachDispatch(ExcelApp.get_Workbooks());

	//打开一个工作簿
	try
	{
		lpDisp = books.Open(filePath, 
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
			vtMissing, vtMissing, vtMissing, vtMissing);
		book.AttachDispatch(lpDisp);
	}
	catch(...)
	{
		return FALSE;
	}
	//获取工作表集合 
	sheets.AttachDispatch(book.get_Worksheets());

	// 	//得到当前活跃sheet
	// 	//如果有单元格正处于编辑状态中，此操作不能返回，会一直等待
	// 	lpDisp=book.get_ActiveSheet();
	// 	sheet.AttachDispatch(lpDisp); 
	return TRUE;
}

//获取一个指定名称的sheet，如果不存则返回false，存在则选中该sheet
BOOL ExcelProcess::getSheet( CString sheetName )
{
	sheets.AttachDispatch(book.get_Sheets());
	try
	{
		lpDisp = sheets.get_Item(_variant_t(sheetName));
		sheet.AttachDispatch(lpDisp);
	}
	catch(...)
	{
// 		lpDisp = sheets.Add(vtMissing, vtMissing, _variant_t(1), vtMissing);
// 		sheet.AttachDispatch(lpDisp);
// 		sheet.put_Name(sheetName);
		return FALSE;
	}
	return TRUE;
}

//创建sheet
void ExcelProcess::createSheet(CString sheetName)
{
	lpDisp = sheets.Add(vtMissing, vtMissing, _variant_t(1), vtMissing);
	sheet.AttachDispatch(lpDisp);
	sheet.put_Name(sheetName);
}

//获取存在的sheet(不指定名称)
BOOL ExcelProcess::getActiveSheet()
{
	//sheets.AttachDispatch(book.get_Sheets());
	try
	{
		lpDisp = book.get_ActiveSheet();
		sheet.AttachDispatch(lpDisp);
	}
	catch(...)
	{
		return FALSE;
	}
	return TRUE;
}

//删除sheet
BOOL ExcelProcess::deleteSheet( CString sheetName )
{
	try
	{
		//删除系统默认产生的三个sheet
		sheet.AttachDispatch(sheets.get_Item(_variant_t(sheetName)));
		sheet.Delete();
	}
	catch(...)
	{
		//TODO
		return FALSE;
	}
	return TRUE;
}

/************************************************************************/
/* 设置单元格数据有效性为数字                                           */
/************************************************************************/
BOOL ExcelProcess::setCellsTypeToNum( CString cellBegin, CString cellEnd, int min, int max, int defValue )
{

	range = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	LPDISPATCH result = range.get_Validation();
	validation.AttachDispatch(result);
	validation.Delete();
	validation.Add(XlDVType::xlValidateWholeNumber, _variant_t(XlDVAlertStyle::xlValidAlertStop),
		_variant_t(XlFormatConditionOperator::xlBetween), _variant_t(min), _variant_t(max));
	range.put_Value2(_variant_t(defValue));
	setCellsAlignLeft(range);
	return TRUE;
}

//重载
BOOL ExcelProcess::setCellsTypeToNum( CString cellBegin, CString cellEnd, int min, int max )
{
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	LPDISPATCH result = range.get_Validation();
	validation.AttachDispatch(result);
	validation.Delete();
	validation.Add(XlDVType::xlValidateWholeNumber, _variant_t(XlDVAlertStyle::xlValidAlertStop),
		_variant_t(XlFormatConditionOperator::xlBetween), _variant_t(min), _variant_t(max));
	setCellsAlignLeft(range);
	return TRUE;
}

//重载
BOOL ExcelProcess::setCellsTypeToNum( CString cellBegin, CString cellEnd, double min, double max )
{
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	LPDISPATCH result = range.get_Validation();
	validation.AttachDispatch(result);
	validation.Delete();
	validation.Add(XlDVType::xlValidateDecimal, _variant_t(XlDVAlertStyle::xlValidAlertStop),
		_variant_t(XlFormatConditionOperator::xlBetween), _variant_t(min), _variant_t(max));
	setCellsAlignLeft(range);
	return TRUE;
}

/************************************************************************/
/* 设置单元格数据有效性为字符串数组                                     */
/************************************************************************/
BOOL ExcelProcess::setCellsToStringList( CString cellBegin, CString cellEnd, const std::vector<CString>& strList, unsigned int defValueIndex /*= 0*/ )
{
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	CString strDef = strList.at(defValueIndex);
	CString str;
	for(unsigned int i = 0; i < strList.size(); i++)
	{
		str += strList.at(i);
		str += ',';
	}
	range.put_Value2(_variant_t(strDef));
	LPDISPATCH result = range.get_Validation();
	validation.AttachDispatch(result);
	validation.Delete();
	validation.Add(XlDVType::xlValidateList, _variant_t(XlDVAlertStyle::xlValidAlertStop), _variant_t(XlFormatConditionOperator::xlBetween), _variant_t(str), _variant_t(NULL));
	setCellsAlignLeft(range);
	return TRUE;
}

//重载
BOOL ExcelProcess::setCellsToStringList( CString cellBegin, CString cellEnd,const std::vector<CString>& strList )
{
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	CString str;
	for(unsigned int i = 0; i < strList.size(); i++)
	{
		str += strList.at(i);
		str += ',';
	}
	LPDISPATCH result = range.get_Validation();
	validation.AttachDispatch(result);
	validation.Delete();
	validation.Add(XlDVType::xlValidateList, _variant_t(XlDVAlertStyle::xlValidAlertStop), _variant_t(XlFormatConditionOperator::xlBetween), _variant_t(str), _variant_t(NULL));
	setCellsAlignLeft(range);
	return TRUE;
}

//重载
BOOL ExcelProcess::setCellsToStringList(CString cellBegin, CString cellEnd, CString sheetName, CString valueCellBegin, CString valueCellEnd)
{
	valueCellBegin.Insert(0, _T("$"));
	valueCellEnd.Insert(0, _T("$"));
	int index = 0;
	for(index = 0; index < valueCellBegin.GetLength(); index++)
	{
		if(valueCellBegin.GetAt(index) >= _T('0') && valueCellBegin.GetAt(index) <= _T('9'))
			break;
	}
	valueCellBegin.Insert(index, _T("$"));
	index = 0;
	for(index = 0; index < valueCellEnd.GetLength(); index++)
	{
		if(valueCellEnd.GetAt(index) >= _T('0') && valueCellEnd.GetAt(index) <= _T('9'))
			break;
	}
	valueCellEnd.Insert(index, _T("$"));
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	CString str;
	str.Format(_T("=%s!%s:%s"), sheetName, valueCellBegin, valueCellEnd);

	LPDISPATCH result = range.get_Validation();
	validation.AttachDispatch(result);
	validation.Delete();
	validation.Add(XlDVType::xlValidateList, _variant_t(XlDVAlertStyle::xlValidAlertStop), _variant_t(XlFormatConditionOperator::xlBetween), _variant_t(str), _variant_t(NULL));
	setCellsAlignLeft(range);
	return TRUE;
}

//重载
BOOL ExcelProcess::setCellsToStringList(CString cellBegin, CString cellEnd, const CString* strList, unsigned int len, unsigned int defValueIndex)
{
	if(strList == nullptr || defValueIndex >= len)
		return FALSE;
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	CString str;
	CString strDef = strList[defValueIndex];
	for(unsigned int i = 0; i < len; i++)
	{
		str += strList[i];
		str += ',';
	}
	range.put_Value2(_variant_t(strDef));
	LPDISPATCH result = range.get_Validation();
	validation.AttachDispatch(result);
	validation.Delete();
	validation.Add(XlDVType::xlValidateList, _variant_t(XlDVAlertStyle::xlValidAlertStop), _variant_t(XlFormatConditionOperator::xlBetween), _variant_t(str), _variant_t(NULL));
	setCellsAlignLeft(range);
	return TRUE;
}

//重载
BOOL ExcelProcess::setCellsToStringList(CString cellBegin, CString cellEnd, const CString* strList, unsigned int len)
{
	if(strList == nullptr)
		return FALSE;
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	CString str;
	for(unsigned int i = 0; i < len; i++)
	{
		str += strList[i];
		str += ',';
	}
	LPDISPATCH result = range.get_Validation();
	validation.AttachDispatch(result);
	validation.Delete();
	validation.Add(XlDVType::xlValidateList, _variant_t(XlDVAlertStyle::xlValidAlertStop), _variant_t(XlFormatConditionOperator::xlBetween), _variant_t(str), _variant_t(NULL));
	setCellsAlignLeft(range);
	return TRUE;
}

//根据Excel公式设置单元格的数据有效性
void ExcelProcess::setCellsToFormula(CString cellBegin, CString cellEnd, CString formulaName)
{
	range = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	LPDISPATCH result = range.get_Validation();
	validation.AttachDispatch(result);
	validation.Delete();
	validation.Add(XlDVType::xlValidateList, _variant_t(XlDVAlertStyle::xlValidAlertStop), _variant_t(XlFormatConditionOperator::xlBetween),
		_variant_t(formulaName), vtMissing);
	setCellsAlignLeft(range);
}


/************************************************************************/
/* 设置单元格颜色，颜色索引见 Excel颜色对照表.doc                       */
/************************************************************************/
BOOL ExcelProcess::setCellsColor(CString cellBegin, CString cellEnd, int colorIndex)
{
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	LPDISPATCH result = range.get_Interior();
	interior.AttachDispatch(result);
	interior.put_ColorIndex(_variant_t(colorIndex));
	return TRUE;
}

/************************************************************************/
/* 设置一片单元格的值
/* cellBegin 起点坐标
/* 范围的长宽由nums决定
/************************************************************************/
BOOL ExcelProcess::setCellsValue( CString cellBegin, vector<vector<CString>>& nums )
{
	COleSafeArray safeArr;
	//第一、二行为标题，从第三行开始去获取列数
	//DWORD numElements[] = {nums.size(), nums.size() > 1 ? nums.at(nums.size() - 1).size() : nums.at(0).size()};
	DWORD numElements[] = {nums.size(), nums.at(0).size()};
	//创建二维数组
	safeArr.Create(VT_BSTR , 2, numElements);

	CString cellEnd = getEndCell(cellBegin, nums.size(), nums.at(0).size());
	range = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));

	long index[2] = {0, 0};
	long lFirstLBound = 0;
	long lFirstUBound = 0;
	long lSecondLBound = 0;
	long lSecondUBound = 0;
	//获取行的下限
	safeArr.GetLBound(1, &lFirstLBound);
	//获取行的上限
	safeArr.GetUBound(1, &lFirstUBound);
	//获取列的下限
	safeArr.GetLBound(2, &lSecondLBound);
	//获取列的上限
	safeArr.GetUBound(2, &lSecondUBound);
	//遍历
	for(ULONG i = lFirstLBound; i < nums.size(); i++)
	{
		index[0] = i;
		for(ULONG j = lSecondLBound; j < nums.at(i).size(); j++)
		{
			index[1] = j;
			//CString lElement = *((CString*)nums + cols*i + j);
			//CString lElement = nums[i][j];
			CString lElement = nums.at(i).at(j);
			BSTR strTmp = lElement.AllocSysString();
			safeArr.PutElement(index, strTmp);
			SysFreeString(strTmp);
		}
	}
	range.put_Value2((VARIANT) safeArr);
	setCellsAlignLeft(range);
	return TRUE;
}

//重载，主要用于设置序号
BOOL ExcelProcess::setCellsValue( CString cellBegin, vector<int>& nums )
{
	COleSafeArray safeArr;
	DWORD numElements[] = {nums.size(), 1};
	//创建二维数组
	safeArr.Create(VT_I4 , 2, numElements);
	//只有一列数据
	CString cellEnd = getEndCell(cellBegin, nums.size(), 1);
	range = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));

	long index[2] = {0, 0};
	long lFirstLBound = 0;
	long lFirstUBound = 0;
	long lSecondLBound = 0;
	long lSecondUBound = 0;
	//获取行的下限
	safeArr.GetLBound(1, &lFirstLBound);
	//获取行的上限
	safeArr.GetUBound(1, &lFirstUBound);
	//获取列的下限
	safeArr.GetLBound(2, &lSecondLBound);
	//获取列的上限
	safeArr.GetUBound(2, &lSecondUBound);
	//遍历
	for(ULONG i = lFirstLBound; i < nums.size(); i++)
	{
		index[0] = i;
		index[1] = lSecondLBound;
		int val = nums.at(i);
		safeArr.PutElement(index, &val);
	}
	range.put_Value2((VARIANT) safeArr);
	setCellsAlignLeft(range);
	setCellsBold(cellBegin, cellEnd, TRUE);
	return TRUE;
}

//重载 暂时不用
BOOL ExcelProcess::setCellsValue( CString cellBegin, long** nums, int rows, int cols )
{
	COleSafeArray safeArr;
	DWORD numElements[] = {rows, cols};
	//创建二维数组
	safeArr.Create(VT_I4 , 2, numElements);

	CString cellEnd = getEndCell(cellBegin, rows, cols);
	range = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));

	long index[2] = {0, 0};
	long lFirstLBound = 0;
	long lFirstUBound = 0;
	long lSecondLBound = 0;
	long lSecondUBound = 0;
	//获取行的下限
	safeArr.GetLBound(1, &lFirstLBound);
	//获取行的上限
	safeArr.GetUBound(1, &lFirstUBound);
	//获取列的下限
	safeArr.GetLBound(2, &lSecondLBound);
	//获取列的上限
	safeArr.GetUBound(2, &lSecondUBound);
	//遍历
	for(long i = lFirstLBound; i <= lFirstUBound; i++)
	{
		index[0] = i;
		for(long j = lSecondLBound; j <= lSecondUBound; j++)
		{
			index[1] = j;
			long lElement = *((long*)nums + cols*i + j);
			safeArr.PutElement(index, &lElement);
		}
	}
	/*VARIANT varWrite = (VARIANT) safeArr;*/
	range.put_Value2((VARIANT) safeArr);
	setCellsAlignLeft(range);
	return TRUE;
}

//重载 暂时不用
BOOL ExcelProcess::setCellsValue( CString cellBegin, CString** nums, int rows, int cols )
{
	COleSafeArray safeArr;
	DWORD numElements[] = {rows, cols};
	//创建二维数组
	safeArr.Create(VT_BSTR , 2, numElements);

	CString cellEnd = getEndCell(cellBegin, rows, cols);
	range = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));

	long index[2] = {0, 0};
	long lFirstLBound = 0;
	long lFirstUBound = 0;
	long lSecondLBound = 0;
	long lSecondUBound = 0;
	//获取行的下限
	safeArr.GetLBound(1, &lFirstLBound);
	//获取行的上限
	safeArr.GetUBound(1, &lFirstUBound);
	//获取列的下限
	safeArr.GetLBound(2, &lSecondLBound);
	//获取列的上限
	safeArr.GetUBound(2, &lSecondUBound);
	//遍历
	for(long i = lFirstLBound; i <= lFirstUBound; i++)
	{
		index[0] = i;
		for(long j = lSecondLBound; j <= lSecondUBound; j++)
		{
			index[1] = j;
			//CString lElement = *((CString*)nums + cols*i + j);
			CString lElement = nums[i][j];
			BSTR strTmp = lElement.AllocSysString();
			safeArr.PutElement(index, strTmp);
			SysFreeString(strTmp);
		}
	}
	range.put_Value2((VARIANT) safeArr);
	setCellsAlignLeft(range);
	return TRUE;
}

//读取一列数据 
//cellBegin 起点坐标 
//count     数量
void ExcelProcess::getColValue(vector<CString>& outData, CString cellBegin, int count)
{
	CString cellEnd = getEndCell(cellBegin, count, 1);
	range = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	VARIANT data = range.get_Value2();
	COleSafeArray safeArr(data);
	long index[2] = {0, 0};
	long lFirstLBound = 0;
	long lFirstUBound = 0;
	long lSecondLBound = 0;
	long lSecondUBound = 0;
	//获取行的下限
	safeArr.GetLBound(1, &lFirstLBound);
	//获取行的上限
	safeArr.GetUBound(1, &lFirstUBound);
	//获取列的下限
	safeArr.GetLBound(2, &lSecondLBound);
	//获取列的上限
	safeArr.GetUBound(2, &lSecondUBound);
	ASSERT(lSecondLBound == 1);
	ASSERT(lSecondUBound == 1);
	VARIANT value;
	for(long i = lFirstLBound; i <= lFirstUBound; i++)
	{
		index[0] = i;
		index[1] = 1;
		safeArr.GetElement(index, &value);
		CString valueStr;
		switch (value.vt)
		{
		case VT_BOOL:
			{
				BOOL t = value.boolVal;
				valueStr = t ? _T("TRUE") : _T("FALSE");
				break;
			}
		case VT_BSTR:
			{
				valueStr = value.bstrVal;
				break;
			}
		case VT_I4:
			{
				int t = value.intVal;
				valueStr.Format(_T("%d"), t);
				break;
			}
		case VT_R8:
			{
				double t = value.dblVal;
				valueStr.Format(_T("%lf"), t);
				break;
			}
		default:
			break;
		}
		outData.push_back(valueStr);
	}
}

//设置一列数据
BOOL ExcelProcess::setColValue( CString cellBegin, const vector<CString>& nums )
{
	COleSafeArray safeArr;
	DWORD numElements[] = {nums.size(), 1};
	//创建二维数组
	safeArr.Create(VT_BSTR , 2, numElements);
	CString cellEnd = getEndCell(cellBegin, nums.size(), 1);
	range = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	long index[2] = {0, 0};
	long lFirstLBound = 0;
	long lFirstUBound = 0;
	long lSecondLBound = 0;
	long lSecondUBound = 0;

	//获取行的下限
	safeArr.GetLBound(1, &lFirstLBound);
	//获取行的上限
	safeArr.GetUBound(1, &lFirstUBound);
	//获取列的下限
	safeArr.GetLBound(2, &lSecondLBound);
	//获取列的上限
	safeArr.GetUBound(2, &lSecondUBound);
	//遍历
	for(ULONG i = lFirstLBound; i < nums.size(); i++)
	{
		index[0] = i;
		index[1] = 0;
		CString lElement = nums.at(i);
		BSTR strTmp = lElement.AllocSysString();
		safeArr.PutElement(index, strTmp);
		SysFreeString(strTmp);
	}
	range.put_Value2((VARIANT) safeArr);
	//setCellsAlignLeft(range);
	return TRUE;
}


//最多支持开头是两个字母
//传入起点坐标，行数和列数，返回结束点坐标
//如传入 A1,2,1 返回 A2  
//如传入 A1,1,2 返回 B2 
CString ExcelProcess::getEndCell( CString cellBegin, int rows, int cols )
{
	BOOL hasTwoChar = FALSE;
	TCHAR beginChar = cellBegin[0];
	if(!((beginChar >= 'A' && beginChar <= 'Z')||(beginChar >= 'a' && beginChar <= 'z')))
	{
		//AfxMessageBox(_T("输入起始点数据错误"));
		return CString("");
	}
	beginChar = cellBegin[1];
	{
		if(beginChar >= '0' && beginChar <= '9')
			beginChar = cellBegin[0];
		else
		{
			beginChar = cellBegin[1];
			hasTwoChar = TRUE;
			if(!((beginChar >= 'A' && beginChar <= 'Z')||(beginChar >= 'a' && beginChar <= 'z')))
			{
				//AfxMessageBox(_T("输入起始点数据错误"));
				return CString("");
			}

		}
	}
	if(cols <= 0)
	{
		//AfxMessageBox(_T("输入的列数需要大于0"));
		return CString("");
	}
	CString numStr;
	if(!hasTwoChar)
		numStr = cellBegin.Mid(1, cellBegin.GetLength() - 1);
	else
		numStr = cellBegin.Mid(2, cellBegin.GetLength() - 1);
	int beginRow = _ttol(numStr);
	int endRow = beginRow + rows - 1;
	TCHAR endChar = beginChar + cols - 1;
	if(beginChar + cols - 1 > _T('Z'))
	{
		if(!hasTwoChar)
		{
			int index = (beginChar + cols - 1)% _T('Z');
			CString a;
			a.Format(_T("%c%c"),_T('A'), (_T('A') + index - 1));
			CString tmp1;
			tmp1.Format(_T("%s%d"), a, endRow);
			return tmp1;
		}
		else
		{
			TCHAR tmpchar = cellBegin[0] + 1;
			int index = (beginChar + cols - 1)% _T('Z');
			CString a;
			a.Format(_T("%c%c"), tmpchar, (_T('A') + index - 1));
			CString tmp1;
			tmp1.Format(_T("%s%d"), a, endRow);
			return tmp1;
		}
	}
	CString tmp;
	if(!hasTwoChar)
		tmp.Format(_T("%c%d"), endChar, endRow);
	else
		tmp.Format(_T("%c%c%d"), cellBegin[0], endChar, endRow);
	return tmp;
}

//获取一片范围的值
//cellBegin 起点坐标
//rows 行数
//cols 列数
void ExcelProcess::getValue( vector<vector<CString>>& outData, CString cellBegin, int rows, int cols)
{
	CString cellEnd = getEndCell(cellBegin, rows, cols);
	range = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	VARIANT data = range.get_Value2();  
	if(rows == 1 && cols == 1)
	{
		vector<CString> tmp;
		tmp.push_back(data.bstrVal);
		outData.push_back(tmp);
		return;
	}
	COleSafeArray safeArr(data);
	
// 	COleSafeArray safeArr;
// 	DWORD numElements[] = {rows, cols};
// 	safeArr.Create(VT_BSTR , 2, numElements);
// 	safeArr = (COleSafeArray)data;

	long index[2] = {0, 0};
	long lFirstLBound = 0;
	long lFirstUBound = 0;
	long lSecondLBound = 0;
	long lSecondUBound = 0;
	//获取行的下限
	safeArr.GetLBound(1, &lFirstLBound);
	//获取行的上限
	safeArr.GetUBound(1, &lFirstUBound);
	//获取列的下限
	safeArr.GetLBound(2, &lSecondLBound);
	//获取列的上限
	safeArr.GetUBound(2, &lSecondUBound);
	VARIANT value;
	//遍历
	for(long i = lFirstLBound; i <= lFirstUBound; i++)
	{
		index[0] = i;
		vector<CString> tmp;
		for(long j = lSecondLBound; j <= lSecondUBound; j++)
		{
			index[1] = j;
			safeArr.GetElement(index, &value);
			CString valueStr;
			switch (value.vt)
			{
			case VT_BOOL:
				{
					BOOL t = value.boolVal;
					valueStr = t ? _T("TRUE") : _T("FALSE");
					break;
				}
			case VT_BSTR:
				{
					valueStr = value.bstrVal;
					break;
				}
			case VT_I4:
				{
					int t = value.intVal;
					valueStr.Format(_T("%d"), t);
					break;
				}
			case VT_R8:
				{
					double t = value.dblVal;
					valueStr.Format(_T("%lf"), t);
					break;
				}
			default:
				break;
			}

			tmp.push_back(valueStr);
		}
		outData.push_back(tmp);
		tmp.clear();
	}
}

/************************************************************************/
/* 保存                                                                 */
/************************************************************************/
void ExcelProcess::saveExcel()
{
	ExcelApp.put_DisplayAlerts(FALSE);
	//book.Close(vtMissing, vtMissing, vtMissing);
	book.Save();
}

//关闭
void ExcelProcess::closeExcel()
{
	//ExcelApp.get_ThisWorkbook();
	COleVariant SaveChanges((short)FALSE), RouteWorkbook((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	book.Close(SaveChanges, vtMissing, RouteWorkbook);
}

/************************************************************************/
/* 另存为                                                               */
/************************************************************************/
void ExcelProcess::saveExcelAs( CString savePath )
{
	savePath.Trim();
	book.SaveAs(_variant_t(savePath), 
		vtMissing, vtMissing, vtMissing ,vtMissing,vtMissing, 0,
		vtMissing,vtMissing,vtMissing,vtMissing, vtMissing);
}

/************************************************************************/
/* 另存为xml                                                            */
/************************************************************************/
void ExcelProcess::savaExcelToXml( CString savePath )
{
	savePath.Trim();
	CString str = savePath.Right(3);
	str.MakeLower();
	if(str != _T("xml"))
	{
		//AfxMessageBox(_T("输入文件名出错"));
		return;
	}
	book.SaveAs(_variant_t(savePath), _variant_t(XlFileFormat::xlXMLSpreadsheet), 
		vtMissing, vtMissing, _variant_t(FALSE),_variant_t(FALSE), 0,
		vtMissing, vtMissing, vtMissing, vtMissing, vtMissing);
}

//分解comBox
//comBox为struct.xml中的opt
BOOL ExcelProcess::splitComBox( CString comBox, std::vector<CString>& strList )
{
	int indexFirst;
	int indexSecond;
	while(1)
	{
		indexFirst = comBox.Find(_T(':'));
		indexSecond = comBox.Find(_T(";"));
		CString tmp = comBox.Mid(0, indexSecond);
		CString key = tmp.Left(indexFirst);
		CString value = tmp.Right(tmp.GetLength() - 1 - indexFirst);
		strList.push_back(value);
		comBox = comBox.Right(comBox.GetLength() - 1 - indexSecond);
		if(comBox.GetLength() == 0)
			break;
	}
	return TRUE;
}
//在这里将combox左边看做key,右边看成value，key和value是一对一的，根据传入的key获取value
BOOL ExcelProcess::getComBoxValue( CString comBox, const CString& key, CString& value )
{
	int indexFirst;
	int indexSecond;
	while(1)
	{
		indexFirst = comBox.Find(_T(':'));
		indexSecond = comBox.Find(_T(";"));
		CString tmp = comBox.Mid(0, indexSecond);
		CString keyStr = tmp.Left(indexFirst);
		CString valueStr = tmp.Right(tmp.GetLength() - 1 - indexFirst);
		if(keyStr == key)
		{
			value = valueStr;
			return TRUE;
		}
		comBox = comBox.Right(comBox.GetLength() - 1 - indexSecond);
		if(comBox.GetLength() == 0)
		{
			return FALSE;
		}
	}
	return TRUE;
}

//在这里将combox左边看做key,右边看成value，key和value是一对一的,通过value查找key，使用在导入功能
BOOL ExcelProcess::getComBoxKey( CString comBox, const CString& value, CString& key  )
{
	int indexFirst;
	int indexSecond;
	while(1)
	{
		indexFirst = comBox.Find(_T(':'));
		indexSecond = comBox.Find(_T(';'));
		CString tmp = comBox.Mid(0, indexSecond);
		CString keyStr = tmp.Left(indexFirst);         
		CString valueStr = tmp.Right(tmp.GetLength() - 1 - indexFirst);
		if(valueStr == value)
		{
			key = keyStr;
			return TRUE;
		}
		int index = valueStr.Find(value);
		if(index != -1 && index >= 2)
		{
			if(isNumber(valueStr.Left(index - 1)))
			{
				key = keyStr;
				return TRUE;
			}
		}
		comBox = comBox.Right(comBox.GetLength() - 1 - indexSecond);
		if(comBox.GetLength() == 0)
		{
			return FALSE;
		}
	}
	return FALSE;
}

//加粗指定范围内单元格的字体
BOOL ExcelProcess::setCellsBold( CString cellBegin, CString cellEnd, BOOL bold )
{
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	LPDISPATCH result = range.get_Font();
	font.AttachDispatch(result);
	font.put_Bold(_variant_t(bold));
	return TRUE;
}

//设置指定范围内单元格的字体
BOOL ExcelProcess::setCellsFont( CString cellBegin, CString cellEnd, CString fontName, int fontSize, BOOL bold /*= FALSE*/ )
{
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	LPDISPATCH result = range.get_Font();
	font.AttachDispatch(result);
	font.put_Bold(_variant_t(bold));
	font.put_Name(_variant_t(fontName));
	font.put_Size(_variant_t(fontSize));

	//office 2003不支持，所以注释掉 不影响结果
	/*font.put_Strikethrough(_variant_t(FALSE));
	font.put_Superscript(_variant_t(FALSE));
	font.put_Subscript(_variant_t(FALSE));
	font.put_OutlineFont(_variant_t(FALSE));
	font.put_Shadow(_variant_t(FALSE));
	font.put_Underline(_variant_t(Excel::XlUnderlineStyle::xlUnderlineStyleNone));
	font.put_ThemeColor(_variant_t(XlThemeColor::xlThemeColorLight1));
	font.put_TintAndShade(_variant_t(0));
	font.put_ThemeFont(_variant_t(XlThemeFont::xlThemeFontNone));*/
	return TRUE;
}


//根据行号设置整行颜色
BOOL ExcelProcess::setRowColor( UINT rowIndex, int colorIndex )
{
	CString strRowIndex;
	strRowIndex.Format(_T("%d:%d"), rowIndex, rowIndex);
	lpDisp = sheet.get_Range(_variant_t(strRowIndex), vtMissing);
	range.AttachDispatch(lpDisp);
	LPDISPATCH tmp = range.get_Interior();
	interior.AttachDispatch(tmp);
	interior.put_ColorIndex(_variant_t(colorIndex));
	return TRUE;
}

//根据单元格获取到行号
long ExcelProcess::getCellRowIndex( CString cellIndex )
{
	int i = 0;
	for(i = 0; i < cellIndex.GetLength(); i++)
	{
		if(cellIndex.GetAt(i) >= '0' && cellIndex.GetAt(i) <= '9')
			break;
	}
	return _ttoi(cellIndex.Mid(i, cellIndex.GetLength() - i));
}

//设置单个单元格的值
void ExcelProcess::setCellValue( CString cellIndex, CString value )
{
	range = sheet.get_Range(_variant_t(cellIndex), _variant_t(cellIndex));
	range.put_Value2(_variant_t(value));
	setCellsAlignLeft(range);
}

//获取单个单元格的值
CString ExcelProcess::getCellValue( CString cellIndex )
{
	range = sheet.get_Range(_variant_t(cellIndex), _variant_t(cellIndex));
	VARIANT value = range.get_Value2();
	if(value.vt == VT_R8)
	{
		int num = (int)value.dblVal;
		CString str;
		str.Format(_T("%d"), num);
		return str;
	}
	return value.bstrVal;
}

//设置单元格格式
void ExcelProcess::setCellsFormat( CString cellBegin, CString cellEnd, CString format )
{
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	range.put_NumberFormatLocal(_variant_t(format));
}

//设置单元格左对齐
void ExcelProcess::setCellsAlignLeft(CRange range )
{
	range.put_HorizontalAlignment(_variant_t(Constants::xlLeft));
	range.put_VerticalAlignment(_variant_t(Constants::xlBottom));
	range.put_WrapText(_variant_t(FALSE));
	range.put_Orientation(_variant_t(0));
	range.put_AddIndent(_variant_t(FALSE));
	range.put_IndentLevel(_variant_t(0));
	range.put_ShrinkToFit(_variant_t(FALSE));
	range.put_ReadingOrder(_variant_t(Constants::xlContext));
	range.put_MergeCells(_variant_t(FALSE));
}

//设置单元格右对齐
void ExcelProcess::setCellsAlignLeft( CString cellBegin, CString cellEnd )
{
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	range.put_HorizontalAlignment(_variant_t(Constants::xlLeft));
	range.put_VerticalAlignment(_variant_t(Constants::xlBottom));
	range.put_WrapText(_variant_t(FALSE));
	range.put_Orientation(_variant_t(0));
	range.put_AddIndent(_variant_t(FALSE));
	range.put_IndentLevel(_variant_t(0));
	range.put_ShrinkToFit(_variant_t(FALSE));
	range.put_ReadingOrder(_variant_t(Constants::xlContext));
	range.put_MergeCells(_variant_t(FALSE));
}

//设置单元格字符长度
void ExcelProcess::setCellsLength( CString cellBegin, CString cellEnd, UINT length )
{
	CString strLength;
	strLength.Format(_T("%d"), length);
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);
	LPDISPATCH result = range.get_Validation();
	validation.AttachDispatch(result);
	validation.Delete();
	validation.Add(XlDVType::xlValidateTextLength, _variant_t(XlDVAlertStyle::xlValidAlertStop), 
		_variant_t(XlFormatConditionOperator::xlEqual),_variant_t(strLength), _variant_t(NULL));

	validation.put_IgnoreBlank(TRUE);
	validation.put_InCellDropdown(TRUE);
	validation.put_IMEMode(XlIMEMode::xlIMEModeNoControl);
	validation.put_ShowInput(TRUE);
	validation.put_ShowError(TRUE);
}

//设置单元格输入的文本长度
void ExcelProcess::setCellsLength( CString cellBegin, CString cellEnd, UINT lengthMin, UINT lengthMax )
{
	CString strLengthMin;
	CString strLengthMax;
	strLengthMin.Format(_T("%d"), lengthMin);
	strLengthMax.Format(_T("%d"), lengthMax);
	lpDisp = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.AttachDispatch(lpDisp);

	LPDISPATCH result = range.get_Validation();
	validation.AttachDispatch(result);
	validation.Delete();
	validation.Add(XlDVType::xlValidateTextLength, _variant_t(XlDVAlertStyle::xlValidAlertStop), 
		_variant_t(XlFormatConditionOperator::xlBetween),_variant_t(strLengthMin), _variant_t(strLengthMax));

	validation.put_IgnoreBlank(TRUE);
	validation.put_InCellDropdown(TRUE);
	validation.put_IMEMode(XlIMEMode::xlIMEModeNoControl);
	validation.put_ShowInput(TRUE);
	validation.put_ShowError(TRUE);
	setCellsAlignLeft(range);
}

//设置单元格格式为文本
void ExcelProcess::setCellsToText( CString cellBegin, CString cellEnd )
{

	COleSafeArray saRet;  
	DWORD numElements = {2};//数组中有2个元素  
	saRet.Create(VT_I4, 1, &numElements);//第一个参数表示存入int，第二个参数表示是一维数组，第三个参数表示数组中有2个元素  
	long index = 0;//数组下标  
	int val = 1;//值  
	saRet.PutElement(&index, &val);//将0下标的值设置为1  
	index++;  
	val = 2;  
	saRet.PutElement(&index, &val);
	CRange tmp = sheet.get_Range(_variant_t(cellBegin), vtMissing);
	range = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.TextToColumns(_variant_t(tmp), XlTextParsingType::xlDelimited, Constants::xlDoubleQuote, 
		_variant_t(FALSE), _variant_t(TRUE), _variant_t(FALSE), _variant_t(FALSE), _variant_t(FALSE), vtMissing,
		vtMissing, _variant_t(saRet), vtMissing, vtMissing, _variant_t(TRUE));
}

//设置列宽
void ExcelProcess::setColumnWidth( CString cellIndex, int width )
{
	range = sheet.get_Range(_variant_t(cellIndex), _variant_t(cellIndex));
	range.put_ColumnWidth(_variant_t(width));
}

//解锁单元格
void ExcelProcess::unlockALL()
{
	range = sheet.get_Cells();
	//range.Activate();
	range.put_Locked(_variant_t(FALSE));
	range.put_FormulaHidden(_variant_t(FALSE));
}

//锁定单元格
void ExcelProcess::lockCells( CString cellBegin, CString cellEnd )
{
	range = sheet.get_Range(_variant_t(cellBegin), _variant_t(cellEnd));
	range.put_Locked(_variant_t(TRUE));
	range.put_FormulaHidden(_variant_t(FALSE));
}

//锁定cellIndex所在行
void ExcelProcess::lockRow(CString cellIndex)
{
	int row = getCellRowIndex(cellIndex);
	CString format;
	format.Format(_T("%d:%d"), row, row);
	range = sheet.get_Range(_variant_t(format), vtMissing);
	range.put_Locked(_variant_t(TRUE));
	range.put_FormulaHidden(_variant_t(FALSE));
}
//锁定cellBegin至cellEnd所在行
void ExcelProcess::lockRows(CString cellBegin, CString cellEnd)
{
	int row1 = getCellRowIndex(cellBegin);
	int row2 = getCellRowIndex(cellEnd);
	CString format;
	format.Format(_T("%d:%d"), row1, row2);
	range = sheet.get_Range(_variant_t(format), vtMissing);
	range.put_Locked(_variant_t(TRUE));
	range.put_FormulaHidden(_variant_t(FALSE));
}

//保护sheet，sheet中被锁定的单元格用户将无法选中
void ExcelProcess::setSheetProtect(CString sheetName)
{
	if(getSheet(sheetName))
	{
		sheet.Protect(vtMissing, _variant_t(TRUE), _variant_t(TRUE), _variant_t(TRUE), vtMissing, vtMissing, _variant_t(TRUE), vtMissing, vtMissing, 
		vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, vtMissing);
		sheet.put_EnableSelection(XlEnableSelection::xlUnlockedCells);
	}
}

//解除保护sheet
void ExcelProcess::setSheetUnprotect()
{
	sheet.Unprotect(vtMissing);
}

//添加公式
void ExcelProcess::addFormula( CString name, CString formula )
{
	names.AttachDispatch(book.get_Names());
	names.Add(_variant_t(name),_variant_t(formula), vtMissing, vtMissing, vtMissing, vtMissing, vtMissing, 
		vtMissing, vtMissing, vtMissing, vtMissing);
}

//获取公式
CString ExcelProcess::getFormula(CString sheetName, CString cellBegin, CString cellEnd, FormulaType type)
{
	cellBegin.Insert(0, _T("$"));
	cellEnd.Insert(0, _T("$"));
	int index = 0;
	for(index = 0; index < cellBegin.GetLength(); index++)
	{
		if(cellBegin.GetAt(index) >= _T('0') && cellBegin.GetAt(index) <= _T('9'))
			break;
	}
	cellBegin.Insert(index, _T("$"));
	index = 0;
	for(index = 0; index < cellEnd.GetLength(); index++)
	{
		if(cellEnd.GetAt(index) >= _T('0') && cellEnd.GetAt(index) <= _T('9'))
			break;
	}
	cellEnd.Insert(index, _T("$"));

	CString formula;
	if(type == FormulaType::cellFromRange)
	{
		CString count;
		count.Format(_T("SUMPRODUCT(N(LEN(%s!%s:%s)>0))"), sheetName, cellBegin, cellEnd);
		//公式：设置单元格的取值是一个范围的连续的值，除去空值
		formula.Format(_T("=OFFSET(%s!%s,,,SUMPRODUCT(N(LEN(%s!%s:%s)>0)),)"), sheetName, cellBegin, sheetName, cellBegin, cellEnd);
		formula.Format(_T("=OFFSET(%s!%s,,,IF(%s > 0, %s, 1),)"), 
			sheetName, cellBegin, count, count);
	}
	return formula;
}

//从A1单元格开始检测表格的范围
//row和column传入一个预期值
void ExcelProcess::getMaxRange(UINT& row, UINT& column)
{
	CString beginCellIndex = _T("A1");
// 	if(row >= 2 && column >= 1)
// 		beginCellIndex = getEndCell(beginCellIndex, row + 1, column);
// 	CString beginCellIndexTmp = beginCellIndex;
	//检测行
	//TODO 由于是逐行检测 在行数较多的时候，效率较低，后续考虑优化
	while(1)
	{
		if(!getCellValue(beginCellIndex).IsEmpty())
		{
			row++;
			beginCellIndex = getEndCell(beginCellIndex, 2, 1);
			
		}
		else
		{
			break;
		}
	}
	//检测列
	//TODO 由于是逐列检测 在列数较多的时候，效率较低，后续考虑优化
	//beginCellIndex = beginCellIndexTmp;
	beginCellIndex = _T("A1");
	while(1)
	{
		if(!getCellValue(beginCellIndex).IsEmpty())
		{
			column++;
			beginCellIndex = getEndCell(beginCellIndex, 1, 2);

		}
		else
		{
			break;
		}
	}
}