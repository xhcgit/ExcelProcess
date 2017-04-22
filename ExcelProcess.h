#pragma once

#include "TLH/mso.h"
using namespace Office;
#include "TLH/vbe6ext.h"
using namespace VBIDE;
#include "TLH/excel.h"
using namespace Excel;

#include "CRange.h"
#include "CApplication.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "CValidation.h"
#include "CInterior.h"
#include "CFont0.h"
#include "CNames.h"

#include <vector>
enum  FormulaType
{
	cellFromRange = 1
};
class ExcelProcess
{
private:
	ExcelProcess();
	// 	ExcelProcess(CString filePath);
	~ExcelProcess();
	static ExcelProcess* excel;
public:
	static ExcelProcess* getInstance();
	static void destroyInstance();
	BOOL init();
	BOOL setExcelExportSheets(CString filePath);
	BOOL setExcelImportSheets(CString filePath); 

	BOOL getSheet(CString sheetName);
	BOOL deleteSheet(CString sheetName);
	BOOL judgeExcelVer(int Ver);
	CString getEndCell(CString cellBegin, int rows, int cols);
	BOOL createServer(CString officeVer);
	BOOL setCellsTypeToNum(CString cellBegin, CString cellEnd, int min, int max, int defValue);
	BOOL setCellsTypeToNum(CString cellBegin, CString cellEnd, int min, int max);
	BOOL setCellsTypeToNum( CString cellBegin, CString cellEnd, double min, double max );
	BOOL setCellsToStringList(CString cellBegin, CString cellEnd, const std::vector<CString>& strList, unsigned int defValueIndex);
	BOOL setCellsToStringList(CString cellBegin, CString cellEnd, const std::vector<CString>& strList);
	BOOL setCellsToStringList(CString cellBegin, CString cellEnd, CString sheetName, CString valueCellBegin, CString valueCellEnd);
	BOOL setCellsToStringList(CString cellBegin, CString cellEnd, const CString* strList, unsigned int len, unsigned int defValueIndex);
	BOOL setCellsToStringList(CString cellBegin, CString cellEnd, const CString* strList, unsigned int len);
	BOOL setCellsColor(CString cellBegin, CString cellEnd, int colorIndex);
	BOOL setCellsValue(CString cellBegin, long** nums, int rows, int cols);
	BOOL setCellsValue(CString cellBegin, CString** nums, int rows, int cols);
	BOOL setCellsValue( CString cellBegin, vector<vector<CString>>& nums );
	BOOL setCellsValue( CString cellBegin, vector<int>& nums );
	BOOL setRowColor(UINT rowIndex, int colorIndex);
	BOOL setCellsBold(CString cellBegin, CString cellEnd, BOOL bold);
	BOOL setCellsFont(CString cellBegin, CString cellEnd, CString fontName, int fontSize, BOOL bold = FALSE);
	BOOL splitComBox(CString comBox, std::vector<CString>& strList);
	BOOL getComBoxValue(CString comBox, const CString& key, CString& value);
	BOOL getComBoxKey( CString comBox, const CString& value, CString& key );

	void getValue( vector<vector<CString>>& data, CString cellBegin, int rows, int cols);
	void saveExcel();
	void saveExcelAs(CString savePath);
	void savaExcelToXml(CString savePath);
	long getCellRowIndex(CString cellIndex);
	void setCellsFormat( CString cellBegin, CString cellEnd, CString format );
	void setCellsAlignLeft(CString cellBegin, CString cellEnd);
	void setCellsAlignLeft(CRange range);
	void setCellsLength(CString cellBegin, CString cellEnd, UINT length);
	void setCellsLength( CString cellBegin, CString cellEnd, UINT lengthMin, UINT lengthMax );
	void setCellValue(CString cellIndex, CString value);
	void setCellsToText(CString cellBegin, CString cellEnd);
	CString getCellValue(CString cellIndex);
	void setView();
	void createSheet(CString sheetName);
	void setColumnWidth(CString cellIndex, int width);
	int getOfficeVer();
	void closeExcel();
	void unlockALL();
	void lockCells( CString cellBegin, CString cellEnd );
	void setSheetProtect(CString sheetName);
	void setSheetUnprotect();
	void lockRow(CString cellIndex);
	void lockRows(CString cellBegin, CString cellEnd);
	void addFormula(CString name, CString formula);
	void setCellsToFormula(CString cellBegin, CString cellEnd, CString formulaName);
	CString getFormula(CString sheetName, CString cellBegin, CString cellEnd, FormulaType type);
	void getMaxRange(UINT& row, UINT& column);
	BOOL openExcelFile( CString filePath );
	void getColValue(vector<CString>& outData, CString cellBegin, int count);
	BOOL createFakeServer();
	BOOL getActiveSheet();
	BOOL setColValue( CString cellBegin, const vector<CString>& nums );
private:
	CApplication ExcelApp;
	CWorkbooks books;
	CWorkbook book;
	CWorksheets sheets;
	CWorksheet sheet;
	CRange range;
	CValidation validation;
	CInterior interior;
	CFont0 font;
	CNames names;
	//CString filePath;
	LPDISPATCH lpDisp;
	//Î±ÔìµÄExcelApp
	CApplication ExcelApp_fake;
	CWorkbooks books_fake;
	int excelVer;
};

