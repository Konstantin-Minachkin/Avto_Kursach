#pragma once
#include "ExlLib.h"
#include <string>

class ExcelReader
{
private: bstr_t way; Excel::_ApplicationPtr pApp; Excel::_WorkbookPtr File; Excel::_WorksheetPtr Sheet; bool flag;
public:
	ExcelReader(CString way1, int list, Excel::_ApplicationPtr pApp); 
	std::string readCell(int strok, int stolb);
	bool getFlag();
	Excel::_WorksheetPtr getSheet();
	~ExcelReader();
};