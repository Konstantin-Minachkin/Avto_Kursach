#pragma once
#include "ExlLib.h"
#include <string>
#include "HelpfullFunc.h"

class ExcelRW
{
private: /*bstr_t way; */ Excel::_ApplicationPtr pApp; Excel::_WorkbookPtr File; Excel::_WorksheetPtr Sheet; bool flag = false;
public:
	ExcelRW(/*CString way1, */ Excel::_ApplicationPtr pApp1);
	CString readCell(int strok, int stolb, bool a);
	std::string readCell(int strok, int stolb);
	void writeCell(int strok, int stolb, CString text, bool bold = false, bool italic = false, bool perenos = true); //текст пишется в юникоде
	void writeCell(int strok, int stolb, double numb, bool bold = false, bool italic = false, int koef_okrugl = 3);
	Excel::_WorkbookPtr getFile();
	Excel::_WorksheetPtr getSheet();
	void setSheet(int list);
	bool getFlag();
	~ExcelRW();
};