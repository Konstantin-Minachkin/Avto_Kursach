#include "stdafx.h"
#include "ExcelReader.h"

ExcelReader::ExcelReader(CString way1, int list, Excel::_ApplicationPtr pApp1)
{
	CT2CA tmp(way1);
	std::string s(tmp);
	way = (s.c_str());
	pApp = pApp1;
	flag = false;
	try {
		File = pApp->Workbooks->Open(way);
	}
	catch (...)
	{
		flag = true;
	}
	if (!flag)
	{
		Sheet = File->Worksheets->Item[list];//sheet теперь указывает на лист №list

		pApp->PutVisible(0, FALSE);
		pApp->PutDisplayAlerts(0, FALSE);//отключить уведомления
	}

}

std::string ExcelReader::readCell(int strok, int stolb)
{
	Excel::RangePtr cell = Sheet->Cells->Item[strok][stolb];
	std::string text = (char *)_bstr_t(cell->Text);

	return text;
}

bool ExcelReader::getFlag()
{
	return flag;
}

Excel::_WorksheetPtr ExcelReader::getSheet()
{
	return Sheet;
}

ExcelReader::~ExcelReader()
{
	//	pApp->PutScreenUpdating(0, true);
	//	pApp->PutCalculation(0, Excel::xlCalculationAutomatic);
	//	pApp->PutEnableEvents(true); //Отключаем отслеживание событий	
	//	pApp->PutStatusBar(0, true);
}