#include "stdafx.h"
#include "ExcelRW.h"

ExcelRW::ExcelRW(/*CString way1,*/Excel::_ApplicationPtr pApp1)
{
	/*CT2CA tmp(way1);
	std::string s(tmp);
	way = s.c_str(); */
	pApp = pApp1;
	try {
		File = pApp->Workbooks->Add(); //добавление новой книги
	}
	catch (...)
	{
		flag = true;
	}
	if (!flag)
	{
		Sheet = File->Worksheets->Item[1];

		pApp->PutVisible(0, FALSE);
		pApp->PutDisplayAlerts(0, FALSE);//отключить уведомления
	}
}

std::string ExcelRW::readCell(int strok, int stolb)
{
	Excel::RangePtr cell = Sheet->Cells->Item[strok][stolb];
	std::string text = (char *)_bstr_t(cell->Value2);
	return text;
}

CString ExcelRW::readCell(int strok, int stolb, bool a)
{
	Excel::RangePtr cell = Sheet->Cells->Item[strok][stolb];
	CString cstr;
	cstr = cell->Text;
	return cstr;
}

void ExcelRW::writeCell(int strok, int stolb, double numb, bool bold, bool italic, int koef_okrugl)
{
	Excel::RangePtr cell = Sheet->Cells->Item[strok][stolb];
	std::wstring text_not(round_my(numb, koef_okrugl));
	bstr_t a = text_not.c_str();
	cell->NumberFormat = "0.00";
	cell->Value2 = numb;
	cell->Value2;
	if (bold) cell->Font->Bold = true;
	if (italic) cell->Font->Italic = true;
}

void ExcelRW::writeCell(int strok, int stolb, CString text, bool bold, bool italic, bool perenos) 
{
	Excel::RangePtr cell = Sheet->Cells->Item[strok][stolb];
	if (perenos) cell->WrapText = TRUE;
	std::wstring text_not(text);
	bstr_t a = text_not.c_str();
	cell->Value2 = a;
	if (bold) cell->Font->Bold = true;
	if (italic) cell->Font->Italic = true;
}

Excel::_WorkbookPtr ExcelRW::getFile()
{
	return File;
}

Excel::_WorksheetPtr ExcelRW::getSheet()
{
	return Sheet;
}

void ExcelRW::setSheet(int list)
{
	Sheet = File->Worksheets->Item[list];
}

bool ExcelRW::getFlag()
{
	return flag;
}

ExcelRW::~ExcelRW()
{
	/* файл пока не сохраняем
	way += ".xlsx";
	newFile->SaveAs(way, Excel::xlWorkbookDefault, "", "", FALSE, FALSE, Excel::XlSaveAsAccessMode::xlNoChange); //сохранение книги по адресу
	way += " - файл создан;";
	if (newFile != NULL) MessageBox(NULL, way, L" ", MB_OK | MB_ICONINFORMATION);
	*/
}