#include "stdafx.h"
#include "WordRW.h"

WordRW::WordRW(CString way1, Word::_ApplicationPtr wordApp1)
{
	wordApp = wordApp1;
	way = way1;
	try {
		file = wordApp->Documents->Open(&_variant_t(way)); //если в диспетчере задач уже есть процесс word такой, то будет бесконечная загрузка
	}
	catch (...)
	{
		err_flag = true;
	}
	if (!err_flag)
	{
		wordApp->PutVisible(false);
		wordApp->PutDisplayAlerts(Word::WdAlertLevel::wdAlertsNone);
	}
	else
	{
		MessageBox(NULL, L"Ошибка при открытии шаблона", L" ", MB_OK | MB_ICONERROR);
	}
}

WordRW::~WordRW()
{
	file->SaveAs(&_variant_t(way));
	if (file == NULL) MessageBox(NULL, L"Ошибка при изменении шаблона", L" ", MB_OK | MB_ICONERROR);
}

bool WordRW::getFlag()
{
	return err_flag;
}

void WordRW::setWay(CString way)
{
	this->way = way;
}

std::string WordRW::getWay()
{
	const char* buf = way;
	int bstrlen = way.length();
	std::string STDString(buf ? buf : "", bstrlen);
	return STDString;
}

void WordRW::write(std::string temp, bstr_t zakladka)
{
	Word::RangePtr bookmark; //место для вставки документа
	bookmark = this->file->Bookmarks->Item(&_variant_t(zakladka))->Range; // нахождение некого места в документе по закладке
	bookmark->Select();
	_bstr_t a = temp.c_str();
	this->wordApp->Selection->TypeText(a);
}

void WordRW::write(CString temp, bstr_t zakladka)
{
	Word::RangePtr bookmark; //место для вставки документа
	bookmark = this->file->Bookmarks->Item(&_variant_t(zakladka))->Range; // нахождение некого места в документе по закладке
	bookmark->Select();
	std::wstring text(temp); //перевод в wstring из CString
	_bstr_t a = text.c_str();
	this->wordApp->Selection->TypeText(a);
}