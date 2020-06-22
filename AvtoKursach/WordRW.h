#pragma once
#include "DocLib.h"
#include "HelpfullFunc.h"
#include <string>

class WordRW
{
public: bstr_t way; Word::_ApplicationPtr wordApp; Word::_DocumentPtr file; bool err_flag = false;
public:
	WordRW(CString way1, Word::_ApplicationPtr wordApp1);
	void setWay(CString way);
	std::string getWay();
	bool getFlag();
	void write(std::string temp, bstr_t zakladka); //запись в utf8
	void write(CString temp, bstr_t zakladka); //запись в юникоде
	~WordRW();
};