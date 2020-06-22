
// AvtoKursach.h : главный файл заголовка для приложения PROJECT_NAME
//

#pragma once

#ifndef __AFXWIN_H__
	#error "включить stdafx.h до включения этого файла в PCH"
#endif

#include "resource.h"		// основные символы
#include "ExlLib.h"
#include "DocLib.h"


// CAvtoKursachApp:
// О реализации данного класса см. AvtoKursach.cpp
//

class CAvtoKursachApp : public CWinApp
{
public:
	CAvtoKursachApp();
	Excel::_ApplicationPtr exApp;
	Word::_ApplicationPtr wordApp;

// Переопределение
public:
	virtual BOOL InitInstance();
	BOOL PreTranslateMessage(MSG* pMsg);
// Реализация

	DECLARE_MESSAGE_MAP()
};

extern CAvtoKursachApp theApp;