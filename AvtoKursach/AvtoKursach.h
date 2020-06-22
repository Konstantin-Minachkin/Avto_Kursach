
// AvtoKursach.h : ������� ���� ��������� ��� ���������� PROJECT_NAME
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�������� stdafx.h �� ��������� ����� ����� � PCH"
#endif

#include "resource.h"		// �������� �������
#include "ExlLib.h"
#include "DocLib.h"


// CAvtoKursachApp:
// � ���������� ������� ������ ��. AvtoKursach.cpp
//

class CAvtoKursachApp : public CWinApp
{
public:
	CAvtoKursachApp();
	Excel::_ApplicationPtr exApp;
	Word::_ApplicationPtr wordApp;

// ���������������
public:
	virtual BOOL InitInstance();
	BOOL PreTranslateMessage(MSG* pMsg);
// ����������

	DECLARE_MESSAGE_MAP()
};

extern CAvtoKursachApp theApp;