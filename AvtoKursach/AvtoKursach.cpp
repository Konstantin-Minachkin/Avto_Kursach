
// AvtoKursach.cpp : Îïðåäåëÿåò ïîâåäåíèå êëàññîâ äëÿ ïðèëîæåíèÿ.
//

#include "stdafx.h"
#include "AvtoKursach.h"
#include "MainDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAvtoKursachApp

BEGIN_MESSAGE_MAP(CAvtoKursachApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// ñîçäàíèå CAvtoKursachApp

CAvtoKursachApp::CAvtoKursachApp()
{
	// ïîääåðæêà äèñïåò÷åðà ïåðåçàãðóçêè
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: äîáàâüòå êîä ñîçäàíèÿ,
	// Ðàçìåùàåò âåñü âàæíûé êîä èíèöèàëèçàöèè â InitInstance
}


// Åäèíñòâåííûé îáúåêò CAvtoKursachApp

CAvtoKursachApp theApp;

BOOL CAvtoKursachApp::PreTranslateMessage(MSG* pMsg) //ïî íàæàòèþ íà enter íàæèìàåòñÿ êíîïêà äàëüøå
{
	if (pMsg->message == WM_KEYDOWN)
	{
		if (pMsg->wParam == VK_RETURN)
		{
			HWND hwnd = FindWindow(  NULL, L"AvtoKursach");
			HWND hButton = FindWindowEx(hwnd, NULL, NULL, L"Дальше");
			::SendMessage(hButton, BM_CLICK, NULL, NULL);
			return TRUE; // Do not process further 
		}
	}

	return CWinApp::PreTranslateMessage(pMsg);
}
// èíèöèàëèçàöèÿ CAvtoKursachApp

BOOL CAvtoKursachApp::InitInstance()
{
	// InitCommonControlsEx() òðåáóåòñÿ äëÿ Windows XP, åñëè ìàíèôåñò
	// ïðèëîæåíèÿ èñïîëüçóåò ComCtl32.dll âåðñèè 6 èëè áîëåå ïîçäíåé âåðñèè äëÿ âêëþ÷åíèÿ
	// ñòèëåé îòîáðàæåíèÿ.  Â ïðîòèâíîì ñëó÷àå áóäåò âîçíèêàòü ñáîé ïðè ñîçäàíèè ëþáîãî îêíà.
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// Âûáåðèòå ýòîò ïàðàìåòð äëÿ âêëþ÷åíèÿ âñåõ îáùèõ êëàññîâ óïðàâëåíèÿ, êîòîðûå íåîáõîäèìî èñïîëüçîâàòü
	// â âàøåì ïðèëîæåíèè.
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();

	// Ñîçäàòü äèñïåò÷åð îáîëî÷êè, â ñëó÷àå, åñëè äèàëîãîâîå îêíî ñîäåðæèò
	// ïðåäñòàâëåíèå äåðåâà îáîëî÷êè èëè êàêèå-ëèáî åãî ýëåìåíòû óïðàâëåíèÿ.
	CShellManager *pShellManager = new CShellManager;

	// Àêòèâàöèÿ âèçóàëüíîãî äèñïåò÷åðà "Êëàññè÷åñêèé Windows" äëÿ âêëþ÷åíèÿ ýëåìåíòîâ óïðàâëåíèÿ MFC
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));

	// Ñòàíäàðòíàÿ èíèöèàëèçàöèÿ
	// Åñëè ýòè âîçìîæíîñòè íå èñïîëüçóþòñÿ è íåîáõîäèìî óìåíüøèòü ðàçìåð
	// êîíå÷íîãî èñïîëíÿåìîãî ôàéëà, íåîáõîäèìî óäàëèòü èç ñëåäóþùèõ
	// êîíêðåòíûõ ïðîöåäóð èíèöèàëèçàöèè, êîòîðûå íå òðåáóþòñÿ
	// Èçìåíèòå ðàçäåë ðååñòðà, â êîòîðîì õðàíÿòñÿ ïàðàìåòðû
	// TODO: ñëåäóåò èçìåíèòü ýòó ñòðîêó íà ÷òî-íèáóäü ïîäõîäÿùåå,
	// íàïðèìåð íà íàçâàíèå îðãàíèçàöèè
	SetRegistryKey(_T("Ëîêàëüíûå ïðèëîæåíèÿ, ñîçäàííûå ñ ïîìîùüþ ìàñòåðà ïðèëîæåíèé"));

	MainDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		//äëÿ íîðìàëüíîãî âûõîäà
		// TODO: Ââåäèòå êîä äëÿ îáðàáîòêè çàêðûòèÿ äèàëîãîâîãî îêíà
		delete dlg.excelIO;
		delete dlg.wordIO;
		exApp->PutVisible(0, true);
		exApp->PutDisplayAlerts(0, true);//îòêëþ÷èòü óâåäîìëåíèÿ
		exApp->Quit(); 
		wordApp->PutVisible(true);
		wordApp->PutDisplayAlerts(Word::WdAlertLevel::wdAlertsAll);
		wordApp->Quit();
		//  ñ ïîìîùüþ êíîïêè "ÎÊ"
	}
	else if (nResponse == IDCANCEL)
	{
		//äëÿ îáðàáîòêè îøèáîê
		// TODO: Ââåäèòå êîä äëÿ îáðàáîòêè çàêðûòèÿ äèàëîãîâîãî îêíà
		exApp->PutVisible(0, true);
		exApp->PutDisplayAlerts(0, true);//îòêëþ÷èòü óâåäîìëåíèÿ
		exApp->Quit();
		wordApp->PutVisible(true);
		wordApp->PutDisplayAlerts(Word::WdAlertLevel::wdAlertsAll);
		wordApp->Quit();
		//  ñ ïîìîùüþ êíîïêè "Îòìåíà"
	}
	else if (nResponse == -1)
	{
		exApp->PutVisible(0, true);
		exApp->PutDisplayAlerts(0, true);//îòêëþ÷èòü óâåäîìëåíèÿ
		exApp->Quit();
		wordApp->PutVisible(true);
		wordApp->PutDisplayAlerts(Word::WdAlertLevel::wdAlertsAll);
		wordApp->Quit();
		TRACE(traceAppMsg, 0, "Ïðåäóïðåæäåíèå. Íå óäàëîñü ñîçäàòü äèàëîãîâîå îêíî, ïîýòîìó ðàáîòà ïðèëîæåíèÿ íåîæèäàííî çàâåðøåíà.\n");
		TRACE(traceAppMsg, 0, "Ïðåäóïðåæäåíèå. Ïðè èñïîëüçîâàíèè ýëåìåíòîâ óïðàâëåíèÿ MFC äëÿ äèàëîãîâîãî îêíà íåâîçìîæíî #define _AFX_NO_MFC_CONTROLS_IN_DIALOGS.\n");
	}

	// Óäàëèòü äèñïåò÷åð îáîëî÷êè, ñîçäàííûé âûøå.
	if (pShellManager != NULL)
	{
		delete pShellManager;
	}

#ifndef _AFXDLL
	ControlBarCleanUp();
#endif

	// Ïîñêîëüêó äèàëîãîâîå îêíî çàêðûòî, âîçâðàòèòå çíà÷åíèå FALSE, ÷òîáû ìîæíî áûëî âûéòè èç
	//  ïðèëîæåíèÿ âìåñòî çàïóñêà ãåíåðàòîðà ñîîáùåíèé ïðèëîæåíèÿ.
	return FALSE;
}
