
// MainDlg.h : файл заголовка
//

#pragma once
#include "afxcmn.h"
#include <string>
#include "ExcelRW.h"
#include "WordRW.h"
#include "HelpfullFunc.h"
#include "afxwin.h"

// диалоговое окно MainDlg
class MainDlg : public CDialogEx
{
// Создание
public:
	MainDlg(CWnd* pParent = NULL);	// стандартный конструктор

// Данные диалогового окна
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_AVTOKURSACH_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// поддержка DDX/DDV


// Реализация
protected:
	HICON m_hIcon;

	// Созданные функции схемы сообщений
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	
	DECLARE_MESSAGE_MAP()
public:
	CListCtrl file_list;
	int page;
	int tableHeight; //для хранения высоты таблиц, чтобы можно было норм границы указывать
	afx_msg void OnBnClickedSolve();
	CString label;
	CString users_answer;
	CString wordWay = L"D:\\Two.docx"; //путь к шаюлону, потом надо сделать еще одно окно (типа настройки), куда можно было бы загрузить шаблон, чтоб прога запомнила к нему путь
	ExcelRW* excelIO;
	WordRW* wordIO;
	int counter; //кол-во строк всяких предметов
	int i; // для проверки того, когда это кол-во строк будет достигнуто
	double sum_glob;
	int result;
	CEdit editText;
	//для уебана компилятора блять ебать
	
	afx_msg void OnBnClickedBack();
};
