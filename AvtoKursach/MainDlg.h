
// MainDlg.h : ���� ���������
//

#pragma once
#include "afxcmn.h"
#include <string>
#include "ExcelRW.h"
#include "WordRW.h"
#include "HelpfullFunc.h"
#include "afxwin.h"

// ���������� ���� MainDlg
class MainDlg : public CDialogEx
{
// ��������
public:
	MainDlg(CWnd* pParent = NULL);	// ����������� �����������

// ������ ����������� ����
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_AVTOKURSACH_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// ��������� DDX/DDV


// ����������
protected:
	HICON m_hIcon;

	// ��������� ������� ����� ���������
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	
	DECLARE_MESSAGE_MAP()
public:
	CListCtrl file_list;
	int page;
	int tableHeight; //��� �������� ������ ������, ����� ����� ���� ���� ������� ���������
	afx_msg void OnBnClickedSolve();
	CString label;
	CString users_answer;
	CString wordWay = L"D:\\Two.docx"; //���� � �������, ����� ���� ������� ��� ���� ���� (���� ���������), ���� ����� ���� �� ��������� ������, ���� ����� ��������� � ���� ����
	ExcelRW* excelIO;
	WordRW* wordIO;
	int counter; //���-�� ����� ������ ���������
	int i; // ��� �������� ����, ����� ��� ���-�� ����� ����� ����������
	double sum_glob;
	int result;
	CEdit editText;
	//��� ������ ����������� ����� �����
	
	afx_msg void OnBnClickedBack();
};
