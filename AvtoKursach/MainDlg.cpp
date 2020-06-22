
// MainDlg.cpp : файл реализации
//

#include "stdafx.h"
#include "AvtoKursach.h"
#include "MainDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// Диалоговое окно CAboutDlg используется для описания сведений о приложении

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Данные диалогового окна
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // поддержка DDX/DDV

// Реализация
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// диалоговое окно MainDlg



MainDlg::MainDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_AVTOKURSACH_DIALOG, pParent)
	, label(_T(""))
	, users_answer(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void MainDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_label, label);
	DDX_Text(pDX, IDC_answer, users_answer);
	DDX_Control(pDX, IDC_answer, editText);
}

BEGIN_MESSAGE_MAP(MainDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(BTN_Solve, &MainDlg::OnBnClickedSolve)
	ON_BN_CLICKED(BTN_BACK, &MainDlg::OnBnClickedBack)
END_MESSAGE_MAP()


// обработчики сообщений MainDlg

BOOL MainDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Добавление пункта "О программе..." в системное меню.

	// IDM_ABOUTBOX должен быть в пределах системной команды.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	// Задает значок для этого диалогового окна.  Среда делает это автоматически,
	//  если главное окно приложения не является диалоговым
	SetIcon(m_hIcon, TRUE);			// Крупный значок
	SetIcon(m_hIcon, FALSE);		// Мелкий значок

	//создание excel и word
	HRESULT hRes = E_FAIL;
	CoInitialize(NULL);
	hRes = theApp.exApp.CreateInstance("Excel.Application");
	if (FAILED(hRes))
	{
		MessageBox(L"Excel не найден!", L"Ошибка", MB_OK);
		OnCancel();
	}
	hRes = theApp.wordApp.CreateInstance("Word.Application");
	if (FAILED(hRes))
	{
		MessageBox(L"Word не найден!", L"Ошибка", MB_OK);
		OnCancel();
	}

	//инициализация excel в случае ошибки функция не продолжается
	excelIO = new ExcelRW(theApp.exApp);
	if (excelIO->getFlag())
	{
		MessageBox(L"Не удалось создать excel файл", L"Ошибка", MB_OK | MB_ICONERROR);
		delete excelIO;
		OnCancel(); //по идее должен после этого выйти из функции
	}

	//прога просит пользователя выбрать место где лежит шаблон
	bool flag = true;
	while (flag)
	{
		CFileDialog open(true);
		if (open.DoModal() == IDOK)
		{
			CString file_name;
			file_name = open.m_ofn.lpstrFile;
			if (CFileFind().FindFile(file_name) == TRUE)
			{
				if (file_name.Find(L".docx", 0) != -1)
				{
					wordWay = open.GetPathName();
					wordIO = new WordRW(wordWay, theApp.wordApp);
					flag = false;
				}
				else {
					MessageBox(L"Файл не .docx", L"Ошибка", MB_OK | MB_ICONERROR);
				}
			}
			else {
				MessageBox(L"Такого файла не существует", L"Ошибка", MB_OK | MB_ICONERROR);
			}
		}
		else {
			MessageBox(L"Ты обязан выбрать word файл в качестве шаблона", L"Ясно?", MB_OK | MB_ICONINFORMATION);
		}
	}

	//проверяет, откроется ли файл без ошибок
	if (wordIO->getFlag())
	{
		MessageBox(L"Не удалось создать word файл", L"Ошибка", MB_OK | MB_ICONERROR);
		delete wordIO;
		delete excelIO;
		OnCancel(); //по идее должен после этого выйти из функции
	}
	//если все норм, то уже в IDOK будет просиходит удаление word и excel IO
	label = L"Введите вариант";
	i = 1;
	page = 0;

	UpdateData(FALSE);
	return TRUE;  // возврат значения TRUE, если фокус не передан элементу управления
}

void MainDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// При добавлении кнопки свертывания в диалоговое окно нужно воспользоваться приведенным ниже кодом,
//  чтобы нарисовать значок.  Для приложений MFC, использующих модель документов или представлений,
//  это автоматически выполняется рабочей областью.

void MainDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // контекст устройства для рисования

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Выравнивание значка по центру клиентского прямоугольника
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Нарисуйте значок
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// Система вызывает эту функцию для получения отображения курсора при перемещении
//  свернутого окна.
HCURSOR MainDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void MainDlg::OnBnClickedSolve()
{
	CString temp_C;
	std::string temp;
	Excel::RangePtr myRange;
	Word::RangePtr tableRange;
	double sum = 0;
	int koef_okrugl = 2;

	UpdateData(TRUE); //считывание того, что понаписал пользователь в users_answer
	temp = cstr_to_str(users_answer);

	switch (page)
	{
	case 0: //вариант
		if (temp.find_first_not_of("0123456789") == std::string::npos)
		{
			wordIO->write(users_answer, L"a"); //вставка текста по названию закладки
			page++;
			//переход к следующей просьбе
			label = L"Введите название продукта";
			users_answer = L"";
			UpdateData(FALSE);
		}
		else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 1: //обработка названия продукта
		wordIO->write(users_answer, L"k"); //вставка текста по названию закладки
		wordIO->write(users_answer, L"k2"); //при нажатии назад надо бы делать так , чтобы все что было вставлено в закладку удалялось
		wordIO->write(users_answer, L"k3");
		page++;
		//переход к следующей просьбе
		label = L"Введите количество сырья и основных материалов";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 2: //кол-во сырья
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789") == std::string::npos) 
			{
				counter = std::stoi(temp) * 3; //перевод из string в int, умножаем на 3 тк нужно для каждого сырья вводить 3 параметра
				if (counter > 0)
				{
					//начинаем писать шаблон для excel

					//предустановка 
					excelIO->getSheet()->Cells->Font->Name = "Times New Roman";
					excelIO->getSheet()->Cells->Font->Size = 12;
					excelIO->getSheet()->Cells->VerticalAlignment = 2;
					excelIO->getSheet()->Cells->NumberFormat = "@"; 
					myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][1]][excelIO->getSheet()->Cells->Item[2][6]];
					myRange->HorizontalAlignment = 3;

					//ширина колламнс для первой таблицы
					myRange = excelIO->getSheet()->Columns->Item[1];
					myRange->ColumnWidth = 39;
					myRange = excelIO->getSheet()->Columns->Item[2];
					myRange->ColumnWidth = 13.29;
					myRange = excelIO->getSheet()->Columns->Item[3];
					myRange->ColumnWidth = 18.43;
					myRange = excelIO->getSheet()->Columns->Item[4];
					myRange->ColumnWidth = 23.14;
					myRange = excelIO->getSheet()->Columns->Item[5];
					myRange->ColumnWidth = 27.43;
					myRange = excelIO->getSheet()->Columns->Item[6];
					myRange->ColumnWidth = 18.86;
					//высота 
					myRange = excelIO->getSheet()->Rows->Item[1];
					myRange->RowHeight = 16.5;
					myRange = excelIO->getSheet()->Rows->Item[2];
					myRange->RowHeight = 33.75;
					myRange = excelIO->getSheet()->Rows->Item[3];
					myRange->RowHeight = 17.25;

					myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][1]][excelIO->getSheet()->Cells->Item[1][2]];
					myRange->Merge();
					myRange->Value2 = L"Статья затрат";

					myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][3]][excelIO->getSheet()->Cells->Item[1][4]];
					myRange->Merge();
					myRange->Value2 = L"На 1 т";

					myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][5]][excelIO->getSheet()->Cells->Item[1][6]];
					myRange->Merge();
					myRange->Value2 = L"На весь выпуск";

					excelIO->writeCell(2, 2, L"Цена, руб.");
					excelIO->writeCell(2, 3, L"Норма расхода, ед./т.");
					excelIO->writeCell(2, 4, L"Сумма, руб./т.");
					excelIO->writeCell(2, 5, L"Норма расхода, млн. ед.");
					excelIO->writeCell(2, 6, L"Сумма, млн. руб.");
					excelIO->writeCell(3, 1, L"1. Сырье и основные материалы, т:");

					tableHeight = 4; //пока что высота таблицы такая
					page++;
					//переход к следующей просьбе
					label = L"Введите название 1-го сырья";
					users_answer = L"";
					UpdateData(FALSE);
				}
				else MessageBox(L"Сырья должно быть больше 0", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 3: //сырье
		if (i <= counter)
		{
			if (i % 3 == 1) //вводим название
			{
				if (temp != "" && temp != " ")
				{
					//записать в excel
					excelIO->writeCell(tableHeight, 1, users_answer);

					//переход к следующей просьбе
					label.Format(L"Введите цену %d-го сырья", i / 3 + 1);
					users_answer = L"";
					UpdateData(FALSE);
					i++;
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else if (i % 3 == 2) //вводим цену
			{
				if (temp != "" && temp != " ")
				{
					if (temp.find_first_not_of("0123456789..") == std::string::npos) //тут английская и русская точки
					{
						excelIO->writeCell(tableHeight, 2, users_answer);
						//переход к следующей просьбе
						label.Format(L"Введите норму расхода %d-го сырья", i / 3 + 1);
						users_answer = L"";
						UpdateData(FALSE);
						if (counter - i == 1) page++; //это чтобы пользователь два раза не жал на кнопку дальше, после ввода нормы расхода последнего сырья
						i++;
					}
					else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else //вводим норму расхода
			{
				if (temp != "" && temp != " ")
				{
					if (temp.find_first_not_of("0123456789..") == std::string::npos)
					{
						myRange = excelIO->getSheet()->Rows->Item[tableHeight];
						myRange->RowHeight = 19.25;
						excelIO->writeCell(tableHeight++, 3, users_answer);
						if (counter - i != 0) {
							label.Format(L"Введите название %d-го сырья", i / 3 + 1);
							users_answer = L"";
							UpdateData(FALSE);
						}
						i++;
					}
					else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
		}
		break;

	case 4: //обработка ввода нормы расхода последнего сырья
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				//записать в excel, тут дубликат того что было в последнем блоке "запись в excel"
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 19.25;
				excelIO->writeCell(tableHeight++, 3, users_answer);
				//пишем дальше
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 24;
				excelIO->writeCell(tableHeight++, 1, L"Итого сырья и основных материалов", true);
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 25.5;
				excelIO->writeCell(tableHeight++, 1, L"2. Вспомогательные материалы, т:");
				label = L"Введите кол-во вспомогательных материалов";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 5: //обработка, кол-во вспомогат материалов
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789") == std::string::npos)
			{
				counter = std::stoi(temp) * 3; //перевод из string в int, умножаем на 3 тк нужно вводить 3 параметра
				if (counter > 0)
				{
					i = 1;
					page++;
					//переход к следующей просьбе
					label = L"Введите название 1-го вспомогательного материала";
					users_answer = L"";
					UpdateData(FALSE);
				}
				else MessageBox(L"Материалов должно быть больше 0", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 6://вспомогательные материалы
		if (i <= counter)
		{
			if (i % 3 == 1) //вводим название
			{
				if (temp != "" && temp != " ")
				{
					//записать в excel 
					excelIO->writeCell(tableHeight, 1, users_answer);
					//переход к следующей просьбе
					label.Format(L"Введите цену %d-го материала", i / 3 + 1);
					users_answer = L"";
					UpdateData(FALSE);
					i++;
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else if (i % 3 == 2) //вводим цену
			{
				if (temp != "" && temp != " ")
				{
					if (temp.find_first_not_of("0123456789..") == std::string::npos)
					{
						excelIO->writeCell(tableHeight, 2, users_answer);
						//переход к следующей просьбе
						label.Format(L"Введите норму расхода %d-го материала", i / 3 + 1);
						users_answer = L"";
						UpdateData(FALSE);
						if (counter - i == 1) page++; //это чтобы пользователь два раза не жал на кнопку дальше, после ввода нормы расхода последнего сырья
						i++;
					}
					else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else //вводим норму расхода
			{
				if (temp != "" && temp != " ")
				{
					if (temp.find_first_not_of("0123456789..") == std::string::npos)
					{
						myRange = excelIO->getSheet()->Rows->Item[tableHeight];
						myRange->RowHeight = 19.25;
						excelIO->writeCell(tableHeight++, 3, users_answer);
						if (counter - i != 0) {
							label.Format(L"Введите название %d-го материала", i / 3 + 1);
							users_answer = L"";
							UpdateData(FALSE);
						}
						i++;
					}
					else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
		}
		break;

	case 7://обработка ввода нормы расхода последнего материала
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				//записать в excel, тут дубликат того что было в последнем блоке "запись в excel"
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 19.25;
				excelIO->writeCell(tableHeight++, 3, users_answer);
				//пишем дальше
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 23;
				excelIO->writeCell(tableHeight++, 1, L"Итого вспомогательных материалов", true);
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 18;
				excelIO->writeCell(tableHeight++, 1, L"3. Энергозатраты");
				label = L"Введите кол-во видов энергозатрат";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 8://обработка, кол-во видов энергозатрат
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789") == std::string::npos)
			{
				counter = std::stoi(temp) * 3; //перевод из string в int, умножаем на 3 тк нужно вводить 3 параметра
				if (counter > 0)
				{
					i = 1;
					page++;
					//переход к следующей просьбе
					label = L"Введите название 1-го вида энергозатрат";
					users_answer = L"";
					UpdateData(FALSE);
				}
				else MessageBox(L"Видов энергозатрат должно быть больше 0", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 9://обработка видов энергозатрат
		if (i <= counter)
		{
			if (i % 3 == 1) //вводим название
			{
				if (temp != "" && temp != " ")
				{
					//записать в excel 
					excelIO->writeCell(tableHeight, 1, users_answer);
					//переход к следующей просьбе
					label.Format(L"Введите цену %d-го вида энергозатрат", i / 3 + 1);
					users_answer = L"";
					UpdateData(FALSE);
					i++;
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else if (i % 3 == 2) //вводим цену
			{
				if (temp != "" && temp != " ")
				{
					if (temp.find_first_not_of("0123456789..") == std::string::npos)
					{
						//записать в excel 
						excelIO->writeCell(tableHeight, 2, users_answer);
						//переход к следующей просьбе
						label.Format(L"Введите норму расхода %d-го вида энергозатрат", i / 3 + 1);
						users_answer = L"";
						UpdateData(FALSE);
						if (counter - i == 1) page++; //это чтобы пользователь два раза не жал на кнопку дальше, после ввода нормы расхода последнего сырья
						i++;
					}
					else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else //вводим норму расхода
			{
				if (temp != "" && temp != " ")
				{
					if (temp.find_first_not_of("0123456789..") == std::string::npos)
					{
						myRange = excelIO->getSheet()->Rows->Item[tableHeight];
						myRange->RowHeight = 19.25;
						excelIO->writeCell(tableHeight++, 3, users_answer);
						if (counter - i != 0) {
							label.Format(L"Введите название %d-го вида энергозатрат", i / 3 + 1);
							users_answer = L"";
							UpdateData(FALSE);
						}
						i++;
					}
					else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
		}
		break;

	case 10://обработка ввода нормы расхода последнего вида энергозатрат
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				//записать в excel, тут дубликат того что было в последнем блоке "запись в excel"
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 19.25;
				excelIO->writeCell(tableHeight++, 3, users_answer);
				//пишем дальше
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 19.25;
				excelIO->writeCell(tableHeight++, 1, L"Итого энергозатрат", true);
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 19.25;
				excelIO->writeCell(tableHeight, 1, L"4. Фонд оплаты труда");
				label = L"Введите сумму в фонд оплаты труда";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 11://обработка суммы фонда оплаты труда
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				excelIO->writeCell(tableHeight++, 4, users_answer);
				label = L"Введите процент от ФОП отчислений в обязательные страховые фонды";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 12://обработка процента от ФОП
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				double counter1 = str_to_d(temp); //использую counter уже просто как int переменную
				if (counter1 <= 100)
				{
					excelIO->writeCell(tableHeight, 7, users_answer); //запоминаем процентик
					wordIO->write(temp, L"proc");
					temp_C = L"5. Отчисления в обязательные страховые фонды(" + users_answer + L"% от ФОП)";
					myRange = excelIO->getSheet()->Rows->Item[tableHeight];
					myRange->RowHeight = 39.75;
					excelIO->writeCell(tableHeight++, 1, temp_C); //запишем процентик
					myRange = excelIO->getSheet()->Rows->Item[tableHeight];
					myRange->RowHeight = 34.5;
					excelIO->writeCell(tableHeight++, 1, L"Итого фонд оплаты труда с отчислениями", true);
					label = L"Введите расходы на содержание и эксплуатацию оборудования";
					users_answer = L"";
					UpdateData(FALSE);
					page++;
				}
				else MessageBox(L"Процентов не бывает больше 100", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 13://обработка расходов на эксплуатацию
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				excelIO->writeCell(tableHeight, 4, users_answer);
				myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[tableHeight][6]][excelIO->getSheet()->Cells->Item[tableHeight + 1][6]];
				myRange->Merge();
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 36;
				excelIO->writeCell(tableHeight++, 1, L"6. Расходы на содержание и эксплуатацию оборудования,");
				label = L"Введите расходы на аммортизацию оборудования";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 14://обработка расходов на аммортизацию
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				excelIO->writeCell(tableHeight, 4, users_answer);
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 19.25;
				excelIO->writeCell(tableHeight++, 1, L"в т. ч. амортизация");
				label = L"Введите цеховые расходы";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 15://обработка цеховых расходов
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				excelIO->writeCell(tableHeight, 4, users_answer);
				wordIO->write(temp, L"cex_rasx");
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 19.25;
				excelIO->writeCell(tableHeight++, 1, L"7. Цеховые расходы");
				label = L"Введите общехозяйственные расходы";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 16://обработка Общехозяйственные расходы
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				wordIO->write(temp, L"obshexoz");
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 19.25;
				excelIO->writeCell(tableHeight++, 1, L"8. Цеховая себестоимость");
				excelIO->writeCell(tableHeight, 4, users_answer);
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 19.25;
				excelIO->writeCell(tableHeight++, 1, L"9. Общехозяйственные расходы");
				label = L"Введите внепроизводственные расходы";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 17://обработка Внепроизводственные расходы
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				wordIO->write(temp, L"vneproizv");
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 31.5;
				excelIO->writeCell(tableHeight++, 1, L"10. Производственная себестоимость");
				excelIO->writeCell(tableHeight, 4, users_answer);
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 19.25;
				excelIO->writeCell(tableHeight++, 1, L"11. Внепроизводственные расходы");
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 19.25;
				excelIO->writeCell(tableHeight++, 1, L"12. Полная себестоимость"); //tableHeight находится за таблицей в следующей ячейке

				myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][2]][excelIO->getSheet()->Cells->Item[1][4]];
				myRange->EntireColumn->AutoFit();

				//граница ячеек
				myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][1]][excelIO->getSheet()->Cells->Item[tableHeight - 1][6]];
				myRange->Borders->Weight = Excel::xlThin;

				//копирование таблиц в word
				myRange->Copy();
				tableRange = wordIO->file->Bookmarks->Item(&_variant_t(L"Tabl1"))->Range;
				tableRange->PasteExcelTable(FALSE, TRUE, FALSE); //вставка в это место (курсор при этом не перемещается!!)
				tableHeight++;//запишем еще другие параметры через строку после таблицы

				result = MessageBox(L"Ввести сразу кап вложения или ввести список оборудования?", L"", MB_YESNO | MB_ICONQUESTION);
				if (result == IDYES)
				{
					label = L"Введите капитальные вложения в основное технологическое оборудование";
					users_answer = L"";
					UpdateData(FALSE);
					page++;
				}
				else if (result == IDNO)
				{
					label = L"Введите кол-во оборудования";
					users_answer = L"";
					UpdateData(FALSE);
					page = 30; //идем на спец страницу где будем считать объемы производства
				}
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 18://обработка капитальные вложения в основное технологическое оборудование 
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				wordIO->write(temp, L"b"); 
				wordIO->write(temp, L"b2");	
				excelIO->writeCell(tableHeight++, 1, users_answer);
				label = L"Введите объем производства";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 19://обработка объем производства
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				wordIO->write(temp, L"c");
				wordIO->write(temp, L"c2");
				excelIO->writeCell(tableHeight++, 1, users_answer);
				label = L"Введите процент увеличения объема производства на";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 20://обработка процент увеличения объема производства на
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				double counter1 = str_to_d(temp);
				if (counter1 <= 100)
				{
					wordIO->write(temp, L"d");
					wordIO->write(temp, L"d2");
					excelIO->writeCell(tableHeight++, 1, users_answer);
					label = L"Введите процент сокращения норм расхода по основному виду исходного сырья на ";
					users_answer = L"";
					UpdateData(FALSE);
					page++;
				}
				else MessageBox(L"Процентов не бывает больше 100", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 21://обработка процент сокращения норм расхода по основному виду исходного сырья на
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				double counter1 = str_to_d(temp);
				if (counter1 <= 100)
				{
					wordIO->write(temp, L"f");
					excelIO->writeCell(tableHeight++, 1, users_answer);
					label = L"Введите процент сокращения норм расхода по энергетическим ресурсам на  ";
					users_answer = L"";
					UpdateData(FALSE);
					page++;
				}
				else MessageBox(L"Процентов не бывает больше 100", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 22://обработка процент сокращения норм расхода по энергетическим ресурсам 
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				wordIO->write(temp, L"g");
				excelIO->writeCell(tableHeight++, 1, users_answer);
				label = L"Введите увеличение численности производственных рабочих на";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 23://обработка увеличение численности производственных рабочих 
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789") == std::string::npos)
			{
				wordIO->write(temp, L"e");
				wordIO->write(temp, L"e2");
				excelIO->writeCell(tableHeight++, 1, users_answer);
				label = L"Введите оклад рабочих";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 24: //обработка оклад рабочих
		wordIO->write(temp, L"h");
		wordIO->write(temp, L"h2");
		excelIO->writeCell(tableHeight, 1, users_answer);
		label = L"Введите норму годовых амортизационных отчислений";
		users_answer = L"";
		UpdateData(FALSE);
		page++;
		break;

	case 25://обработка годовых отчислений
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				double counter1 = str_to_d(temp);
				if (counter1 <= 100)
				{
					excelIO->writeCell(tableHeight + 22, 3, users_answer); //ввели норму отчислений
					label = L"Введите изменение цены ресурсов (в процентах)";
					users_answer = L"";
					UpdateData(FALSE);
					page++;
				}
				else MessageBox(L"Процентов не бывает больше 100", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 26://обработка цены ресурсов
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				double counter1 = str_to_d(temp);
				if (counter1 <= 100)
				{
					excelIO->writeCell(tableHeight - 6, 2, users_answer);
					label = L"Введите коэффициент выручки";
					users_answer = L"";
					UpdateData(FALSE);
					page++;
				}
				else MessageBox(L"Процентов не бывает больше 100", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 27://обработка коэфф выручки
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				excelIO->writeCell(tableHeight - 5, 2, users_answer);
				label = L"Введите ставку дисконтирования";
				users_answer = L"";
				UpdateData(FALSE);
				page++;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 28://обработка ставки дисконтирования
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				wordIO->write(temp, L"stavka_dicsont");
				excelIO->writeCell(tableHeight - 4, 2, users_answer);
				page++;
		//припишем в таблицу несколько вычислений и в ворд их тоже, а потом опять её скопируем
		tableHeight = 4; //просматриваем заново
		sum = 0;
		double itogo; 
		do {
			double k = str_to_d(excelIO->readCell(tableHeight, 2)) * str_to_d(excelIO->readCell(tableHeight, 3)); //рассчитаем
			sum += k;
			excelIO->writeCell(tableHeight++, 4, k);//запишем в excel
		} while (excelIO->readCell(tableHeight, 1, true) != L"Итого сырья и основных материалов");
		
		excelIO->writeCell(tableHeight, 4, sum);
		tableHeight += 2;
		itogo = sum;
		sum = 0;
		do {
			double k = str_to_d(excelIO->readCell(tableHeight, 2)) * str_to_d(excelIO->readCell(tableHeight, 3)); //рассчитаем
			sum += k;
			excelIO->writeCell(tableHeight++, 4, k);//запишем в excel
		} while (excelIO->readCell(tableHeight, 1, true) != L"Итого вспомогательных материалов");
		excelIO->writeCell(tableHeight, 4, sum);//запишем в excel
		tableHeight += 2;
		itogo += sum;
		sum = 0;
		do {
			double k = str_to_d(excelIO->readCell(tableHeight, 2)) * str_to_d(excelIO->readCell(tableHeight, 3)); //рассчитаем
			sum += k;
			excelIO->writeCell(tableHeight++, 4, k);//запишем в excel
		} while (excelIO->readCell(tableHeight, 1, true) != L"Итого энергозатрат");
		
		excelIO->writeCell(tableHeight, 4, sum);//запишем в excel
		tableHeight += 2;
		itogo += sum;
		sum = str_to_d(excelIO->readCell(tableHeight - 1, 4)) * str_to_d(excelIO->readCell(tableHeight, 7)) / 100;
		excelIO->writeCell(tableHeight++, 4, sum);
		sum += str_to_d(excelIO->readCell(tableHeight - 2, 4));
		excelIO->writeCell(tableHeight, 4, sum); //итого в фонд оплаты труда
		itogo += sum; //итого фонд оплаты
		itogo += str_to_d(excelIO->readCell(tableHeight + 1, 4));//Расходы на содержание и эксплуатацию оборудования
		tableHeight += 4;
		itogo += str_to_d(excelIO->readCell(tableHeight - 1, 4)); //читаем цеховые расходы
		excelIO->writeCell(tableHeight, 4, itogo); //пишем цеховую себестоимость
		tableHeight += 2;
		itogo += str_to_d(excelIO->readCell(tableHeight - 1, 4));//общехозяйственные расходы
		excelIO->writeCell(tableHeight++, 4, itogo); //пишем  Производственная себестоимость
		itogo += str_to_d(excelIO->readCell(tableHeight++, 4));//внепроизводственные расходы
		excelIO->writeCell(tableHeight, 4, itogo); //пишем  полную себестоимость

		myRange = excelIO->getSheet()->Cells->Item[1][4];
		myRange->EntireColumn->AutoFit();
		myRange = excelIO->getSheet()->Cells->Item[1][2];
		myRange->EntireColumn->AutoFit();

		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][1]][excelIO->getSheet()->Cells->Item[tableHeight][4]];
		myRange->Copy();
		tableRange = wordIO->file->Bookmarks->Item(&_variant_t(L"Tabl2"))->Range;
		tableRange->PasteExcelTable(FALSE, TRUE, FALSE);

		//переходим к рисованию второй таблицы
		tableHeight += 10;
		myRange = excelIO->getSheet()->Rows->Item[tableHeight];
		myRange->RowHeight = 26;
		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[tableHeight][1]][excelIO->getSheet()->Cells->Item[tableHeight][2]];
		myRange->HorizontalAlignment = 3;
		excelIO->writeCell(tableHeight, 1, L"Статья затрат", true);
		excelIO->writeCell(tableHeight++, 2, L"Доля в полной себестоимости (%)", true);
		excelIO->writeCell(tableHeight, 1, L"1. Сырьё и основные материалы");
		//ищем сырье
		counter = 4;
		{
			double k;
			do { counter++; } while (excelIO->readCell(counter, 1, true) != L"Итого сырья и основных материалов");
			sum = str_to_d(excelIO->readCell(counter, 4)) / itogo * 100;
			temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(counter, 4)), koef_okrugl)) + "/" + cstr_to_str(round_my(itogo, koef_okrugl)) 
				+ "*100% = " + cstr_to_str(round_my(sum, koef_okrugl)) + "%";
			wordIO->write(temp, L"Form1");
			k = sum;
			excelIO->writeCell(tableHeight++, 2, sum);

			excelIO->writeCell(tableHeight, 1, L"2. Вспомогательные материалы");
			do { counter++; } while (excelIO->readCell(counter, 1, true) != L"Итого вспомогательных материалов");
			sum = str_to_d(excelIO->readCell(counter, 4)) / itogo * 100;
			temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(counter, 4)), koef_okrugl)) + "/" + cstr_to_str(round_my(itogo, koef_okrugl)) 
				+ "*100% = " + cstr_to_str(round_my(sum, koef_okrugl)) + "%";
			k += sum;
			wordIO->write(temp, L"Form2");
			wordIO->write(cstr_to_str(round_my(k, koef_okrugl)), L"j");
			excelIO->writeCell(tableHeight++, 2, sum);
		}
		excelIO->writeCell(tableHeight, 1, L"3. Энергозатраты");
		//ищем энергозатраты
		do { counter++; } while (excelIO->readCell(counter, 1, true) != L"Итого энергозатрат");
		sum = str_to_d(excelIO->readCell(counter, 4)) / itogo * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(counter, 4)), koef_okrugl)) + "/" + cstr_to_str(round_my(itogo, koef_okrugl)) 
			+ "*100% = " + cstr_to_str(round_my(sum, koef_okrugl)) + "%";
		wordIO->write(temp, L"Form3");
		wordIO->write(cstr_to_str(round_my(sum, koef_okrugl)), L"i");
		excelIO->writeCell(tableHeight++, 2, sum);

		excelIO->writeCell(tableHeight, 1, L"4. Фонд оплаты труда с отчислениями");
		//ищем фонд оплаты
		do { counter++; } while (excelIO->readCell(counter, 1, true) != L"Итого фонд оплаты труда с отчислениями");
		sum = str_to_d(excelIO->readCell(counter, 4)) / itogo * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(counter, 4)), koef_okrugl)) + "/" + cstr_to_str(round_my(itogo, koef_okrugl)) 
			+ "*100% = " + cstr_to_str(round_my(sum, koef_okrugl)) + "%";
		wordIO->write(temp, L"Form4");
		excelIO->writeCell(tableHeight++, 2, sum);

		excelIO->writeCell(tableHeight, 1, L"5. Расходы на содержание и эксплуатацию оборудования");
		counter++;
		sum = str_to_d(excelIO->readCell(counter, 4)) / itogo * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(counter, 4)), koef_okrugl)) + "/" + cstr_to_str(round_my(itogo, koef_okrugl)) 
			+ "*100% = " + cstr_to_str(round_my(sum, koef_okrugl)) + "%";
		wordIO->write(temp, L"Form5");
		excelIO->writeCell(tableHeight++, 2, sum);

		excelIO->writeCell(tableHeight, 1, L"6. Цеховые расходы");
		counter += 2;
		sum = str_to_d(excelIO->readCell(counter, 4)) / itogo * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(counter, 4)), koef_okrugl)) + "/" + cstr_to_str(round_my(itogo, koef_okrugl)) 
			+ "*100% = " + cstr_to_str(round_my(sum, koef_okrugl)) + "%";
		wordIO->write(temp, L"Form6");
		excelIO->writeCell(tableHeight++, 2, sum);

		excelIO->writeCell(tableHeight, 1, L"Итого: цеховая себестоимость", false, true);
		counter++;
		sum = str_to_d(excelIO->readCell(counter, 4)) / itogo * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(counter, 4)), koef_okrugl)) + "/" + cstr_to_str(round_my(itogo, koef_okrugl)) 
			+ "*100% = " + cstr_to_str(round_my(sum, koef_okrugl)) + "%";
		wordIO->write(temp, L"Form7");
		excelIO->writeCell(tableHeight++, 2, sum);

		excelIO->writeCell(tableHeight, 1, L"7. Общезаводские расходы");
		counter++;
		sum = str_to_d(excelIO->readCell(counter, 4)) / itogo * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(counter, 4)), koef_okrugl)) + "/" + cstr_to_str(round_my(itogo, koef_okrugl)) 
			+ "*100% = " + cstr_to_str(round_my(sum, koef_okrugl)) + "%";
		wordIO->write(temp, L"Form8");
		excelIO->writeCell(tableHeight++, 2, sum);

		excelIO->writeCell(tableHeight, 1, L"Итого: производственная себестоимость", false, true);
		counter++;
		sum = str_to_d(excelIO->readCell(counter, 4)) / itogo * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(counter, 4)), koef_okrugl)) + "/" + cstr_to_str(round_my(itogo, koef_okrugl)) 
			+ "*100% = " + cstr_to_str(round_my(sum, koef_okrugl)) + "%";
		wordIO->write(temp, L"Form9");
		excelIO->writeCell(tableHeight++, 2, sum);

		excelIO->writeCell(tableHeight, 1, L"8. Внепроизводственные расходы");
		counter++;
		sum = str_to_d(excelIO->readCell(counter, 4)) / itogo * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(counter, 4)), koef_okrugl)) + "/" + cstr_to_str(round_my(itogo, koef_okrugl)) 
			+ "*100% = " + cstr_to_str(round_my(sum, koef_okrugl)) + "%";
		wordIO->write(temp, L"Form10");
		excelIO->writeCell(tableHeight++, 2, sum);

		excelIO->writeCell(tableHeight, 1, L"Итого: полная себестоимость", false, true);
		counter++;
		sum = str_to_d(excelIO->readCell(counter, 4)) / itogo * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(counter, 4)), koef_okrugl)) + "/" + cstr_to_str(round_my(itogo, koef_okrugl)) 
			+ "*100% = " + cstr_to_str(round_my(sum, koef_okrugl)) + "%";
		wordIO->write(temp, L"Form11");
		excelIO->writeCell(tableHeight, 2, L"100");

		myRange = excelIO->getSheet()->Columns->Item[1]; myRange->ColumnWidth = 45;
		myRange = excelIO->getSheet()->Columns->Item[2]; myRange->ColumnWidth = 42;

		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[tableHeight - 10][1]][excelIO->getSheet()->Cells->Item[tableHeight][2]];
		myRange->Rows->RowHeight = 19.25;

		myRange = excelIO->getSheet()->Cells->Item[1][2]; myRange->EntireColumn->AutoFit();

		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[tableHeight - 11][1]][excelIO->getSheet()->Cells->Item[tableHeight][2]];
		myRange->Borders->Weight = Excel::xlThin;

		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[tableHeight - 11][1]][excelIO->getSheet()->Cells->Item[tableHeight][2]];
		myRange->Copy();
		tableRange = wordIO->file->Bookmarks->Item(&_variant_t(L"Tabl3"))->Range;
		tableRange->PasteExcelTable(FALSE, TRUE, FALSE);

		//расчет капитальных вложений, 4 таблица
		tableHeight += 2;
		myRange = excelIO->getSheet()->Rows->Item[tableHeight];
		myRange->RowHeight = 75.75;
		excelIO->writeCell(tableHeight, 1, L"Оборудование", true);
		excelIO->writeCell(tableHeight, 2, L"Стоимость, млн. руб.", true);
		excelIO->writeCell(tableHeight, 3, L"Норма годовых амортизационных отчислений, %", true);
		excelIO->writeCell(tableHeight++, 4, L"Годовая сумма амортизационных отчислений, млн. руб.", true);

		myRange = excelIO->getSheet()->Rows->Item[tableHeight];
		myRange->RowHeight = 25.25;
		excelIO->writeCell(tableHeight, 1, L"1. Основное технологическое оборудование");
		excelIO->writeCell(tableHeight, 3, L"---------");
		excelIO->writeCell(tableHeight, 4, L"---------");
		excelIO->writeCell(tableHeight, 2, str_to_cstr(excelIO->readCell(tableHeight - 22, 1)));
		tableHeight++;

		myRange = excelIO->getSheet()->Rows->Item[tableHeight];
		myRange->RowHeight = 48;
		excelIO->writeCell(tableHeight, 1, L"2. Вспомогательное (40% от стоимости основного технологического оборудования)");
		excelIO->writeCell(tableHeight, 3, L"---------");
		excelIO->writeCell(tableHeight, 4, L"---------");
		sum = str_to_d(excelIO->readCell(tableHeight - 1, 2)) * 0.4;
		excelIO->writeCell(tableHeight, 2, sum);
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 1, 2)), koef_okrugl)) + "* 0.4 = " 
			+ cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form12");
		tableHeight++;

		myRange = excelIO->getSheet()->Rows->Item[tableHeight];
		myRange->RowHeight = 48;
		excelIO->writeCell(tableHeight, 1, L"3. Прочее (10% от стоимости основного технологического оборудования)");
		excelIO->writeCell(tableHeight, 3, L"---------");
		excelIO->writeCell(tableHeight, 4, L"---------");
		sum = str_to_d(excelIO->readCell(tableHeight - 2, 2)) * 0.1;
		excelIO->writeCell(tableHeight, 2, sum);
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 2, 2)), koef_okrugl)) + "* 0.1 = " 
			+ cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form13");
		tableHeight++;

		myRange = excelIO->getSheet()->Rows->Item[tableHeight];
		myRange->RowHeight = 25.25;
		excelIO->writeCell(tableHeight, 1, L"Итого: общая стоимость оборудования", true);
		excelIO->writeCell(tableHeight, 3, L"---------");
		excelIO->writeCell(tableHeight, 4, L"---------");
		sum = str_to_d(excelIO->readCell(tableHeight - 3, 2)) + str_to_d(excelIO->readCell(tableHeight - 2, 2)) + str_to_d(excelIO->readCell(tableHeight - 1, 2));
		excelIO->writeCell(tableHeight, 2, sum);
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 3, 2)), koef_okrugl)) + " + " + excelIO->readCell(tableHeight - 2, 2) 
			+ " + " + cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 1, 2)), koef_okrugl)) + " = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form14");
		tableHeight++;

		myRange = excelIO->getSheet()->Rows->Item[tableHeight];
		myRange->RowHeight = 32.25;
		excelIO->writeCell(tableHeight, 1, L"Транспортные расходы (20% от общей стоимости оборудования)");
		excelIO->writeCell(tableHeight, 3, L"---------");
		excelIO->writeCell(tableHeight, 4, L"---------");
		sum = str_to_d(excelIO->readCell(tableHeight - 1, 2)) * 0.2;
		excelIO->writeCell(tableHeight, 2, sum);
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 1, 2)), koef_okrugl)) + "* 0.2 = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form15");
		tableHeight++;

		myRange = excelIO->getSheet()->Rows->Item[tableHeight];
		myRange->RowHeight = 35.25;
		excelIO->writeCell(tableHeight, 1, L"Монтажные работы (40% от общей стоимости оборудования)");
		excelIO->writeCell(tableHeight, 3, L"---------");
		excelIO->writeCell(tableHeight, 4, L"---------");
		sum = str_to_d(excelIO->readCell(tableHeight - 2, 2)) * 0.4;
		excelIO->writeCell(tableHeight, 2, sum);
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 2, 2)), koef_okrugl)) + "* 0.4 = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form16");
		tableHeight++;

		myRange = excelIO->getSheet()->Rows->Item[tableHeight];
		myRange->RowHeight = 35.25;
		excelIO->writeCell(tableHeight, 1, L"Итого: капитальные вложения в оборудование", true);
		sum = str_to_d(excelIO->readCell(tableHeight - 3, 2)) + str_to_d(excelIO->readCell(tableHeight - 2, 2)) + str_to_d(excelIO->readCell(tableHeight - 1, 2));
		excelIO->writeCell(tableHeight, 2, sum);
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 3, 2)), koef_okrugl)) + " + " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 2, 2)), koef_okrugl)) + " + " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 1, 2)), koef_okrugl)) + " = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form17");
		sum = str_to_d(excelIO->readCell(tableHeight, 2)) * str_to_d(excelIO->readCell(tableHeight, 3)) / 100;
		excelIO->writeCell(tableHeight, 4, sum);
		excelIO->writeCell(tableHeight + 2, 4, sum); 
		tableHeight++;

		myRange = excelIO->getSheet()->Rows->Item[tableHeight];
		myRange->RowHeight = 43;
		excelIO->writeCell(tableHeight, 1, L"Затраты на прирост оборотных средств (10% от капитальных вложений в оборудование)");
		excelIO->writeCell(tableHeight, 3, L"---------");
		excelIO->writeCell(tableHeight, 4, L"---------");
		sum = str_to_d(excelIO->readCell(tableHeight - 1, 2)) * 0.1; 
		excelIO->writeCell(tableHeight, 2, sum);
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 1, 2)), koef_okrugl)) + "* 0,1 = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form18");
		tableHeight++;

		myRange = excelIO->getSheet()->Rows->Item[tableHeight];
		myRange->RowHeight = 17;
		excelIO->writeCell(tableHeight, 1, L"Всего: капитальные вложения", true);
		excelIO->writeCell(tableHeight, 3, L"---------");
		sum = str_to_d(excelIO->readCell(tableHeight - 2, 2)) + str_to_d(excelIO->readCell(tableHeight - 1, 2));
		excelIO->writeCell(tableHeight, 2, sum);
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 1, 2)), koef_okrugl)) + " + " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 2, 2)), koef_okrugl)) + " = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form19");

		myRange = excelIO->getSheet()->Columns->Item[1]; myRange->ColumnWidth = 45;
		myRange = excelIO->getSheet()->Columns->Item[2]; myRange->ColumnWidth = 22;
		myRange = excelIO->getSheet()->Columns->Item[3]; myRange->ColumnWidth = 19; myRange->AutoFit();
		myRange = excelIO->getSheet()->Columns->Item[4]; myRange->ColumnWidth = 24; myRange->AutoFit();

		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[tableHeight - 9][1]][excelIO->getSheet()->Cells->Item[tableHeight][4]];
		myRange->Borders->Weight = Excel::xlThin;

		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[tableHeight - 9][1]][excelIO->getSheet()->Cells->Item[tableHeight][4]];
		myRange->Copy();
		tableRange = wordIO->file->Bookmarks->Item(&_variant_t(L"Tabl4"))->Range;
		tableRange->PasteExcelTable(FALSE, TRUE, FALSE);
		// 5 таблица
		tableHeight -= 26;
		sum = str_to_d(excelIO->readCell(tableHeight - 3, 1)) * (str_to_d(excelIO->readCell(tableHeight - 2, 1)) + 100) / 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 2, 1)), koef_okrugl)) + "%) = " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 3, 1)), koef_okrugl)) + " * 1." 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 2, 1)), koef_okrugl)) + " = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form20");
		excelIO->writeCell(tableHeight, 2, sum); //годовой объем производаства будет записан в ячейку после ячейки с процентам сокращения по энерго ресурсам

		//создали копию таблицы в 1,8
		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][1]][excelIO->getSheet()->Cells->Item[tableHeight - 6][6]];
		myRange->Copy();
		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][8]][excelIO->getSheet()->Cells->Item[tableHeight - 6][13]];
		myRange->PasteSpecial(Excel::xlPasteAll, Excel::xlPasteSpecialOperationNone);
		{
			tableHeight -= 4;
			double j2 = str_to_d(excelIO->readCell(tableHeight + 1, 1));
			double j5 = str_to_d(excelIO->readCell(tableHeight, 2));
			double j6 = str_to_d(excelIO->readCell(tableHeight + 3, 1));
			double j7 = str_to_d(excelIO->readCell(tableHeight + 4, 1));
			double j8 = str_to_d(excelIO->readCell(tableHeight + 5, 1));
			double j9 = str_to_d(excelIO->readCell(tableHeight + 6, 1));
			double j10 = str_to_d(excelIO->readCell(tableHeight + 4, 2));
			double j12 = str_to_d(excelIO->readCell(tableHeight + 2, 2));
			double d47 = str_to_d(excelIO->readCell(tableHeight + 28, 4));
			//начинаем строить таблицу другую, тк иначе дальше не продвинемся
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 33.75;
			excelIO->writeCell(tableHeight, 4, L"Статья затрат", true);
			excelIO->writeCell(tableHeight, 5, L"Методика расчёта", true);
			excelIO->writeCell(tableHeight++, 6, L"Изменение, млн.руб.", true);
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 42.75;
			excelIO->writeCell(tableHeight, 4, L"1. Амортизация оборудования и транспортных средств");
			excelIO->writeCell(tableHeight, 5, L"Сумма годовых амортизационных отчислений");
			excelIO->writeCell(tableHeight++, 6, d47);	
			itogo = d47;
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 35.5;
			excelIO->writeCell(tableHeight, 4, L"2. Эксплуатация оборудования и транспортных средств");
			excelIO->writeCell(tableHeight, 5, L"3,5% стоимости вводимого оборудования");
			sum = str_to_d(excelIO->readCell(tableHeight + 28, 2))*0.035;
			excelIO->writeCell(tableHeight++, 6, sum);
			itogo += sum;
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 37.5;
			excelIO->writeCell(tableHeight, 4, L"3. Текущий ремонт оборудования и транспортных средств");
			excelIO->writeCell(tableHeight, 5, L"7% стоимости вводимого оборудования");
			sum = str_to_d(excelIO->readCell(tableHeight + 27, 2))*0.07;			
			itogo += sum;
			excelIO->writeCell(tableHeight++, 6, sum);
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 49.5;
			excelIO->writeCell(tableHeight, 4, L"4. Внутризаводское перемещение грузов");
			excelIO->writeCell(tableHeight, 5, L"30% от затрат на текущий ремонт оборудования и транспортных средств");
			sum = str_to_d(excelIO->readCell(tableHeight - 1, 6))*0.3;
			//itogo += sum;
			excelIO->writeCell(tableHeight++, 6, sum);
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 23.75;
			excelIO->writeCell(tableHeight, 4, L"5. Прочие расходы");
			excelIO->writeCell(tableHeight, 5, L"3% от суммы статей 1-3");
			sum = itogo * 0.03;
			excelIO->writeCell(tableHeight++, 6, sum);			
			itogo += sum + str_to_d(excelIO->readCell(tableHeight - 2, 6));
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 21;
			excelIO->writeCell(tableHeight, 4, L"Итого: изменение РСЭО", true);
			excelIO->writeCell(tableHeight, 6, itogo);
			double c57 = itogo;

			myRange = excelIO->getSheet()->Columns->Item[4]; myRange->ColumnWidth = 41;
			myRange = excelIO->getSheet()->Columns->Item[5]; myRange->ColumnWidth = 35;
			myRange = excelIO->getSheet()->Columns->Item[6]; myRange->ColumnWidth = 19; myRange->EntireColumn->AutoFit();
			
			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[tableHeight - 6][4]][excelIO->getSheet()->Cells->Item[tableHeight][6]];
			myRange->Borders->Weight = Excel::xlThin;

			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[tableHeight - 6][4]][excelIO->getSheet()->Cells->Item[tableHeight][6]];
			myRange->Copy();
			tableRange = wordIO->file->Bookmarks->Item(&_variant_t(L"Tabl6"))->Range;
			tableRange->PasteExcelTable(FALSE, TRUE, FALSE);
			//
			myRange = excelIO->getSheet()->Rows->Item[1]; myRange->RowHeight = 16.5;
			myRange = excelIO->getSheet()->Rows->Item[2]; myRange->RowHeight = 50.25;
			myRange = excelIO->getSheet()->Rows->Item[3]; myRange->RowHeight = 33.75;
			myRange = excelIO->getSheet()->Rows->Item[4]; myRange->RowHeight = 19.25;
			
			myRange = excelIO->getSheet()->Columns->Item[8]; myRange->ColumnWidth = 32.43;
			myRange = excelIO->getSheet()->Columns->Item[9]; myRange->ColumnWidth = 12.43;
			myRange = excelIO->getSheet()->Columns->Item[10]; myRange->ColumnWidth = 12.5;
			myRange = excelIO->getSheet()->Columns->Item[11]; myRange->ColumnWidth = 12.5;
			myRange = excelIO->getSheet()->Columns->Item[12]; myRange->ColumnWidth = 12.5;
			myRange = excelIO->getSheet()->Columns->Item[13]; myRange->ColumnWidth = 12.5;
			
			tableHeight = 4;
			sum = 0;
			itogo = 0;
			//пересчет значений для каждой строчки
			do
			{
				myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 19.25;
				double k = str_to_d(excelIO->readCell(tableHeight, 9)) * (100 - j5) / 100;
				excelIO->writeCell(tableHeight, 9, k);
				k = str_to_d(excelIO->readCell(tableHeight, 10)) * (100 - j6) / 100;
				excelIO->writeCell(tableHeight, 10, k);
				k = str_to_d(excelIO->readCell(tableHeight, 9)) * str_to_d(excelIO->readCell(tableHeight, 10));
				sum += k;
				excelIO->writeCell(tableHeight, 11, k);
				k = str_to_d(excelIO->readCell(tableHeight, 10)) * j2;
				excelIO->writeCell(tableHeight, 12, k);
				k = str_to_d(excelIO->readCell(tableHeight, 11)) * j10;
				excelIO->writeCell(tableHeight++, 13, k);
			} while (excelIO->readCell(tableHeight, 8, true) != L"Итого сырья и основных материалов");
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 33;
			myRange = excelIO->getSheet()->Rows->Item[tableHeight + 1]; myRange->RowHeight = 31.5;
			excelIO->writeCell(tableHeight, 11, sum);
			excelIO->writeCell(tableHeight, 13, sum*j10);
			tableHeight += 2;
			itogo += sum;
			sum = 0;

			do
			{
				myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 19.25;
				double k = str_to_d(excelIO->readCell(tableHeight, 9)) * (100 - j5) / 100;
				excelIO->writeCell(tableHeight, 9, k);
				k = str_to_d(excelIO->readCell(tableHeight, 10)) * (100 - j6) / 100;
				excelIO->writeCell(tableHeight, 10, k);
				k = str_to_d(excelIO->readCell(tableHeight, 9)) * str_to_d(excelIO->readCell(tableHeight, 10));
				sum += k;
				excelIO->writeCell(tableHeight, 11, k);
				k = str_to_d(excelIO->readCell(tableHeight, 10)) * j2;
				excelIO->writeCell(tableHeight, 12, k);
				k = str_to_d(excelIO->readCell(tableHeight, 11)) * j10;
				excelIO->writeCell(tableHeight++, 13, k);
			} while (excelIO->readCell(tableHeight, 8, true) != L"Итого вспомогательных материалов");
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 33;
			myRange = excelIO->getSheet()->Rows->Item[tableHeight + 1]; myRange->RowHeight = 22;
			excelIO->writeCell(tableHeight, 11, sum);
			excelIO->writeCell(tableHeight, 13, sum*j10);
			tableHeight += 2;
			itogo += sum;
			sum = 0;

			do
			{
				myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 19.25;
				double k = str_to_d(excelIO->readCell(tableHeight, 9)) * (100 - j5) / 100;
				excelIO->writeCell(tableHeight, 9, k);
				k = str_to_d(excelIO->readCell(tableHeight, 10)) * (100 - j7) / 100;
				excelIO->writeCell(tableHeight, 10, k);
				k = str_to_d(excelIO->readCell(tableHeight, 9)) * str_to_d(excelIO->readCell(tableHeight, 10));
				sum += k;
				excelIO->writeCell(tableHeight, 11, k);
				k = str_to_d(excelIO->readCell(tableHeight, 10)) * j2;
				excelIO->writeCell(tableHeight, 12, k);
				k = str_to_d(excelIO->readCell(tableHeight, 11)) * j10;
				excelIO->writeCell(tableHeight++, 13, k);
			} while (excelIO->readCell(tableHeight, 8, true) != L"Итого энергозатрат");
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 19.25;
			myRange = excelIO->getSheet()->Rows->Item[tableHeight + 1]; myRange->RowHeight = 19.25;
			excelIO->writeCell(tableHeight, 11, sum);
			excelIO->writeCell(tableHeight++, 13, sum*j10);
			itogo += sum;
			//фонд оплаты труда
			{
				std::string a;
				sum = str_to_d(excelIO->readCell(tableHeight, 4)) * j2;
				temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 11)), koef_okrugl)) + " * " 
					+ cstr_to_str(round_my(j2, koef_okrugl)) + " = " + cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form21");
				temp = cstr_to_str(round_my(j8, koef_okrugl)) + " * " + cstr_to_str(round_my(j9, koef_okrugl)) + " * 12 = " 
					+ cstr_to_str(round_my(j8 * j9 * 12, koef_okrugl));
				wordIO->write(temp, L"Form22");
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " + " + cstr_to_str(round_my(j8 * j9 * 12, koef_okrugl)) + " = " 
					+ cstr_to_str(round_my(sum + j8 * j9 * 12, koef_okrugl));
				wordIO->write(temp, L"Form23");
				sum = sum + j8 * j9 * 12;
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " / " 
					+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight + 16, 2)), koef_okrugl)) + " = " 
					+ cstr_to_str(round_my(sum / str_to_d(excelIO->readCell(tableHeight + 16, 2)), koef_okrugl));
				wordIO->write(temp, L"Form24");

				sum = str_to_d(excelIO->readCell(tableHeight, 4)) * j2 + (j8 * j9 * 12) / 1000000; 
				excelIO->writeCell(tableHeight, 13, sum);
				sum = str_to_d(excelIO->readCell(tableHeight, 13)) / j10;
				excelIO->writeCell(tableHeight++, 11, sum);
			}
			//отчисления в страховые фонды
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 36;
			sum = str_to_d(excelIO->readCell(tableHeight, 7)) / 100;
			temp = " * " + cstr_to_str(round_my(sum, koef_okrugl)) + " = " 
				+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 1, 11)), koef_okrugl)) + " * " 
				+ cstr_to_str(round_my(sum, koef_okrugl)) + " = ";
			sum = str_to_d(excelIO->readCell(tableHeight - 1, 11)) * str_to_d(excelIO->readCell(tableHeight, 7)) / 100;
			excelIO->writeCell(tableHeight, 11, sum);
			temp += cstr_to_str(round_my(sum, koef_okrugl));
			wordIO->write(temp, L"Form26");
			sum = str_to_d(excelIO->readCell(tableHeight, 7)) / 100;
			temp = " * " + cstr_to_str(round_my(sum, koef_okrugl)) + " = " + cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 1, 13)), koef_okrugl)) 
				+ " * " + cstr_to_str(round_my(sum, koef_okrugl)) + " = ";
			sum = str_to_d(excelIO->readCell(tableHeight - 1, 13)) * str_to_d(excelIO->readCell(tableHeight, 7)) / 100;
			temp += cstr_to_str(round_my(sum, koef_okrugl));
			wordIO->write(temp, L"Form25");
			sum = str_to_d(excelIO->readCell(tableHeight, 11)) * j10;
			excelIO->writeCell(tableHeight++, 13, sum);
			//Итого фонд оплаты труда с отчислениями
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 36;
			sum = str_to_d(excelIO->readCell(tableHeight - 1, 11)) + str_to_d(excelIO->readCell(tableHeight - 2, 11));
			excelIO->writeCell(tableHeight, 11, sum);
			itogo += sum;
			sum = str_to_d(excelIO->readCell(tableHeight, 11)) * j10;
			excelIO->writeCell(tableHeight++, 13, sum);
			//6. Расходы на содержание и эксплуатацию оборудования,
			myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 39;
			sum = str_to_d(excelIO->readCell(tableHeight, 4)) * j2;
			temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 4)), koef_okrugl)) + " * " + cstr_to_str(round_my(j2, koef_okrugl)) 
				+ " = " + cstr_to_str(round_my(sum, koef_okrugl));
			wordIO->write(temp, L"Form27");
			temp = cstr_to_str(round_my(sum, koef_okrugl)) + " + " + cstr_to_str(round_my(c57, koef_okrugl)) + " = ";
			sum += c57;
			temp += cstr_to_str(round_my(sum, koef_okrugl));
			wordIO->write(temp, L"Form28");
			temp = cstr_to_str(round_my(sum, koef_okrugl)) + " / " + cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight + 13, 2)), koef_okrugl)) + " = ";
			sum = sum / str_to_d(excelIO->readCell(tableHeight + 13, 2));
			temp += cstr_to_str(round_my(sum, koef_okrugl));
			wordIO->write(temp, L"Form29");

			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[tableHeight][13]][excelIO->getSheet()->Cells->Item[tableHeight + 1][13]];
			myRange->Merge();
			sum = str_to_d(excelIO->readCell(tableHeight, 11)) * j2 + c57;
			excelIO->writeCell(tableHeight, 13, sum);
			sum = str_to_d(excelIO->readCell(tableHeight, 13)) / j10;
			excelIO->writeCell(tableHeight++, 11, sum);
			itogo += sum;
			//аммортизация
			{
				myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 22;
				sum = str_to_d(excelIO->readCell(tableHeight, 4)) * j2 / 0.04;
				temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 4)), koef_okrugl)) + " * " + cstr_to_str(round_my(j2, koef_okrugl)) + " / 0.04 = " + cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form45");
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " * 1.1 = ";
				sum = sum * 1.1;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form46");
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " * 0.022 = ";
				sum = sum * 0.022;
				double ni_baz = sum;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form47");
				sum = str_to_d(excelIO->readCell(tableHeight, 4)) * j2 / 0.04 * 1.1;
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " + " + cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight + 8, 1)), koef_okrugl)) + " = ";
				sum += str_to_d(excelIO->readCell(tableHeight + 8, 1));
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form48");
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " * 0,022 = ";
				sum = sum * 0.022;
				double ni_proect = sum;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form49");

				sum = (str_to_d(excelIO->readCell(tableHeight, 11)) * j2 + d47) / j10;
				excelIO->writeCell(tableHeight++, 11, sum);
				//itogo += sum;
				//цеховые расходы
				myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 22;
				temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 4)), koef_okrugl)) + " * " 
					+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight + 8, 1)), koef_okrugl)) + " = ";
				sum = str_to_d(excelIO->readCell(tableHeight, 4)) * str_to_d(excelIO->readCell(tableHeight + 8, 1));
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form30");
				temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 5, 4)), koef_okrugl)) + " * 0,1 = ";
				sum = str_to_d(excelIO->readCell(tableHeight - 5, 4)) * 0.1;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form31");
				temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight - 5, 11)), koef_okrugl)) + " * 0,1 = ";
				sum = str_to_d(excelIO->readCell(tableHeight - 5, 11)) * 0.1;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form32");
				sum = str_to_d(excelIO->readCell(tableHeight - 5, 4)) * 0.1;
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " - ";
				sum = str_to_d(excelIO->readCell(tableHeight - 5, 11)) * 0.1;
				temp += cstr_to_str(round_my(sum, koef_okrugl)) + " = ";
				sum = str_to_d(excelIO->readCell(tableHeight - 5, 4)) * 0.1 - sum;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form33");
				temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 4)), koef_okrugl)) + " - " + cstr_to_str(round_my(sum, koef_okrugl)) + " = ";
				sum = str_to_d(excelIO->readCell(tableHeight, 4)) - sum;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form34");
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " * " + cstr_to_str(round_my(j10, koef_okrugl)) + " = ";
				sum = sum * j10;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form35");
				sum = str_to_d(excelIO->readCell(tableHeight, 11)) - (str_to_d(excelIO->readCell(tableHeight - 5, 4)) * 0.1 - str_to_d(excelIO->readCell(tableHeight - 5, 11)) * 0.1);
				excelIO->writeCell(tableHeight, 11, sum);
				itogo += sum;
				sum = str_to_d(excelIO->readCell(tableHeight, 11)) * j10;
				excelIO->writeCell(tableHeight++, 13, sum);
				//8. Цеховая себестоимость
				myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 22;
				excelIO->writeCell(tableHeight, 11, itogo);
				sum = str_to_d(excelIO->readCell(tableHeight, 11)) * j10;
				excelIO->writeCell(tableHeight++, 13, sum);
				//Общехоз расходы
				myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 22;
				sum = str_to_d(excelIO->readCell(tableHeight, 4)) * str_to_d(excelIO->readCell(tableHeight + 6, 1));
				temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 4)), koef_okrugl)) + " * " 
					+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight + 6, 1)), koef_okrugl)) + " = " 
					+ cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form36");
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " * 0.05 * " + cstr_to_str(round_my(j10, koef_okrugl)) + " / " 
					+ cstr_to_str(round_my(j2, koef_okrugl)) + " + " + cstr_to_str(round_my(sum, koef_okrugl)) + " * 0.95 = ";
				sum = sum * 0.05 * j10 / j2 + sum * 0.95;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form37");
				temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 4)), koef_okrugl)) + " * 0.05 + " 
					+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 4)), koef_okrugl)) + " * 0.95 * " 
					+ cstr_to_str(round_my(j2, koef_okrugl)) + " / " + cstr_to_str(round_my(j10, koef_okrugl)) + " = ";
				sum = str_to_d(excelIO->readCell(tableHeight, 4)) * 0.05 + str_to_d(excelIO->readCell(tableHeight, 4)) * 0.95 * j2 / j10;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form38");
				itogo += sum;
				excelIO->writeCell(tableHeight, 11, sum);
				sum = str_to_d(excelIO->readCell(tableHeight, 11)) * j10;
				excelIO->writeCell(tableHeight++, 13, sum);
				//10. Производственная себестоимость
				myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 42;
				excelIO->writeCell(tableHeight, 11, itogo);
				sum = str_to_d(excelIO->readCell(tableHeight, 11)) * j10;
				excelIO->writeCell(tableHeight++, 13, sum);
				//внепроизводств расходы
				myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 36.75;
				temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 4)), koef_okrugl)) + " * " + cstr_to_str(round_my(j2, koef_okrugl)) + " = ";
				sum = str_to_d(excelIO->readCell(tableHeight, 4)) * j2;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form39");
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " * 0.9 * " + cstr_to_str(round_my(j10, koef_okrugl)) + " / " 
					+ cstr_to_str(round_my(j2, koef_okrugl)) + " + " + cstr_to_str(round_my(sum, koef_okrugl)) + " * 0,1 = ";
				sum = sum * 0.9 * j10 / j2 + sum * 0.1;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form40");
				temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 4)), koef_okrugl)) + " * 0,9 + " 
					+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 4)), koef_okrugl)) + " * 0,1 * "	+ cstr_to_str(round_my(j2, koef_okrugl)) 
					+ " * " + cstr_to_str(round_my(j10, koef_okrugl)) + " = ";
				sum = str_to_d(excelIO->readCell(tableHeight, 4)) * 0.9 + str_to_d(excelIO->readCell(tableHeight, 4)) * 0.1 * j2 / j10;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form41");
				itogo += sum;
				excelIO->writeCell(tableHeight, 11, sum);
				sum = str_to_d(excelIO->readCell(tableHeight, 11)) * j10;
				excelIO->writeCell(tableHeight++, 13, sum);
				//12. Полная себестоимость
				myRange = excelIO->getSheet()->Rows->Item[tableHeight]; myRange->RowHeight = 26;
				sum = str_to_d(excelIO->readCell(tableHeight, 4)) * str_to_d(excelIO->readCell(tableHeight + 3, 2));
				temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 4)), koef_okrugl)) + " * " 
					+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight + 3, 2)), koef_okrugl)) + " = " 
					+ cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form42");
				temp = "(" + cstr_to_str(round_my(sum, koef_okrugl)) + " - " + cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 4)), koef_okrugl)) 
					+ ") * " + cstr_to_str(round_my(j2, koef_okrugl)) + " = ";
				sum = (sum - str_to_d(excelIO->readCell(tableHeight, 4))) * j2;
				double p_baz = sum;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form43");
				sum = str_to_d(excelIO->readCell(tableHeight, 4)) * str_to_d(excelIO->readCell(tableHeight + 3, 2));
				temp = "(" + cstr_to_str(round_my(sum, koef_okrugl)) + " - " + cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight, 11)), koef_okrugl)) + " * " + cstr_to_str(round_my(j10, koef_okrugl)) + " = ";
				sum = (sum - str_to_d(excelIO->readCell(tableHeight, 11))) * j10;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				double p_proect = sum;
				wordIO->write(temp, L"Form44");

				excelIO->writeCell(tableHeight, 11, itogo);
				sum = str_to_d(excelIO->readCell(tableHeight, 11)) * j10;
				excelIO->writeCell(tableHeight, 13, sum);//остались в последней строке таблицы

				sum = p_baz - ni_baz;
				temp = cstr_to_str(round_my(p_baz, koef_okrugl)) + " - " + cstr_to_str(round_my(ni_baz, koef_okrugl)) + " = " 
					+ cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form50");
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " * 0,2 = ";
				sum = sum * 0.2;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form52");
				sum = p_proect - ni_proect;
				temp = cstr_to_str(round_my(p_proect, koef_okrugl)) + " - " + cstr_to_str(round_my(ni_proect, koef_okrugl)) + " = " 
					+ cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form51");
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " * 0,2 = ";
				sum = sum * 0.2;
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form53");
				sum = p_baz - ni_proect - (p_baz - ni_baz) * 0.2;
				temp = cstr_to_str(round_my(p_baz, koef_okrugl)) + " - " + cstr_to_str(round_my(ni_proect, koef_okrugl)) + " - " 
					+ cstr_to_str(round_my((p_baz - ni_baz) * 0.2, koef_okrugl)) + " = " + cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form54");
				sum = p_proect - ni_proect - (p_proect - ni_proect) * 0.2;
				temp = cstr_to_str(round_my(p_proect, koef_okrugl)) + " - " + d_to_str(ni_proect) + " - " 
					+ cstr_to_str(round_my((p_proect - ni_proect) * 0.2, koef_okrugl)) + " = " + cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form55");
				temp = cstr_to_str(round_my(sum, koef_okrugl)) + " - " + cstr_to_str(round_my(p_baz - ni_proect - (p_baz - ni_baz) * 0.2, koef_okrugl)) 
					+ " = ";
				sum = sum - (p_baz - ni_proect - (p_baz - ni_baz) * 0.2);
				temp += cstr_to_str(round_my(sum, koef_okrugl));
				wordIO->write(temp, L"Form56");
			}
			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][9]][excelIO->getSheet()->Cells->Item[1][13]];
			myRange->EntireColumn->AutoFit();

			//копирование в 4.1 только до энергозатрат
			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][8]][excelIO->getSheet()->Cells->Item[tableHeight - 11][13]];
			myRange->Copy();
			tableRange = wordIO->file->Bookmarks->Item(&_variant_t(L"Tabl5"))->Range;
			tableRange->PasteExcelTable(FALSE, TRUE, FALSE);
			//копирование в 4.6
			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][8]][excelIO->getSheet()->Cells->Item[tableHeight][13]];
			myRange->Copy();
			tableRange = wordIO->file->Bookmarks->Item(&_variant_t(L"Tabl7"))->Range;
			tableRange->PasteExcelTable(FALSE, TRUE, FALSE);
			//таблица 8
			myRange = excelIO->getSheet()->Rows->Item[1];
			myRange->RowHeight = 16.5;
			myRange = excelIO->getSheet()->Rows->Item[2];
			myRange->RowHeight = 51;
			myRange = excelIO->getSheet()->Rows->Item[3];
			myRange->RowHeight = 17.25;
			myRange = excelIO->getSheet()->Rows->Item[4];
			myRange->RowHeight = 37.5;
			myRange = excelIO->getSheet()->Rows->Item[5];
			myRange->RowHeight = 37.5;
			myRange = excelIO->getSheet()->Rows->Item[6];
			myRange->RowHeight = 25.5;
			myRange = excelIO->getSheet()->Rows->Item[7];
			myRange->RowHeight = 25.5;
			myRange = excelIO->getSheet()->Rows->Item[8];
			myRange->RowHeight = 25.5;
			myRange = excelIO->getSheet()->Rows->Item[9];
			myRange->RowHeight = 25.5;
			myRange = excelIO->getSheet()->Rows->Item[10];
			myRange->RowHeight = 37.5;
			myRange = excelIO->getSheet()->Rows->Item[11];
			myRange->RowHeight = 37.5;
			myRange = excelIO->getSheet()->Rows->Item[12];
			myRange->RowHeight = 19.25;
			myRange = excelIO->getSheet()->Columns->Item[15];
			myRange->ColumnWidth = 19.57;
			myRange = excelIO->getSheet()->Columns->Item[16];
			myRange->ColumnWidth = 11.71;
			myRange = excelIO->getSheet()->Columns->Item[17];
			myRange->ColumnWidth = 19.43;
			myRange = excelIO->getSheet()->Columns->Item[18];
			myRange->ColumnWidth = 9.6;
			myRange = excelIO->getSheet()->Columns->Item[19];
			myRange->ColumnWidth = 18.43;
			myRange = excelIO->getSheet()->Columns->Item[20];
			myRange->ColumnWidth = 9;
			myRange = excelIO->getSheet()->Columns->Item[21];
			myRange->ColumnWidth = 19.57;
			myRange = excelIO->getSheet()->Columns->Item[22];
			myRange->ColumnWidth = 10;
			myRange = excelIO->getSheet()->Columns->Item[23];
			myRange->ColumnWidth = 18;
			excelIO->writeCell(1, 15, L"Статья затрат");
			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][15]][excelIO->getSheet()->Cells->Item[3][15]];
			myRange->Merge();
			excelIO->writeCell(1, 15, L"Статья затрат");
			excelIO->writeCell(1, 16, L"Переменные затраты", true);
			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][16]][excelIO->getSheet()->Cells->Item[1][19]];
			myRange->Merge();
			excelIO->writeCell(1, 20, L"Постоянные затраты", true);
			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][20]][excelIO->getSheet()->Cells->Item[1][23]];
			myRange->Merge();
			excelIO->writeCell(2, 16, L"До внедрения", true);
			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[2][16]][excelIO->getSheet()->Cells->Item[2][17]];
			myRange->Merge();
			excelIO->writeCell(2, 18, L"После внедрения");
			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[2][18]][excelIO->getSheet()->Cells->Item[2][19]];
			myRange->Merge();
			excelIO->writeCell(3, 16, L"Доля, %");
			excelIO->writeCell(3, 17, L"Сумма, млн. руб.");
			excelIO->writeCell(3, 18, L"Доля, %");
			excelIO->writeCell(3, 19, L"Сумма, млн. руб.");
			excelIO->writeCell(3, 20, L"Доля, %");
			excelIO->writeCell(3, 21, L"Сумма, млн. руб.");
			excelIO->writeCell(3, 22, L"Доля, %");
			excelIO->writeCell(3, 23, L"Сумма, млн. руб.");
			excelIO->writeCell(4, 15, L"Основные материалы");
			excelIO->writeCell(5, 15, L"Вспомогательные материалы");
			excelIO->writeCell(6, 15, L"Энергозатраты");
			excelIO->writeCell(7, 15, L"ФОТ и отчисления");
			excelIO->writeCell(8, 15, L"РСЭО(без амортизации)");
			excelIO->writeCell(9, 15, L"Цеховые расходы");
			excelIO->writeCell(10, 15, L"Общезаводские расходы");
			excelIO->writeCell(11, 15, L"Внепроизводственные расходы");
			excelIO->writeCell(12, 15, L"Итого :", true);
			excelIO->writeCell(4, 16, L"100");
			excelIO->writeCell(5, 16, L"100");
			excelIO->writeCell(6, 16, L"50");
			excelIO->writeCell(7, 16, L"40");
			excelIO->writeCell(8, 16, L"45");
			excelIO->writeCell(9, 16, L"5");
			excelIO->writeCell(10, 16, L"5");
			excelIO->writeCell(11, 16, L"90");
			//
			excelIO->writeCell(4, 18, L"100");
			excelIO->writeCell(5, 18, L"100");
			excelIO->writeCell(6, 18, L"50");
			excelIO->writeCell(7, 18, L"40");
			excelIO->writeCell(8, 18, L"45");
			excelIO->writeCell(9, 18, L"5");
			excelIO->writeCell(10, 18, L"5");
			excelIO->writeCell(11, 18, L"90");
			//
			excelIO->writeCell(4, 20, L"0");
			excelIO->writeCell(5, 20, L"0");
			excelIO->writeCell(6, 20, L"50");
			excelIO->writeCell(7, 20, L"60");
			excelIO->writeCell(8, 20, L"55");
			excelIO->writeCell(9, 20, L"95");
			excelIO->writeCell(10, 20, L"95");
			excelIO->writeCell(11, 20, L"10");
			//
			excelIO->writeCell(4, 22, L"0");
			excelIO->writeCell(5, 22, L"0");
			excelIO->writeCell(6, 22, L"50");
			excelIO->writeCell(7, 22, L"60");
			excelIO->writeCell(8, 22, L"55");
			excelIO->writeCell(9, 22, L"95");
			excelIO->writeCell(10, 22, L"95");
			excelIO->writeCell(11, 22, L"10");

			tableHeight = 4;
			do { tableHeight++; } while (excelIO->readCell(tableHeight, 1, true) != L"Итого сырья и основных материалов");
			sum = str_to_d(excelIO->readCell(tableHeight, 4)) * j2;
			excelIO->writeCell(4, 17, sum);
			excelIO->writeCell(4, 19, str_to_d(excelIO->readCell(tableHeight, 13))); //считали q6 (не обязательно)
			do { tableHeight++; } while (excelIO->readCell(tableHeight, 1, true) != L"Итого вспомогательных материалов");
			sum = str_to_d(excelIO->readCell(tableHeight, 4)) * j2;
			excelIO->writeCell(5, 17, sum);
			excelIO->writeCell(5, 19, str_to_d(excelIO->readCell(tableHeight, 13))); //считали q10
			do { tableHeight++; } while (excelIO->readCell(tableHeight, 1, true) != L"Итого энергозатрат");
			sum = str_to_d(excelIO->readCell(tableHeight, 4)) * j2 / 2; //должно быть умножение на 50%, но тк доли НЕ меняются, то можно и так оставить
			excelIO->writeCell(6, 17, sum);
			excelIO->writeCell(6, 21, sum);
			sum = str_to_d(excelIO->readCell(tableHeight, 13)) / 2;
			excelIO->writeCell(6, 19, sum);
			excelIO->writeCell(6, 23, sum);
			//фот и отчисления
			tableHeight += 3;
			sum = str_to_d(excelIO->readCell(tableHeight, 4)) * 0.4 * j2;
			excelIO->writeCell(7, 17, sum);
			sum = str_to_d(excelIO->readCell(tableHeight, 13)) * 0.4;
			excelIO->writeCell(7, 19, sum);
			sum = str_to_d(excelIO->readCell(7, 17)) * 0.6 / 0.4;
			excelIO->writeCell(7, 21, sum);
			sum = str_to_d(excelIO->readCell(7, 19)) * 0.6 / 0.4;
			excelIO->writeCell(7, 23, sum);
			//рсэо
			tableHeight += 2;
			sum = str_to_d(excelIO->readCell(tableHeight - 1, 4)) * j2 * 0.45;
			excelIO->writeCell(8, 17, sum);
			sum = str_to_d(excelIO->readCell(tableHeight - 1, 13)) * 0.45;
			excelIO->writeCell(8, 19, sum);
			sum = str_to_d(excelIO->readCell(8, 17)) * 0.55 / 0.45;
			excelIO->writeCell(8, 21, sum);
			sum = str_to_d(excelIO->readCell(8, 19)) * 0.55 / 0.45;
			excelIO->writeCell(8, 23, sum);
			//Цеховые расходы
			tableHeight++; //20 строка (для алгоритма нет)
			sum = str_to_d(excelIO->readCell(tableHeight, 4)) * 0.05 * j2;
			excelIO->writeCell(9, 17, sum);
			sum = str_to_d(excelIO->readCell(tableHeight, 13)) * 0.05;
			excelIO->writeCell(9, 19, sum);
			sum = str_to_d(excelIO->readCell(9, 17)) * 0.95 / 0.05;
			excelIO->writeCell(9, 21, sum);
			sum = str_to_d(excelIO->readCell(9, 19)) * 0.95 / 0.05;
			excelIO->writeCell(9, 23, sum);
			//Общезаводские расходы
			tableHeight += 2;
			sum = str_to_d(excelIO->readCell(tableHeight, 4)) * 0.05 * j2;
			excelIO->writeCell(10, 17, sum);
			sum = str_to_d(excelIO->readCell(tableHeight, 13)) * 0.05;
			excelIO->writeCell(10, 19, sum);
			sum = str_to_d(excelIO->readCell(10, 17)) * 0.95 / 0.05;
			excelIO->writeCell(10, 21, sum);
			sum = str_to_d(excelIO->readCell(10, 19)) * 0.95 / 0.05;
			excelIO->writeCell(10, 23, sum);
			//Внепроизводственные расходы
			tableHeight += 2;
			sum = str_to_d(excelIO->readCell(tableHeight, 4)) * 0.9 * j2;
			excelIO->writeCell(11, 17, sum);
			sum = str_to_d(excelIO->readCell(tableHeight, 13)) * 0.9;
			excelIO->writeCell(11, 19, sum);
			sum = str_to_d(excelIO->readCell(11, 17)) * 0.1 / 0.9;
			excelIO->writeCell(11, 21, sum);
			sum = str_to_d(excelIO->readCell(11, 19)) * 0.1 / 0.9;
			excelIO->writeCell(11, 23, sum);
			//итого
			sum = str_to_d(excelIO->readCell(4, 17)) + str_to_d(excelIO->readCell(5, 17)) + str_to_d(excelIO->readCell(6, 17))
				+ str_to_d(excelIO->readCell(7, 17)) + str_to_d(excelIO->readCell(8, 17)) + str_to_d(excelIO->readCell(9, 17))
				+ str_to_d(excelIO->readCell(10, 17)) + str_to_d(excelIO->readCell(11, 17));
			excelIO->writeCell(12, 17, sum); 

			sum = str_to_d(excelIO->readCell(4, 19)) + str_to_d(excelIO->readCell(5, 19))
				+ str_to_d(excelIO->readCell(6, 19)) + str_to_d(excelIO->readCell(7, 19)) + str_to_d(excelIO->readCell(8, 19))
				+ str_to_d(excelIO->readCell(9, 19)) + str_to_d(excelIO->readCell(10, 19)) + str_to_d(excelIO->readCell(11, 19));
			excelIO->writeCell(12, 19, sum);
			sum = str_to_d(excelIO->readCell(6, 21)) + str_to_d(excelIO->readCell(7, 21)) + str_to_d(excelIO->readCell(8, 21))
				+ str_to_d(excelIO->readCell(9, 21)) + str_to_d(excelIO->readCell(10, 21)) + str_to_d(excelIO->readCell(11, 21));
			excelIO->writeCell(12, 21, sum);
			sum = str_to_d(excelIO->readCell(6, 23)) + str_to_d(excelIO->readCell(7, 23)) + str_to_d(excelIO->readCell(8, 23))
				+ str_to_d(excelIO->readCell(9, 23)) + str_to_d(excelIO->readCell(10, 23)) + str_to_d(excelIO->readCell(11, 23));
			excelIO->writeCell(12, 23, sum);

			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][15]][excelIO->getSheet()->Cells->Item[12][23]];
			myRange->Borders->Weight = Excel::xlThin;
			
			myRange = excelIO->getSheet()->Cells->Item[1][17];
			myRange->EntireColumn->AutoFit();
			myRange = excelIO->getSheet()->Cells->Item[1][19];
			myRange->EntireColumn->AutoFit();
			myRange = excelIO->getSheet()->Cells->Item[1][20];
			myRange->EntireColumn->AutoFit();
			myRange = excelIO->getSheet()->Cells->Item[1][23];
			myRange->EntireColumn->AutoFit();

			myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[1][15]][excelIO->getSheet()->Cells->Item[12][23]];
			myRange->Copy();
			tableRange = wordIO->file->Bookmarks->Item(&_variant_t(L"Tabl8"))->Range;
			tableRange->PasteExcelTable(FALSE, TRUE, FALSE);
		}
		//таблица 9
		myRange = excelIO->getSheet()->Rows->Item[14]; myRange->RowHeight = 19.25;
		myRange = excelIO->getSheet()->Rows->Item[15]; myRange->RowHeight = 19.25;
		myRange = excelIO->getSheet()->Rows->Item[16]; myRange->RowHeight = 19.25;
		myRange = excelIO->getSheet()->Rows->Item[17]; myRange->RowHeight = 19.25;
		myRange = excelIO->getSheet()->Rows->Item[18]; myRange->RowHeight = 19.25;
		myRange = excelIO->getSheet()->Rows->Item[19]; myRange->RowHeight = 19.25;
		myRange = excelIO->getSheet()->Rows->Item[20]; myRange->RowHeight = 19.25;
		myRange = excelIO->getSheet()->Rows->Item[21]; myRange->RowHeight = 36.75;
		myRange = excelIO->getSheet()->Rows->Item[22]; myRange->RowHeight = 24.25;
		myRange = excelIO->getSheet()->Rows->Item[23]; myRange->RowHeight = 19.25;
		myRange = excelIO->getSheet()->Rows->Item[24]; myRange->RowHeight = 33.75;
		myRange = excelIO->getSheet()->Rows->Item[25]; myRange->RowHeight = 26.25;
		myRange = excelIO->getSheet()->Rows->Item[26]; myRange->RowHeight = 49.5;
		myRange = excelIO->getSheet()->Rows->Item[27]; myRange->RowHeight = 26.25;
		myRange = excelIO->getSheet()->Rows->Item[28]; myRange->RowHeight = 35.25;
		myRange = excelIO->getSheet()->Rows->Item[29]; myRange->RowHeight = 35.5;
		myRange = excelIO->getSheet()->Rows->Item[30]; myRange->RowHeight = 31.5;
		myRange = excelIO->getSheet()->Rows->Item[31]; myRange->RowHeight = 30.75;
		myRange = excelIO->getSheet()->Rows->Item[32]; myRange->RowHeight = 39.75;
		myRange = excelIO->getSheet()->Rows->Item[33]; myRange->RowHeight = 26.25;
		myRange = excelIO->getSheet()->Rows->Item[34]; myRange->RowHeight = 21;
		myRange = excelIO->getSheet()->Rows->Item[35]; myRange->RowHeight = 21;
		myRange = excelIO->getSheet()->Rows->Item[36]; myRange->RowHeight = 33;
		myRange = excelIO->getSheet()->Rows->Item[37]; myRange->RowHeight = 45.75;
		myRange = excelIO->getSheet()->Rows->Item[38]; myRange->RowHeight = 48;
		myRange = excelIO->getSheet()->Rows->Item[39]; myRange->RowHeight = 41;
		myRange = excelIO->getSheet()->Rows->Item[40]; myRange->RowHeight = 37.5;
		myRange = excelIO->getSheet()->Rows->Item[41]; myRange->RowHeight = 35.25;
		myRange = excelIO->getSheet()->Rows->Item[42]; myRange->RowHeight = 65.25;
		myRange = excelIO->getSheet()->Rows->Item[43]; myRange->RowHeight = 63;

		myRange = excelIO->getSheet()->Columns->Item[15]; myRange->ColumnWidth = 35;
		myRange = excelIO->getSheet()->Columns->Item[16]; myRange->ColumnWidth = 20;
		myRange = excelIO->getSheet()->Columns->Item[17]; myRange->ColumnWidth = 20;

		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[14][15]][excelIO->getSheet()->Cells->Item[15][15]];
		myRange->Merge();
		excelIO->writeCell(14, 15, L"Поток реальных денег", true);
		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[14][16]][excelIO->getSheet()->Cells->Item[14][17]];
		myRange->Merge();
		excelIO->writeCell(14, 16, L"Номер периода времени", true);
		excelIO->writeCell(15, 16, L"0", true);
		excelIO->writeCell(15, 17, L"1", true);
		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[16][15]][excelIO->getSheet()->Cells->Item[16][17]];
		myRange->Merge();
		myRange->HorizontalAlignment = 2;
		excelIO->writeCell(16, 15, L"От операционной деятельности", false, true);
		excelIO->writeCell(17, 15, L"1. Объём продаж, млн. т"); 

		tableHeight = 22; //мин размер таблицы
		while (excelIO->readCell(tableHeight, 1, true) != L"12. Полная себестоимость") { tableHeight++; }
		excelIO->writeCell(17, 16, str_to_d(excelIO->readCell(tableHeight + 3, 1))); //j2
		excelIO->writeCell(17, 17, str_to_d(excelIO->readCell(tableHeight + 6, 2))); //j10

		excelIO->writeCell(18, 15, L"2. Цена, руб. / т");
		sum = str_to_d(excelIO->readCell(tableHeight, 4)) * str_to_d(excelIO->readCell(tableHeight + 3, 2));
		excelIO->writeCell(18, 16, sum);
		excelIO->writeCell(18, 17, sum);

		excelIO->writeCell(19, 15, L"3. Выручка[(1)*(2)], млн.руб.");
		sum = str_to_d(excelIO->readCell(17, 16)) * str_to_d(excelIO->readCell(18, 16));
		excelIO->writeCell(19, 16, sum);
		sum = str_to_d(excelIO->readCell(17, 17)) * str_to_d(excelIO->readCell(18, 17));
		excelIO->writeCell(19, 17, sum);

		excelIO->writeCell(20, 15, L"4. Внереализационные доходы");
		excelIO->writeCell(20, 16, L"---------");
		excelIO->writeCell(20, 17, L"---------");

		excelIO->writeCell(21, 15, L"5. Переменные затраты(сырьё, материалы и др.), млн.руб.");
		excelIO->writeCell(21, 16, str_to_d(excelIO->readCell(12, 17)));
		excelIO->writeCell(21, 17, str_to_d(excelIO->readCell(12, 19)));
		
		excelIO->writeCell(22, 15, L"6. Постоянные затраты, млн.руб.");
		excelIO->writeCell(22, 16, str_to_d(excelIO->readCell(12, 21)));
		excelIO->writeCell(22, 17, str_to_d(excelIO->readCell(12, 23)));

		excelIO->writeCell(23, 15, L"7. Амортизация зданий, млн.руб.");
		excelIO->writeCell(23, 16, L"---------");
		excelIO->writeCell(23, 17, L"---------");

		excelIO->writeCell(24, 15, L"8. Амортизация оборудования, млн.руб.");
		excelIO->writeCell(24, 16, str_to_d(excelIO->readCell(tableHeight - 6, 4)));
		excelIO->writeCell(24, 17, str_to_d(excelIO->readCell(tableHeight - 6, 11)));

		excelIO->writeCell(25, 15, L"9.Проценты по кредитам, млн.руб.");
		excelIO->writeCell(25, 16, L"---------");
		excelIO->writeCell(25, 17, L"---------");

		excelIO->writeCell(26, 15, L"10. Прибыль до вычета налогов[(3) + (4) - (5) - (6) - (7) - (8) - (9)], млн.руб.");
		sum = str_to_d(excelIO->readCell(19, 16)) - str_to_d(excelIO->readCell(21, 16)) - str_to_d(excelIO->readCell(22, 16))
			- str_to_d(excelIO->readCell(24, 16));
		excelIO->writeCell(26, 16, sum);
		sum = str_to_d(excelIO->readCell(19, 17)) - str_to_d(excelIO->readCell(21, 17)) - str_to_d(excelIO->readCell(22, 17))
			- str_to_d(excelIO->readCell(24, 17));
		excelIO->writeCell(26, 17, sum);

		//для начала просчитаем ОПФ
		sum = str_to_d(excelIO->readCell(tableHeight - 6, 4))* str_to_d(excelIO->readCell(tableHeight + 3, 1)) / 0.04 * 1.1;
		excelIO->writeCell(tableHeight + 7, 2, sum);

		excelIO->writeCell(27, 15, L"11. Налог на имущество, млн.руб.");
		sum = sum * 0.022;
		excelIO->writeCell(27, 16, sum);
		sum = (str_to_d(excelIO->readCell(tableHeight + 7, 2)) + str_to_d(excelIO->readCell(tableHeight + 24, 2))) * 0.022;
		excelIO->writeCell(27, 17, sum);

		excelIO->writeCell(28, 15, L"12. Налогооблагаемая прибыль[(10) - (11)], млн.руб.");
		sum = str_to_d(excelIO->readCell(26, 16)) - str_to_d(excelIO->readCell(27, 16));
		excelIO->writeCell(28, 16, sum);
		sum = sum * 0.2;
		excelIO->writeCell(29, 16, sum);
		sum = str_to_d(excelIO->readCell(26, 17)) - str_to_d(excelIO->readCell(27, 17));
		excelIO->writeCell(28, 17, sum);
		sum = sum * 0.2;
		excelIO->writeCell(29, 17, sum);
		excelIO->writeCell(29, 15, L"13. Налог на прибыль[0, 2 * (12)], млн.руб.");

		excelIO->writeCell(30, 15, L"14. Проектируемый чистый доход[(12) - (13)]");
		sum = str_to_d(excelIO->readCell(28, 16)) - str_to_d(excelIO->readCell(29, 16));
		excelIO->writeCell(30, 16, sum);
		sum = str_to_d(excelIO->readCell(28, 17)) - str_to_d(excelIO->readCell(29, 17));
		excelIO->writeCell(30, 17, sum);

		excelIO->writeCell(31, 15, L"15. Амортизация[(7) + (8)], млн.руб.");
		excelIO->writeCell(31, 16, str_to_d(excelIO->readCell(24, 16)));
		excelIO->writeCell(31, 17, str_to_d(excelIO->readCell(24, 17)));

		excelIO->writeCell(32, 15, L"16. Чистый поток от операций(Rt - Зt)[(14) + (15)], млн.руб.");
		sum = str_to_d(excelIO->readCell(30, 16)) + str_to_d(excelIO->readCell(31, 16));
		excelIO->writeCell(32, 16, sum);
		sum = str_to_d(excelIO->readCell(30, 17)) + str_to_d(excelIO->readCell(31, 17));
		excelIO->writeCell(32, 17, sum);

		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[33][15]][excelIO->getSheet()->Cells->Item[33][17]];
		myRange->Merge();
		myRange->HorizontalAlignment = 2;
		excelIO->writeCell(33, 15, L"От инвестиционной деятельности"); 

		excelIO->writeCell(34, 15, L"17. Земля, млн.руб.");
		excelIO->writeCell(34, 16, L"---------");
		excelIO->writeCell(34, 17, L"---------");

		excelIO->writeCell(35, 15, L"18. Здания, сооружения, млн.руб.");
		excelIO->writeCell(35, 16, L"---------");
		excelIO->writeCell(35, 17, L"---------");

		excelIO->writeCell(36, 15, L"19. Машины, оборудование, инструмент, инвентарь, млн.руб.");
		excelIO->writeCell(36, 16, L"---------");
		excelIO->writeCell(36, 17, str_to_d(excelIO->readCell(tableHeight + 32, 2)));

		excelIO->writeCell(37, 15, L"20. Нематериальные активы, млн.руб.");
		excelIO->writeCell(37, 16, L"---------");
		excelIO->writeCell(37, 17, L"---------");

		excelIO->writeCell(38, 15, L"21. Итого вложений в основной капитал[(17) + (18) + (19) + (20)], млн.руб.");
		excelIO->writeCell(38, 16, L"---------");
		excelIO->writeCell(38, 17, str_to_d(excelIO->readCell(36, 17)));
	
		excelIO->writeCell(39, 15, L"22. Прирост оборотного капитала, млн.руб.");
		excelIO->writeCell(39, 16, L"---------");
		excelIO->writeCell(39, 17, str_to_d(excelIO->readCell(tableHeight + 32, 4)));

		excelIO->writeCell(40, 15, L"23. Всего инвестиций(Kt)[(21) + (22)], млн.руб.");
		excelIO->writeCell(40, 16, L"---------");
		sum = str_to_d(excelIO->readCell(38, 17)) + str_to_d(excelIO->readCell(39, 17));
		excelIO->writeCell(40, 17, sum);

		excelIO->writeCell(41, 15, L"24. Поток наличности(Rt - Зt - Kt)[(16) - (23)], млн.руб.");
		sum = str_to_d(excelIO->readCell(32, 16)) - str_to_d(excelIO->readCell(40, 17));
		excelIO->writeCell(41, 16, sum);
		excelIO->writeCell(42, 16, sum);
		excelIO->writeCell(43, 16, sum);
		sum = str_to_d(excelIO->readCell(32, 17)) - str_to_d(excelIO->readCell(40, 17));
		excelIO->writeCell(41, 17, sum);

		excelIO->writeCell(42, 15, L"25. Дисконтированный поток наличности((Rt - Зt) / (1 + E)t - (Kt / (1 + E)t), млн.руб.");
		sum = sum / (1 + str_to_d(excelIO->readCell(tableHeight + 4, 2))) - str_to_d(excelIO->readCell(40, 17)) / (1 + str_to_d(excelIO->readCell(tableHeight + 4, 2)));
		excelIO->writeCell(42, 17, sum);
		//костыль тк знак сигма должен быть в юникоде, для создания юникодных строчек, нужно перед ними ставить знак L
		myRange = excelIO->getSheet()->Cells->Item[43][15];
		myRange->Value2 = L"26. Накопленный дисконтированный поток наличности(Ʃ(Rt - Зt) / (1 + E)t - (ƩKt / (1 + E)t), млн.руб.";
		sum = str_to_d(excelIO->readCell(42, 16)) + str_to_d(excelIO->readCell(42, 17));
		excelIO->writeCell(43, 17, sum);

		myRange = excelIO->getSheet()->Cells->Item[1][17];
		myRange->EntireColumn->AutoFit();
		myRange = excelIO->getSheet()->Cells->Item[1][19];
		myRange->EntireColumn->AutoFit();

		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[14][15]][excelIO->getSheet()->Cells->Item[43][17]];
		myRange->Borders->Weight = Excel::xlThin;

		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[14][15]][excelIO->getSheet()->Cells->Item[43][17]];
		myRange->Copy();
		tableRange = wordIO->file->Bookmarks->Item(&_variant_t(L"Tabl9"))->Range;
		tableRange->PasteExcelTable(FALSE, TRUE, FALSE);

		sum = str_to_d(excelIO->readCell(36, 17)) / (str_to_d(excelIO->readCell(30, 17)) - str_to_d(excelIO->readCell(30, 16)));
		temp = d_to_str(str_to_d(excelIO->readCell(36, 17))) + " / (" + cstr_to_str(round_my(str_to_d(excelIO->readCell(30, 17)), koef_okrugl)) 
			+ " - " + cstr_to_str(round_my(str_to_d(excelIO->readCell(30, 16)), koef_okrugl)) + ") = " + d_to_str(sum); //d_to_str(str_to_d( - нужно чтобы числа писалиьс не огромные в ворд
		wordIO->write(temp, L"Form57");

		temp = d_to_str(sum) + " * 365 = ";
		sum *= 365;
		temp += d_to_str(sum);
		wordIO->write(temp, L"Form58");

		//10 таблица
		myRange = excelIO->getSheet()->Rows->Item[14]; myRange->RowHeight = 36.75;
		myRange = excelIO->getSheet()->Rows->Item[15]; myRange->RowHeight = 19.25;
		myRange = excelIO->getSheet()->Rows->Item[16]; myRange->RowHeight = 19.25;
		myRange = excelIO->getSheet()->Rows->Item[17]; myRange->RowHeight = 36.75;
		myRange = excelIO->getSheet()->Rows->Item[18]; myRange->RowHeight = 38.25;
		myRange = excelIO->getSheet()->Rows->Item[19]; myRange->RowHeight = 35.25;
		myRange = excelIO->getSheet()->Rows->Item[20]; myRange->RowHeight = 22.5;
		myRange = excelIO->getSheet()->Rows->Item[21]; myRange->RowHeight = 23.5;
		myRange = excelIO->getSheet()->Rows->Item[22]; myRange->RowHeight = 24.25;
		myRange = excelIO->getSheet()->Rows->Item[23]; myRange->RowHeight = 24;
		myRange = excelIO->getSheet()->Rows->Item[24]; myRange->RowHeight = 24;
		myRange = excelIO->getSheet()->Rows->Item[25]; myRange->RowHeight = 26.25;
		myRange = excelIO->getSheet()->Rows->Item[26]; myRange->RowHeight = 39.75;
		myRange = excelIO->getSheet()->Rows->Item[27]; myRange->RowHeight = 26.25;
		myRange = excelIO->getSheet()->Rows->Item[28]; myRange->RowHeight = 35.25;

		myRange = excelIO->getSheet()->Columns->Item[19]; myRange->ColumnWidth = 40;
		myRange = excelIO->getSheet()->Columns->Item[20]; myRange->ColumnWidth = 19;
		myRange = excelIO->getSheet()->Columns->Item[21]; myRange->ColumnWidth = 19;

		excelIO->writeCell(14, 19, L"Показатель", true);
		excelIO->writeCell(14, 20, L"До внедрения", true);
		excelIO->writeCell(14, 21, L"После внедрения", true);
		excelIO->writeCell(15, 19, L"1. Годовой объём производства в выражении");
		excelIO->writeCell(16, 19, L"     натуральном, млн. т");
		excelIO->writeCell(16, 20, str_to_d(excelIO->readCell(tableHeight + 3, 1)));
		excelIO->writeCell(16, 21, str_to_d(excelIO->readCell(tableHeight + 6, 2)));

		excelIO->writeCell(17, 19, L"     стоимостном, млн. руб.");
		excelIO->writeCell(17, 20, str_to_d(excelIO->readCell(19, 16)));
		excelIO->writeCell(17, 21, str_to_d(excelIO->readCell(19, 17)));

		excelIO->writeCell(18, 19, L"2. Стоимость основных производственных фондов, млн. руб.");
		excelIO->writeCell(18, 20, str_to_d(excelIO->readCell(tableHeight + 7, 2)));
		sum = str_to_d(excelIO->readCell(tableHeight + 7, 2)) + str_to_d(excelIO->readCell(tableHeight + 2, 1));
		excelIO->writeCell(18, 21, sum);

		excelIO->writeCell(19, 19, L"3. Себестоимость единицы продукции, руб/т");
		excelIO->writeCell(19, 20, str_to_d(excelIO->readCell(tableHeight, 4)));
		excelIO->writeCell(19, 21, str_to_d(excelIO->readCell(tableHeight, 11)));

		excelIO->writeCell(20, 19, L"4. Себестоимость годового выпуска, млн. руб/год");
		sum = str_to_d(excelIO->readCell(tableHeight, 4)) * str_to_d(excelIO->readCell(tableHeight + 3, 1));
		excelIO->writeCell(20, 20, sum);
		excelIO->writeCell(20, 21, str_to_d(excelIO->readCell(tableHeight, 13)));

		excelIO->writeCell(21, 19, L"5. Прибыль от реализации, млн. руб/год");
		excelIO->writeCell(21, 20, str_to_d(excelIO->readCell(26, 16)));
		excelIO->writeCell(21, 21, str_to_d(excelIO->readCell(26, 17)));

		excelIO->writeCell(22, 19, L"6. Чистая прибыль, млн. руб/год");
		excelIO->writeCell(22, 20, str_to_d(excelIO->readCell(30, 16)));
		excelIO->writeCell(22, 21, str_to_d(excelIO->readCell(30, 17)));

		excelIO->writeCell(23, 19, L"7. Рентабельность производства, %");
		sum = (str_to_d(excelIO->readCell(18, 16)) - str_to_d(excelIO->readCell(tableHeight, 4))) * str_to_d(excelIO->readCell(tableHeight + 3, 1))
			/ str_to_d(excelIO->readCell(tableHeight + 7, 2)) * 100;
		excelIO->writeCell(23, 20, sum);
		sum = (str_to_d(excelIO->readCell(18, 17)) - str_to_d(excelIO->readCell(tableHeight, 11)))* str_to_d(excelIO->readCell(tableHeight + 6, 2))
			/ (str_to_d(excelIO->readCell(tableHeight + 7, 2)) + str_to_d(excelIO->readCell(tableHeight + 2, 1))) * 100;
		excelIO->writeCell(23, 21, sum);

		excelIO->writeCell(24, 19, L"8. Рентабельность продукции, %");
		sum = (str_to_d(excelIO->readCell(18, 16)) - str_to_d(excelIO->readCell(tableHeight, 4))) / str_to_d(excelIO->readCell(tableHeight, 4)) * 100;
		excelIO->writeCell(24, 20, sum);
		sum = (str_to_d(excelIO->readCell(18, 17)) - str_to_d(excelIO->readCell(tableHeight, 11))) / str_to_d(excelIO->readCell(tableHeight, 11)) * 100;
		excelIO->writeCell(24, 21, sum);

		excelIO->writeCell(25, 19, L"9. Рентабельность продаж, %");
		sum = (str_to_d(excelIO->readCell(18, 16)) - str_to_d(excelIO->readCell(tableHeight, 4))) / str_to_d(excelIO->readCell(18, 16)) * 100;
		excelIO->writeCell(25, 20, sum);
		sum = (str_to_d(excelIO->readCell(18, 17)) - str_to_d(excelIO->readCell(tableHeight, 11))) / str_to_d(excelIO->readCell(18, 17)) * 100;
		excelIO->writeCell(25, 21, sum);

		excelIO->writeCell(26, 19, L"10. Уровень затрат на 1 руб. товарной продукции, руб.");
		sum = str_to_d(excelIO->readCell(tableHeight, 4)) / str_to_d(excelIO->readCell(18, 16));
		excelIO->writeCell(26, 20, sum);
		sum = str_to_d(excelIO->readCell(tableHeight, 11)) / str_to_d(excelIO->readCell(18, 17));
		excelIO->writeCell(26, 21, sum);

		excelIO->writeCell(27, 19, L"11. Срок окупаемости, лет");
		excelIO->writeCell(27, 20, L"---------");
		sum = str_to_d(excelIO->readCell(36, 17)) / (str_to_d(excelIO->readCell(30, 17)) - str_to_d(excelIO->readCell(30, 16)));
		excelIO->writeCell(27, 21, sum);

		excelIO->writeCell(28, 19, L"12. Критический объём реализации (точка без убыточности проекта), т");
		sum = str_to_d(excelIO->readCell(12, 21)) / (str_to_d(excelIO->readCell(18, 16)) - str_to_d(excelIO->readCell(12, 17)) * str_to_d(excelIO->readCell(tableHeight + 3, 1)));
		excelIO->writeCell(28, 20, sum);
		sum = str_to_d(excelIO->readCell(12, 23)) / (str_to_d(excelIO->readCell(18, 17)) - str_to_d(excelIO->readCell(12, 19)) * str_to_d(excelIO->readCell(tableHeight + 6, 2)));
		excelIO->writeCell(28, 21, sum);

		myRange = excelIO->getSheet()->Cells->Item[1][20];
		myRange->EntireColumn->AutoFit();
		myRange = excelIO->getSheet()->Cells->Item[1][23];
		myRange->EntireColumn->AutoFit();

		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[14][19]][excelIO->getSheet()->Cells->Item[28][21]];
		myRange->Borders->Weight = Excel::xlThin;

		myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[14][19]][excelIO->getSheet()->Cells->Item[28][21]];
		myRange->Copy();
		tableRange = wordIO->file->Bookmarks->Item(&_variant_t(L"Tabl10"))->Range;
		tableRange->PasteExcelTable(FALSE, TRUE, FALSE);

		//7 пункт
		sum = str_to_d(excelIO->readCell(22, 16)) / (str_to_d(excelIO->readCell(18, 16)) - str_to_d(excelIO->readCell(21, 16)) * str_to_d(excelIO->readCell(tableHeight + 3, 1)));
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(22, 16)), koef_okrugl)) + " / (" + cstr_to_str(round_my(str_to_d(excelIO->readCell(18, 16)), koef_okrugl)) 
			+ " - " + cstr_to_str(round_my(str_to_d(excelIO->readCell(21, 16)), koef_okrugl)) + " * " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight + 3, 1)), koef_okrugl)) + ") = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form59");
		wordIO->write(cstr_to_str(round_my(sum, koef_okrugl)), L"l");
		sum = str_to_d(excelIO->readCell(22, 17)) / (str_to_d(excelIO->readCell(18, 17)) - str_to_d(excelIO->readCell(21, 17)) * str_to_d(excelIO->readCell(tableHeight + 6, 2)));
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(22, 17)), koef_okrugl)) + " / (" 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(18, 17)), koef_okrugl)) + " - " + cstr_to_str(round_my(str_to_d(excelIO->readCell(21, 17)), koef_okrugl)) 
			+ " * " + cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight + 6, 2)), koef_okrugl)) + ") = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form60");
		wordIO->write(excelIO->readCell(tableHeight + 6, 2), L"m");
		wordIO->write(cstr_to_str(round_my(sum, koef_okrugl)), L"n");
		sum = (str_to_d(excelIO->readCell(18, 16)) - str_to_d(excelIO->readCell(19, 20))) / str_to_d(excelIO->readCell(19, 20)) * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(18, 16)), koef_okrugl)) + " - " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(19, 20)), koef_okrugl)) + " / " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(19, 20)), koef_okrugl)) + " * 100% = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form61");
		sum = (str_to_d(excelIO->readCell(18, 16)) - str_to_d(excelIO->readCell(19, 21))) / str_to_d(excelIO->readCell(19, 21)) * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(18, 16)), koef_okrugl)) + " - " + cstr_to_str(round_my(str_to_d(excelIO->readCell(19, 21)), koef_okrugl)) 
			+ " / " + cstr_to_str(round_my(str_to_d(excelIO->readCell(19, 21)), koef_okrugl)) + " * 100% = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form62");
		sum = (str_to_d(excelIO->readCell(18, 16)) - str_to_d(excelIO->readCell(19, 20))) / str_to_d(excelIO->readCell(18, 16)) * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(18, 16)), koef_okrugl)) + " - " + cstr_to_str(round_my(str_to_d(excelIO->readCell(19, 20)), koef_okrugl)) 
			+ " / " + cstr_to_str(round_my(str_to_d(excelIO->readCell(18, 16)), koef_okrugl)) + " * 100% = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form63");
		sum = (str_to_d(excelIO->readCell(18, 16)) - str_to_d(excelIO->readCell(19, 21))) / str_to_d(excelIO->readCell(18, 16)) * 100;
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(18, 16)), koef_okrugl)) + " - " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(19, 21)), koef_okrugl)) + " / " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(18, 16)), koef_okrugl)) + " * 100% = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form64");
		sum = (str_to_d(excelIO->readCell(18, 16)) - str_to_d(excelIO->readCell(19, 20))) * str_to_d(excelIO->readCell(tableHeight + 3, 1)) / str_to_d(excelIO->readCell(tableHeight + 7, 2)) * 100;
		temp = "(" + cstr_to_str(round_my(str_to_d(excelIO->readCell(18, 16)), koef_okrugl)) + " - " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(19, 20)), koef_okrugl)) + ") * " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight + 3, 1)), koef_okrugl)) + " / "
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight + 7, 2)), koef_okrugl)) + " * 100% = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form65");
		sum = (str_to_d(excelIO->readCell(18, 16)) - str_to_d(excelIO->readCell(19, 20))) * str_to_d(excelIO->readCell(tableHeight + 6, 2)) 
			/ str_to_d(excelIO->readCell(tableHeight + 2, 1)) * 100;
		temp = "(" + cstr_to_str(round_my(str_to_d(excelIO->readCell(18, 16)), koef_okrugl)) + " - " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(19, 20)), koef_okrugl)) + ") * " 
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight + 6, 2)), koef_okrugl)) + " / "
			+ cstr_to_str(round_my(str_to_d(excelIO->readCell(tableHeight + 2, 1)), koef_okrugl)) + " * 100% = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form66");
		sum = str_to_d(excelIO->readCell(19, 20)) / str_to_d(excelIO->readCell(18, 16));
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(19, 20)), koef_okrugl)) + " / " + cstr_to_str(round_my(str_to_d(excelIO->readCell(18, 16)), koef_okrugl)) + " = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form67");
		sum = str_to_d(excelIO->readCell(19, 21)) / str_to_d(excelIO->readCell(18, 16));
		temp = cstr_to_str(round_my(str_to_d(excelIO->readCell(19, 21)), koef_okrugl)) + " / " + cstr_to_str(round_my(str_to_d(excelIO->readCell(18, 16)), koef_okrugl)) + " = " + cstr_to_str(round_my(sum, koef_okrugl));
		wordIO->write(temp, L"Form68");
		OnOK();
		}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 30: //кол-во оборудования
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789") == std::string::npos)
			{
				counter = std::stoi(temp) * 4; //перевод из string в int, умножаем на 4 тк нужно для каждого сырья вводить 4 параметра
				if (counter > 0)
				{
					i = 1;
					tableHeight += 32; //перешли вниз к той таблице
					result = tableHeight; // для запоминания начала таблицы

					excelIO->writeCell(tableHeight, 1, L"Наименование", true);
					excelIO->writeCell(tableHeight, 2, L"Маркировка", true);
					excelIO->writeCell(tableHeight, 3, L"Колличество", true);
					excelIO->writeCell(tableHeight++, 4, L"Цена за единицу", true);

					page++;
					label = L"Введите название 1-го оборудования";
					users_answer = L"";
					UpdateData(FALSE);
				}
				else MessageBox(L"Колличетсво должно быть больше 0", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;

	case 31://обработка оборудования
		if (i <= counter)
		{
			if (i % 4 == 1) //вводим название
			{
				if (temp != "" && temp != " ")
				{
					//записать в excel 
					excelIO->writeCell(tableHeight, 1, users_answer);
					//переход к следующей просьбе
					label.Format(L"Введите маркировку %d-го оборудования", i / 4 + 1);
					users_answer = L"";
					UpdateData(FALSE);
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else if (i % 4 == 2) //вводим название
			{
				if (temp != "" && temp != " ")
				{
					//записать в excel 
					excelIO->writeCell(tableHeight, 2, users_answer);
					//переход к следующей просьбе
					label.Format(L"Введите колличество %d-го оборудования", i / 4 + 1);
					users_answer = L"";
					UpdateData(FALSE);
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else if (i % 4 == 3) //вводим колличество
			{
				if (temp != "" && temp != " ")
				{
					if (temp.find_first_not_of("0123456789..") == std::string::npos)
					{
						excelIO->writeCell(tableHeight, 3, users_answer);
						label.Format(L"Введите цену за единицу для %d-го оборудования", i / 4 + 1);
						users_answer = L"";
						UpdateData(FALSE);
						if (counter - i == 1) page++; //это чтобы пользователь два раза не жал на кнопку дальше, после ввода нормы расхода последнего сырья
					}
					else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			else //вводим цену за единицу
			{
				if (temp != "" && temp != " ")
				{
					if (temp.find_first_not_of("0123456789..") == std::string::npos)
					{
						myRange = excelIO->getSheet()->Rows->Item[tableHeight];
						myRange->RowHeight = 19.25;
						excelIO->writeCell(tableHeight, 4, users_answer);

						sum_glob += str_to_d(excelIO->readCell(tableHeight, 3)) * str_to_d(excelIO->readCell(tableHeight, 4));//считаем общие кап вложения
						tableHeight++;
						if (counter - i != 0) {
							label.Format(L"Введите название %d-го вида оборудования", i / 4 + 1);
							users_answer = L"";
							UpdateData(FALSE);
						}
					}
					else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
				}
				else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
			}
			i++;
		}
		break;

	case 32://обработка ввода цены последнего оборудования
		if (temp != "" && temp != " ")
		{
			if (temp.find_first_not_of("0123456789..") == std::string::npos)
			{
				//записать в excel, тут дубликат того что было в последнем блоке "запись в excel"
				myRange = excelIO->getSheet()->Rows->Item[tableHeight];
				myRange->RowHeight = 19.25;
				excelIO->writeCell(tableHeight, 4, users_answer);
				sum_glob += str_to_d(excelIO->readCell(tableHeight, 3)) * str_to_d(excelIO->readCell(tableHeight, 4));//считаем общие кап вложения
				myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[result][1]][excelIO->getSheet()->Cells->Item[tableHeight][3]];
				myRange->Borders->Weight = Excel::xlThin;

				myRange = excelIO->getSheet()->Range[excelIO->getSheet()->Cells->Item[result][1]][excelIO->getSheet()->Cells->Item[tableHeight][3]];
				myRange->Copy();
				tableRange = wordIO->file->Bookmarks->Item(&_variant_t(L"Tabl1_maybe"))->Range;
				tableRange->PasteExcelTable(FALSE, TRUE, FALSE);

				//возвращаемся обратно
				tableHeight -= 34;
				while (excelIO->readCell(tableHeight, 1, true) != L"12. Полная себестоимость")
				{ tableHeight--; }
				tableHeight += 2;
				excelIO->writeCell(tableHeight++, 1, sum_glob);
				wordIO->write(d_to_str(sum_glob), L"b");
				wordIO->write(d_to_str(sum_glob), L"b2");

				label = L"Введите объем производства";
				users_answer = L"";
				UpdateData(FALSE);
				page = 19;
			}
			else MessageBox(L"Вводите только цифры", L"Ошибка", MB_OK | MB_ICONERROR);
		}
		else MessageBox(L"Введи хоть что-нибудь", L"Ошибка", MB_OK | MB_ICONERROR);
		break;
	}
	editText.SetFocus();
}

void MainDlg::OnBnClickedBack()
{
	switch (page)
	{
	case 1:
		//wordIO->write(temp, L"k2"); //при нажатии назад надо бы делать так , чтобы все что было вставлено в закладку удалялось
		page--;
		label = L"Введите вариант";
		users_answer = L"";
		UpdateData(FALSE);
		break;
	case 2:
		//wordIO->write(temp, L"k2"); //при нажатии назад надо бы делать так , чтобы все что было вставлено в закладку удалялось
		page--;
		label = L"Введите название продукта";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 3:
		if (i == 1)
		{
			page--;
			label = L"Введите количество сырья и основных материалов";
		}
		else if (i % 3 == 1) //вводим название
		{
			i--;
			label.Format(L"Введите норму расхода %d-го сырья", i / 3 + 1);
		}
		else if (i % 3 == 2) //вводим цену
		{
			i--;
			label.Format(L"Введите название %d-го сырья", i / 3 + 1);
		}
		else //вводим норму расхода
		{
			i--;
			label.Format(L"Введите цену %d-го сырья", i / 3 + 1);
		}
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 4:
		i--;
		page--;
		label.Format(L"Введите цену %d-го сырья", i / 3 + 1);
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 5: //обработка, кол-во вспомогат материалов
		page--;
		label = L"Введите кол-во вспомогательных материалов";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 6:
		if (i == 1)
		{
			page--;
			label = L"Введите количество сырья и основных материалов";
		}
		else if (i % 3 == 1) //вводим название
		{
			i--;
			label.Format(L"Введите норму расхода %d-го материала", i / 3 + 1);
		}
		else if (i % 3 == 2) //вводим цену
		{
			i--;
			label.Format(L"Введите название %d-го материала", i / 3 + 1);
		}
		else //вводим норму расхода
		{
			i--;
			label.Format(L"Введите цену %d-го материала", i / 3 + 1);
		}
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 7:
		i--;
		page--;
		label.Format(L"Введите цену %d-го материала", i / 3 + 1);
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 8://обработка, кол-во видов энергозатрат
		page--;
		label.Format(L"Введите норму расхода %d-го материала", i / 3 + 1);
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 9:
		if (i == 1)
		{
			page--;
			label = L"Введите кол-во видов энергозатрат";
		}
		else if (i % 3 == 1) //вводим название
		{
			i--;
			label.Format(L"Введите норму расхода %d-го вида энергозатрат", i / 3 + 1);
		}
		else if (i % 3 == 2) //вводим цену
		{
			i--;
			label.Format(L"Введите название %d-го вида энергозатрат", i / 3 + 1);
		}
		else //вводим норму расхода
		{
			i--;
			label.Format(L"Введите цену %d-го вида энергозатрат", i / 3 + 1);
		}
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 10:
		i--;
		page--;
		label.Format(L"Введите цену %d-го вида энергозатрат", i / 3 + 1);
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 11:
		page--;
		label.Format(L"Введите норму расхода %d-го вида энергозатрат", i / 3 + 1);
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 12://обработка суммы фонда оплаты труда
		page--;
		label = L"Введите сумму в фонд оплаты труда";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 13: // назад к обработке процента от ФОП
		page--;
		label = L"Введите процент от ФОП отчислений в обязательные страховые фонды";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 14://назад к обработке расходов на эксплуатацию
		page--;
		label = L"Введите расходы на содержание и эксплуатацию оборудования";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 15:
		page--;
		label = L"Введите расходы на аммортизацию оборудования";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 16:
		page--;
		label = L"Введите цеховые расходы";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 17:
		page--;
		label = L"Введите общехозяйственные расходы";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 18:
		page--;
		label = L"Введите внепроизводственные расходы";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 19:
		if (result == IDYES)
		{
			page--;
			label = L"Введите капитальные вложения в основное технологическое оборудование";
			users_answer = L"";
			UpdateData(FALSE);
			break;
		}
		else if (result == IDNO)
		{
			i--;
			page = 32;
			label.Format(L"Введите цену за единицу для %d-го оборудования", i / 4 + 1);
			users_answer = L"";
			UpdateData(FALSE);
			break;
		}
		break;

	case 20:
		page--;
		label = L"Введите объем производства";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 21:
		page--;
		label = L"Введите процент увеличения объема производства на";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 22:
		page--;
		label = L"Введите процент сокращения норм расхода по основному виду исходного сырья на ";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 23:
		page--;
		label = L"Введите процент сокращения норм расхода по энергетическим ресурсам на  ";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 24:
		page--;
		label = L"Введите увеличение численности производственных рабочих на";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 25:
		page--;
		label = L"Введите оклад рабочих";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 26:
		page--;
		label = L"Введите норму годовых амортизационных отчислений";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 27:
		page--;
		label = L"Введите изменение цены ресурсов (в процентах)";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 28:
		page--;
		label = L"Введите коэффициент выручки";
		users_answer = L"";
		UpdateData(FALSE);
		break; //от ставки дисконтирования назад уже не вернутся, значит case29 нет

	case 30: //кол-во оборудования
		page = 17;
		label = L"Введите внепроизводственные расходы";
		users_answer = L"";
		UpdateData(FALSE);
		break;

	case 31://обработка оборудования
		if (i == 1)
		{
			page--;
			label = L"Введите кол-во оборудования";
		}
		else if (i % 4 == 1) //вводим название
		{
			i--;
			label.Format(L"Введите цену за единицу для %d-го оборудования", i / 4 + 1);
		}
		else if (i % 4 == 2) //вводим название
		{
			i--;
			label.Format(L"Введите название %d-го вида оборудования", i / 4 + 1);
		}
		else if (i % 4 == 3) //вводим колличество 
		{
			i--;
			label.Format(L"Введите маркировку %d-го оборудования", i / 4 + 1);
		}
		else
		{
			i--;
			label.Format(L"Введите колличество %d-го оборудования", i / 4 + 1);
		}
		users_answer = L"";
		UpdateData(FALSE);
		break;


	case 32://обработка ввода цены последнего оборудования
		i--;
		page--;
		label.Format(L"Введите колличество %d-го оборудования", i / 4 + 1);
		users_answer = L"";
		UpdateData(FALSE);
		break;
	}
}
