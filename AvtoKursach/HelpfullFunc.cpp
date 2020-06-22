#include "stdafx.h"
#include "HelpfullFunc.h"

bool hasSuffix(const CString& Cname, const std::string &suffix) //����������, ������������ �� ���� �� suffix
{
	CT2CA tmp(Cname);
	std::string name(tmp);
	return name.size() >= suffix.size() && name.compare(name.size() - suffix.size(), suffix.size(), suffix) == 0;
}

std::string cstr_to_str(const CString& cstr)
{
	CT2CA tmp(cstr);
	std::string str(tmp);
	return str;
}

CString str_to_cstr(const std::string& str)
{
	CString cstr;
	cstr = str.c_str();
	return cstr;
}

bstr_t str_to_bstr_t(std::string stroka)
{
	CString a = str_to_cstr(stroka);
	bstr_t b = a;
	return b;
}

std::string d_to_str(double a) 
{
	std::string c = std::to_string(a);
	return c;
}

CString d_to_cstr(double a)
{
	std::string c = std::to_string(a);
	CString b = str_to_cstr(c);
	return b;
}

double cstr_to_d(CString a)
{
	wchar_t * stopString;
	double b = wcstod(a, &stopString);
	return b;
}

double str_to_d(std::string a) //��� �������� ����� � �������
{
	double b;
	int flag = a.find_first_of(",,");
	if (flag != std::string::npos)
	{
		b = std::stod(a);
		std::string a2 = a.substr(flag + 1, a.size() - flag + 1); //���� ���������� ���� � n1 �� n2 ������������ new = old.substr(n1, n2 - n1 + 1);
		b += std::stod(a2) / pow(10, a2.size());
		return b;
	}
	return std::stod(a);
}

CString round_my(double a, int numb)
{
	//����� ������ ���������� �� numb ������
	//�������� ����� �� 10 � ������� ������� ������ ����� ���������
	//��������� 0.5 
	//�������� � �������������
	//��������� �� 10 � ������� ������� ������ ����� ���������
	int b = a * pow(10, numb) + 0.5;
	std::string str = d_to_str(b / pow(10, numb));

	int flag = str.find_first_of("..");
	int k = numb;
	while (str[flag + k] == '0' && k > 0) k--; 
	if (k != 0)
	{
		str = str.substr(0, flag + k + 1);
	}
	else {
		str = str.substr(0, flag - 1 + 1);
	}
	return str_to_cstr(str);
}