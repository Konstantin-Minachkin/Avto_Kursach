#pragma once
#include <string>
#include <sstream>
#include <cmath>

bool hasSuffix(const CString& Cname, const std::string &suffix); //определяет, оканчивается ли файл на suffix
std::string cstr_to_str(const CString& cstr);
CString str_to_cstr(const std::string& str);
bstr_t str_to_bstr_t(std::string stroka);
std::string d_to_str(double a);
double str_to_d(std::string a);
double cstr_to_d(CString a);
CString d_to_cstr(double a);
CString round_my(double a, int numb);