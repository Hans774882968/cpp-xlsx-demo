#include <bits/stdc++.h>
#include <direct.h>
using namespace std::filesystem;
using string = std::string;
// Copyright 2024 hans774882968

#ifndef D__HELLOWORLD_CPP_407_CPP_XLSX_DEMO_UTILS_H_
#define D__HELLOWORLD_CPP_407_CPP_XLSX_DEMO_UTILS_H_

static const char out_dir_name[] = "outs";
static char* cwd_str = getcwd(NULL, 0);

int check_or_mkdir(string path);

void mk_out_dir();

string get_out_xlsx_path(string filename);

#endif  // D__HELLOWORLD_CPP_407_CPP_XLSX_DEMO_UTILS_H_
