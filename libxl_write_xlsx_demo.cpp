#include <bits/stdc++.h>
#include <direct.h>
#include "libxl.h"
#include "utils.h"
using namespace libxl;
using namespace std::filesystem;
using string = std::string;
// Copyright 2024 hans774882968

const char out_dir_name[] = "outs";
char* cwd_str = getcwd(NULL, 0);

string get_out_dir_path() {
  path p = path(cwd_str) / ".." / out_dir_name;
  check_or_mkdir(p.string());  // 调用 canonical 前必须先创建文件夹
  path res = canonical(p);
  return res.string();
}

const char* out_dir_path = get_out_dir_path().c_str();

string get_out_xlsx_path() {
  path p = path(cwd_str) / ".." / out_dir_name / "libxl_hw.xlsx";
  return p.string();
}

void write_hw_xlsx() {
  Book* book = xlCreateXMLBook();
  Sheet* sheet = book->addSheet("Sheet1");
  if (sheet) {
    // 没氪金的话，建议从第二行开始写
    sheet->writeStr(1, 0, "Hello libxl");
    sheet->writeNum(1, 1, 114514);
    sheet->writeBool(1, 2, true);
    sheet->writeBool(1, 3, false);
  }
  string out_xlsx_path_str = get_out_xlsx_path();
  const char* out_xlsx_path = out_xlsx_path_str.c_str();
  if (!book->save(out_xlsx_path)) {
    std::cout << book->errorMessage() << std::endl;
  } else {
    string out_xlsx_path_canonical = canonical(out_xlsx_path_str).string();
    std::cout << "write " << out_xlsx_path_canonical << " success" << std::endl;
  }
  book->release();
}

int main(int, char**) {
  write_hw_xlsx();
  return 0;
}
