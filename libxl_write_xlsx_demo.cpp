#include <bits/stdc++.h>
#include <direct.h>
#include "libxl.h"
#include "utils.h"
using namespace libxl;
using namespace std::filesystem;
using string = std::string;
// Copyright 2024 hans774882968

void write_hw_xlsx() {
  Book* book = xlCreateXMLBook();
  Sheet* sheet = book->addSheet("Sheet1");
  if (sheet) {
    // 没氪金的话，建议从第二行开始写
    sheet->writeStr(1, 0, "Hello libxl");
    sheet->writeNum(1, 1, 114000);
    sheet->writeNum(1, 2, 514);
    sheet->writeBool(1, 3, true);
    sheet->writeBool(1, 4, false);
    sheet->writeFormula(1, 5, "SUM(B2:C2)");  // 114514
  }
  string out_xlsx_path_str = get_out_xlsx_path("libxl_hw.xlsx");
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
  mk_out_dir();
  write_hw_xlsx();
  return 0;
}
