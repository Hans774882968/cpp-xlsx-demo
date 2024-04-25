#include <bits/stdc++.h>
#include <direct.h>
#include "libxl.h"
using namespace libxl;
using namespace std::filesystem;
using string = std::string;
// Copyright 2024 hans774882968

char* cwd_str = getcwd(NULL, 0);

void read_data(Sheet* sheet) {
  int ans = 0;
  for (int row = sheet->firstRow(); row < sheet->lastRow(); ++row) {
    for (int col = sheet->firstCol(); col < sheet->lastCol(); ++col) {
      CellType cellType = sheet->cellType(row, col);
      if (cellType == CELLTYPE_NUMBER) {
        double d = sheet->readNum(row, col);
        ans += d;
      }
      // 实测 libxl 把A1单元格改了，要求你氪金
      if (cellType == CELLTYPE_STRING) {
        std::cout << row << " " << col << " " << sheet->readStr(row, col)
                  << std::endl;
      }
    }
  }
  std::cout << "ans = " << ans << std::endl;
}

string get_inp_xlsx_path() {
  path p = canonical(path(cwd_str) / ".." / "inp" / "example.xlsx");
  return p.string();
}

void read_example_xlsx() {
  string xlsx_path_str = get_inp_xlsx_path();
  const char* xlsx_path = xlsx_path_str.c_str();

  Book* book = xlCreateXMLBook();
  if (book->load(xlsx_path)) {
    Sheet* sheet = book->getSheet(0);
    if (sheet) read_data(sheet);
  }
  book->release();
}

int main(int, char**) {
  read_example_xlsx();
  return 0;
}
