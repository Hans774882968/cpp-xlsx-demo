#include <bits/stdc++.h>
#include <direct.h>
#include <OpenXLSX.hpp>
using namespace std::filesystem;
using string = std::string;
using namespace OpenXLSX;
// Copyright 2024 hans774882968

char* cwd_str = getcwd(NULL, 0);

void read_data(XLWorksheet wks) {
  int ans = 0;
  int rc = wks.rowCount(), cc = wks.columnCount();
  std::cout << rc << " " << cc << std::endl;  // 6 5
  for (int row = 1; row <= rc; ++row) {
    for (int col = 1; col <= cc; ++col) {
      XLCellValue cell_value = wks.cell(row, col).value();
      if (cell_value.type() == XLValueType::Integer) {
        ans += cell_value.get<int>();
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
  string xlsx_path = get_inp_xlsx_path();
  XLDocument doc;
  doc.open(xlsx_path);
  auto wks = doc.workbook().worksheet("Sheet1");
  read_data(wks);
}

int main(int, char**) {
  read_example_xlsx();
  return 0;
}
