#include <bits/stdc++.h>
#include <direct.h>
#include <OpenXLSX.hpp>
#include "utils.h"
using namespace std::filesystem;
using string = std::string;
using namespace OpenXLSX;
// Copyright 2024 hans774882968

void write_date(XLWorksheet wks) {
  std::tm tm;
  tm.tm_year = 124;
  tm.tm_mon = 3;
  tm.tm_mday = 25;
  tm.tm_hour = 22;
  tm.tm_min = 42;
  tm.tm_sec = 6;
  XLDateTime dt(tm);
  wks.cell("H1").value() = dt;
}

void write_hw_xlsx() {
  string out_xlsx_path = get_out_xlsx_path("opxl_hw.xlsx");

  XLDocument doc;
  doc.create(out_xlsx_path);
  auto wks = doc.workbook().worksheet("Sheet1");
  wks.cell("A1").value() = "Hello, OpenXLSX!";
  wks.cell("B1").value() = 114510;
  wks.cell("C1").value() = 4;
  wks.cell("D1").value() = true;
  wks.cell("E1").value() = false;
  wks.cell("F1").formula() = "SUM(B1:C1)";
  wks.cell("G1").formula() = "EXP(C1)";
  write_date(wks);
  XLCellValue b1 = wks.cell("B1").value();
  std::cout << "Cell B1(" << b1.typeAsString() << "): " << b1.get<int>()
            << std::endl;
  XLFormula f1 = wks.cell("F1").formula();
  std::cout << "Formula in F1: " << f1.get() << std::endl;
  XLFormula g1 = wks.cell("G1").formula();
  string g1_s = wks.cell("G1").formula();
  std::cout << "Formula in G1: " << g1.get() << " " << g1_s << std::endl;
  auto h1_dt = wks.cell("H1").value().get<XLDateTime>();
  auto h1_dt_serial = h1_dt.serial();
  auto h1_tmo = h1_dt.tm();
  char h1_time_str[32];
  asctime_s(h1_time_str, 32, &h1_tmo);
  // H1: 45407.9 Thu Apr 25 22:42:06 2024
  std::cout << "H1: " << h1_dt_serial << " " << h1_time_str << std::endl;

  doc.save();
  doc.close();

  string out_xlsx_path_canonical = canonical(out_xlsx_path).string();
  std::cout << "write " << out_xlsx_path_canonical << " success" << std::endl;
}

int main(int, char**) {
  mk_out_dir();
  write_hw_xlsx();
  return 0;
}
