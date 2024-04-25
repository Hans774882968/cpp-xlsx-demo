#include "utils.h"
// Copyright 2024 hans774882968

int check_or_mkdir(string path) {
  const char* path_c_str = path.c_str();
  if (access(path_c_str, 0) == -1) {
    return mkdir(path_c_str);
  }
  return 0;
}

void mk_out_dir() {
  path p = path(cwd_str) / ".." / out_dir_name;
  check_or_mkdir(p.string());
}

string get_out_xlsx_path(string filename) {
  path p = path(cwd_str) / ".." / out_dir_name / filename;
  return p.string();
}
