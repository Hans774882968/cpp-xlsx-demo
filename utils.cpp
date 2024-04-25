#include "utils.h"
// Copyright 2024 hans774882968

int check_or_mkdir(string path) {
  const char* path_c_str = path.c_str();
  if (access(path_c_str, 0) == -1) {
    return mkdir(path_c_str);
  }
  return 0;
}
