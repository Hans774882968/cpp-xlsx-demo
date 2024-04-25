[TOC]

# C++操作xlsx初体验：libxl、OpenXLSX

## 引言

C++操作xlsx着实麻烦，所以为此精心产出~~水~~一篇文章也是很OK的吧。

[本文GitHub传送门](https://github.com/Hans774882968/cpp-xlsx-demo)。**作者：[hans774882968](https://blog.csdn.net/hans774882968)以及[hans774882968](https://juejin.cn/user/1464964842528888)以及[hans774882968](https://www.52pojie.cn/home.php?mod=space&uid=1906177)**

本文打算体验的库：

- [libxl](https://www.libxl.com/)。要氪金，否则使用的功能会受限。这里我就不氪了。
- [OpenXLSX](https://github.com/troldal/OpenXLSX)。

## 环境

1. Windows10 64位
2. `g++.exe (x86_64-win32-seh-rev1, Built by MinGW-Builds project) 13.1.0`
3. VSCode+CMake插件+CMake Tools插件。本项目是用CMake Tools插件创建的，过程不赘述。

### 如何下载高版本g++

咕果一下，找到一个[现成的](https://whitgit.whitworth.edu/tutorials/installing_mingw_64)。

### VSCode修改CMake插件使用的g++版本

左侧导航栏点击CMake的图标，左上角的`PROJECT STATUS`处有一个树形组件，其中有`Configure`节点，节点下有一项`GCC 13.1.0 x86_64-w64-mingw32`，hover会出现铅笔，点击即可修改。若没有看到你想选的编译器，可以先选择`[Scan for kits]`让插件先帮你找下。

## libxl

libxl官方文档并没有给出VSCode+CMake Tools如何配置，但可以参考[eclipsecpp的配置文档](https://www.libxl.com/eclipsecpp.html)。添加的配置如下：

```cmake
# 路径必须用左斜杠分隔
include_directories(D:/libxl-4.3.0.14/include_cpp)
target_link_libraries(libxl_read_xlsx_demo D:/libxl-4.3.0.14/lib64/libxl.lib)
target_link_libraries(libxl_write_xlsx_demo D:/libxl-4.3.0.14/lib64/libxl.lib)
```

另外需要注意，一定记得把`bin64/libxl.dll`复制到exe文件所在的文件夹下，否则能编译成功，但运行时会静默失败。

示例代码看官方文档就会写了，这里就不贴出来了，只列出一些注意点。

一、`Book* book = xlCreateXMLBook();`用于读写07及以后版本的xlsx。

二、cmake生成的二进制文件在`build`文件夹下，而我的示例xlsx在根目录下，所以考虑先`getcwd`再拼接出所求路径：

```cpp
#include <bits/stdc++.h>
#include <direct.h>
// 省略其他头文件
using namespace std::filesystem;

string get_inp_xlsx_path() {
  char* cwd_buffer = getcwd(NULL, 0);
  path p = canonical(path(cwd_buffer) / ".." / "inp" / "example.xlsx");
  return p.string();
}
```

三、调用`canonical`前需要保证路径的文件/文件夹存在。为此，我们引入了`utils.h utils.cpp`，并修改`cmake`：

```cmake
add_executable(libxl_write_xlsx_demo libxl_write_xlsx_demo.cpp utils.cpp utils.h)
```

## OpenXLSX

TODO