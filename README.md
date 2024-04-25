[TOC]

# C++操作xlsx初体验：OpenXLSX（建议使用）、libxl

## 引言

C++操作xlsx着实麻烦，所以为此精心创作~~水~~一篇文章也是很OK的吧。

[本文GitHub传送门](https://github.com/Hans774882968/cpp-xlsx-demo)。**作者：[hans774882968](https://blog.csdn.net/hans774882968)以及[hans774882968](https://juejin.cn/user/1464964842528888)以及[hans774882968](https://www.52pojie.cn/home.php?mod=space&uid=1906177)**

本文52pojie：https://www.52pojie.cn/thread-1917834-1-1.html

本文juejin：https://juejin.cn/post/7361687968518635554

本文CSDN：https://blog.csdn.net/hans774882968/article/details/138202684

本文打算体验的库：

- [libxl](https://www.libxl.com/)。要氪金，否则使用的功能会受限。这里我就不氪了。
- [OpenXLSX](https://github.com/troldal/OpenXLSX)。

体验结果：建议使用`OpenXLSX`。

## 环境

1. Windows10 64位
2. `g++.exe (x86_64-win32-seh-rev1, Built by MinGW-Builds project) 13.1.0`
3. VSCode+CMake插件+CMake Tools插件。本项目是用CMake Tools插件创建的，过程不赘述。

### 如何下载高版本g++

咕果一下，找到一个[现成的](https://whitgit.whitworth.edu/tutorials/installing_mingw_64)。

### VSCode修改CMake Tools插件使用的g++版本

左侧导航栏点击CMake的图标，左上角的`PROJECT STATUS`处有一个树形组件，其中有`Configure`节点，节点下有一项`GCC 13.1.0 x86_64-w64-mingw32`，hover会出现铅笔，点击即可修改。若没有看到你想选的编译器，可以先选择`[Scan for kits]`让插件先帮你找下。

### 执行哪个二进制文件

同上找到树形组件，其中有`Debug`和`Launch`节点，hover会出现铅笔，点击即可修改你想执行的二进制文件。

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

参考[其GitHub](https://github.com/troldal/OpenXLSX)的README，我选用了最简单的方式（和[参考链接1](https://blog.csdn.net/weixin_44569865/article/details/131623832)一样的方式）：先复制源码到项目根目录，然后配置cmake：

```cmake
add_executable(opxl_read_xlsx_demo opxl_read_xlsx_demo.cpp)
add_executable(opxl_write_xlsx_demo opxl_write_xlsx_demo.cpp utils.cpp utils.h)

add_subdirectory(OpenXLSX)

target_link_libraries(opxl_read_xlsx_demo OpenXLSX::OpenXLSX)
target_link_libraries(opxl_write_xlsx_demo OpenXLSX::OpenXLSX)
```

最后导入：

```cpp
#include <OpenXLSX.hpp>
using namespace OpenXLSX;
```

于是就遇到了报错：`uint32_t`等未定义，需要导入`cstdint`。我搜索报错信息，只找到了[这个issue](https://github.com/troldal/OpenXLSX/issues/242)。于是我就按照报错提示，给若干`.hpp`手动添加了`#include <cstdint>`。幸运的是，添加后就能编译成功了。

示例代码看[官方Demo](https://github.com/troldal/OpenXLSX/blob/master/Examples)就会写了，这里就不贴出来了，只列出一些注意点。

一、官方Demo似乎没有给读取Excel的API，这里说一下：

```cpp
void read_example_xlsx() {
  string xlsx_path = get_inp_xlsx_path();
  XLDocument doc;
  doc.open(xlsx_path);
  auto wks = doc.workbook().worksheet("Sheet1");
  read_data(wks);
}
```

二、和libxl不同，OpenXLSX写入的公式不会在打开Excel时自己计算好结果。libxl官方说这不是bug，是feature：

> Nota that OpenXLSX does not calculate the results of a formula. If you add a formula to a spreadsheet using OpenXLSX, you have to open the spreadsheet in the Excel application in order to see the results of the calculation.

三、同理，日期写入后打开Excel文档，也是一个浮点数，需要手动调格式才能看到日期。

## 参考资料

1. 配置OpenXLSX：https://blog.csdn.net/weixin_44569865/article/details/131623832