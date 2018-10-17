# 需求背景

开发中经常会碰到导出excel的需求，通常导出的格式大部分都是“表头+循环数据+表尾”的形式，其中可能还包含字体大小颜色，边框，背景颜色，合并单元格等场景。

# 分析

初级的开发者可能会想到用poi根据实际需求去实现，然后一个格子一个格子的去设置格式和填数据再导出，这种方式只适合应付单个导出需求，当多个导出需求时弊端明显。<br>
可能有人会使用alibaba的EasyExcel来处理导出，他的思路是使用对象注解的形势去设置每个单元格的格式，但是我觉得这种方式也不够简便<br>

# 实现思路

类似于写页面一样，先定好模板，然后填数。<br>
1、本工程实现的思路是，使用excel先画好一个模板，模板中写入一些变量和个性化标签。<br>
2、首次导出的时候，系统会读取模板内容（包括样式，数据等），并将模板信息以结构化的形式缓存在内存中。<br>
3、获取业务数据，按模板的格式填充数据自动进行导出。<br>

# 核心代码

参考 com.staconfree.poi.write.sxssf.TestParser
```
String templateFilePath = "D://template.xlsx";
String exportFilePath = "D://export.xlsx";
SxssfExcelParser excelParse = new SxssfExcelParser(templateFilePath, exportFilePath);
Map<String,Object> dataStack = getBizData();// 此处获取业务代码
 excelParse.setTemplateSheetName("Sheet1");
 excelParse.parse(dataStack);
 excelParse.endSheet();
 excelParse.setTemplateSheetName("Sheet2");
 excelParse.parse(dataStack);
 excelParse.endSheet();
 excelParse.flush();
```
# 演示效果

运行源码中的TestParser，**（ps：第一次导出由于要读模板和反射数据对象，所以比较慢，之后这些数据都会缓存就不会慢了）**<br>
会在 roy-poi/target/test-classes/ 下看到两个文件 template.xlsx 和 test_parser.xlsx 。<br>
- 一个是模板文件，如下：<br>
![sheet1](https://github.com/staconfree/roy-poi/raw/master/readme_pic/template-sheet1.png)
![sheet2](https://github.com/staconfree/roy-poi/raw/master/readme_pic/template-sheet2.png)
```
关键的模板格式解释：
一、
$$BEGINLOOP{"name":"brandList","mergedbaseon":[[1,2,3],[4]]}
代表循环brandList对象，后面 mergedbaseon 的含义是，如果相邻两行的1,2,3列数据完全相同，则这三列相邻的数据自动合并单元格，同理相邻两行的第4列如果相同，则第四列相邻相同的数据也自动合并单元格
二、
${#index+1} 代表循环体下标+1
三、
$$yyyy-$$MM-$$dd $$HH:$$mm:$$ss 自动打印系统时间
```
- 另一个是导出文件，如下：<br>
![export1](https://github.com/staconfree/roy-poi/raw/master/readme_pic/export-sheet1.png)
![export2](https://github.com/staconfree/roy-poi/raw/master/readme_pic/export-sheet2.png)
<br>从效果中可以看到，导出的格式跟模板格式一模一样，细看导出代码，可以发现预留了WriteCallBack接口，可以对导出的每个单元格做个性化的设置，如上面的合计那行做了横向的合并单元格

# 如何使用

