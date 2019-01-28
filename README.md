### poi读取和写入office使用方法

---

> POI介绍

        Apache POI是Apache软件基金会的开源项目，POI提供API给Java程序员对Microsoft Office格式文档度和写的功能。.NET的开发人员可以利用NPOI（POI for .NET）来存取Microsoft Office文档的功能。

> POI常用类说明

+ HSSF提供读写Microsoft `Execl` `XLS` 格式文档的功能
+ XSSF提供读写Microsoft `Execl OOXML` `XLSX` 格式文档的功能
+ HWPF提供读写Microsoft `Word` `DOC` 格式档案的功能
+ XWPF提供读写Microsoft `Word` `DOCX` 格式文档的功能
+ HSLF提供读写Microsoft `PowerPoint` 格式文档的功能
+ HDGF提供读写Microsoft `Visio` 格式文档的功能
+ HPBF提供读写Microsoft `Publisher` 格式文档的功能
+ HSMF提供读写Microsoft `Outlook` 格式文档的功能

> POI结构说明

| 类名                 | 说明              |
|:------------------:|:---------------:|
| HSSFWorkbook       | Excel的文档对象      |
| HSSFSheet          | Excel的Sheet     |
| HSSFRow            | Excel的行         |
| HSSFCell           | Excel的单元格       |
| HSSFFont           | Excel的字体        |
| HSSFDataFormat     | Excel单元格的日期格式   |
| HSSFHeader         | Excel文档Sheet的页眉 |
| HSSFFooter         | Excel文档Sheet的页脚 |
| HSSFCellStyle      | Excel单元格样式      |
| HSSFDateUtil       | 日期              |
| HSSFPrintSetup     | 打印              |
| HSSFErrorConstants | 错误信息表           |
