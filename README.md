### poi读取和写入office使用方法

---

> POI介绍

        Apache POI是Apache软件基金会的开源项目，POI提供API给Java程序员对Microsoft Office格式文档度和写的功能。.NET的开发人员可以利用NPOI（POI for .NET）来存取Microsoft Office文档的功能。

> POI结构说明

     HSSF提供读写Microsoft `Execl` `XLS` 格式文档的功能
    
     XSSF提供读写Microsoft `Execl OOXML` `XLSX` 格式文档的功能
    
     HWPF提供读写Microsoft `Word` `DOC` 格式档案的功能
    
     XWPF提供读写Microsoft `Word` `DOCX` 格式文档的功能
    
     HSLF提供读写Microsoft `PowerPoint` 格式文档的功能
    
     HDGF提供读写Microsoft `Visio` 格式文档的功能
    
     HPBF提供读写Microsoft `Publisher` 格式文档的功能
    
     HSMF提供读写Microsoft `Outlook` 格式文档的功能

> POI常用类说明

| 类名             | 说明              |
|:--------------:|:---------------:|
| Workbook       | Excel的文档对象      |
| Sheet          | Excel的Sheet     |
| Row            | Excel的行         |
| Cell           | Excel的单元格       |
| Font           | Excel的字体        |
| DataFormat     | Excel单元格的日期格式   |
| Header         | Excel文档Sheet的页眉 |
| Footer         | Excel文档Sheet的页脚 |
| CellStyle      | Excel单元格样式      |
| DateUtil       | 日期              |
| PrintSetup     | 打印              |
| ErrorConstants | 错误信息表           |
