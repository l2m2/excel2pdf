# excel2pdf
用Office.Excel将Excel转为PDF (C#实现).

这个工程生成的是一个命令行工具。转换示例：
```
excel2pdf.exe -i D:/1.xlsx -o D:/1.pdf
```

## 更新适配

- 升级为多平台工程, net7.0 和 net48
- 不是必须装 Office2013+ 才能使用, 只装 WPS 也能调用这个工具


## 常见问题

- 若遇到下面的错误，需要删除一个注册表值
    ```
    未经处理的异常:  System.InvalidCastException: 无法将类型为“System.__ComObject”的 COM 对象强制转换
    为接口类型“Microsoft.Office.Interop.Excel.Application”。此操作失败的原因是对 IID 为“{000208D5-000
    0-0000-C000-000000000046}”的接口的 COM 组件调用 QueryInterface 因以下错误而失败: 不支持此接口 (异常
    来自 HRESULT:0x80004002 (E_NOINTERFACE))。
    在 excel2pdf.Program.ExportWorkbookToPdf(String workbookPath, String outputPath)
    在 excel2pdf.Program.Main(String[] args)
    ```
    解决方案：
    ```
    Remove the two keys for Excel:
    HKCR\TypeLib\{00020813-0000-0000-C000-000000000046}\1.8
    HKCR\TypeLib\{00020813-0000-0000-C000-000000000046}\1.9
    ```
    https://www.inflectra.com/support/knowledgebase/kb180.aspx

## References
- https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.exportasfixedformat