using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.IO;
using System.Runtime.InteropServices;

public enum OfficeAppType
{
    Word,
    Excel,
    PPT
}

public interface IOfficeApplication : IDisposable
{
    void OpenDocument(string filePath);
    void SaveAsPDF(string toFilePath);
    void CloseDocument();
}

public class WordApplication : IOfficeApplication
{
    private Microsoft.Office.Interop.Word.Application _application;
    private Document? _document;
    public bool IsPrintRevisions { get; set; } = true;

    public WordApplication()
    {
        _application = new Microsoft.Office.Interop.Word.Application() { Visible = false, DisplayAlerts = WdAlertLevel.wdAlertsNone };
    }

    public void OpenDocument(string filePath)
    {
        _document = _application.Documents.Open(filePath, ReadOnly: true);
    }

    public void SaveAsPDF(string toFilePath)
    {
        if (_document != null)
        {
            var originalShowRevisions = _document.ShowRevisions;
            _document.ShowRevisions = IsPrintRevisions;
            _document.SaveAs2(toFilePath, WdSaveFormat.wdFormatPDF);
            _document.ShowRevisions = originalShowRevisions;
        }
    }

    public void CloseDocument()
    {
        if (_document != null)
        {
            _document.Close(WdSaveOptions.wdDoNotSaveChanges);
            Marshal.ReleaseComObject(_document);
            _document = null;
        }
    }

    public void Dispose()
    {
        _application.Quit();
        Marshal.ReleaseComObject(_application);
    }
}

public class ExcelApplication : IOfficeApplication
{
    private Microsoft.Office.Interop.Excel.Application _application;
    private Workbook? _workbook;
    public bool IsConvertOneSheetOnePDF { get; set; } = true;
    public bool IsPrintToOneColumn { get; set; } = false;

    public ExcelApplication()
    {
        _application = new Microsoft.Office.Interop.Excel.Application() { Visible = false, DisplayAlerts = false };
    }

    public void OpenDocument(string filePath)
    {
        _workbook = _application.Workbooks.Open(filePath, ReadOnly: true);
    }

    public void SaveAsPDF(string toFilePath)
    {
        if (_workbook != null)
        {
            if (IsConvertOneSheetOnePDF)
            {
                // 每个sheet保存为单独的PDF
                if (_workbook.Sheets.Count == 1)
                {
                    var worksheet = _workbook.Sheets[1];
                    if (IsPrintToOneColumn)
                    {
                        SetPrintToOneColumn(worksheet);
                    }
                    worksheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, toFilePath);
                    Marshal.ReleaseComObject(worksheet);
                }
                else
                {
                    for (int i = 1; i <= _workbook.Sheets.Count; i++)
                    {
                        var worksheet = _workbook.Sheets[i];
                        if (IsPrintToOneColumn)
                        {
                            SetPrintToOneColumn(worksheet);
                        }
                        var directory = Path.GetDirectoryName(toFilePath);
                        if (directory == null) throw new ArgumentException("文件没有目录", nameof(toFilePath));
                        worksheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, Path.Combine(directory, $"{Path.GetFileNameWithoutExtension(toFilePath)}_sheet{i}.pdf"));
                        Marshal.ReleaseComObject(worksheet);
                    }
                }
            }
            else
            {
                // 所有sheet保存为一个PDF
                if (IsPrintToOneColumn)
                {
                    // 为所有工作表设置一列宽打印
                    for (int i = 1; i <= _workbook.Sheets.Count; i++)
                    {
                        var worksheet = _workbook.Sheets[i];
                        SetPrintToOneColumn(worksheet);
                        Marshal.ReleaseComObject(worksheet);
                    }
                }
                _workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, toFilePath);
            }
        }
    }

    private void SetPrintToOneColumn(Worksheet worksheet)
    {
        try
        {
            // 设置页面设置为适合一页宽
            worksheet.PageSetup.Zoom = false;
            worksheet.PageSetup.FitToPagesWide = 1;
            worksheet.PageSetup.FitToPagesTall = 1; // 0表示自动调整高度
        }
        catch (Exception ex)
        {
            // 如果设置失败，记录错误但不中断转换
            System.Diagnostics.Debug.WriteLine($"设置一列宽打印失败: {ex.Message}");
        }
    }

    public void CloseDocument()
    {
        if (_workbook != null)
        {
            _workbook.Close(false);
            Marshal.ReleaseComObject(_workbook);
            _workbook = null;
        }
    }

    public void Dispose()
    {
        _application.Quit();
        Marshal.ReleaseComObject(_application);
    }
}

public class PowerPointApplication : IOfficeApplication
{
    private Microsoft.Office.Interop.PowerPoint.Application _application;
    private Presentation? _presentation;

    public PowerPointApplication()
    {
        _application = new Microsoft.Office.Interop.PowerPoint.Application();
    }

    public void OpenDocument(string filePath)
    {
        _presentation = _application.Presentations.Open(filePath, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);
    }

    public void SaveAsPDF(string toFilePath)
    {
        _presentation?.ExportAsFixedFormat(toFilePath, PpFixedFormatType.ppFixedFormatTypePDF);
    }

    public void CloseDocument()
    {
        if (_presentation != null)
        {
            _presentation.Close();
            Marshal.ReleaseComObject(_presentation);
            _presentation = null;
        }
    }

    public void Dispose()
    {
        _application.Quit();
        Marshal.ReleaseComObject(_application);
    }
}
