using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.IO;

namespace DOC2PDF
{
    public class WorkClass
    {
        private object missing = System.Reflection.Missing.Value;
        private XlFixedFormatType paramExportFormat = XlFixedFormatType.xlTypePDF;
        private XlFixedFormatQuality paramExportQuality = XlFixedFormatQuality.xlQualityStandard;
        private bool paramOpenAfterPublish = false;
        private bool paramIncludeDocProps = true;
        private bool paramIgnorePrintAreas = true;

        private String _file;
        private String _targetFolder;
        private Microsoft.Office.Interop.Word.Application _appWord = null;
        private Microsoft.Office.Interop.Excel.Application _appExcel = null;

        public WorkClass(String file, String targetFolder, bool isWord)
        {
            _file = file;
            _targetFolder = targetFolder;
            if (isWord)
            {
                _appWord = new Microsoft.Office.Interop.Word.Application();
            }else
            {
                _appExcel = new Microsoft.Office.Interop.Excel.Application();
            }
            
        }

        public WorkClass(String file, String targetFolder)
        {
            _file = file;
            _targetFolder = targetFolder;
        }

        public void work(Object stateInfo)
        {
            try
            {
                var fileName = Path.GetFullPath(_file);

                if (_appWord != null)
                {
                    Microsoft.Office.Interop.Word.Document wordDocument = _appWord.Documents.Open(fileName);
                    try
                    {
                        wordDocument = _appWord.Documents.Open(fileName, ReadOnly: false, Visible: false);
                        wordDocument.ExportAsFixedFormat(GetPdfFileName(_file, _targetFolder), WdExportFormat.wdExportFormatPDF);

                    }
                    finally
                    {
                        wordDocument.Close(ref missing, ref missing, ref missing);
                        _appWord.Quit(SaveChanges: false, OriginalFormat: false, RouteDocument: false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_appWord);
                    }
                }
                else
                {
                    Microsoft.Office.Interop.Excel.Workbook excelDocument = _appExcel.Workbooks.Open(fileName);
                    try
                    {
                        excelDocument = _appExcel.Workbooks.Open(fileName, ReadOnly: false);
                        excelDocument.ExportAsFixedFormat(paramExportFormat,
                            GetPdfFileName(_file, _targetFolder), paramExportQuality,
                            paramIncludeDocProps, paramIgnorePrintAreas, missing,
                            missing, paramOpenAfterPublish,
                            missing);

                    }
                    finally
                    {
                        excelDocument.Close(false, missing, missing);
                        _appExcel.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_appExcel);
                    }
                }
            } catch(Exception e)
            {
                //
            }           
        }

        private String GetPdfFileName(String file, String targetFolder)
        {
            var pdfFileName = Path.GetFileName(Path.ChangeExtension(file, "pdf"));
            return String.Format(@"{0}\{1}", targetFolder, pdfFileName);
        }
    }
}
