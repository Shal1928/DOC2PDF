using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

namespace doc2pdfconsole
{
    public class WorkClass
    {
        private object _missing = System.Reflection.Missing.Value;
        private const XlFixedFormatType ParamExportFormat = XlFixedFormatType.xlTypePDF;
        private const XlFixedFormatQuality ParamExportQuality = XlFixedFormatQuality.xlQualityStandard;
        private const bool ParamOpenAfterPublish = false;
        private const bool ParamIncludeDocProps = true;
        private const bool ParamIgnorePrintAreas = true;

        private readonly string _file;
        private readonly string _targetFolder;
        private readonly Microsoft.Office.Interop.Word.Application _appWord;
        private readonly Microsoft.Office.Interop.Excel.Application _appExcel;

        public event Program.MainDelegate OnUpdate;

        public int ThreadId { get; set; }
        public long Duration { get; set; }

        public WorkClass(string file, string targetFolder, bool isWord)
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

        public WorkClass(string file, string targetFolder)
        {
            _file = file;
            _targetFolder = targetFolder;
        }

        public void Work(object stateInfo)
        {
            var stopWatch = new Stopwatch();
            var logPart = new StringBuilder();
            var fileSize = 0L;
            string log;
            try
            {

                var fileName = Path.GetFullPath(_file);
                fileSize = new FileInfo(fileName).Length * 1000;

                ThreadId = Thread.CurrentThread.ManagedThreadId;
                log = $"Thread {ThreadId} {Thread.CurrentThread.Name} work with {_file} {fileSize} kb started.";
                logPart.AppendLine(log);
                Console.WriteLine(log);

                stopWatch.Start();
                if (_appWord != null)
                {
                    var wordDocument = _appWord.Documents.Open(fileName);
                    try
                    {
                        wordDocument = _appWord.Documents.Open(fileName, ReadOnly: false, Visible: false);
                        wordDocument.ExportAsFixedFormat(GetPdfFileName(_file, _targetFolder),
                            WdExportFormat.wdExportFormatPDF);

                    }
                    finally
                    {
                        wordDocument.Close(ref _missing, ref _missing, ref _missing);
                        _appWord.Quit(SaveChanges: false, OriginalFormat: false, RouteDocument: false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_appWord);
                    }
                }
                else
                {
                    var excelDocument = _appExcel.Workbooks.Open(fileName);
                    try
                    {
                        excelDocument = _appExcel.Workbooks.Open(fileName, ReadOnly: false);
                        excelDocument.ExportAsFixedFormat(ParamExportFormat,
                            GetPdfFileName(_file, _targetFolder), ParamExportQuality,
                            ParamIncludeDocProps, ParamIgnorePrintAreas, _missing,
                            _missing, ParamOpenAfterPublish,
                            _missing);

                    }
                    finally
                    {
                        excelDocument.Close(false, _missing, _missing);
                        _appExcel.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_appExcel);
                    }
                }
            }
            catch (Exception)
            {
                //
            }
            finally
            {
                Duration = stopWatch.Elapsed.Seconds;
                log = $"Thread {Thread.CurrentThread.ManagedThreadId} {Thread.CurrentThread.Name} work with {_file} {fileSize} kb finished by {Duration} ms.";
                logPart.AppendLine(log);
                Console.WriteLine(log);
                OnOnUpdate(new ResultWork(ThreadId, logPart.ToString(), Duration));
            }
        }

        private static string GetPdfFileName(string file, string targetFolder)
        {
            var pdfFileName = Path.GetFileName(Path.ChangeExtension(file, "pdf"));
            return $@"{targetFolder}\{pdfFileName}";
        }

        protected virtual void OnOnUpdate(ResultWork resultWork)
        {
            OnUpdate?.Invoke(this, resultWork);
        }
    }
}
