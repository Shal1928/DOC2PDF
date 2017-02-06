using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace doc2pdfconsole
{
    public class Program
    {
        private const string InputDir = @"C:\OfficeDocs";
        private const string OutputDir = @"C:\PDFDocs";

        private static readonly Dictionary<int, ResultWork> Dic = new Dictionary<int, ResultWork>();
        public delegate void MainDelegate(WorkClass sender, ResultWork resultWork);

        private static void Main()
        {
            try
            {
                FileStream filestream = new FileStream($@"{OutputDir}\out.txt", FileMode.Create);
                var streamwriter = new StreamWriter(filestream)
                {
                    AutoFlush = true
                };
                Console.SetOut(streamwriter);
                Console.SetError(streamwriter);

                var filesEnumerate = Directory.EnumerateFiles(InputDir);
                var i = 0;
                ThreadPool.SetMaxThreads(10, 10);

                var files = filesEnumerate as string[] ?? filesEnumerate.ToArray();
                Console.WriteLine($@"Input: {InputDir}; Count: {files.Count()}");
                Console.WriteLine($@"Output: {OutputDir}");
                var oneHour = new TimeSpan(0, 1, 0);
                var mainStopWatch = new Stopwatch();
                mainStopWatch.Start();
                foreach (var file in files)
                {
                    var officeExt = Path.GetExtension(file);
                    if (officeExt != ".docx" && officeExt != ".doc" && officeExt != ".xlsx" && officeExt != ".xls")
                        continue;

                    i++;

                    var workClass = new WorkClass(file, OutputDir, officeExt == ".docx" || officeExt == ".doc");
                    workClass.OnUpdate += ThreadMessages;
                    ThreadPool.QueueUserWorkItem(workClass.Work);

                    if (mainStopWatch.Elapsed.Duration() >= oneHour)
                    {
                        Console.WriteLine($"Finally count by one hour is {i}");
                        var average = (60d * 60d) / i;
                        Console.WriteLine($"Average {average:n2} sec.");
                        break; 
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                
                foreach (var k in Dic.Keys)
                {
                    ResultWork r;
                    if (Dic.TryGetValue(k, out r))
                    {
                        r.LogPart.AppendLine($"Duration {r.Duration} sec.");
                        // ReSharper disable once PossibleLossOfFraction
                        double average = r.Duration / r.Count;
                        r.LogPart.AppendLine($"Total {r.Count}; Average {average:n2} sec.");
                        File.WriteAllText($@"{OutputDir}\{k}.txt", r.LogPart.ToString());
                    }
                    
                }

                //Console.ReadLine();
            }
        }

        private static void ThreadMessages(WorkClass sender, ResultWork resultWork)
        {
            var id = resultWork.ThreadId;
            ResultWork r;
            if (Dic.TryGetValue(id, out r))
            {
                r.AddResultWork(resultWork);
            }
            else
            {
                Dic.Add(id, resultWork);
            }
        }
    }
}
