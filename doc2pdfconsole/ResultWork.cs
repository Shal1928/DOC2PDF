using System.Text;

namespace doc2pdfconsole
{
    public class ResultWork
    {
        public ResultWork(int id, string logPart, long duration)
        {
            Count = 0;
            ThreadId = id;
            LogPart = new StringBuilder();
            LogPart.AppendLine(logPart);
            Duration = duration;
        }

        public int ThreadId { get; set; }

        public StringBuilder LogPart { get; set; }

        public long Duration { get; set; }

        public void AddResultWork(ResultWork resultWork)
        {
            Count++;
            Duration += resultWork.Duration;
            LogPart.AppendLine(resultWork.LogPart.ToString());
        }

        public int Count { get; private set; }
    }
}
