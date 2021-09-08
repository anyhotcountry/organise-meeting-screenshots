using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScreenshotBuddy
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            var outlookAccess = new OutlookAccess();
            var dirInfo = new DirectoryInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Screenshots"));
            var invalids = Path.GetInvalidFileNameChars();
            foreach (var file in dirInfo.EnumerateFiles())
            {
                var evt = outlookAccess.GetEvents(file.CreationTime).FirstOrDefault();
                if (evt != null)
                {
                    var evtFolder = string.Empty;
                    try
                    {
                        evtFolder = string.Join("_", evt.Split(invalids, StringSplitOptions.RemoveEmptyEntries)).Trim();
                        var fullPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Screenshots", evtFolder);
                        Directory.CreateDirectory(fullPath);
                        Console.WriteLine($"{file.Name} => {fullPath}");
                        file.MoveTo(Path.Combine(fullPath, $"{file.CreationTime:yyyy-MM-dd} - {file.Name}"));
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"{evtFolder} - {e.Message}");
                    }
                }
            }
        }
    }
}
