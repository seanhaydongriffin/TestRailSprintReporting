using System;
using System.IO;
using System.Reflection;

namespace SharedProject
{
    class Log
    {
        static public string path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + "\\" + Assembly.GetCallingAssembly().GetName().Name + ".log";


        static public void Initialise(string log_path)
        {
            if (log_path != null)

                path = log_path.ReplaceIgnoreCase("/LOG:", "");

            if (System.IO.File.Exists(path))

                System.IO.File.Delete(path);
        }

        static public void WriteLine(string message)
        {
            Console.WriteLine(message);

            using (StreamWriter w = System.IO.File.AppendText(path))
            {
                w.WriteLine($"{System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")} : {message}");
            }
        }
        static public void Reset()
        {
            if (System.IO.File.Exists(path))

                System.IO.File.Delete(path);
        }

        static public void Debug(string message)
        {
            using (StreamWriter w = System.IO.File.AppendText(path))
            {
                w.Write("\r\nLog Entry : ");
                w.WriteLine($"{System.DateTime.Now.ToLongTimeString()} {System.DateTime.Now.ToLongDateString()}");
                w.WriteLine("  :");
                w.WriteLine($"  :{message}");
                w.WriteLine("-------------------------------");
            }
        }


    }
}
