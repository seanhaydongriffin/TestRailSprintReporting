using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace SharedProject
{
    public static class Command
    {

        public static bool Execute(string filename, string arguments, string working_directory)
        {
            ProcessStartInfo _processStartInfo = new ProcessStartInfo();
            _processStartInfo.WorkingDirectory = @working_directory;
            _processStartInfo.FileName = @filename;
            _processStartInfo.Arguments = arguments;
            System.Diagnostics.Process myProcess = System.Diagnostics.Process.Start(_processStartInfo);

            return true;
        }


        public static string Execute(string filename, string arguments, string working_directory, Encoding StandardOutputEncoding)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.FileName = @filename;
            process.StartInfo.Arguments = arguments;
            process.StartInfo.WorkingDirectory = @working_directory;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.StandardOutputEncoding = StandardOutputEncoding; // ie. Encoding.UTF8;
            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            return output;
        }

        public static string Execute(string filename, string arguments, string working_directory, string standard_input, Encoding standard_input_encoding, Encoding standard_output_encoding)
        {
            if (File.IsValidFilename(filename) && !working_directory.Equals(""))

                filename = System.IO.Path.Combine(working_directory, filename);

            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.FileName = @filename;
            process.StartInfo.Arguments = arguments;
            process.StartInfo.WorkingDirectory = @working_directory;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.StandardOutputEncoding = standard_output_encoding; // ie. Encoding.UTF8;
            process.Start();
            StreamWriter utf8Writer = new StreamWriter(process.StandardInput.BaseStream, standard_input_encoding); // ie. Encoding.UTF8;
            utf8Writer.Write(standard_input);
            utf8Writer.Close();
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            return output;
        }

    }
}
