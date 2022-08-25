using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace SharedProject
{
    class File
    {

        public static bool Exists(String filename)
        {
            return System.IO.File.Exists(filename);
        }



        public static String read(String filename)
        {
            var result_str = "";

            if (Exists(filename))

                result_str = System.IO.File.ReadAllText(filename);

            return result_str;
        }



        public static bool overwrite(String filename, String data)
        {
            try
            {
                System.IO.File.WriteAllText(@filename, data);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }


        public static bool append(String filename, String text)
        {
            System.IO.File.AppendAllText(@filename, text);
            //            System.IO.File.AppendAllText(@filename, text + Environment.NewLine);

            return true;
        }

        public static bool remove_bytes_from_end(String filename, long num_of_bytes)
        {

            FileInfo fi = new FileInfo(filename);
            FileStream fs = fi.Open(FileMode.Open);
            fs.SetLength(Math.Max(0, fi.Length - num_of_bytes));
            fs.Close();

            return true;
        }

        public static bool delete(String directory_path, String filename)
        {
            bool result = true;

            var dir = new DirectoryInfo(directory_path);

            try
            {
                foreach (var file in dir.EnumerateFiles(filename))
                {
                    try
                    {
                        file.Delete();
                    }
                    catch (Exception)
                    {
                        result = false;
                    }
                }
            }
            catch (DirectoryNotFoundException)
            {
                result = false;
            }

            return result;
        }



        public static void CopyFiles(String SourceFilePath, String TargetFilePath, bool overwrite, string searchPattern)
        {
            DirectoryInfo source = new DirectoryInfo(SourceFilePath);
            DirectoryInfo destination = new DirectoryInfo(TargetFilePath);
            FileInfo[] files = source.GetFiles(searchPattern);

            //this section is what's really important for your application.
            foreach (FileInfo file in files)
            {
                file.CopyTo(destination.FullName + "\\" + file.Name, overwrite);
            }
        }


        public static bool Copy(String SourceFilePath, String TargetFilePath)
        {
            bool result = true;

            try
            {
                System.IO.Directory.CreateDirectory(Path.GetDirectoryName(TargetFilePath));
            }
            catch (Exception)
            {
            }

            try
            {
                System.IO.File.Copy(SourceFilePath, TargetFilePath, true);
            }
            catch (Exception)
            {
                result = false;
            }

            return result;
        }

        public static bool CopyLatest(String SourceDirectory, String TargetDirectory)
        {
            bool result = true;

            try
            {
                var directory = new DirectoryInfo("C:\\MyDirectory");
                var myFile = directory.GetFiles().OrderByDescending(f => f.LastWriteTime).First();
                Copy(myFile.FullName, "R:\\");
            }
            catch (Exception)
            {
                result = false;
            }

            return result;
        }

        public static bool CopyLatestFromDownloads(String SourceDirectory, String TargetDirectory)
        {
            bool result = true;

            try
            {
                var directory = new DirectoryInfo("C:\\MyDirectory");
                var myFile = directory.GetFiles().OrderByDescending(f => f.LastWriteTime).First();
                Copy(myFile.FullName, "R:\\");
            }
            catch (Exception)
            {
                result = false;
            }

            return result;
        }

        //public static int RoboCopy(TextBox tb, String source_path, String target_path, String filename = "")
        //{
            
	       // if (filename.Length > 0)

		      //  filename = " \"" + filename + "\"";

        //    //tb.LogMessage("Executing robocopy.exe " + "\"" + source_path + "\" \"" + target_path + "\"" + filename + " /sec /purge /is /it /r:15 /w:1 /np /ns /nc /nfl /ndl /njh /njs");

        //    var _processStartInfo = new ProcessStartInfo();
        //    _processStartInfo.FileName = "robocopy.exe";
        //    _processStartInfo.Arguments = "\"" + source_path + "\" \"" + target_path + "\"" + filename + " /sec /purge /is /it /r:15 /w:1 /np /ns /nc /nfl /ndl /njh /njs";
        //    var process = Process.Start(_processStartInfo);
        //    process.WaitForExit();
        //    var result = process.ExitCode;
        //    return result;
        //}


        public static bool IsValidFilename(string testName)
        {
            Regex containsABadCharacter = new Regex("[" + Regex.Escape(new string(System.IO.Path.GetInvalidPathChars())) + "]");
            if (containsABadCharacter.IsMatch(testName)) { return false; };

            // other checks for UNC, drive-path format, etc

            return true;
        }


    }
}
