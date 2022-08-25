using System;
using System.IO;

namespace SharedProject
{
    public static class DateTime
    {

        public static (int LongYear, int ShortYear, int Quarter) GetNowQuarterInfo()
        {
            var now = System.DateTime.Now;
            var year = now.Month > 6 ? now.Year + 1 : now.Year;
            var quarter = 4;

            if (now.Month >= 7 && now.Month <= 9)
                quarter = 1;
            else if (now.Month >= 10 && now.Month <= 12)
                quarter = 2;
            else if (now.Month >= 1 && now.Month <= 3)
                quarter = 3;

            return (year, year % 100, quarter);
        }

        public static bool create(String directory_path)
        {
            System.IO.Directory.CreateDirectory(directory_path);
            return true;
        }


        public static void delete(this FileSystemInfo fileSystemInfo)
        {
            var directoryInfo = fileSystemInfo as DirectoryInfo;

            if (directoryInfo != null)
            {
                try
                {
                    foreach (FileSystemInfo childInfo in directoryInfo.GetFileSystemInfos())

                        delete(childInfo);
                } catch (Exception)
                {
                }
            }

            try
            {
                fileSystemInfo.Attributes = FileAttributes.Normal;
                fileSystemInfo.Delete();
            }
            catch (Exception)
            {
            }
        }

        public static string GetRightPartOfPath(string path, string startAfterPart)
        {
            // use the correct seperator for the environment
            var pathParts = path.Split(Path.DirectorySeparatorChar);

            // this assumes a case sensitive check. If you don't want this, you may want to loop through the pathParts looking
            // for your "startAfterPath" with a StringComparison.OrdinalIgnoreCase check instead
            int startAfter = Array.IndexOf(pathParts, startAfterPart);

            if (startAfter == -1)
            {
                // path not found
                return null;
            }

            // try and work out if last part was a directory - if not, drop the last part as we don't want the filename
            //var lastPartWasDirectory = pathParts[pathParts.Length - 1].EndsWith(Path.DirectorySeparatorChar.ToString());
            return string.Join(
                Path.DirectorySeparatorChar.ToString(),
                pathParts, startAfter,
                pathParts.Length - startAfter); // - (lastPartWasDirectory ? 0 : 1));
        }

        public static System.DateTime UnixTimeStampToDateTime(double unixTimeStamp)
        {
            // Unix timestamp is seconds past epoch
            System.DateTime dateTime = new System.DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            dateTime = dateTime.AddSeconds(unixTimeStamp).ToLocalTime();
            return dateTime;
        }
    }
}
