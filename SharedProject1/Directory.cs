using System;
using System.IO;

namespace SharedProject
{
    public static class Directory
    {

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


    }
}
