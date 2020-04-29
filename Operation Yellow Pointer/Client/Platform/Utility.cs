using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using FileMode = System.IO.FileMode;

namespace Operation_Yellow_Pointer.Client.Platform
{
    internal class Utility
    {
        internal static string DataPath { private get; set; }
        internal static string DesktopPath { private get; set; }
        internal static string ErrorLog { private get; set; }
        internal static bool AdminMode { get; set; }
        internal static string Temp { private get; set; }

        //Valid returns: 'windows', 'macos' (includes iOS), 'linux' (includes Android), 'unrecognized'
        internal static string DetermineOs()
        {
            var windir = Environment.GetEnvironmentVariable("windir");
            if (!string.IsNullOrEmpty(windir) && windir.Contains(@"\") && Directory.Exists(windir))
            {
                return "windows";
            }
            if (!File.Exists(@"/proc/sys/kernel/ostype"))
                return File.Exists(@"/System/Library/CoreServices/SystemVersion.plist")
                    ? "macos"
                    : "unrecognized";
            var osType = File.ReadAllText(@"/proc/sys/kernel/ostype");
            return osType.StartsWith("Linux", StringComparison.OrdinalIgnoreCase)
                ? "linux"
                : "unrecognized";
        }

        public static void ExecuteCommand(string command)
        {
            var proc = new Process
            {
                StartInfo =
                {
                    FileName = "/bin/bash",
                    Arguments = "-c \" " + command + " \"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true
                }
            };
            proc.Start();

            while (!proc.StandardOutput.EndOfStream)
            {
                Console.WriteLine(proc.StandardOutput.ReadLine());
            }
        }

        #region ErrorLogging
        internal static void WriteToLog(string message, string fileName)
        {
            var path = DesktopPath + fileName;
            var file = new StreamWriter(path, true);
            file.WriteLine(message);
            file.Close();
        }
        internal static void WriteToLog(string ticker, string message, string fileName)
        {
            WriteToLog(ticker + ": " + message, fileName);
        }
        internal static void WriteToErrorLog(Exception exception)
        {
            var time = DateTime.Now.ToString(CultureInfo.CurrentCulture);
            var message =
                time + 
                ": " +
                exception.Message +
                "Source: " + exception.Source +
                "Inner Exception: " + exception.InnerException +
                "Stacktrace: " + exception.StackTrace;
            File.AppendAllText(ErrorLog, message);
        }
        internal static void WriteToErrorLog(string error)
        {
            var time = DateTime.Now.ToString(CultureInfo.CurrentCulture);
            var message =
                time + 
                ": " +
                error + 
                "\n";
            File.AppendAllText(ErrorLog, message);
        }
        #endregion

        internal static void CheckIfFileIsLocked(FileInfo file)
        {
            var fileUnlocked = false;
            do
            {
                if (IsFileLocked(file))
                {
                    Console.WriteLine("The file you are trying to import is currently open with Excel");
                    Console.WriteLine("Please close the file in Excel and press any key to continue");
                    Console.ReadKey();
                }
                else
                    fileUnlocked = true;
            } while (!fileUnlocked);
        }

        public static void BackUpDatabase()
        {
            Console.WriteLine("Modification to main database requested. Would you like to back it up now?");
            Console.WriteLine("1. Yes");
            Console.WriteLine("2. No");
            var result = Console.ReadLine();
            switch (result)
            {
                case "1":
                    Console.WriteLine("Backing up...");
                    var backup = new BackupUtil(DataPath);
                    var id = backup.Backup();
                    if (!backup.UploadSuccessful)
                    {
                        Console.WriteLine("Backup unsuccessful. Please try again later.");
                        Console.WriteLine("Press any key to continue...");
                        Console.ReadKey();
                        backup.PurgeTempBackupFiles(id);
                    }
                    break;
                case "2":
                    break;
            }
        }

        public static void PurgeTempDirectory()
        {
            var directory = new DirectoryInfo(Temp);

            foreach (var file in directory.GetFiles())
            {
                try
                {
                    file.Delete();
                }
                catch (Exception e)
                {
                    if (IsFileLocked(file))
                    {
                        WriteToErrorLog(file.FullName + " was detected as being locked");
                    }
                    else
                    {
                        WriteToErrorLog(e);
                    }
                }
            }
        }

        #region StringManipulation

        internal static List<string> RemoveWhiteSpace(List<string> array)
        {
            for (var i = 0; i < array.Count; i++)
            {
                array[i] = array[i].Replace(" ", string.Empty);
            }
            return array;
        }

        internal static string[] RemoveWhiteSpace(string[] array)
        {
            for (var i = 0; i < array.Length; i++)
            {
                array[i] = array[i].Replace(" ", string.Empty);
            }
            return array;
        }

        internal static string UppercaseFirst(string s)
        {
            return char.ToUpper(s[0]) + s.Substring(1);
        }

        #endregion
        
        internal class WebDownload : WebClient
        {
            private int Timeout { get; }
            public WebDownload() : this(5 * 1000) { } //5 seconds
            private WebDownload(int timeout)
            {
                Timeout = timeout;
            }
            protected override WebRequest GetWebRequest(Uri address)
            {
                var request = base.GetWebRequest(address);
                if (request != null)
                {
                    request.Timeout = Timeout;
                }
                return request;
            }
        }

        #region InternalMethods

        private static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;
            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                // File is locked because it's open or does not exist
                return true;
            }
            finally
            {
                stream?.Close();
            }
            return false;
        }

        #endregion
    }
}
