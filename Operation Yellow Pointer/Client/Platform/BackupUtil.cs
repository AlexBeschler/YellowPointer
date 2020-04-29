using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Amazon;
using Amazon.S3;
using Amazon.S3.Model;
using Amazon.S3.Transfer;
using LiteDB;
using Newtonsoft.Json;

namespace Operation_Yellow_Pointer.Client.Platform
{
    public class BackupUtil
    {
        public bool UploadSuccessful;

        //AWS Fields
        private const string BackupBucket = "yellowpointerdatabasebackup";
        //private const string BackupBucket = "yellowpointertest";
        private const string ErrorBucket = "yellowpointererror";
        private const string AccessKey = "";
        private const string SecretKey = "";

        private readonly AmazonS3Config _s3Config;
        private readonly string _database;
        private readonly string _temp;
        private readonly string _error;
        private readonly char _seperator;

        public BackupUtil(string dataPath)
        {
            _s3Config = new AmazonS3Config
            {
                MaxErrorRetry = 2,
                RegionEndpoint = RegionEndpoint.USEast1,
                Timeout = TimeSpan.FromSeconds(5)
            };
            _seperator = Utility.DetermineOs() == "windows" ? '\\' : '/';
            _database = dataPath + _seperator + "Database.db";
            _temp = dataPath + _seperator + "temp" + _seperator;
            _error = dataPath + "ErrorLog" + _seperator + "Error.txt";
        }

        #region BackupMethods

        public string Backup()
        {
            var date = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");
            var id = Guid.NewGuid().ToString();
            int count;
            using (var db = new LiteDatabase(_database))
            {
                var dbKeyRatio = db.GetCollection<KeyRatio>("keyRatios");
                count = dbKeyRatio.Count();
            }
            //Make copy of database
            File.Copy(_database, _temp + id + ".db");

            //Append to index file on server
            var list = GetIndexListing() ?? new List<Index>();
            //Add entry
            var index = new Index
            {
                Date = date,
                Guid = id,
                DatabaseEntries = count
            };
            list.Add(index);
            //Serialize
            var convertedJson = JsonConvert.SerializeObject(list, Formatting.Indented);
            //Write to file
            File.WriteAllText(_temp + "index.json", convertedJson);

            Compress(new FileInfo(_temp + id + ".db"));

            using (var client = new AmazonS3Client(AccessKey, SecretKey, _s3Config))
            {
                var fileTransfer = new TransferUtility(client);

                //Upload database
                var databaseRequest = new TransferUtilityUploadRequest
                {
                    FilePath = _temp + id + ".db.gz",
                    BucketName = BackupBucket
                };
                databaseRequest.UploadProgressEvent += UploadFileProgressCallback;
                //Upload index
                var indexRequest = new TransferUtilityUploadRequest
                {
                    FilePath = _temp + "index.json",
                    BucketName = BackupBucket,
                    CannedACL = S3CannedACL.PublicRead
                };

                if (!AttemptUpload(fileTransfer, databaseRequest))
                {
                    Console.WriteLine();
                    Console.WriteLine("Error: Failed to upload.");
                    return id;
                }
                if (!AttemptUpload(fileTransfer, indexRequest))
                {
                    Console.WriteLine();
                    Console.WriteLine("Error: Failed to upload index. Program will exit.");
                    Console.WriteLine("Please contact support.");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    //Created copy of database in current installation directory
                    Environment.Exit(1);
                }
                Console.WriteLine();
                File.Delete(_temp + id + ".db.gz");
                File.Delete(_temp + "index.json");
                UploadSuccessful = true;
            }
            return id;
        }

        public bool RestoreBackup(string id)
        {
            using (var client = new AmazonS3Client(AccessKey, SecretKey, _s3Config))
            {
                var request = new GetObjectRequest
                {
                    BucketName = BackupBucket,
                    Key = id + ".db.gz"
                };
                using (var response = client.GetObject(request))
                {
                    response.WriteObjectProgressEvent += DownloadFileProgressCallback;
                    var success = AttemptDownload(response, _database + ".gz");
                    if (!success) return false;
                    Decompress(new FileInfo(_database + ".gz"));
                    File.Delete(_database + ".gz");
                    return true;
                }
            }
        }

        public bool UploadErrorLogs()
        {
            var date = DateTime.Now.ToString("MM-dd-yyyy HH-mm-ss");
            Compress(new FileInfo(_error), "ErrorLogs " + date + ".txt");
            using (var client = new AmazonS3Client(AccessKey, SecretKey, _s3Config))
            {
                var fileTransfer = new TransferUtility(client);

                //Upload database
                var databaseRequest = new TransferUtilityUploadRequest
                {
                    FilePath = new FileInfo(_error).Directory + _seperator.ToString() + "ErrorLogs " + date + ".txt.gz",
                    BucketName = ErrorBucket
                };
                databaseRequest.UploadProgressEvent += UploadFileProgressCallback;

                if (!AttemptUpload(fileTransfer, databaseRequest))
                {
                    Console.WriteLine();
                    Console.WriteLine("Error: Failed to upload.");
                    return false;
                }
                Console.WriteLine();
                File.Delete(_error + ".gz");
                return true;
            }
        }

        #endregion
        
        #region Helpers

        public class Index
        {
            public string Date { get; set; }
            public string Guid { get; set; }
            public int DatabaseEntries { get; set; }
        }

        public List<Index> GetIndexListing()
        {
            List<Index> list = null;
            try
            {
                using (var webClient = new Utility.WebDownload())
                {
                    webClient.DownloadFile("https://s3.amazonaws.com/" + BackupBucket + "/index.json",
                        _temp + "index.json");
                }
                list =
                    JsonConvert.DeserializeObject<List<Index>>(File.ReadAllText(_temp + "index.json")) ??
                    new List<Index>();
                File.Delete(_temp + "index.json");
            }
            catch (Exception)
            {
                // ignored
            }
            return list;
        }

        public void PurgeTempBackupFiles(string databaseId)
        {
            try
            {
                File.Delete(_temp + databaseId + ".db");
                File.Delete(_temp + "index.json");
            }
            catch (Exception)
            {
                // ignored
            }
        }

        #endregion

        #region FileTransferMethods

        private bool AttemptUpload(TransferUtility fileTransfer, TransferUtilityUploadRequest request)
        {
            const int timeoutTimes = 4;
            var currentTimeout = 1;
            var success = false;
            do
            {
                try
                {
                    fileTransfer.Upload(request);
                    success = true;
                    return true;
                }
                catch (Exception)
                {
                    const int maxtimeout = timeoutTimes - 1;
                    Console.WriteLine("Attempting to reupload file (" + currentTimeout + "/" + maxtimeout + " times)");
                    currentTimeout++;
                }
            } while (currentTimeout < timeoutTimes || success);
            return false;
        }

        private bool AttemptDownload(GetObjectResponse response, string downloadLocation)
        {
            const int timeoutTimes = 4;
            var currentTimeout = 1;
            var success = false;
            do
            {
                try
                {
                    response.WriteResponseStreamToFile(downloadLocation);
                    success = true;
                    return true;
                }
                catch (Exception)
                {
                    const int maxtimeout = timeoutTimes - 1;
                    Console.WriteLine("Attempting to reupload file (" + currentTimeout + "/" + maxtimeout + " times)");
                    currentTimeout++;
                }
            } while (currentTimeout < timeoutTimes || success);
            return false;
        }

        #endregion

        #region Callbacks

        private void UploadFileProgressCallback(object sender, UploadProgressArgs e)
        {
            var uploaded = e.TransferredBytes / Math.Pow(1024, 2);
            var total = e.TotalBytes / Math.Pow(1024, 2);
            Console.Write("\r" + Math.Round(uploaded, 2) + "/" + Math.Round(total, 2) + " MB (" + e.PercentDone + "%)");
        }
        private void DownloadFileProgressCallback(object sender, WriteObjectProgressArgs e)
        {
            var downloaded = e.TransferredBytes / Math.Pow(1024, 2);
            var total = e.TotalBytes / Math.Pow(1024, 2);
            Console.Write("\r" + Math.Round(downloaded, 2) + "/" + Math.Round(total, 2) + " MB (" + e.PercentDone + "%)");
        }

        #endregion

        #region CompressionMethods

        private void Compress(FileInfo file)
        {
            // Get the stream of the source file.
            using (var inFile = file.OpenRead())
            {
                // Prevent compressing hidden and already compressed files.
                if (!((File.GetAttributes(file.FullName) & FileAttributes.Hidden) != FileAttributes.Hidden & file.Extension != ".gz"))
                    return;
                // Create the compressed file
                using (var outFile = File.Create(file.FullName + ".gz"))
                {
                    using (var compress = new GZipStream(outFile, CompressionMode.Compress))
                    {
                        // Copy the source file into the compression stream.
                        inFile.CopyTo(compress);
                    }
                }
            }
        }

        private void Compress(FileInfo file, string fileName)
        {
            // Get the stream of the source file.
            using (var inFile = file.OpenRead())
            {
                // Prevent compressing hidden and already compressed files.
                if (!((File.GetAttributes(file.FullName) & FileAttributes.Hidden) != FileAttributes.Hidden &
                      file.Extension != ".gz"))
                    return;
                // Create the compressed file
                using (var outFile = File.Create(file.Directory + _seperator.ToString() + fileName + ".gz"))
                {
                    using (var compress = new GZipStream(outFile, CompressionMode.Compress))
                    {
                        // Copy the source file into the compression stream.
                        inFile.CopyTo(compress);
                    }
                }
            }
        }

        private void Decompress(FileInfo file)
        {
            using (var originalFileStream = file.OpenRead())
            {
                var currentFileName = file.FullName;
                var newFileName = currentFileName.Remove(currentFileName.Length - file.Extension.Length);

                using (var decompressedFileStream = File.Create(newFileName))
                {
                    using (var decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress))
                    {
                        decompressionStream.CopyTo(decompressedFileStream);
                    }
                }
            }
        }

        #endregion
        
    }
}
