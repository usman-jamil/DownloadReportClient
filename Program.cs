using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Net;

namespace DownloadReportClient
{
    public static class DateTimeExtensions
    {
        public static DateTime StartOfWeek(this DateTime dt, DayOfWeek startOfWeek)
        {
            int diff = (7 + (dt.DayOfWeek - startOfWeek)) % 7;
            return dt.AddDays(-1 * diff).Date;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string strSite = ConfigurationManager.AppSettings["Site"];
            string strLibraryUrl = ConfigurationManager.AppSettings["LibraryUrl"];
            string strFileFormat = ConfigurationManager.AppSettings["FileFormat"];
            string strFileDateFormat = ConfigurationManager.AppSettings["FileDateFormat"];
            string strUpdateDay = ConfigurationManager.AppSettings["UpdateDay"];
            string strFolderLocation = ConfigurationManager.AppSettings["FolderLocation"];

            LogToFile("Starting Execution", "###############################################");

            string strLastOperationResult = System.IO.File.ReadAllText(string.Format("{0}\\LastOperationResult.txt", strFolderLocation));
            HttpStatusCode lastOperationResult = (HttpStatusCode)Enum.Parse(typeof(HttpStatusCode), strLastOperationResult);
            if (lastOperationResult == HttpStatusCode.Unauthorized)
            {
                LogToFile("UnAuthorized", "The last execution resulted in 401. The program will not continue. Please modify LastOperationResult.txt and changed text back to OK once the authorization issue is resolved");
                LogToFile("Ending Execution", "###############################################");
                return;
            }

            HttpStatusCode responseCode = HttpStatusCode.OK;
            DateTime serverTime = GetServerTime(strSite, out responseCode);

            if (responseCode == HttpStatusCode.OK && serverTime != DateTime.MinValue)
            {
                LogToFile("Fetch Server Time", "Completed: " + serverTime.ToString());
                DayOfWeek updateDay = (DayOfWeek)Enum.Parse(typeof(DayOfWeek), strUpdateDay);
                DateTime thursday = serverTime.StartOfWeek(DayOfWeek.Thursday);
                DateTime lastThursday = thursday.AddDays(-7);
                DateTime nextThursday = thursday.AddDays(7);
                LogToFile("This Thursday", thursday.ToString());
                LogToFile("Last Thursday", lastThursday.ToString());
                LogToFile("Next Thursday", nextThursday.ToString());

                string nextFileName = strFileFormat.Replace("##FileDateFormat##", nextThursday.ToString(strFileDateFormat));
                string newFileName = strFileFormat.Replace("##FileDateFormat##", thursday.ToString(strFileDateFormat));
                string oldFileName = strFileFormat.Replace("##FileDateFormat##", lastThursday.ToString(strFileDateFormat));

                string[] files = new string[] { nextFileName, newFileName, oldFileName };

                string timeStamp = System.IO.File.ReadAllText(string.Format("{0}\\Modified.txt", strFolderLocation));
                DateTime lastModifiedTimeStamp = new DateTime(Convert.ToInt64(timeStamp));
                int itemID = -1;

                foreach (string file in files)
                {
                    LogToFile("Reading File", file);
                    HttpStatusCode existsStatusCode = HttpStatusCode.OK;
                    DateTime lastModified = DateTime.MinValue;
                    bool fileExists = FileExists(strSite, strLibraryUrl, file, out existsStatusCode, out lastModified, out itemID);

                    if (!fileExists)
                    {
                        LogToFile("File Not Found", file + " does not exist");
                        continue;
                    }

                    if (existsStatusCode != HttpStatusCode.OK)
                    {
                        LogToFile("Web Response Error", existsStatusCode.ToString() + " is the response");
                        continue;
                    }

                    if (lastModified <= lastModifiedTimeStamp)
                    {
                        LogToFile("No Change", file + " has not changed");
                        break;
                    }

                    LogToFile("Processing File", file + " is being processed");
                    ProcessFile(strSite, strLibraryUrl, file, itemID, lastModified);
                    break;
                }
            }

            LogToFile("Ending Execution", "###############################################");
        }

        private static void ProcessFile(string strSite, string strLibraryUrl, string newFileName, int itemID, DateTime lastModified)
        {
            try
            {
                using (ClientContext clientContext = new ClientContext(strSite))
                {
                    List list = clientContext.Web.GetList(strLibraryUrl);
                    ListItem item = list.GetItemById(itemID);
                    File file = item.File;
                    clientContext.Load(file);
                    ClientResult<System.IO.Stream> stream = file.OpenBinaryStream();
                    clientContext.Load(item);
                    clientContext.Load(list);
                    clientContext.ExecuteQuery();

                    System.IO.Stream memoryStream = ReadFully(stream.Value);
                    string extension = newFileName.Substring(newFileName.LastIndexOf("."));
                    string strFileName = newFileName.Replace(extension, string.Empty);
                    string fileName = string.Format("{0} - {1}{2}",
                                                    newFileName,
                                                    lastModified.ToString("yyyy.MM.dd HH-mm"),
                                                    extension);
                    string strFolderLocation = ConfigurationManager.AppSettings["FolderLocation"];
                    using (System.IO.FileStream fileStream = System.IO.File.Create(strFolderLocation + "\\" + fileName, (int)memoryStream.Length))
                    {
                        // Initialize the bytes array with the stream length and then fill it with data 
                        byte[] bytesInStream = new byte[memoryStream.Length];
                        memoryStream.Read(bytesInStream, 0, bytesInStream.Length);
                        // Use write method to write to the file specified above 
                        fileStream.Write(bytesInStream, 0, bytesInStream.Length);

                        System.IO.File.WriteAllText(string.Format("{0}\\Modified.txt", strFolderLocation), Convert.ToString(lastModified.Ticks));
                        System.IO.File.WriteAllText(string.Format("{0}\\LastOperationResult.txt", strFolderLocation), "OK");
                    }
                }
            }
            catch (WebException ex)
            {
                var webResponse = ex.Response as HttpWebResponse;
                LogToFile("Web Response Exception", "Status Code: " + webResponse.StatusCode.ToString() + " Exception: " + ex.Message);
                string strFolderLocation = ConfigurationManager.AppSettings["FolderLocation"];
                System.IO.File.WriteAllText(string.Format("{0}\\LastOperationResult.txt", strFolderLocation), webResponse.StatusCode.ToString());
            }
            catch (Exception ex)
            { LogToFile("Exception", "Exception: " + ex.Message); }
        }

        private static System.IO.Stream ReadFully(System.IO.Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return new System.IO.MemoryStream(ms.ToArray()); ;
            }
        } 

        private static DateTime GetServerTime(string siteUrl, out HttpStatusCode statusCode)
        {
            DateTime dt = DateTime.MinValue;
            statusCode = HttpStatusCode.OK;

            using (ClientContext context = new ClientContext(siteUrl))
            {
                try
                {
                    Web web = context.Web;
                    context.Load(web, l => l.RegionalSettings);
                    ClientResult<System.DateTime> cr = web.RegionalSettings.TimeZone.UTCToLocalTime(DateTime.Now.ToUniversalTime());
                    context.ExecuteQuery();

                    dt = cr.Value;
                }
                catch (WebException ex)
                {
                    var webResponse = ex.Response as HttpWebResponse;
                    LogToFile("Web Response Exception", "Status Code: " + webResponse.StatusCode.ToString() + " Exception: " + ex.Message);
                    string strFolderLocation = ConfigurationManager.AppSettings["FolderLocation"];
                    System.IO.File.WriteAllText(string.Format("{0}\\LastOperationResult.txt", strFolderLocation), webResponse.StatusCode.ToString());
                }
                catch (Exception ex)
                { LogToFile("Exception", "Exception: " + ex.Message); }
            }

            return dt;
        }

        private static bool FileExists(string siteUrl, string libraryUrl, string name, out HttpStatusCode statusCode, out DateTime dateModified, out int itemID)
        {
            bool exists = false;
            statusCode = HttpStatusCode.OK;
            dateModified = DateTime.MinValue;
            itemID = -1;

            try
            {
                using (ClientContext clientContext = new ClientContext(siteUrl))
                {
                    List list = clientContext.Web.GetList(libraryUrl);
                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = string.Format(@"<View Scope='RecursiveAll'>
                                    <Query><Where><Eq><FieldRef Name='FileLeafRef' /><Value Type='File'>{0}</Value></Eq></Where>
                                    </Query>
                                </View>", name);
                    ListItemCollection listItems = list.GetItems(camlQuery);
                    clientContext.Load(listItems);
                    clientContext.ExecuteQuery();
                    exists = listItems.Count > 0;

                    ListItem item = listItems[0];
                    dateModified = Convert.ToDateTime(item["Modified"]).ToLocalTime();
                    itemID = item.Id;
                }
            }
            catch (WebException ex)
            {
                var webResponse = ex.Response as HttpWebResponse;
                LogToFile("Web Response Exception", "Status Code: " + webResponse.StatusCode.ToString() + " Exception: " + ex.Message);
                string strFolderLocation = ConfigurationManager.AppSettings["FolderLocation"];
                System.IO.File.WriteAllText(string.Format("{0}\\LastOperationResult.txt", strFolderLocation), webResponse.StatusCode.ToString());
            }
            catch (Exception ex)
            { LogToFile("Exception", "Exception: " + ex.Message); }

            return exists;
        }

        public static void LogToFile(string title, string message)
        {
            string strFolderLocation = ConfigurationManager.AppSettings["FolderLocation"];
            string filePath = string.Format("{0}\\Log.txt", strFolderLocation);

            bool exists = System.IO.File.Exists(filePath);
            if (!exists)
            {
                string header = string.Format("{0,-20}{1,-30}{2}" + Environment.NewLine,
                                                "Time Stamp",
                                                "Title",
                                                "Message");
                System.IO.File.WriteAllText(filePath, header);
            }

            title = title.Length > 25 ? string.Format("{0}...", title.Substring(0, 25)) : title;

            string output = string.Format("{0,-20}{1,-30}{2}" + Environment.NewLine,
                DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString(),
                title,
                message);

            System.IO.File.AppendAllText(filePath, output);
        }
    }
}