using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using MSC=Microsoft.SharePoint.Client;
using Microsoft.SharePoint;
using System.Configuration;
using System.Data.OleDb;
using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Net;
using System.Xml;
using Microsoft.SharePoint.Client;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Threading;
using System.Windows.Forms;

namespace UploadToSP
{
    public class Program
    {
        public static string sharePointSite = "https://share.philips.com/sites/STS020170706072003/";
        public static string documentLibraryName = "Documents";
        public static MSC.ClientContext context;
        static void Main(string[] args)
        {
            new SharePointHandler().Upload();
        }
    }


    public class SharePointHandler : Program
    {
      
        public SharePointHandler() { }


        public void Upload()
        {
            try
            {
                using (context = new MSC.ClientContext(sharePointSite))
                {

                    SecureString s = new SecureString();
                    //s.
                    MSC.SharePointOnlineCredentials cred = new MSC.SharePointOnlineCredentials(ConfigurationManager.AppSettings["UsrName"], getPassword(ConfigurationManager.AppSettings["PassWord"]));
                    context.Credentials = cred;
                    var list = context.Web.Lists.GetByTitle(documentLibraryName);
                    context.Load(list);

                    var root = list.RootFolder;
                    context.Load(root);
                    context.ExecuteQuery();

                    // ADDITION
                    string SourceDocPath = ConfigurationManager.AppSettings["SourceDocsPath"];

                    DirectoryInfo dInfo = new DirectoryInfo(SourceDocPath);
                    FileInfo[] ListofFiles = dInfo.GetFiles();
                    List<linkIdentifier> listofLinks = new List<linkIdentifier>();
                    XmlDocument doc = new XmlDocument();
                    doc.Load("Links.xml");
                    XmlNodeList listXml = doc.GetElementsByTagName("link");
                    foreach (XmlNode n1 in listXml)
                    {
                        linkIdentifier id = new linkIdentifier();
                        id.rowIndex = Convert.ToInt32(n1["rowIndex"].InnerText);
                        id.colIndex = Convert.ToInt32(n1["colIndex"].InnerText);
                        id.SheetName = n1["SheetName"].InnerText;
                        listofLinks.Add(id);
                    }

                    foreach (FileInfo fileInstance in ListofFiles)
                    {
                        bool IsgoodLink = false;
                        string path = fileInstance.FullName;

                        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                        Workbook wb = excel.Workbooks.Open(path);

                        //***********************LINK CHECK*****************************************
                        //Read the first cell
                        foreach (linkIdentifier identifier in listofLinks)
                        {
                           

                            Worksheet excelSheet = wb.Sheets[identifier.SheetName];
                            string test = excelSheet.Cells[identifier.rowIndex, identifier.colIndex].Formula;
                           
                            test = test.Split(',')[0].TrimEnd("\"".ToCharArray());
                            String[] pathList = test.Split('/');
                            
                            try
                            {
                                if (test.Contains(".aspx"))
                                {
                                    //LinkCheck(test);
                                    IsgoodLink = CheckLink(pathList, cred);

                                }
                                else
                                {
                                    IsgoodLink=CheckLink(pathList, cred);
                                }
                            }
                            catch (MSC.ServerException e)
                            {
                                if (e.ServerErrorTypeName == "System.IO.FileNotFoundException")
                                    IsgoodLink = false;
                                wb.Close();
                                IsgoodLink = false;
                            }
                            if (IsgoodLink == false)
                            {
                                Console.WriteLine("File {0} is having deadlinks.", fileInstance.Name);
                                wb.Close();
                                return;
                            }
                           
                        }
                        wb.Close();
                        //***********************LINK CHECK*****************************************


                        string tempdir = fileInstance.Name;
                            tempdir = tempdir.Substring("2019.craft ".Length);
                            tempdir = tempdir.Trim(' ');
                            tempdir = tempdir.Remove((tempdir.Length - ".xlsm".Length));
                            String ParentDirectoryName = tempdir.Split('-')[0];
                            ParentDirectoryName = ParentDirectoryName.Trim();
                            string ChildDirectoryName = tempdir.Split('-')[1];
                            ChildDirectoryName = ChildDirectoryName.Trim();
                            try
                            {
                                MSC.ListItemCreationInformation information = new MSC.ListItemCreationInformation();
                                string targetFolder = ConfigurationManager.AppSettings["RootFolder"];
                                if (ConfigurationManager.AppSettings["Testing"] == "1")
                                    targetFolder = ConfigurationManager.AppSettings["RootFolderTest"]; ;
                                information.FolderUrl = list.RootFolder.ServerRelativeUrl + targetFolder + ParentDirectoryName;
                                MSC.Folder parentFolder = list.RootFolder.Folders.Add(information.FolderUrl);
                                context.Load(parentFolder);
                                context.ExecuteQuery();
                                information.FolderUrl = information.FolderUrl + "/" + ChildDirectoryName;

                                MSC.Folder childDirectory = list.RootFolder.Folders.Add(information.FolderUrl);
                                context.Load(childDirectory);
                                context.ExecuteQuery();


                            if (IsgoodLink)
                            {
                                string filePath = fileInstance.FullName;
                                FileStream documentStream = System.IO.File.OpenRead(filePath);
                                byte[] info = new byte[documentStream.Length];
                                documentStream.Read(info, 0, (int)documentStream.Length);
                                string fileURL = information.FolderUrl + "/" + fileInstance.Name;

                                MSC.FileCreationInformation fileCreationInformation = new MSC.FileCreationInformation();
                                fileCreationInformation.Overwrite = true;
                                fileCreationInformation.Content = info;
                                fileCreationInformation.Url = fileURL;
                                try
                                {
                                    Microsoft.SharePoint.Client.File f = context.Web.GetFileByServerRelativeUrl(fileURL);
                                    context.Load(f);
                                    context.ExecuteQuery();
                                    f.CheckOut();
                                }
                                catch (Microsoft.SharePoint.Client.ServerException ex)
                                {
                                    if (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                                    {
                                        Console.WriteLine("File is not found for Checkout");
                                    }
                                }
                                Microsoft.SharePoint.Client.File uploadFile = list.RootFolder.Files.Add(fileCreationInformation);


                                uploadFile.CheckIn("Improvement Plan", MSC.CheckinType.MajorCheckIn);

                                context.Load(uploadFile, w => w.MajorVersion, w => w.MinorVersion);
                                context.ExecuteQuery();
                                Console.WriteLine("Document {0} is uploaded and checked in into SharePoint", fileURL);
                            }

                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
                            }

                         }
                    }
                
            }
            catch (Exception ex)
            {
                new EventLog().WriteEntry(ex.Message, EventLogEntryType.Error);
                return;
            }
        }

       

        public bool CheckLink(String[] pathList, MSC.SharePointOnlineCredentials cred)
        {
            string baseSubPath1;
            string baseSubPath2;
            string documentLibrary;
            if((pathList[3]==":p:") ||(pathList[4]=="r"))
            {
                baseSubPath1 = pathList[5];
                baseSubPath2 = pathList[6];
                documentLibrary= pathList[7];
            }
            else
            {
                baseSubPath1 = pathList[3];
                baseSubPath2 = pathList[4];
                documentLibrary = pathList[5];
            }
            string baseURL = pathList[0] + "//" + pathList[2] + "/" + baseSubPath1 + "/" + baseSubPath2 + "/";
            
            bool isPathFound = false;
            using (ClientContext webInstance = new ClientContext(baseURL))
            {
                
                if (documentLibrary == "SitePages")
                    documentLibrary = "Site Pages";
                documentLibrary = documentLibrary.Replace("%20", " ");
                webInstance.Credentials = cred;
                List listLink = webInstance.Web.Lists.GetByTitle(documentLibrary);
                webInstance.Web.Context.Load(listLink);
                webInstance.Web.Context.Load(listLink.RootFolder);
                webInstance.Web.Context.Load(listLink.RootFolder.Folders);
                webInstance.Web.Context.Load(listLink.RootFolder.Files);
                webInstance.Web.Context.ExecuteQuery();
                FolderCollection fcol = listLink.RootFolder.Folders;
                List<string> lstFile = new List<string>();
                string filetobeChecked = pathList[pathList.Count() - 1];
                if(filetobeChecked.Contains('?'))
                filetobeChecked=filetobeChecked.Split('?')[0];
                filetobeChecked = filetobeChecked.Replace("%20", " ");
                //Check in the Root Path
                FileCollection fileCol = listLink.RootFolder.Files;
                foreach (MSC.File file in fileCol)
                {
                    if (file.Name == filetobeChecked)
                    {
                        isPathFound = true;
                        Console.WriteLine(" expected file {0}  found", filetobeChecked);
                        break;
                    }
                }

                if (!isPathFound)
                {
                    foreach (Folder f1 in fcol)
                    {
                        webInstance.Web.Context.Load(f1.Files);
                        webInstance.Web.Context.ExecuteQuery();
                        fileCol = f1.Files;
                        foreach (MSC.File file in fileCol)
                        {
                            if (file.Name == filetobeChecked)
                            {
                                isPathFound = true;
                                Console.WriteLine(" expected file {0}  found", filetobeChecked);
                                break;
                            }
                        }
                        if (isPathFound)
                        {
                            break;
                        }
                    }
                }

            }
            return isPathFound;
        }

        public void Dispose() { }

        public SecureString getPassword(string pwd)
        {
            SecureString secure = new SecureString();
            char[] cs = pwd.ToCharArray();
            foreach (char c in cs)
            {
                secure.AppendChar(c);
            }
            return secure;
        }

        public struct linkIdentifier
        {
           public int rowIndex;
           public int colIndex;
           public string SheetName;
            
        }

       

        public bool CheckForLinks(FileInfo fileTobeChecked,List<linkIdentifier> list)
        {
            bool IsgoodLink = false;
            string path = fileTobeChecked.FullName;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = excel.Workbooks.Open(path);
        

            //Read the first cell
            foreach (linkIdentifier identifier in list)
            {
                
                Worksheet excelSheet = wb.Sheets[identifier.SheetName];
                string test = excelSheet.Cells[identifier.rowIndex, identifier.colIndex].Formula;
                test = test.Remove(0, 12);
                test = test.Split(',')[0].TrimEnd("\"".ToCharArray());
                
                try
                {
                    if (test.Contains(".aspx"))
                    {
                        MSC.Folder isFolderExist = context.Web.GetFolderByServerRelativeUrl(test);
                       
                    }
                    else
                    {
                     
                        MSC.File isFileExist = context.Web.GetFileByServerRelativeUrl(test);
                        context.Load(isFileExist);
                        context.ExecuteQuery();
                        if (context.Web.GetFileByServerRelativeUrl(test).Exists)
                            IsgoodLink = true;
                    }
                }
                catch( MSC.ServerException e)
                {
                    if(e.ServerErrorTypeName== "System.IO.FileNotFoundException")
                    IsgoodLink = false;
                    
                }

                //using (var web = new SPSite("https://share.philips.com/sites/STS020170706072003/SitePages/Home.aspx").OpenWeb())
                //{
                //    var folderExists = CheckFolderExists(web, "Attachments/CMR_2000");
                //}

                //CookieContainer cookies = new CookieContainer();
                //HttpWebRequest webRequest = (HttpWebRequest)HttpWebRequest.Create(test);
                //webRequest.UserAgent = @"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.106 Safari/537.36";
                //webRequest.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials;
                //webRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                //webRequest.Method = "GET";
                //webRequest.CookieContainer = cookies;
                //HttpWebResponse _webResponse;

                //try
                //{
                //    _webResponse = (HttpWebResponse)webRequest.GetResponse();
                //    _webResponse.GetResponseStream();
                //    if (_webResponse.StatusCode == HttpStatusCode.OK)
                //        IsgoodLink = true;
                //}
                //catch (Exception e)
                //{
                //    IsgoodLink = false;
                //    return false; //could not connect to the internet (maybe) 
                //}
            }
            wb.Close();
            excel.ActiveWorkbook.Close();
            return IsgoodLink;
        }

       

    }

}
