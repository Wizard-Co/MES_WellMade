using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Windows.Media.Imaging;

//*******************************************************************************
//프로그램명    FTP_EX.cs
//메뉴ID        
//설명          공통 메서드
//작성일        2016.12.06
//개발자        김은미
//*******************************************************************************
// 변경일자     변경자      요청자      요구사항ID          요청 및 작업내용
//*******************************************************************************
//
//
//*******************************************************************************

namespace WizMes_WellMade
{
    public class FTP_EX
    {
        private string host = null;
        private string user = null;
        private string pass = null;

        private FtpWebRequest ftpRequest = null;
        private FtpWebResponse ftpResponse = null;
        private Stream ftpStream = null;
        private int bufferSize = 2048;

        /* Construct Object */
        public FTP_EX(string hostIP, string userName, string password) { host = hostIP; user = userName; pass = password; }

        // 20191210 둘리 - 테스트
        public byte[] getFileToByte(string remoteFile)
        {
            byte[] file = null;

            try
            {
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + remoteFile);
                ftpRequest.Credentials = new NetworkCredential(user, pass);

                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                Stream responseStream = ftpResponse.GetResponseStream();

                file = ToByteArray(responseStream);

                responseStream.Close();
            }
            catch (Exception ex9)
            {
                System.Windows.MessageBox.Show("바이트 변환 실패임여");
            }

            return file;
        }
        public Byte[] ToByteArray(Stream stream)
        {
            MemoryStream ms = new MemoryStream();
            byte[] chunk = new byte[4096];
            int bytesRead;
            while ((bytesRead = stream.Read(chunk, 0, chunk.Length)) > 0)
            {
                ms.Write(chunk, 0, bytesRead);
            }

            return ms.ToArray();
        }
        // 20191210 둘리 - 파일 업로드 테스트
        public bool UploadUsingByte(byte[] buffer, string FileName)
        {
            bool flag = false;

            try
            {
                FtpWebRequest ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + FileName);
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;

                Stream reqStream = ftpRequest.GetRequestStream();
                reqStream.Write(buffer, 0, buffer.Length);
                reqStream.Close();
            }
            catch (Exception ex9)
            {
                System.Windows.MessageBox.Show("저장 실패임여");
            }

            return flag;
        }
        // 20191211 둘리 - 파일 업로드 - 리스트로 넣어봅세
        public bool UploadUsingByteList(Dictionary<string, byte[]> lstFileByte, string FolderName)
        {
            bool flag = false;

            try
            {
                foreach (string FileName in lstFileByte.Keys)
                {
                    FtpWebRequest ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + FolderName + "/" + FileName);
                    ftpRequest.Credentials = new NetworkCredential(user, pass);
                    ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;

                    Stream reqStream = ftpRequest.GetRequestStream();
                    reqStream.Write(lstFileByte[FileName], 0, lstFileByte[FileName].Length);
                    reqStream.Close();
                }
            }
            catch (Exception ex9)
            {
                System.Windows.MessageBox.Show("저장 실패임여");
            }

            return flag;
        }

        // 20191211 둘리 → 요거는 FTP Server File Path 를 받아서, 그걸 다운로드 한 다음에 그걸로다가 바로 업로드 시켜버리는 것이여
        public bool UploadUsingFtpServerFilePath(Dictionary<string, string> lstFtpFilePath, string FolderName)
        {
            bool flag = false;

            try
            {
                foreach (string FileName in lstFtpFilePath.Keys)
                {
                    // 일단 다운로드 먼저 받아야 쓰것어
                    WebClient ftpClient = new WebClient();
                    ftpClient.Credentials = new NetworkCredential(user, pass);
                    byte[] imageByte = ftpClient.DownloadData(host + "/" + lstFtpFilePath[FileName] + "/" + FileName);

                    // 이것은 업로드여
                    FtpWebRequest ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + FolderName + "/" + FileName);
                    ftpRequest.Credentials = new NetworkCredential(user, pass);
                    ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;

                    Stream reqStream = ftpRequest.GetRequestStream();
                    reqStream.Write(imageByte, 0, imageByte.Length);
                    ftpRequest = null;
                    reqStream.Close();
                }
            }
            catch (Exception ex9)
            {
                System.Windows.MessageBox.Show("저장 실패임여");
            }

            return flag;
        }



        /* Download File */
        public bool download(string remoteFile, string localFile, bool showMsg = true)
        {
            try
            {
                /* Create an FTP Request */
                //string url = HttpUtility.UrlEncode("#");
                //url = remoteFile.Replace("#", url);

                //ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + url);
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + remoteFile);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Get the FTP Server's Response Stream */
                ftpStream = ftpResponse.GetResponseStream();
                /* Open a File Stream to Write the Downloaded File */
                FileStream localFileStream = new FileStream(localFile, FileMode.Create);
                /* Buffer for the Downloaded Data */
                byte[] byteBuffer = new byte[bufferSize];
                int bytesRead = ftpStream.Read(byteBuffer, 0, bufferSize);
                /* Download the File by Writing the Buffered Data Until the Transfer is Complete */
                try
                {
                    while (bytesRead > 0)
                    {
                        localFileStream.Write(byteBuffer, 0, bytesRead);
                        bytesRead = ftpStream.Read(byteBuffer, 0, bufferSize);
                    }
                }
                catch (Exception ex)
                {
                    if (showMsg == true) System.Windows.MessageBox.Show("1" + ex.Message + " / " + ex.Source);
                    throw;
                }
                /* Resource Cleanup */
                localFileStream.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                return true;
            }
            catch (Exception ex)
            {
                if (showMsg == true) System.Windows.MessageBox.Show("2" + ex.Message + " / " + ex.Source);
                return false;
                //throw ex;
            }
        }

        /* Upload File */
        public bool upload(string remoteFile, string localFile)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + remoteFile);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;
                /* Establish Return Communication with the FTP Server */
                ftpStream = ftpRequest.GetRequestStream();
                /* Open a File Stream to Read the File for Upload */
                FileStream localFileStream = new FileStream(localFile, FileMode.Open);
                /* Buffer for the Upload Data */
                byte[] byteBuffer = new byte[bufferSize];
                int bytesSent = localFileStream.Read(byteBuffer, 0, bufferSize);
                /* Upload the File by Sending the Buffered Data Until the Transfer is Complete */
                try
                {
                    while (bytesSent != 0)
                    {
                        ftpStream.Write(byteBuffer, 0, bytesSent);
                        bytesSent = localFileStream.Read(byteBuffer, 0, bufferSize);
                    }
                }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                localFileStream.Close();
                ftpStream.Close();
                ftpRequest = null;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            return false;
        }


        /* Upload Files */
        public bool uploads(string remoteFile, List<string> strName, List<string> localFile)
        {
            try
            {
                for (int i = 0; i < strName.Count; i++)
                {
                    /* Create an FTP Request */
                    ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + remoteFile + strName[i]);
                    /* When in doubt, use these options */
                    ftpRequest.UseBinary = true;
                    ftpRequest.UsePassive = true;
                    ftpRequest.Timeout = -1;
                    ftpRequest.KeepAlive = false;
                    /* Log in to the FTP Server with the User Name and Password Provided */
                    ftpRequest.Credentials = new NetworkCredential(user, pass);
                    /* Specify the Type of FTP Request */
                    ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;
                    /* Open a File Stream to Read the File for Upload */
                    FileStream fs = File.OpenRead(localFile[i] + strName[i]);
                    byte[] buffer = new byte[fs.Length];
                    fs.Read(buffer, 0, buffer.Length);
                    fs.Close();
                    Stream requestStr = ftpRequest.GetRequestStream();
                    requestStr.Write(buffer, 0, buffer.Length);
                    requestStr.Close();
                    requestStr.Flush();
                    FtpWebResponse response = (FtpWebResponse)ftpRequest.GetResponse();
                    response.Close();
                    //File.Delete(localFile[i] + strName[i]);
                }
                ftpRequest = null;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            return false;
        }



        /// <summary>
        /// upload 할 file의 name and Full Path를 string 배열에 넣는 방식으로 
        /// upload할 모든 string 배열을 List에 담아주면 모두 업로드
        /// </summary>
        public bool UploadTempFilesToFTP(List<string[]> remoteFiles)
        {
            List<string[]> fileList;
            try
            {
                fileList = remoteFiles;
                for (int i = 0; i < fileList.Count; i++)
                {
                    string FileName = fileList[i][0];
                    string localFile = fileList[i][1];
                    FtpWebRequest ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + FileName.Replace("#", "%23"));
                    ftpRequest.Credentials = new NetworkCredential(user, pass);
                    ftpRequest.UseBinary = true;
                    ftpRequest.UsePassive = true;
                    if (i == fileList.Count - 1)
                    {
                        ftpRequest.KeepAlive = true;
                    }
                    else
                    {
                        ftpRequest.KeepAlive = false;
                    }

                    ftpRequest.Timeout = -1;

                    ftpRequest.Method = WebRequestMethods.Ftp.UploadFile;
                    string file = FileName.Substring(FileName.IndexOf('/') + 1);
                    FileStream fs = File.OpenRead(localFile + file);
                    byte[] buffer = new byte[fs.Length];
                    fs.Read(buffer, 0, buffer.Length);
                    fs.Close();
                    Stream requestStr = ftpRequest.GetRequestStream();
                    requestStr.Write(buffer, 0, buffer.Length);
                    requestStr.Close();
                    requestStr.Flush();
                    FtpWebResponse response = (FtpWebResponse)ftpRequest.GetResponse();
                    response.Close();
                }
                Console.WriteLine("Uploaded Successfully to Temp folder");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Upload failed. {0}", ex.Message);
                return false;
            }

        }

        /* Delete File */
        public bool delete(string deleteFile)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)WebRequest.Create(host + "/" + deleteFile);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.DeleteFile;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Resource Cleanup */
                ftpResponse.Close();
                ftpRequest = null;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            return false;
        }

        //2016.12.05.kem
        /* Remove Directory */
        public bool removeDir(string deleteDirectory)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)WebRequest.Create(host + "/" + deleteDirectory);
                ftpRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;

                List<string> lines = new List<string>();

                using (FtpWebResponse listResponse = (FtpWebResponse)ftpRequest.GetResponse())
                using (Stream listStream = listResponse.GetResponseStream())
                using (StreamReader listReader = new StreamReader(listStream))
                {
                    while (!listReader.EndOfStream)
                    {
                        lines.Add(listReader.ReadLine());
                    }
                }

                foreach (string line in lines)
                {
                    string[] tokens = line.Split(new[] { ' ' }, 9, StringSplitOptions.RemoveEmptyEntries);
                    string name = tokens[3].ToString(); //배열이 8이었다 3으로 바꿈
                    string permissions = tokens[0].ToString();

                    string fileUrl = host + "/" + deleteDirectory + "/" + name;

                    FtpWebRequest deleteRequest = (FtpWebRequest)WebRequest.Create(fileUrl);
                    deleteRequest.Method = WebRequestMethods.Ftp.DeleteFile;
                    deleteRequest.Credentials = new NetworkCredential(user, pass);

                    deleteRequest.GetResponse();
                }

                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)WebRequest.Create(host + "/" + deleteDirectory);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.RemoveDirectory;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Resource Cleanup */
                ftpResponse.Close();
                ftpRequest = null;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            return false;
        }

        /* Rename File */
        public bool rename(string currentFileNameAndPath, string newFileName)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)WebRequest.Create(host + "/" + currentFileNameAndPath);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.Rename;
                /* Rename the File */
                ftpRequest.RenameTo = newFileName;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Resource Cleanup */
                ftpResponse.Close();
                ftpRequest = null;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            return false;
        }

        ///* Create a New Directory on the FTP Server */
        //public bool createDirectory(string newDirectory)
        //{
        //    try
        //    {
        //        /* Create an FTP Request */
        //        ftpRequest = (FtpWebRequest)WebRequest.Create(host + "/" + newDirectory);
        //        /* Log in to the FTP Server with the User Name and Password Provided */
        //        ftpRequest.Credentials = new NetworkCredential(user, pass);
        //        /* When in doubt, use these options */
        //        ftpRequest.UseBinary = true;
        //        ftpRequest.UsePassive = true;
        //        ftpRequest.KeepAlive = true;
        //        /* Specify the Type of FTP Request */
        //        ftpRequest.Method = WebRequestMethods.Ftp.MakeDirectory;
        //        /* Establish Return Communication with the FTP Server */
        //        ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
        //        /* Resource Cleanup */
        //        ftpResponse.Close();
        //        ftpRequest = null;
        //        return true;
        //    }
        //    catch (Exception ex) { Console.WriteLine(ex.ToString()); }
        //    return false;
        //}

        public bool createDirectory(string newDirectory, bool createParent = true)
        {
            try
            {
                // 디렉토리가 이미 존재하는지 확인
                if (directoryExists(newDirectory))
                {
                    return true; // 이미 존재하면 성공으로 간주
                }

                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)WebRequest.Create(host + "/" + newDirectory);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.MakeDirectory;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Resource Cleanup */
                ftpResponse.Close();
                ftpRequest = null;
                return true;
            }
            catch (WebException ex)
            {
                FtpWebResponse response = (FtpWebResponse)ex.Response;
                if (response.StatusCode == FtpStatusCode.ActionNotTakenFileUnavailable && createParent)
                {
                    /* 상위 디렉토리를 만들어야 하는 경우 */
                    if (createParentDirectories())
                    {
                        /*상위 폴더 생성 성공 시, 다시 현재 폴더 생성 시도 (재귀 호출 한 번만) */
                        return createDirectory(newDirectory, false); // false로 설정하여 재귀 호출 제한
                    }
                }
                else if (response.StatusCode == FtpStatusCode.ActionNotTakenFileUnavailable)
                {
                    // 이미 createParent=false인 상태에서 재시도 실패
                    Console.WriteLine($"디렉토리 생성 실패: {newDirectory}");
                }
                Console.WriteLine($"FTP Error: {response.StatusCode} - {response.StatusDescription}");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return false;
            }
        }

        // 디렉토리 존재 여부 확인 메서드 추가
        private bool directoryExists(string directory)
        {
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(host + "/" + directory);
                request.Credentials = new NetworkCredential(user, pass);
                request.Method = WebRequestMethods.Ftp.ListDirectory;

                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                {
                    return true; // 폴더가 존재하고 접근 가능
                }
            }
            catch
            {
                return false; // 폴더가 존재하지 않거나 접근 불가
            }
        }


        private bool createParentDirectories()
        {
            try
            {
                // 프로토콜 이후의 시작 위치를 찾습니다
                int protocolIndex = host.IndexOf("://");
                // 포트 시작 위치를 찾습니다
                int colonIndex = host.IndexOf(":", protocolIndex + 3);
                // 포트 다음의 첫 슬래시 위치를 찾습니다
                int slashIndex = host.IndexOf("/", colonIndex > -1 ? colonIndex : protocolIndex + 3);

                if (slashIndex == -1) return true; // 기본 경로가 없는 경우

                // host에서 기본 경로 추출 
                string basePath = host.Substring(slashIndex + 1);
                string[] pathParts = basePath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

                // 상위 경로들 순차적으로 생성
                string currentPath = host.Substring(0, slashIndex + 1);

                foreach (string part in pathParts)
                {
                    currentPath += part + "/";
                    try
                    {
                        FtpWebRequest request = (FtpWebRequest)WebRequest.Create(currentPath.TrimEnd('/'));
                        request.Method = WebRequestMethods.Ftp.MakeDirectory;
                        request.Credentials = new NetworkCredential(user, pass);
                        request.UseBinary = true;
                        request.UsePassive = true;
                        request.KeepAlive = true;

                        using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                        {
                            // 폴더가 성공적으로 생성되거나 이미 존재하는 경우
                        }
                    }
                    catch (WebException ex)
                    {
                        // 폴더가 이미 존재하는 경우는 무시하고 계속 진행
                        FtpWebResponse response = (FtpWebResponse)ex.Response;
                        if (response.StatusCode != FtpStatusCode.ActionNotTakenFileUnavailable)
                        {
                            Console.WriteLine($"Error creating directory {currentPath}: {response.StatusDescription}");
                            return false;
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in createParentDirectories: {ex.Message}");
                return false;
            }
        }



        /* Get the Date/Time a File was Created */
        public string getFileCreatedDateTime(string fileName)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + fileName);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.GetDateTimestamp;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Establish Return Communication with the FTP Server */
                ftpStream = ftpResponse.GetResponseStream();
                /* Get the FTP Server's Response Stream */
                StreamReader ftpReader = new StreamReader(ftpStream);
                /* Store the Raw Response */
                string fileInfo = null;
                /* Read the Full Response Stream */
                try { fileInfo = ftpReader.ReadToEnd(); }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                ftpReader.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                /* Return File Created Date Time */
                return fileInfo;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            /* Return an Empty string Array if an Exception Occurs */
            return "";
        }

        /* Get the Size of a File */
        public string getFileSize(string fileName)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + fileName);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.GetFileSize;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Establish Return Communication with the FTP Server */
                ftpStream = ftpResponse.GetResponseStream();
                /* Get the FTP Server's Response Stream */
                StreamReader ftpReader = new StreamReader(ftpStream);
                /* Store the Raw Response */
                string fileInfo = null;
                /* Read the Full Response Stream */
                try { while (ftpReader.Peek() != -1) { fileInfo = ftpReader.ReadToEnd(); } }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                ftpReader.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                /* Return File Size */
                return fileInfo;
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            /* Return an Empty string Array if an Exception Occurs */
            return "";
        }

        /* List Directory Contents File/Folder Name Only */
        public string[] directoryListSimple(string directory)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + directory);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Establish Return Communication with the FTP Server */
                ftpStream = ftpResponse.GetResponseStream();
                /* Get the FTP Server's Response Stream */
                StreamReader ftpReader = new StreamReader(ftpStream);
                /* Store the Raw Response */
                string directoryRaw = null;
                /* Read Each Line of the Response and Append a Pipe to Each Line for Easy Parsing */
                try { while (ftpReader.Peek() != -1) { directoryRaw += ftpReader.ReadLine() + "|"; } }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                ftpReader.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                /* Return the Directory Listing as a string Array by Parsing 'directoryRaw' with the Delimiter you Append (I use | in This Example) */
                try { string[] directoryList = directoryRaw.Split("|".ToCharArray()); return directoryList; }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            /* Return an Empty string Array if an Exception Occurs */
            return new string[] { "" };
        }

        /* List Directory Contents in Detail (Name, Size, Created, etc.) */
        public string[] directoryListDetailed(string directory)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + directory);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Establish Return Communication with the FTP Server */
                ftpStream = ftpResponse.GetResponseStream();
                /* Get the FTP Server's Response Stream */
                StreamReader ftpReader = new StreamReader(ftpStream);
                /* Store the Raw Response */
                string directoryRaw = null;
                /* Read Each Line of the Response and Append a Pipe to Each Line for Easy Parsing */
                try { while (ftpReader.Peek() != -1) { directoryRaw += ftpReader.ReadLine() + "|"; } }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                ftpReader.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                /* Return the Directory Listing as a string Array by Parsing 'directoryRaw' with the Delimiter you Append (I use | in This Example) */
                try { string[] directoryList = directoryRaw.Split("|".ToCharArray()); return directoryList; }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            /* Return an Empty string Array if an Exception Occurs */
            return new string[] { "" };
        }

        /* List Directory Contents File/Folder Name Only */
        public string[] directoryListSimple(string directory, Encoding encoding)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + directory);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Establish Return Communication with the FTP Server */
                ftpStream = ftpResponse.GetResponseStream();
                /* Get the FTP Server's Response Stream */
                StreamReader ftpReader = new StreamReader(ftpStream, encoding);
                /* Store the Raw Response */
                string directoryRaw = null;
                /* Read Each Line of the Response and Append a Pipe to Each Line for Easy Parsing */
                try { while (ftpReader.Peek() != -1) { directoryRaw += ftpReader.ReadLine() + "|"; } }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                ftpReader.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                /* Return the Directory Listing as a string Array by Parsing 'directoryRaw' with the Delimiter you Append (I use | in This Example) */
                try { string[] directoryList = directoryRaw.Split("|".ToCharArray()); return directoryList; }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            /* Return an Empty string Array if an Exception Occurs */
            return new string[] { "" };
        }

        /* List Directory Contents in Detail (Name, Size, Created, etc.) */
        public string[] directoryListDetailed(string directory, Encoding encoding)
        {
            try
            {
                /* Create an FTP Request */
                ftpRequest = (FtpWebRequest)FtpWebRequest.Create(host + "/" + directory);
                /* Log in to the FTP Server with the User Name and Password Provided */
                ftpRequest.Credentials = new NetworkCredential(user, pass);
                /* When in doubt, use these options */
                ftpRequest.UseBinary = true;
                ftpRequest.UsePassive = true;
                ftpRequest.KeepAlive = true;
                /* Specify the Type of FTP Request */
                ftpRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                /* Establish Return Communication with the FTP Server */
                ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                /* Establish Return Communication with the FTP Server */
                ftpStream = ftpResponse.GetResponseStream();
                /* Get the FTP Server's Response Stream */
                StreamReader ftpReader = new StreamReader(ftpStream, encoding);
                /* Store the Raw Response */
                string directoryRaw = null;
                /* Read Each Line of the Response and Append a Pipe to Each Line for Easy Parsing */
                try { while (ftpReader.Peek() != -1) { directoryRaw += ftpReader.ReadLine() + "|"; } }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
                /* Resource Cleanup */
                ftpReader.Close();
                ftpStream.Close();
                ftpResponse.Close();
                ftpRequest = null;
                /* Return the Directory Listing as a string Array by Parsing 'directoryRaw' with the Delimiter you Append (I use | in This Example) */
                try { string[] directoryList = directoryRaw.Split("|".ToCharArray()); return directoryList; }
                catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            }
            catch (Exception ex) { Console.WriteLine(ex.ToString()); }
            /* Return an Empty string Array if an Exception Occurs */
            return new string[] { "" };
        }

        /// <summary>
        /// 그냥 삭제가능으로 가보자
        /// </summary>
        /// <param name="serverUri"></param>
        /// <param name="ftpUsername"></param>
        /// <param name="ftpPassword"></param>
        /// <returns></returns>
        public bool DeleteFileOnFtpServer(Uri serverUri)
        {
            try
            {
                // The serverUri parameter should use the ftp:// scheme.
                // It contains the name of the server file that is to be deleted.
                // Example: ftp://contoso.com/someFile.txt.
                // 

                if (serverUri.Scheme != Uri.UriSchemeFtp)
                {
                    return false;
                }
                // Get the object used to communicate with the server.
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(serverUri);
                request.Credentials = new NetworkCredential(user, pass);
                request.Method = WebRequestMethods.Ftp.DeleteFile;

                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                //Console.WriteLine("Delete status: {0}", response.StatusDescription);
                response.Close();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// ftp경로를 가지고 Bitmap 정보 리턴한다
        /// </summary>
        /// <param name="ftpFilePath"></param>
        /// <returns></returns>
        public BitmapImage DrawingImageByByte(string ftpFilePath)
        {
            BitmapImage image = new BitmapImage();

            try
            {
                WebClient ftpClient = new WebClient();
                ftpClient.Credentials = new NetworkCredential(user, pass);
                byte[] imageByte = ftpClient.DownloadData(ftpFilePath);

                //MemoryStream mStream = new MemoryStream();
                //mStream.Write(imageByte, 0, Convert.ToInt32(imageByte.Length));

                using (MemoryStream stream = new MemoryStream(imageByte))
                {
                    image.BeginInit();
                    image.StreamSource = stream;
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.EndInit();
                    image.Freeze();
                }

            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("1" + ex.Message + " / " + ex.Source);
                //throw ex;
            }

            return image;
        }
    }
}
