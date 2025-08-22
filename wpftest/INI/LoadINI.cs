using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;

namespace WizMes_WellMade
{
    public class LoadINI
    {
        public static StringBuilder server = new StringBuilder();

        public static StringBuilder Database = new StringBuilder();

        public static StringBuilder FTPPort = new StringBuilder();
        public static StringBuilder FileSvr = new StringBuilder();
        public static StringBuilder FtpImagePath = new StringBuilder();

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal,
                                                        int size, string filePath);

        public void loadINI()
        {
            //wizard.ini 파일 없으면
            if (!File.Exists("c:/windows/wizard.ini"))
            {
                WritePrivateProfileString("SQLServer", "server", "AFTServer,20140", "wizard.ini");
                WritePrivateProfileString("SQLServer", "database", "MES_nAFT ", "wizard.ini");
                MessageBox.Show("ini 파일이 없어서 생성하였습니다.. DB 연결 실패");
                Environment.Exit(0);
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }

            //ini 읽기
            GetPrivateProfileString("SQLServer", "server", "NOT_FOUND", server, 255, "wizard.ini");
            //GetPrivateProfileString("SQLServer", "server", "", retId, 256, "wizard.ini");
            GetPrivateProfileString("SQLServer", "Database", "NOT_FOUND", Database, 256, "wizard.ini");

            //FTP정보 얻기
            GetPrivateProfileString("FTPINFO", "FTPPort", "NOT_FOUND", FTPPort, 255, "wizard.ini");
            GetPrivateProfileString("FTPINFO", "FileSvr", "NOT_FOUND", FileSvr, 256, "wizard.ini");
            GetPrivateProfileString("FTPINFO", "FTPIMAGEPATH", "NOT_FOUND", FtpImagePath, 256, "wizard.ini");
        }



    }
}
