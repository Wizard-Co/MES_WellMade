using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;

namespace WizMes_WellMade
{
    class BookMarkINI
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal,
                                                        int size, string filePath);

        public static StringBuilder myBookMarkMenu = new StringBuilder();
        public string[] strBookMarkMenu;

        //설정값 얻기
        public void GetBookMarkINI()
        {

            //wizard.ini 파일 없으면
            if (!File.Exists("c:/windows/wizard.ini"))
            {
                WritePrivateProfileString("BookMark", "BookMarkMenuQ", "", "wizard.ini");
                MessageBox.Show("ini 파일이 없어서 생성하였습니다.");
                Environment.Exit(0);
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }

            //ini 읽기
            GetPrivateProfileString("BookMark", "BookMarkMenuQ", "NOT_FOUND", myBookMarkMenu, 255, "wizard.ini");
            ConvertFontSetting();
        }

        //설정값 저장
        public void WriteBookMarkINI()
        {
            WritePrivateProfileString("BookMark", "BookMarkMenuQ", myBookMarkMenu.ToString(), "wizard.ini");
        }

        private void ConvertFontSetting()
        {
            if (!myBookMarkMenu.ToString().Equals("NOT_FOUND"))
            {
                strBookMarkMenu = myBookMarkMenu.ToString().Split('/');
            }
        }
    }
}
