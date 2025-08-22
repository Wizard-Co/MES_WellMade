using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Media;

namespace WizMes_WellMade
{
    class SettingINI
    {
        public static StringBuilder server = new StringBuilder();
        public static StringBuilder Database = new StringBuilder();

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal,
                                                        int size, string filePath);

        public static StringBuilder myFontSize = new StringBuilder();
        public static StringBuilder myFontFamily = new StringBuilder();
        public static StringBuilder myFontStyle = new StringBuilder();
        public static StringBuilder myFontWeight = new StringBuilder();

        public static StringBuilder myMainScale = new StringBuilder();
        public static StringBuilder myChildScale = new StringBuilder();

        public double setFontSize { get; set; }
        public FontFamily setFontFamily { get; set; }
        public FontStyle setFontStyle { get; set; }
        public FontWeight setFontWeight { get; set; }

        public double setMainScale { get; set; }
        public double setChildScale { get; set; }

        //설정값 얻기
        public void GetSettingINI()
        {

            //wizard.ini 파일 없으면
            if (!File.Exists("c:/windows/wizard.ini"))
            {
                WritePrivateProfileString("FontSetting", "FontSize", "15", "wizard.ini");
                WritePrivateProfileString("FontSetting", "FontFamily", "맑은고딕", "wizard.ini");
                WritePrivateProfileString("FontSetting", "FontStyle", "Normal", "wizard.ini");
                WritePrivateProfileString("FontSetting", "FontWeight", "Normal", "wizard.ini");
                MessageBox.Show("ini 파일이 없어서 생성하였습니다.");
                Environment.Exit(0);
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }

            //ini 읽기
            GetPrivateProfileString("FontSetting", "FontSize", "NOT_FOUND", myFontSize, 255, "wizard.ini");
            //GetPrivateProfileString("SQLServer", "server", "", retId, 256, "wizard.ini");
            GetPrivateProfileString("FontSetting", "FontFamily", "NOT_FOUND", myFontFamily, 256, "wizard.ini");
            GetPrivateProfileString("FontSetting", "FontStyle", "NOT_FOUND", myFontStyle, 256, "wizard.ini");
            GetPrivateProfileString("FontSetting", "FontWeight", "NOT_FOUND", myFontWeight, 256, "wizard.ini");

            GetPrivateProfileString("SizeScale", "MainScale", "NOT_FOUND", myMainScale, 256, "wizard.ini");
            GetPrivateProfileString("SizeScale", "ChildScale", "NOT_FOUND", myChildScale, 256, "wizard.ini");
            ConvertFontSetting();
        }

        //설정값 저장
        public void WriteSettingINI()
        {
            WritePrivateProfileString("FontSetting", "FontSize", myFontSize.ToString(), "wizard.ini");
            WritePrivateProfileString("FontSetting", "FontFamily", myFontFamily.ToString(), "wizard.ini");
            WritePrivateProfileString("FontSetting", "FontStyle", myFontStyle.ToString(), "wizard.ini");
            WritePrivateProfileString("FontSetting", "FontWeight", myFontWeight.ToString(), "wizard.ini");

            WritePrivateProfileString("SizeScale", "MainScale", myMainScale.ToString(), "wizard.ini");
            WritePrivateProfileString("SizeScale", "ChildScale", myChildScale.ToString(), "wizard.ini");
        }

        private void ConvertFontSetting()
        {
            if (myFontSize.ToString().Equals("NOT_FOUND"))
            {
                setFontSize = 15.0;
            }
            else
            {
                setFontSize = Convert.ToDouble(myFontSize.ToString());
            }

            if (myFontFamily.ToString().Equals("NOT_FOUND"))
            {
                setFontFamily = new FontFamily("맑은 고딕");
            }
            else
            {
                setFontFamily = new FontFamily(myFontFamily.ToString());
            }

            if (myFontStyle.ToString().Equals("NOT_FOUND"))
            {
                setFontStyle = FontStyles.Normal;
            }
            else
            {
                setFontStyle = SetFontStyle(myFontStyle.ToString());
            }

            if (myFontWeight.ToString().Equals("NOT_FOUND"))
            {
                setFontWeight = FontWeights.Normal;
            }
            else
            {
                setFontWeight = SetFontWeight(myFontWeight.ToString());
            }

            if (myMainScale.ToString().Equals("NOT_FOUND"))
            {
                setMainScale = 1.0;
            }
            else
            {
                setMainScale = Convert.ToDouble(myMainScale.ToString());
            }

            if (myChildScale.ToString().Equals("NOT_FOUND"))
            {
                setChildScale = 1.0;
            }
            else
            {
                setChildScale = Convert.ToDouble(myChildScale.ToString());
            }
        }

        private FontStyle SetFontStyle(string value)
        {
            FontStyle fontStyle;

            switch (value.ToString())
            {
                case "Italic":
                    fontStyle = FontStyles.Italic;
                    break;
                case "Normal":
                    fontStyle = FontStyles.Normal;
                    break;
                case "Oblique":
                    fontStyle = FontStyles.Oblique;
                    break;
                default:
                    fontStyle = FontStyles.Normal;
                    break;
            }

            return fontStyle;
        }

        private FontWeight SetFontWeight(string value)
        {
            FontWeight fontWeight;

            switch (value.ToString())
            {
                case "Bold":
                    fontWeight = FontWeights.Bold;
                    break;
                case "Normal":
                    fontWeight = FontWeights.Normal;
                    break;
                case "Light":
                    fontWeight = FontWeights.Light;
                    break;
                default:
                    fontWeight = FontWeights.Normal;
                    break;
            }

            return fontWeight;
        }
    }
}
