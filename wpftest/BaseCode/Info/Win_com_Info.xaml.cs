using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_info_Info_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_Info : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        #region 전역변수, 초기설정(ftp)

        string DateToday = DateTime.Today.ToString();
        string yyyyMMdd = string.Empty;     // DB활용 용도의 년월일.

        private FTP_EX _ftp = null;
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Info";
        //string FTP_ADDRESS = "ftp://192.168.0.4/Info";
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/Info";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";
        private List<UploadFileInfo> _listFileInfo = new List<UploadFileInfo>();

        internal struct UploadFileInfo          //FTP.
        {
            public string Filename { get; set; }
            public FtpFileType Type { get; set; }
            public DateTime LastModifiedTime { get; set; }
            public long Size { get; set; }
            public string Filepath { get; set; }
        }
        internal enum FtpFileType
        {
            None,
            DIR,
            File
        }

        #endregion

        public Win_com_Info()
        {
            InitializeComponent();
        }

        // 첫 로드시.
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);

            string yyyy = DateToday.Substring(0, 4);
            string MM = DateToday.Substring(5, 2);
            string dd = DateToday.Substring(8, 2);
            yyyyMMdd = yyyy + MM + dd;
            txtAllNotice.Text = " ▶ 오늘은 " + yyyy + "년 " + MM + "월 " + dd + "일 입니다.";       // 로드 기본설정

            SetFactoryPlace(); // 매출사업장. 콤보박스.

            // 매출사업장 기본 체크 되도록
            chkWorkplace.IsChecked = true;
            cboWorkplace.SelectedIndex = 0;

            Fill_UP_AllNotice();        // 오늘날짜의 전체용 공지 확인 + 첨부문서 있다면 그리드 표시. (로드시 자동)  
            Fill_UP_PersonNotice();     // 오늘날짜의 개별용 공지 확인. (로드시 자동)


        }


        #region 조회용 체크박스 세팅 
        //매출사업장
        private void lblWorkplace_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkWorkplace.IsChecked == true)
            {
                chkWorkplace.IsChecked = false;
            }
            else
            {
                chkWorkplace.IsChecked = true;
            }
        }
        //매출사업장
        private void chkWorkplace_Checked(object sender, RoutedEventArgs e)
        {
            cboWorkplace.IsEnabled = true;
        }
        //매출사업장
        private void chkWorkplace_Unchecked(object sender, RoutedEventArgs e)
        {
            cboWorkplace.IsEnabled = false;
        }

        #endregion


        #region 콤보박스 구성
        // 사업장 콤보박스 구성.
        private void SetFactoryPlace()
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Company_sCompanyInfo", sqlParameter, false);

                DataTable dt = null;
                dt = ds.Tables[0];
                if (dt.Rows.Count == 0)
                {
                    dt.Clear();
                    return;
                }
                else
                {
                    cboWorkplace.ItemsSource = dt.DefaultView;
                    this.cboWorkplace.DisplayMemberPath = "KCompany";
                    this.cboWorkplace.SelectedValuePath = "CompanyID";
                }
                //DataStore.Instance.CloseConnection(); 2021-09-13
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 콤보박스 세팅 : " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }
        #endregion


        #region 조회 // 상단조회 // 하단조회 // 첨부파일 ftp 자료조회

        // 전체공지용 오늘날짜 업로드.
        private void Fill_UP_AllNotice()
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sCompanyID", "");
                sqlParameter.Add("SDate", yyyyMMdd);
                sqlParameter.Add("EDate", yyyyMMdd);

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Info_sInfoByDate", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {
                        dt.Clear();
                        return;
                    }
                    else
                    {
                        dgAttachFile.Items.Clear();

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            txtAllNotice.Text = txtAllNotice.Text + "\r\n";

                            string ShowNotice = dt.Rows[i]["Info"].ToString();
                            txtAllNotice.Text = txtAllNotice.Text + ShowNotice;

                            ////////////////////////////////////////////////////////////////////////////
                            //텍스트 표시 하고, 첨부문서는 데이터그리드에 따로 add 하고,
                            ///////////////////////////////////////////////////////////////////////////
                            ///

                            if (dt.Rows[i]["PartFile"].ToString() != string.Empty)
                            {
                                var data = new GetDatagridItems { ColInfoID = dt.Rows[i]["InfoID"].ToString(), ColAttachFile = dt.Rows[i]["PartFile"].ToString(), ColAttachPath = dt.Rows[i]["PartPath"].ToString() };
                                dgAttachFile.Items.Add(data);
                            }
                            if (dt.Rows[i]["AttachFile1"].ToString() != string.Empty)
                            {
                                var data = new GetDatagridItems { ColInfoID = dt.Rows[i]["InfoID"].ToString(), ColAttachFile = dt.Rows[i]["AttachFile1"].ToString(), ColAttachPath = dt.Rows[i]["AttachPath1"].ToString() };
                                dgAttachFile.Items.Add(data);
                            }
                            if (dt.Rows[i]["AttachFile2"].ToString() != string.Empty)
                            {
                                var data = new GetDatagridItems { ColInfoID = dt.Rows[i]["InfoID"].ToString(), ColAttachFile = dt.Rows[i]["AttachFile2"].ToString(), ColAttachPath = dt.Rows[i]["AttachPath2"].ToString() };
                                dgAttachFile.Items.Add(data);
                            }
                            if (dt.Rows[i]["AttachFile3"].ToString() != string.Empty)
                            {
                                var data = new GetDatagridItems { ColInfoID = dt.Rows[i]["InfoID"].ToString(), ColAttachFile = dt.Rows[i]["AttachFile3"].ToString(), ColAttachPath = dt.Rows[i]["AttachPath3"].ToString() };
                                dgAttachFile.Items.Add(data);
                            }
                        }
                    }
                    dt.Clear();
                }
                ds.Clear();
                //DataStore.Instance.CloseConnection(); 2021-09-13
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 전체공지 조회" + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }


        //개별공지
        private void Fill_UP_PersonNotice()
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("Date", yyyyMMdd);
                sqlParameter.Add("UserID", MainWindow.CurrentUser);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Info_sInfoUserByUserID", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {
                        dt.Clear();
                        return;
                    }
                    else
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            string ShowNotice = dt.Rows[i]["Info"].ToString();
                            txtPersonNotice.Text = ShowNotice;
                            txtPersonNotice.Text = txtPersonNotice.Text + "\r\n";
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 개별공지 조회 : " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }


        #endregion


        #region 첨부파일 데이터 그리드 내 다운로드 버튼 클릭

        // 그리드 내 첨부파일 다운로드 버튼 클릭.  ★
        private void btn_GridAttach_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 다운로드 하시겠습니까?", "다운로드 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                try
                {
                    Button btnSender = sender as Button;
                    GetDatagridItems senderAllNotice = btnSender.DataContext as GetDatagridItems;

                    string FilePath = senderAllNotice.ColInfoID.Trim();
                    string FileName = senderAllNotice.ColAttachFile.Trim();

                    string str_path = FTP_ADDRESS + '/' + FilePath;     //풀 경로.

                    _ftp = new FTP_EX(str_path, FTP_ID, FTP_PASS);

                    string str_remotepath = FileName;
                    string str_localpath = str_localpath = LOCAL_DOWN_PATH + "\\" + FileName;

                    DirectoryInfo DI = new DirectoryInfo(LOCAL_DOWN_PATH);      // Temp 폴더가 없는 컴터라면, 만들어 줘야지.
                    if (DI.Exists == false)
                    {
                        DI.Create();
                    }

                    FileInfo file = new FileInfo(str_localpath);
                    if (file.Exists)
                    {
                        file.Delete();
                    }

                    _ftp.download(str_remotepath, str_localpath);

                    ProcessStartInfo proc = new ProcessStartInfo(str_localpath);
                    proc.UseShellExecute = true;
                    Process.Start(proc);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("파일이 존재하지 않습니다.\r관리자에게 문의해주세요.");
                    return;
                }
            }
        }



        private bool MakeFileInfoList(string[] simple, string[] detail, string str_NoticeID)
        {
            bool tf_return = false;
            foreach (string filename in simple)
            {
                foreach (string info in detail)
                {
                    if (info.Contains(filename) == true)
                    {

                        if (MakeFileInfoList(filename, info, str_NoticeID) == true)
                        {

                            tf_return = true;
                        }
                    }
                }
            }
            return tf_return;
        }

        private bool MakeFileInfoList(string simple, string detail, string strCompare)
        {
            UploadFileInfo info = new UploadFileInfo();
            info.Filename = simple;
            info.Filepath = detail;

            if (simple.Length > 0)
            {
                string[] tokens = detail.Split(new[] { ' ' }, 9, StringSplitOptions.RemoveEmptyEntries);
                string name = tokens[3].ToString();         // 2017.03.16  허윤구.  토근 배열이 8자리로 되어 있었는데 에러가 나길래 확인해 보니 4자리 배열로 나오길래 바꾸었습니다.
                string permissions = tokens[2].ToString();      // premission도 배열 0번이 아니라 배열 2번인데...;;


                if (permissions.Contains("D") == true)          // 대문자 D로 표시해야 합니다.
                {
                    info.Type = FtpFileType.DIR;
                }
                else
                {
                    info.Type = FtpFileType.File;
                }

                if (info.Type == FtpFileType.File)
                {
                    info.Size = Convert.ToInt64(detail.Substring(17, detail.LastIndexOf(simple) - 17).Trim());      // 사이즈가 중요한가?
                }

                _listFileInfo.Add(info);

                if (string.Compare(simple, strCompare, false) == 0)
                    return true;
            }
            return false;
        }


        #endregion


        #region 닫기버튼

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }
        #endregion


        #region 새로고침
        // 새로고침.
        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            string yyyy = DateToday.Substring(0, 4);
            string MM = DateToday.Substring(5, 2);
            string dd = DateToday.Substring(8, 2);

            txtAllNotice.Text = string.Empty;
            txtAllNotice.Text = " ▶ 오늘은 " + yyyy + "년 " + MM + "월 " + dd + "일 입니다.";

            txtPersonNotice.Text = string.Empty;

            Fill_UP_AllNotice();
            Fill_UP_PersonNotice();
        }

        #endregion


    }


    public class GetDatagridItems
    {
        public string ColInfoID { get; set; }
        public string ColAttachFile { get; set; }
        public string ColAttachPath { get; set; }
    }

}

