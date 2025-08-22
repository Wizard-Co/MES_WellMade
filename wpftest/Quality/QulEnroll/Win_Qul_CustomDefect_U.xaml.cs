using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_CustomDefect_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_CustomDefect_U : UserControl
    {
        #region 전역변수 설정

        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        // 추가저장인지 / 수정저장인지 구별하는 용도입니다.
        string ButtonTag = string.Empty;

        int Wh_Ar_SelectedLastIndex = 0;        // 그리드 마지막 선택 줄 임시저장 그릇

        // FTP 활용모음.
        private FTP_EX _ftp = null;
        List<string[]> listFtpFile = new List<string[]>();



        string FullPath1 = string.Empty;
        string FullPath2 = string.Empty;
        string FullPath3 = string.Empty;

        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/CustomDefect";
        //string FTP_ADDRESS = "ftp://HKserver:210/ImageData/CustomDefect";

        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/CustomDefect";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";

        #endregion

        public Win_Qul_CustomDefect_U()
        {
            InitializeComponent();
        }

        // 첫 로드시.
        private void Win_Qul_CustomDefect_U_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            First_Step();
            ComboBoxSetting();
        }


        #region 첫 스텝 // 날짜용 버튼 // 조회용 체크박스 세팅 // 그룹박스 날짜 체크박스 세팅
        // 첫 스텝
        private void First_Step()
        {
            //chkSearchDay.IsChecked = true;
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            rbnOccurDay.IsChecked = true;
            rbnChoice1.IsChecked = true;

            txtCustomer.IsEnabled = false;
            btnCustomer.IsEnabled = false;
            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
            txtArticleNo.IsEnabled = false;
            btnArticleNo.IsEnabled = false;

            txtDefectContent.IsEnabled = false;
            txtDefectNO.IsEnabled = false;
            cboCloseYN.IsEnabled = false;

            grbBasisInfoBox.IsEnabled = false;
            grbDefectInfoBox.IsEnabled = false;
            grbAttachBox.IsEnabled = false;
            grbReceiptBox.IsEnabled = false;
            grbActionBox.IsEnabled = false;
            cboCloseYN_InGroupBox.IsEnabled = false;
            cboDvlYN_InGroupBox.IsEnabled = false;

            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            EventLabel.Visibility = Visibility.Hidden;
        }

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpToDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = DateTime.Today;
            dtpToDate.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpFromDate.SelectedDate.Value);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpToDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }

        //검사일자
        private void chkSearchDay_Click(object sender, RoutedEventArgs e)
        {
            if (chkSearchDay.IsChecked == true)
            {
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
            else
            {
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }

        }
        //검사일자
        private void chkSearchDay_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkSearchDay.IsChecked == true)
            {
                chkSearchDay.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkSearchDay.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }

        //거래처
        private void chkCustomer_Click(object sender, RoutedEventArgs e)
        {
            if (chkCustomer.IsChecked == true)
            {
                txtCustomer.IsEnabled = true;
                txtCustomer.Focus();
                btnCustomer.IsEnabled = true;
            }
            else
            {
                txtCustomer.IsEnabled = false;
                btnCustomer.IsEnabled = false;
            }

        }
        //거래처
        private void chkCustomer_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomer.IsChecked == true)
            {
                chkCustomer.IsChecked = false;
                txtCustomer.IsEnabled = false;
                btnCustomer.IsEnabled = false;
            }
            else
            {
                chkCustomer.IsChecked = true;
                txtCustomer.IsEnabled = true;
                txtCustomer.Focus();
                btnCustomer.IsEnabled = true;
            }
        }
        //품명
        private void chkArticle_Click(object sender, RoutedEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                txtArticle.IsEnabled = true;
                txtArticle.Focus();
                btnArticle.IsEnabled = true;
            }
            else
            {
                txtArticle.IsEnabled = false;
                btnArticle.IsEnabled = false;
            }
        }
        //품명
        private void chkArticle_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true)
            {
                chkArticle.IsChecked = false;
                txtArticle.IsEnabled = false;
                btnArticle.IsEnabled = false;
            }
            else
            {
                chkArticle.IsChecked = true;
                txtArticle.IsEnabled = true;
                btnArticle.IsEnabled = true;
                txtArticle.Focus();
            }
        }

        //품번
        private void chkArticleNo_Click(object sender, RoutedEventArgs e)
        {
            if (chkArticleNo.IsChecked == true)
            {
                txtArticleNo.IsEnabled = true;
                txtArticleNo.Focus();
                btnArticleNo.IsEnabled = true;
            }
            else
            {
                txtArticleNo.IsEnabled = false;
                btnArticleNo.IsEnabled = false;
            }
        }

        //품번
        private void chkArticleNo_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleNo.IsChecked == true)
            {
                chkArticleNo.IsChecked = false;
                txtArticleNo.IsEnabled = false;
                btnArticleNo.IsEnabled = false;
            }
            else
            {
                chkArticleNo.IsChecked = true;
                txtArticleNo.IsEnabled = true;
                btnArticleNo.IsEnabled = true;
                txtArticleNo.Focus();
            }
        }
        //발생내역
        private void chkDefectContent_Click(object sender, RoutedEventArgs e)
        {
            if (chkDefectContent.IsChecked == true)
            {
                txtDefectContent.IsEnabled = true;
                txtDefectContent.Focus();
            }
            else
            {
                txtDefectContent.IsEnabled = false;
            }
        }
        //발생내역
        private void chkDefectContent_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkDefectContent.IsChecked == true)
            {
                chkDefectContent.IsChecked = false;
                txtDefectContent.IsEnabled = false;
            }
            else
            {
                chkDefectContent.IsChecked = true;
                txtDefectContent.IsEnabled = true;
                txtDefectContent.Focus();
            }
        }
        //불량번호
        private void chkDefectNO_Click(object sender, RoutedEventArgs e)
        {
            if (chkDefectNO.IsChecked == true)
            {
                txtDefectNO.IsEnabled = true;
                txtDefectNO.Focus();
            }
            else
            {
                txtDefectNO.IsEnabled = false;
            }
        }
        //불량번호
        private void chkDefectNO_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkDefectNO.IsChecked == true)
            {
                chkDefectNO.IsChecked = false;
                txtDefectNO.IsEnabled = false;
            }
            else
            {
                chkDefectNO.IsChecked = true;
                txtDefectNO.IsEnabled = true;
                txtDefectNO.Focus();
            }
        }
        //종결여부
        private void chkCloseYN_Click(object sender, RoutedEventArgs e)
        {
            if (chkCloseYN.IsChecked == true)
            {
                cboCloseYN.IsEnabled = true;
            }
            else
            {
                cboCloseYN.IsEnabled = false;
            }
        }
        //종결여부
        private void chkCloseYN_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkCloseYN.IsChecked == true)
            {
                chkCloseYN.IsChecked = false;
                cboCloseYN.IsEnabled = false;
            }
            else
            {
                chkCloseYN.IsChecked = true;
                cboCloseYN.IsEnabled = true;
            }
        }

        ////////////////////////////////////////////////////////////////////////////////////////

        //발생일
        private void chkOccurDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkOccurDate.IsChecked == true)
            {
                if (dtpOccurDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpOccurDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpOccurDate.IsEnabled = true;
            }
            else
            {
                dtpOccurDate.IsEnabled = false;
            }
        }
        //발생일
        private void chkOccurDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkOccurDate.IsChecked == true)
            {
                chkOccurDate.IsChecked = false;
                dtpOccurDate.IsEnabled = false;
            }
            else
            {
                chkOccurDate.IsChecked = true;
                if (dtpOccurDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpOccurDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpOccurDate.IsEnabled = true;
            }
        }
        //통보일
        private void chkNotifyDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkNotifyDate.IsChecked == true)
            {
                if (dtpNotifyDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpNotifyDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpNotifyDate.IsEnabled = true;
            }
            else
            {
                dtpNotifyDate.IsEnabled = false;
            }
        }
        //통보일
        private void chkNotifyDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkNotifyDate.IsChecked == true)
            {
                chkNotifyDate.IsChecked = false;
                dtpNotifyDate.IsEnabled = false;
            }
            else
            {
                chkNotifyDate.IsChecked = true;
                if (dtpNotifyDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpNotifyDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpNotifyDate.IsEnabled = true;
            }
        }
        //회신요청일
        private void chkReplyReqDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkReplyReqDate.IsChecked == true)
            {
                if (dtpReplyReqDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpReplyReqDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpReplyReqDate.IsEnabled = true;
            }
            else
            {
                dtpReplyReqDate.IsEnabled = false;
            }
        }
        //회신요청일
        private void chkReplyReqDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkReplyReqDate.IsChecked == true)
            {
                chkReplyReqDate.IsChecked = false;
                dtpReplyReqDate.IsEnabled = false;
            }
            else
            {
                chkReplyReqDate.IsChecked = true;
                if (dtpReplyReqDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpReplyReqDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpReplyReqDate.IsEnabled = true;
            }
        }
        //회신일
        private void chkReplyDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkReplyDate.IsChecked == true)
            {
                if (dtpReplyDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpReplyDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpReplyDate.IsEnabled = true;
            }
            else
            {
                dtpReplyDate.IsEnabled = false;
            }
        }
        //회신일
        private void chkReplyDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkReplyDate.IsChecked == true)
            {
                chkReplyDate.IsChecked = false;
                dtpReplyDate.IsEnabled = false;
            }
            else
            {
                chkReplyDate.IsChecked = true;
                if (dtpReplyDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpReplyDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpReplyDate.IsEnabled = true;
            }
        }
        //접수일
        private void chkAcptDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkAcptDate.IsChecked == true)
            {
                if (dtpAcptDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpAcptDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpAcptDate.IsEnabled = true;
            }
            else
            {
                dtpAcptDate.IsEnabled = false;
            }
        }
        //접수일
        private void chkAcptDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkAcptDate.IsChecked == true)
            {
                chkAcptDate.IsChecked = false;
                dtpAcptDate.IsEnabled = false;
            }
            else
            {
                chkAcptDate.IsChecked = true;
                if (dtpAcptDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpAcptDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpAcptDate.IsEnabled = true;
            }
        }
        //불량접수, 원인분석 검토일
        private void chkDefectRespectDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkDefectRespectDate.IsChecked == true)
            {
                if (dtpDefectRespectDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpDefectRespectDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpDefectRespectDate.IsEnabled = true;
            }
            else
            {
                dtpDefectRespectDate.IsEnabled = false;
            }
        }
        //불량접수, 원인분석 검토일
        private void chkDefectRespectDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkDefectRespectDate.IsChecked == true)
            {
                chkDefectRespectDate.IsChecked = false;
                dtpDefectRespectDate.IsEnabled = false;
            }
            else
            {
                chkDefectRespectDate.IsChecked = true;
                if (dtpDefectRespectDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpDefectRespectDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpDefectRespectDate.IsEnabled = true;
            }
        }
        //조치예정일
        private void chkCorrExpectDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkCorrExpectDate.IsChecked == true)
            {
                if (dtpCorrExpectDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpCorrExpectDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpCorrExpectDate.IsEnabled = true;
            }
            else
            {
                dtpCorrExpectDate.IsEnabled = false;
            }
        }
        //조치예정일
        private void chkCorrExpectDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkCorrExpectDate.IsChecked == true)
            {
                chkCorrExpectDate.IsChecked = false;
                dtpCorrExpectDate.IsEnabled = false;
            }
            else
            {
                chkCorrExpectDate.IsChecked = true;
                if (dtpCorrExpectDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpCorrExpectDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpCorrExpectDate.IsEnabled = true;
            }
        }
        //조치완료일
        private void chkCorrCompDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkCorrCompDate.IsChecked == true)
            {
                if (dtpCorrCompDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpCorrCompDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpCorrCompDate.IsEnabled = true;
            }
            else
            {
                dtpCorrCompDate.IsEnabled = false;
            }
        }
        //조치완료일
        private void chkCorrCompDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkCorrCompDate.IsChecked == true)
            {
                chkCorrCompDate.IsChecked = false;
                dtpCorrCompDate.IsEnabled = false;
            }
            else
            {
                chkCorrCompDate.IsChecked = true;
                if (dtpCorrCompDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpCorrCompDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpCorrCompDate.IsEnabled = true;
            }
        }
        //시정조치 검토일
        private void chkCorrRespectDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkCorrRespectDate.IsChecked == true)
            {
                if (dtpCorrRespectDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpCorrRespectDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpCorrRespectDate.IsEnabled = true;
            }
            else
            {
                dtpCorrRespectDate.IsEnabled = false;
            }
        }
        //시정조치 검토일
        private void chkCorrRespectDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkCorrRespectDate.IsChecked == true)
            {
                chkCorrRespectDate.IsChecked = false;
                dtpCorrRespectDate.IsEnabled = false;
            }
            else
            {
                chkCorrRespectDate.IsChecked = true;
                if (dtpCorrRespectDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpCorrRespectDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpCorrRespectDate.IsEnabled = true;
            }
        }
        //생산일 from - to
        private void chkCorrProdDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkCorrProdDate.IsChecked == true)
            {
                if (dtpCorrProdFromDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpCorrProdFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                if (dtpCorrProdToDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpCorrProdToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpCorrProdFromDate.IsEnabled = true;
                dtpCorrProdToDate.IsEnabled = true;
            }
            else
            {
                dtpCorrProdFromDate.IsEnabled = false;
                dtpCorrProdToDate.IsEnabled = false;
            }
        }
        //생산일 from - to
        private void chkCorrProdDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkCorrProdDate.IsChecked == true)
            {
                chkCorrProdDate.IsChecked = false;
                dtpCorrProdFromDate.IsEnabled = false;
                dtpCorrProdToDate.IsEnabled = false;
            }
            else
            {
                chkCorrProdDate.IsChecked = true;
                if (dtpCorrProdFromDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpCorrProdFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                if (dtpCorrProdToDate.Text == string.Empty)      // 비어있는 날짜박스를 체크한 경우,
                {
                    dtpCorrProdToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
                }
                dtpCorrProdFromDate.IsEnabled = true;
                dtpCorrProdToDate.IsEnabled = true;
            }
        }

        #endregion


        #region 플러스파인더

        // 플러스파인더 _ 거래처 찾기.
        private void btnCustomer_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomer, 0, "");
        }
        // 플러스파인더 _ 거래처 찾기.
        private void btnCustomer_InGroupBox_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomer_InGroupBox, 0, "");
        }
        // 플러스파인더 _ 품명 찾기
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticle, 77, txtArticle.Text);
        }

        // 품명 키다운 _ 품명 찾기
        private void TxtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticle, 77, txtArticle.Text);
            }
        }


        // 플러스 파인더 품번 이벤트
        private void plusFinder_replyArticle(string buyerArticleNo)
        {
            txtBuyerArticleNo_InGroupBox.Text = buyerArticleNo;
            pf.refEvent -= new PlusFinder.RefEventHandler(plusFinder_replyArticle);
        }

        // 플러스파인더 _ 품명찾기
        private void btnArticle_InGroupBox_Click(object sender, RoutedEventArgs e)
        {
            pf.refEvent += new PlusFinder.RefEventHandler(plusFinder_replyArticle);
            pf.ReturnCode(txtArticle_InGroupBox, 77, txtArticle_InGroupBox.Text);
            dtpOccurDate.Focus();
        }

        // 플러스파인더 _ 품번 찾기
        private void btnArticleNo_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticleNo, 76, txtArticleNo.Text);
        }

        // 품번 키다운 
        private void TxtArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticleNo, 76, txtArticleNo.Text);
            }
        }

        // 플러스파인더 _ 불량유형 찾기
        private void btnDefectSymptom_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtDefectSymptom, 3, "");
            lib.SendK(Key.Tab, this);
        }
        // 플러스파인더 _ 불량유형 찾기
        private void btnDefectRespectSymptom_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtDefectRespectSymptom, 3, "");
            lib.SendK(Key.Tab, this);
        }

        #endregion


        #region 콤보박스 세팅

        //콤보박스 세팅
        private void ComboBoxSetting()
        {
            cboCloseYN.Items.Clear();
            cboCloseYN_InGroupBox.Items.Clear();
            cboDefectReason.Items.Clear();
            cboDefectRespectReason.Items.Clear();
            cboImportantGrade.Items.Clear();
            cboCriticalGrade.Items.Clear();
            cboDvlYN_InGroupBox.Items.Clear();

            DataTable dt = new DataTable();
            dt.Columns.Add("value");
            dt.Columns.Add("display");

            DataRow row0 = dt.NewRow();
            row0["value"] = "Y";
            row0["display"] = "Y";

            DataRow row1 = dt.NewRow();
            row1["value"] = "N";
            row1["display"] = "N";

            dt.Rows.Add(row0);
            dt.Rows.Add(row1);

            this.cboCloseYN.ItemsSource = dt.DefaultView;
            this.cboCloseYN.DisplayMemberPath = "display";
            this.cboCloseYN.SelectedValuePath = "value";
            this.cboCloseYN.SelectedIndex = 0;

            this.cboCloseYN_InGroupBox.ItemsSource = dt.DefaultView;
            this.cboCloseYN_InGroupBox.DisplayMemberPath = "display";
            this.cboCloseYN_InGroupBox.SelectedValuePath = "value";
            this.cboCloseYN_InGroupBox.SelectedIndex = 0;

            dt = new DataTable();
            dt.Columns.Add("value");
            dt.Columns.Add("display");

            row0 = dt.NewRow();
            row0["value"] = "Y";
            row0["display"] = "개발";

            row1 = dt.NewRow();
            row1["value"] = "N";
            row1["display"] = "양산";

            dt.Rows.Add(row0);
            dt.Rows.Add(row1);

            this.cboDvlYN_InGroupBox.ItemsSource = dt.DefaultView;
            this.cboDvlYN_InGroupBox.DisplayMemberPath = "display";
            this.cboDvlYN_InGroupBox.SelectedValuePath = "value";
            this.cboDvlYN_InGroupBox.SelectedIndex = 0;

            ////////////////////////////////////////////////////////////////////////////////////

            ObservableCollection<CodeView> cbDefectReason = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "QULDFRSN", "Y", "", "");
            ObservableCollection<CodeView> cbImportantGrade = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "DEFIMPORT", "Y", "", "");
            ObservableCollection<CodeView> cbCriticalGrade = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "DEFCRITI", "Y", "", "");


            this.cboDefectReason.ItemsSource = cbDefectReason;
            this.cboDefectReason.DisplayMemberPath = "code_name";
            this.cboDefectReason.SelectedValuePath = "code_id";
            this.cboDefectReason.SelectedIndex = -1;
            this.cboDefectRespectReason.ItemsSource = cbDefectReason;
            this.cboDefectRespectReason.DisplayMemberPath = "code_name";
            this.cboDefectRespectReason.SelectedValuePath = "code_id";
            this.cboDefectRespectReason.SelectedIndex = -1;

            this.cboImportantGrade.ItemsSource = cbImportantGrade;
            this.cboImportantGrade.DisplayMemberPath = "code_name";
            this.cboImportantGrade.SelectedValuePath = "code_id";
            this.cboImportantGrade.SelectedIndex = -1;

            this.cboCriticalGrade.ItemsSource = cbCriticalGrade;
            this.cboCriticalGrade.DisplayMemberPath = "code_name";
            this.cboCriticalGrade.SelectedValuePath = "code_id";
            this.cboCriticalGrade.SelectedIndex = -1;

        }

        #endregion


        #region 공통버튼 이벤트
        //공통 사용가능
        private void PublicEnableTrue()
        {
            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;
            EventLabel.Visibility = Visibility.Hidden;

            btnSearch.IsEnabled = true;
            btnAdd.IsEnabled = true;
            btnEdit.IsEnabled = true;
            btnDelete.IsEnabled = true;
            btnExcel.IsEnabled = true;

            //그룹박스와는 별개로 다시 활성화 __ (추가랑 연동되서 바보되지 않도록)
            dtpNotifyDate.IsEnabled = true;
            dtpReplyReqDate.IsEnabled = true;
            dtpReplyDate.IsEnabled = true;
            dtpAcptDate.IsEnabled = true;
            dtpDefectRespectDate.IsEnabled = true;
            dtpCorrExpectDate.IsEnabled = true;
            dtpCorrCompDate.IsEnabled = true;
            dtpCorrRespectDate.IsEnabled = true;
            dtpCorrProdFromDate.IsEnabled = true;
            dtpCorrProdToDate.IsEnabled = true;

            grbBasisInfoBox.IsEnabled = false;
            grbDefectInfoBox.IsEnabled = false;
            grbReceiptBox.IsEnabled = false;
            grbActionBox.IsEnabled = false;
            grbAttachBox.IsEnabled = false;

            cboCloseYN_InGroupBox.IsEnabled = false;
            cboDvlYN_InGroupBox.IsEnabled = false;

            //FTP 버튼.
            btnFileEnroll1.IsEnabled = false;
            btnFileEnroll2.IsEnabled = false;
            btnFileEnroll3.IsEnabled = false;
            btnFileDel1.IsEnabled = false;
            btnFileDel2.IsEnabled = false;
            btnFileDel3.IsEnabled = false;
            btnFileDown1.IsEnabled = false;
            btnFileDown2.IsEnabled = false;
            btnFileDown3.IsEnabled = false;

            //dgdCustomDefect.IsEnabled = true; //메인그리드 사용가능.
            dgdCustomDefect.IsHitTestVisible = true; //메인그리드 사용가능.
        }

        // 공통 버튼이벤트.
        private void PublicEnableFalse()
        {
            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
            EventLabel.Visibility = Visibility.Visible;

            btnSearch.IsEnabled = false;
            btnAdd.IsEnabled = false;
            btnEdit.IsEnabled = false;
            btnDelete.IsEnabled = false;
            btnExcel.IsEnabled = false;

            grbBasisInfoBox.IsEnabled = true;
            grbDefectInfoBox.IsEnabled = true;
            grbReceiptBox.IsEnabled = true;
            grbActionBox.IsEnabled = true;
            grbAttachBox.IsEnabled = true;

            cboCloseYN_InGroupBox.IsEnabled = true;
            cboDvlYN_InGroupBox.IsEnabled = true;

            //불량 id 사용불가.
            txtDefectID.IsReadOnly = true;

            //FTP 버튼.
            btnFileEnroll1.IsEnabled = true;
            btnFileEnroll2.IsEnabled = true;
            btnFileEnroll3.IsEnabled = true;
            btnFileDel1.IsEnabled = true;
            btnFileDel2.IsEnabled = true;
            btnFileDel3.IsEnabled = true;
            btnFileDown1.IsEnabled = false;
            btnFileDown2.IsEnabled = false;
            btnFileDown3.IsEnabled = false;


            //dgdCustomDefect.IsEnabled = false; //메인그리드 못 건드리게.
            dgdCustomDefect.IsHitTestVisible = false; //메인그리드 못 건드리게.

        }

        #endregion


        #region  조회 / 조회용 프로시저

        // 검색 버튼 클릭 시.
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {

            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                Wh_Ar_SelectedLastIndex = 0;
                re_Search(Wh_Ar_SelectedLastIndex);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        /// <summary>
        /// 재검색(수정,삭제,추가 저장후에 자동 재검색)
        /// </summary>
        /// <param name="selectedIndex"></param>
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdCustomDefect.Items.Count > 0)
            {
                dgdCustomDefect.SelectedIndex = selectedIndex;
            }
        }


        private void FillGrid()
        {
            if (dgdCustomDefect.Items.Count > 0)
            {
                dgdCustomDefect.Items.Clear();
            }

            int nchkDate = 0;
            string FromDate = string.Empty;
            string ToDate = string.Empty;
            if (chkSearchDay.IsChecked == true)
            {
                if (rbnOccurDay.IsChecked == true) { nchkDate = 1; }
                if (rbnReceiptDay.IsChecked == true) { nchkDate = 2; }
                if (rbnActionDay.IsChecked == true) { nchkDate = 3; }
                FromDate = dtpFromDate.ToString().Substring(0, 10).Replace("-", "");
                ToDate = dtpToDate.ToString().Substring(0, 10).Replace("-", "");
            }

            string sDvlYN = "Y";
            if (rbnChoice2.IsChecked == true) { sDvlYN = "N"; }

            int nchkCustomID = 0;
            string CustomID = string.Empty;
            string Custom = string.Empty;
            if (chkCustomer.IsChecked == true)
            {
                if (txtCustomer.Tag == null)
                {
                    txtCustomer.Tag = "";
                    if (txtCustomer.Text.Length > 0)
                    {
                        nchkCustomID = 2;
                        Custom = txtCustomer.Text;
                    }
                }
                else
                {
                    nchkCustomID = 1;
                    CustomID = txtCustomer.Tag.ToString();
                }
            }

            //int nchkArticleID = 0;
            //string ArticleID = string.Empty;
            //string Article = string.Empty;
            //if (chkArticle.IsChecked == true)
            //{
            //    if (txtArticle.Tag == null)
            //    {
            //        txtArticle.Tag = "";
            //        if (txtArticle.Text.Length > 0)
            //        {
            //            nchkArticleID = 2;
            //            Article = txtArticle.Text;
            //        }
            //    }
            //    else
            //    {
            //        nchkArticleID = 1;
            //        ArticleID = txtArticle.Tag.ToString();
            //    }
            //}

            int nchkDefectNo = 0;
            string DefectNo = string.Empty;
            if (chkDefectNO.IsChecked == true)
            {
                nchkDefectNo = 1;
                DefectNo = txtDefectNO.Text;
            }

            int nchkDefectContents = 0;
            string DefectContents = string.Empty;
            if (chkDefectContent.IsChecked == true)
            {
                nchkDefectContents = 1;
                DefectContents = txtDefectContent.Text;
            }

            int nchkCloseYN = 0;
            string CloseYN = string.Empty;
            if (chkCloseYN.IsChecked == true)
            {
                nchkCloseYN = 1;
                CloseYN = cboCloseYN.SelectedValue.ToString();
            }


            int chkArticleID = 0;
            string ArticleID = "";

            if (chkArticle.IsChecked == true)
            {
                chkArticleID = 1;
                ArticleID = txtArticle.Tag.ToString();
            }
            if (chkArticleNo.IsChecked == true)
            {
                chkArticleID = 1;
                ArticleID = txtArticleNo.Tag.ToString();
            }

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("nchkDate", nchkDate);             // 인트
                sqlParameter.Add("FromDate", FromDate);
                sqlParameter.Add("ToDate", ToDate);
                sqlParameter.Add("sDvlYN", sDvlYN);                 // 개발_양산 구분 (개발 : Y, 양산 : N)
                sqlParameter.Add("nchkCustomID", nchkCustomID);     // 인트

                sqlParameter.Add("CustomID", CustomID);
                //sqlParameter.Add("Custom", Custom);
                sqlParameter.Add("nchkArticleID", chkArticleID);              // 인트
                sqlParameter.Add("ArticleID", ArticleID);
                //sqlParameter.Add("Article", Article);

                sqlParameter.Add("nchkDefectNo", nchkDefectNo);               //인트
                sqlParameter.Add("DefectNo", DefectNo);
                sqlParameter.Add("nchkDefectContents", nchkDefectContents);         //인트
                sqlParameter.Add("DefectContents", DefectContents);
                sqlParameter.Add("nchkCloseYN", nchkCloseYN);                //인트

                sqlParameter.Add("CloseYN", CloseYN);
                sqlParameter.Add("BuyerArticleNme", chkArticle.IsChecked == true ? txtArticle.Text : "");
                sqlParameter.Add("BuyerArticleNo", chkArticleNo.IsChecked == true ? txtArticleNo.Text : "");
                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Qul_sInspectCustom", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    int i = 0;
                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        // 검색 후, 첨부박스 봉인 해제/.
                        grbAttachBox.IsEnabled = true;
                        btnFileEnroll1.IsEnabled = false;
                        btnFileEnroll2.IsEnabled = false;
                        btnFileEnroll3.IsEnabled = false;
                        btnFileDel1.IsEnabled = false;
                        btnFileDel2.IsEnabled = false;
                        btnFileDel3.IsEnabled = false;
                        btnFileDown1.IsEnabled = false;
                        btnFileDown2.IsEnabled = false;
                        btnFileDown3.IsEnabled = false;

                        //조회결과가 있다면,
                        //dgdCustomDefect.Items.Clear();
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow item in drc)
                        {
                            var Win_Qul_CustomDefect_U_Insert = new Win_Qul_CustomDefect_U_View()
                            {
                                DefectID = item["DefectID"].ToString(),
                                DefectNo = item["DefectNo"].ToString(),
                                CustomID = item["CustomID"].ToString(),
                                KCustom = item["KCustom"].ToString(),
                                ArticleID = item["ArticleID"].ToString(),

                                Article = item["Article"].ToString(),
                                BuyerArticleNo = item["BuyerArticleNo"].ToString(),
                                BuyerModelID = item["BuyerModelID"].ToString(),
                                BuyerModelName = item["BuyerModelName"].ToString(),
                                OccurDate = item["OccurDate"].ToString(),

                                ReplyReqDate = item["ReplyReqDate"].ToString(),
                                NotifyDate = item["NotifyDate"].ToString(),
                                ReplyDate = item["ReplyDate"].ToString(),
                                DefectCheckMan = item["DefectCheckMan"].ToString(),
                                NotifyMan = item["NotifyMan"].ToString(),

                                DefectQty = item["DefectQty"].ToString(),
                                ReDefectQty = item["ReDefectQty"].ToString(),
                                DefectSymtomCode = item["DefectSymtomCode"].ToString(),
                                DefectSymtom = item["DefectSymtom"].ToString(),
                                DefectReasonCode = item["DefectReasonCode"].ToString(),

                                DefectReason = item["DefectReason"].ToString(),
                                ImportantGrade = item["ImportantGrade"].ToString(),
                                ImportantGradeName = item["ImportantGradeName"].ToString(),
                                OccurProcess = item["OccurProcess"].ToString(),
                                OccurMachine = item["OccurMachine"].ToString(),

                                CriticalGrade = item["CriticalGrade"].ToString(),
                                CriticalGradeName = item["CriticalGradeName"].ToString(),
                                DefectContents = item["DefectContents"].ToString(),
                                ReasonImput = item["ReasonImput"].ToString(),
                                AcptDate = item["AcptDate"].ToString(),         // 접수일                            

                                AcptMan = item["AcptMan"].ToString(),
                                DefectRespectDate = item["DefectRespectDate"].ToString(),
                                DefectRespectMan = item["DefectRespectMan"].ToString(),
                                DefectRespectContents = item["DefectRespectContents"].ToString(),
                                CorrExpectDate = item["CorrExpectDate"].ToString(),         // 조치예정일                           

                                CorrCompDate = item["CorrCompDate"].ToString(),
                                CorrRespectDate = item["CorrRespectDate"].ToString(),
                                CorrRespectMan = item["CorrRespectMan"].ToString(),
                                CorrProdFromDate = item["CorrProdFromDate"].ToString(),
                                CorrProdToDate = item["CorrProdToDate"].ToString(),

                                DefectRespectSymtomCode = item["DefectRespectSymtomCode"].ToString(),
                                DefectRespectSymtom = item["DefectRespectSymtom"].ToString(),
                                DefectRespectReasonCode = item["DefectRespectReasonCode"].ToString(),
                                DefectRespectReason = item["DefectRespectReason"].ToString(),
                                CorrContents = item["CorrContents"].ToString(),

                                CorrEOChangeYN = item["CorrEOChangeYN"].ToString(),
                                Corr4MChangeYN = item["Corr4MChangeYN"].ToString(),
                                CorrDesignChangeYN = item["CorrDesignChangeYN"].ToString(),
                                CorrDesignChangeContents = item["CorrDesignChangeContents"].ToString(),
                                CloseYN = item["CloseYN"].ToString(),

                                AttPath1 = item["AttPath1"].ToString(),
                                AttFile1 = item["AttFile1"].ToString(),
                                AttPath2 = item["AttPath2"].ToString(),
                                AttFile2 = item["AttFile2"].ToString(),
                                AttPath3 = item["AttPath3"].ToString(),

                                AttFile3 = item["AttFile3"].ToString(),
                                CreateUserID = item["CreateUserID"].ToString(),

                                DvlYN = sDvlYN
                            };

                            if ((Win_Qul_CustomDefect_U_Insert.OccurDate != string.Empty) && (Win_Qul_CustomDefect_U_Insert.OccurDate != null) && (lib.IsNullOrWhiteSpace(Win_Qul_CustomDefect_U_Insert.OccurDate) != true))   //발생일
                            {
                                Win_Qul_CustomDefect_U_Insert.OccurDate = DateTime.ParseExact(Win_Qul_CustomDefect_U_Insert.OccurDate, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                            }
                            else { Win_Qul_CustomDefect_U_Insert.OccurDate = string.Empty; }

                            if ((Win_Qul_CustomDefect_U_Insert.ReplyReqDate != string.Empty) && (Win_Qul_CustomDefect_U_Insert.ReplyReqDate != null) && (lib.IsNullOrWhiteSpace(Win_Qul_CustomDefect_U_Insert.ReplyReqDate) != true))     //회신요청일
                            {
                                Win_Qul_CustomDefect_U_Insert.ReplyReqDate = DateTime.ParseExact(Win_Qul_CustomDefect_U_Insert.ReplyReqDate, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                            }
                            else { Win_Qul_CustomDefect_U_Insert.ReplyReqDate = string.Empty; }

                            if ((Win_Qul_CustomDefect_U_Insert.NotifyDate != string.Empty) && (Win_Qul_CustomDefect_U_Insert.NotifyDate != null) && (lib.IsNullOrWhiteSpace(Win_Qul_CustomDefect_U_Insert.NotifyDate) != true))     //통보일
                            {
                                Win_Qul_CustomDefect_U_Insert.NotifyDate = DateTime.ParseExact(Win_Qul_CustomDefect_U_Insert.NotifyDate, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                            }
                            else { Win_Qul_CustomDefect_U_Insert.NotifyDate = string.Empty; }

                            if ((Win_Qul_CustomDefect_U_Insert.ReplyDate != string.Empty) && (Win_Qul_CustomDefect_U_Insert.ReplyDate != null) && (lib.IsNullOrWhiteSpace(Win_Qul_CustomDefect_U_Insert.ReplyDate) != true))     //회신(응답)일
                            {
                                Win_Qul_CustomDefect_U_Insert.ReplyDate = DateTime.ParseExact(Win_Qul_CustomDefect_U_Insert.ReplyDate, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                            }
                            else { Win_Qul_CustomDefect_U_Insert.ReplyDate = string.Empty; }

                            if ((Win_Qul_CustomDefect_U_Insert.AcptDate != string.Empty) && (Win_Qul_CustomDefect_U_Insert.AcptDate != null) && (lib.IsNullOrWhiteSpace(Win_Qul_CustomDefect_U_Insert.AcptDate) != true))     //접수일
                            {
                                Win_Qul_CustomDefect_U_Insert.AcptDate = DateTime.ParseExact(Win_Qul_CustomDefect_U_Insert.AcptDate, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                            }
                            else { Win_Qul_CustomDefect_U_Insert.AcptDate = string.Empty; }

                            if ((Win_Qul_CustomDefect_U_Insert.DefectRespectDate != string.Empty) && (Win_Qul_CustomDefect_U_Insert.DefectRespectDate != null) && (lib.IsNullOrWhiteSpace(Win_Qul_CustomDefect_U_Insert.DefectRespectDate) != true))     //불량접수 검토일
                            {
                                Win_Qul_CustomDefect_U_Insert.DefectRespectDate = DateTime.ParseExact(Win_Qul_CustomDefect_U_Insert.DefectRespectDate, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                            }
                            else { Win_Qul_CustomDefect_U_Insert.DefectRespectDate = string.Empty; }

                            if ((Win_Qul_CustomDefect_U_Insert.CorrExpectDate != string.Empty) && (Win_Qul_CustomDefect_U_Insert.CorrExpectDate != null) && (lib.IsNullOrWhiteSpace(Win_Qul_CustomDefect_U_Insert.CorrExpectDate) != true))     //조치 예정일
                            {
                                Win_Qul_CustomDefect_U_Insert.CorrExpectDate = DateTime.ParseExact(Win_Qul_CustomDefect_U_Insert.CorrExpectDate, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                            }
                            else { Win_Qul_CustomDefect_U_Insert.CorrExpectDate = string.Empty; }

                            if ((Win_Qul_CustomDefect_U_Insert.CorrCompDate != string.Empty) && (Win_Qul_CustomDefect_U_Insert.CorrCompDate != null) && (lib.IsNullOrWhiteSpace(Win_Qul_CustomDefect_U_Insert.CorrCompDate) != true))     //조치 완료일
                            {
                                Win_Qul_CustomDefect_U_Insert.CorrCompDate = DateTime.ParseExact(Win_Qul_CustomDefect_U_Insert.CorrCompDate, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                            }
                            else { Win_Qul_CustomDefect_U_Insert.CorrCompDate = string.Empty; }

                            if ((Win_Qul_CustomDefect_U_Insert.CorrRespectDate != string.Empty) && (Win_Qul_CustomDefect_U_Insert.CorrRespectDate != null) && (lib.IsNullOrWhiteSpace(Win_Qul_CustomDefect_U_Insert.CorrRespectDate) != true))     //시정조치 검토일
                            {
                                Win_Qul_CustomDefect_U_Insert.CorrRespectDate = DateTime.ParseExact(Win_Qul_CustomDefect_U_Insert.CorrRespectDate, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                            }
                            else { Win_Qul_CustomDefect_U_Insert.CorrRespectDate = string.Empty; }

                            if ((Win_Qul_CustomDefect_U_Insert.CorrProdFromDate != string.Empty) && (Win_Qul_CustomDefect_U_Insert.CorrProdFromDate != null) && (lib.IsNullOrWhiteSpace(Win_Qul_CustomDefect_U_Insert.CorrProdFromDate) != true))     //시정조치 생산 시작일
                            {
                                Win_Qul_CustomDefect_U_Insert.CorrProdFromDate = DateTime.ParseExact(Win_Qul_CustomDefect_U_Insert.CorrProdFromDate, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                            }
                            else { Win_Qul_CustomDefect_U_Insert.CorrProdFromDate = string.Empty; }

                            if ((Win_Qul_CustomDefect_U_Insert.CorrProdToDate != string.Empty) && (Win_Qul_CustomDefect_U_Insert.CorrProdToDate != null) && (lib.IsNullOrWhiteSpace(Win_Qul_CustomDefect_U_Insert.CorrProdToDate) != true))     //시정조치 생산 마지막일
                            {
                                Win_Qul_CustomDefect_U_Insert.CorrProdToDate = DateTime.ParseExact(Win_Qul_CustomDefect_U_Insert.CorrProdToDate, "yyyyMMdd", null).ToString("yyyy-MM-dd");
                            }
                            else { Win_Qul_CustomDefect_U_Insert.CorrProdToDate = string.Empty; }



                            dgdCustomDefect.Items.Add(Win_Qul_CustomDefect_U_Insert);
                            i++;
                        }
                        tbkIndexCount.Text = "▶검색결과 : " + i + "건";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        #endregion


        #region 메인그리드 연동 _ SHOW DATA
        // 그리드 셀 체인지 이벤트 _ show data.
        private void dgdCustomDefect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //this.DataContext = dgdCustomDefect.SelectedItem as Win_Qul_CustomDefect_U_View;
            var ViewReceiver = dgdCustomDefect.SelectedItem as Win_Qul_CustomDefect_U_View;

            if (ViewReceiver != null)
            {
                //this.DataContext = ViewReceiver;

                txtCustomer_InGroupBox.Tag = ViewReceiver.CustomID;
                txtArticle_InGroupBox.Tag = ViewReceiver.ArticleID;

                Win_Qul_CustomDefect_U_View ProCopy = ViewReceiver.Copy();
                this.DataContext = ProCopy;

                dtpOccurDate.Text = ViewReceiver.OccurDate;
                dtpReplyReqDate.Text = ViewReceiver.ReplyReqDate;
                dtpNotifyDate.Text = ViewReceiver.NotifyDate;
                dtpReplyDate.Text = ViewReceiver.ReplyDate;
                dtpAcptDate.Text = ViewReceiver.AcptDate;
                dtpDefectRespectDate.Text = ViewReceiver.DefectRespectDate;
                dtpCorrExpectDate.Text = ViewReceiver.CorrExpectDate;
                dtpCorrCompDate.Text = ViewReceiver.CorrCompDate;
                dtpCorrRespectDate.Text = ViewReceiver.CorrRespectDate;
                dtpCorrProdFromDate.Text = ViewReceiver.CorrProdFromDate;
                dtpCorrProdToDate.Text = ViewReceiver.CorrProdToDate;

                if (ViewReceiver.AttPath1 != string.Empty) { btnFileDown1.IsEnabled = true; }
                else { btnFileDown1.IsEnabled = false; }

                if (ViewReceiver.AttPath2 != string.Empty) { btnFileDown2.IsEnabled = true; }
                else { btnFileDown2.IsEnabled = false; }

                if (ViewReceiver.AttPath3 != string.Empty) { btnFileDown3.IsEnabled = true; }
                else { btnFileDown3.IsEnabled = false; }
            }
        }

        #endregion


        #region (추가 , 수정 , 삭제 , 저장, 취소) 버튼 이벤트 모음

        //추가 버튼 클릭 시.
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            //1. 그룹박스 데이터 클리어.
            grbBoxDataClear();

            //2. 공통 버튼이벤트
            PublicEnableFalse();
            EventLabel.Content = "자료 입력(추가) 중..";

            chkOccurDate.IsChecked = true;
            dtpOccurDate.Text = DateTime.Now.ToString("yyyy-MM-dd");            //발생일에 대한 기본세팅.(오늘자)

            //3. 각종 정보용 날짜들 사용불가 기본설정.(but, 딴 날짜들은 몰라도 인간적으로 발생일은 있어야지)
            dtpNotifyDate.IsEnabled = false;
            dtpReplyReqDate.IsEnabled = false;
            dtpReplyDate.IsEnabled = false;
            dtpAcptDate.IsEnabled = false;
            dtpDefectRespectDate.IsEnabled = false;
            dtpCorrExpectDate.IsEnabled = false;
            dtpCorrCompDate.IsEnabled = false;
            dtpCorrRespectDate.IsEnabled = false;
            dtpCorrProdFromDate.IsEnabled = false;
            dtpCorrProdToDate.IsEnabled = false;

            ButtonTag = ((Button)sender).Tag.ToString();

            if (dgdCustomDefect.Items.Count > 0)
            {
                Wh_Ar_SelectedLastIndex = dgdCustomDefect.SelectedIndex;
            }
            else
            {
                Wh_Ar_SelectedLastIndex = 0;
            }
            txtDefectNo_InGroupBox.Focus();
        }

        // 수정 버튼 클릭 시.
        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            // 1. 수정할 자격은 있는거야? 조회? 데이터 선택??
            if (dgdCustomDefect.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }
            var OBJ = dgdCustomDefect.SelectedItem as Win_Qul_CustomDefect_U_View;
            if (OBJ == null)
            {
                MessageBox.Show("수정할 항목이 정확히 선택되지 않았습니다.");
                return;
            }

            // 2.공통 버튼이벤트
            PublicEnableFalse();
            EventLabel.Content = "자료입력(작성) 중..";
            Wh_Ar_SelectedLastIndex = dgdCustomDefect.SelectedIndex;


            // 3. 각종 정보용 날짜들 체크세팅.
            if (dtpOccurDate.Text != string.Empty) { chkOccurDate.IsChecked = true; }
            else
            {
                chkOccurDate.IsChecked = true;
                dtpOccurDate.Text = DateTime.Now.ToString("yyyy-MM-dd");            //발생일에 대한 기본세팅.(오늘자)
            }

            if (dtpNotifyDate.Text != string.Empty) { chkNotifyDate.IsChecked = true; }
            else { dtpNotifyDate.IsEnabled = false; }

            if (dtpReplyReqDate.Text != string.Empty) { chkReplyReqDate.IsChecked = true; }
            else { dtpReplyReqDate.IsEnabled = false; }

            if (dtpReplyDate.Text != string.Empty) { chkReplyDate.IsChecked = true; }
            else { dtpReplyDate.IsEnabled = false; }

            if (dtpAcptDate.Text != string.Empty) { chkAcptDate.IsChecked = true; }
            else { dtpAcptDate.IsEnabled = false; }

            if (dtpDefectRespectDate.Text != string.Empty) { chkDefectRespectDate.IsChecked = true; }
            else { dtpDefectRespectDate.IsEnabled = false; }

            if (dtpCorrExpectDate.Text != string.Empty) { chkCorrExpectDate.IsChecked = true; }
            else { dtpCorrExpectDate.IsEnabled = false; }

            if (dtpCorrCompDate.Text != string.Empty) { chkCorrCompDate.IsChecked = true; }
            else { dtpCorrCompDate.IsEnabled = false; }

            if (dtpCorrRespectDate.Text != string.Empty) { chkCorrRespectDate.IsChecked = true; }
            else { dtpCorrRespectDate.IsEnabled = false; }

            if (dtpCorrProdFromDate.Text != string.Empty) { chkCorrProdDate.IsChecked = true; }
            else
            {
                dtpCorrProdFromDate.IsEnabled = false;
                dtpCorrProdToDate.IsEnabled = false;
            }

            ButtonTag = ((Button)sender).Tag.ToString();
        }

        //삭제 버튼 클릭 시.
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            // 1. 삭제할 자격은 있는거야? 조회? 데이터 선택??
            if (dgdCustomDefect.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }
            var OBJ = dgdCustomDefect.SelectedItem as Win_Qul_CustomDefect_U_View;
            if (OBJ == null)
            {
                MessageBox.Show("삭제할 항목이 정확히 선택되지 않았습니다.");
                return;
            }
            MessageBoxResult msgresult = MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                if (dgdCustomDefect.Items.Count > 0 && dgdCustomDefect.SelectedItem != null)
                {
                    Wh_Ar_SelectedLastIndex = dgdCustomDefect.SelectedIndex;
                }

                // 2.  삭제용
                DeleteData();
                dgdCustomDefect.Refresh();

                Wh_Ar_SelectedLastIndex -= 1;
                re_Search(Wh_Ar_SelectedLastIndex);
            }
        }

        //저장 버튼 클릭 시.
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            // 1. 데이터 기입체크.(항목을 제대로 모두 똑바로 넣고 저장버튼을 누르는 거야??) 
            if (grbEnroll_DataCheck() == false) { return; }

            // 2. 저장.
            SaveData(ButtonTag);

            if (ButtonTag == "1")     //1. 추가 > 저장했다면,
            {
                //공통 버튼이벤트
                PublicEnableTrue();
                //그룹박스 데이터 클리어
                grbBoxDataClear();
                if (dgdCustomDefect.Items.Count > 0)
                {
                    re_Search(dgdCustomDefect.Items.Count - 1);
                    dgdCustomDefect.Focus();
                }
                else
                { re_Search(0); }
            }
            else        //2. 수정 > 저장했다면,
            {
                //공통 버튼이벤트
                PublicEnableTrue();
                re_Search(Wh_Ar_SelectedLastIndex);
                dgdCustomDefect.Focus();
            }
        }
        #region 저장 전, 그룹박스 데이터 기입체크

        private bool grbEnroll_DataCheck()
        {
            if (txtCustomer_InGroupBox.Tag == null)
            {
                MessageBox.Show("고객사는 반드시 플러스파인더를 통해 입력하셔야 합니다.");
                return false;
            }
            if (txtArticle_InGroupBox.Tag == null)
            {
                MessageBox.Show("품명은 반드시 플러스파인더를 통해 입력하셔야 합니다.");
                return false;
            }
            if ((chkOccurDate.IsChecked == false) || (dtpOccurDate.Text == string.Empty))
            {
                MessageBox.Show("발생일은 반드시 입력하셔야 합니다.");
                return false;
            }
            if (lib.IsNullOrWhiteSpace(txtDefectCheckMan.Text) == true)
            {
                MessageBox.Show("불량 확인자는 반드시 입력하셔야 합니다.");
                return false;
            }

            if (txtDefectQty.Text != string.Empty)
            {
                if (lib.IsNumOrAnother(txtDefectQty.Text) == false)
                {
                    MessageBox.Show("불량수량은 숫자로만 입력하셔야 합니다.");
                    return false;
                }
            }

            if (txtReDefectQty.Text != string.Empty)
            {
                if (lib.IsNumOrAnother(txtReDefectQty.Text) == false)
                {
                    MessageBox.Show("불량재발횟수는 숫자로만 입력하셔야 합니다.");
                    return false;
                }
            }
            return true;
        }

        #endregion

        //취소 버튼 클릭 시.
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //공통 버튼이벤트
            PublicEnableTrue();

            //그룹박스 데이터 클리어
            grbBoxDataClear();

            if (ButtonTag == "1") // 1. 추가하다가 취소했다면,
            {
                if (dgdCustomDefect.Items.Count > 0)
                {
                    re_Search(Wh_Ar_SelectedLastIndex);
                    dgdCustomDefect.Focus();
                }
                else
                { re_Search(0); }
            }
            else        //2. 수정하다가 취소했다면
            {
                re_Search(Wh_Ar_SelectedLastIndex);
                dgdCustomDefect.Focus();
            }
            ButtonTag = string.Empty;
        }

        #endregion


        #region 그룹박스 데이터 클리어
        // 그룹박스 데이터 클리어 하기.
        private void grbBoxDataClear()
        {
            //1. 기본정보 박스.
            txtDefectID.Text = string.Empty;
            txtDefectNo_InGroupBox.Text = string.Empty;
            txtCustomer_InGroupBox.Text = string.Empty;
            txtCustomer_InGroupBox.Tag = null;
            txtArticle_InGroupBox.Text = string.Empty;
            txtArticle_InGroupBox.Tag = null;
            txtBuyerArticleNo_InGroupBox.Text = string.Empty;

            //2. 불량정보 박스.
            dtpOccurDate.Text = string.Empty;
            dtpNotifyDate.Text = string.Empty;
            dtpReplyReqDate.Text = string.Empty;
            dtpReplyDate.Text = string.Empty;
            txtDefectCheckMan.Text = string.Empty;
            txtNotifyMan.Text = string.Empty;
            txtDefectQty.Text = string.Empty;
            txtReDefectQty.Text = string.Empty;
            txtDefectSymptom.Text = string.Empty;
            txtDefectSymptom.Tag = null;
            cboDefectReason.SelectedIndex = -1;
            cboImportantGrade.SelectedIndex = -1;
            cboCriticalGrade.SelectedIndex = -1;
            txtOccurProcess.Text = string.Empty;
            txtOccurMachine.Text = string.Empty;
            txtDefectContent_InGroupBox.Text = string.Empty;

            //3. 불량 접수, 검토, 원인분석.
            dtpAcptDate.Text = string.Empty;
            dtpDefectRespectDate.Text = string.Empty;
            txtAcptMan.Text = string.Empty;
            txtDefectRespectMan.Text = string.Empty;
            txtReasonImput.Text = string.Empty;
            txtDefectRespectContent.Text = string.Empty;

            //4. 시정조치 박스.
            dtpCorrExpectDate.Text = string.Empty;
            dtpCorrCompDate.Text = string.Empty;
            dtpCorrRespectDate.Text = string.Empty;
            txtCorrRespectMan.Text = string.Empty;
            dtpCorrProdFromDate.Text = string.Empty;
            dtpCorrProdToDate.Text = string.Empty;
            txtDefectRespectSymptom.Text = string.Empty;
            txtDefectRespectSymptom.Tag = null;
            cboDefectRespectReason.SelectedIndex = -1;
            txtCorrContent.Text = string.Empty;

            cboCloseYN_InGroupBox.SelectedIndex = 0;
            cboDvlYN_InGroupBox.SelectedIndex = 0;

            //5. 파일 박스.
            txtAttFile1.Text = string.Empty;
            txtAttPath1.Text = string.Empty;
            txtAttFile2.Text = string.Empty;
            txtAttPath2.Text = string.Empty;
            txtAttFile3.Text = string.Empty;
            txtAttPath3.Text = string.Empty;


            //6. 날짜 체츠박스 해제.
            chkOccurDate.IsChecked = false;
            chkNotifyDate.IsChecked = false;
            chkReplyReqDate.IsChecked = false;
            chkReplyDate.IsChecked = false;
            chkAcptDate.IsChecked = false;
            chkDefectRespectDate.IsChecked = false;
            chkCorrExpectDate.IsChecked = false;
            chkCorrCompDate.IsChecked = false;
            chkCorrRespectDate.IsChecked = false;
            chkCorrProdDate.IsChecked = false;
        }

        #endregion


        #region CRUD // 각종 프로시저 모음

        private void SaveData(string TagNUM)
        {
            try
            {
                List<Procedure> Prolist = new List<Procedure>();
                List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

                if (TagNUM == "1")      // 신규추가입니다.
                {
                    // 신규추가 저장 insert.
                    string OccurDate = string.Empty;
                    if (chkOccurDate.IsChecked == true) { OccurDate = dtpOccurDate.Text.Substring(0, 10).Replace("-", ""); }

                    string ReplyReqDate = string.Empty;
                    if (chkReplyReqDate.IsChecked == true) { ReplyReqDate = dtpReplyReqDate.Text.Substring(0, 10).Replace("-", ""); }

                    string NotifyDate = string.Empty;
                    if (chkNotifyDate.IsChecked == true) { NotifyDate = dtpNotifyDate.Text.Substring(0, 10).Replace("-", ""); }

                    string ReplyDate = string.Empty;
                    if (chkReplyDate.IsChecked == true) { ReplyDate = dtpReplyDate.Text.Substring(0, 10).Replace("-", ""); }

                    double DefectQty = 0;
                    double ReDefectQty = 0;
                    if (txtDefectQty.Text != string.Empty) { DefectQty = Convert.ToDouble(txtDefectQty.Text); }  //(인트 체크완료)
                    if (txtReDefectQty.Text != string.Empty) { ReDefectQty = Convert.ToDouble(txtReDefectQty.Text); }  //(인트 체크완료)

                    string DefectSymtomCode = string.Empty;
                    if (txtDefectSymptom.Tag != null)
                    {
                        DefectSymtomCode = txtDefectSymptom.Tag.ToString();
                    }

                    string DefectReasonCode = string.Empty;
                    if (cboDefectReason.SelectedIndex != -1)
                    {
                        DefectReasonCode = cboDefectReason.SelectedValue.ToString();
                    }

                    string ImportantGrade = string.Empty;
                    if (cboImportantGrade.SelectedIndex != -1)
                    {
                        ImportantGrade = cboImportantGrade.SelectedValue.ToString();
                    }

                    string CriticalGrade = string.Empty;
                    if (cboCriticalGrade.SelectedIndex != -1)
                    {
                        CriticalGrade = cboCriticalGrade.SelectedValue.ToString();
                    }

                    string AcptDate = string.Empty;
                    if (chkAcptDate.IsChecked == true) { AcptDate = dtpAcptDate.Text.Substring(0, 10).Replace("-", ""); }

                    string DefectRespectDate = string.Empty;
                    if (chkDefectRespectDate.IsChecked == true) { DefectRespectDate = dtpDefectRespectDate.Text.Substring(0, 10).Replace("-", ""); }

                    string CorrExpectDate = string.Empty;
                    if (chkCorrExpectDate.IsChecked == true) { CorrExpectDate = dtpCorrExpectDate.Text.Substring(0, 10).Replace("-", ""); }

                    string CorrCompDate = string.Empty;
                    if (chkCorrCompDate.IsChecked == true) { CorrCompDate = dtpCorrCompDate.Text.Substring(0, 10).Replace("-", ""); }

                    string CorrRespectDate = string.Empty;
                    if (chkCorrRespectDate.IsChecked == true) { CorrRespectDate = dtpCorrRespectDate.Text.Substring(0, 10).Replace("-", ""); }

                    string CorrProdFromDate = string.Empty;
                    string CorrProdToDate = string.Empty;
                    if (chkCorrProdDate.IsChecked == true)
                    {
                        CorrProdFromDate = dtpCorrProdFromDate.Text.Substring(0, 10).Replace("-", "");
                        CorrProdToDate = dtpCorrProdToDate.Text.Substring(0, 10).Replace("-", "");
                    }

                    string DefectRespectSymtomCode = string.Empty;
                    if (txtDefectRespectSymptom.Tag != null)
                    {
                        DefectRespectSymtomCode = txtDefectRespectSymptom.Tag.ToString();
                    }

                    string DefectRespectReasonCode = string.Empty;
                    if (cboDefectRespectReason.SelectedIndex != -1)
                    {
                        DefectRespectReasonCode = cboDefectRespectReason.SelectedValue.ToString();
                    }



                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("DefectID", "");
                    sqlParameter.Add("DefectNo", txtDefectNo_InGroupBox.Text);
                    sqlParameter.Add("CustomID", txtCustomer_InGroupBox.Tag.ToString());        //필수기입 체크완료
                    sqlParameter.Add("ArticleID", txtArticle_InGroupBox.Tag.ToString());        //필수기입 체크완료
                    sqlParameter.Add("BuyerModelID", "");

                    sqlParameter.Add("OccurDate", OccurDate);
                    sqlParameter.Add("ReplyReqDate", ReplyReqDate);
                    sqlParameter.Add("NotifyDate", NotifyDate);
                    sqlParameter.Add("ReplyDate", ReplyDate);
                    sqlParameter.Add("DefectCheckMan", txtDefectCheckMan.Text);

                    sqlParameter.Add("NotifyMan", txtNotifyMan.Text);
                    sqlParameter.Add("DefectQty", DefectQty);          //numeric 
                    sqlParameter.Add("ReDefectQty", ReDefectQty);      //numeric 
                    sqlParameter.Add("DefectSymtomCode", DefectSymtomCode);
                    sqlParameter.Add("DefectReasonCode", DefectReasonCode);

                    sqlParameter.Add("ImportantGrade", ImportantGrade);
                    sqlParameter.Add("OccurProcess", txtOccurProcess.Text);
                    sqlParameter.Add("OccurMachine", txtOccurMachine.Text);
                    sqlParameter.Add("CriticalGrade", CriticalGrade);
                    sqlParameter.Add("DefectContents", txtDefectContent_InGroupBox.Text);

                    sqlParameter.Add("AcptDate", AcptDate);
                    sqlParameter.Add("AcptMan", txtAcptMan.Text);
                    sqlParameter.Add("DefectRespectDate", DefectRespectDate);
                    sqlParameter.Add("DefectRespectMan", txtDefectRespectMan.Text);
                    sqlParameter.Add("DefectRespectContents", txtDefectRespectContent.Text);

                    sqlParameter.Add("CorrExpectDate", CorrExpectDate);
                    sqlParameter.Add("CorrCompDate", CorrCompDate);
                    sqlParameter.Add("CorrRespectDate", CorrRespectDate);
                    sqlParameter.Add("CorrRespectMan", txtCorrRespectMan.Text);
                    sqlParameter.Add("CorrProdFromDate", CorrProdFromDate);

                    sqlParameter.Add("CorrProdToDate", CorrProdToDate);
                    sqlParameter.Add("DefectRespectSymtomCode", DefectRespectSymtomCode);
                    sqlParameter.Add("DefectRespectReasonCode", DefectRespectReasonCode);
                    sqlParameter.Add("CorrContents", txtCorrContent.Text);
                    sqlParameter.Add("CorrEOChangeYN", "");
                    sqlParameter.Add("Corr4MChangeYN", "");

                    sqlParameter.Add("CorrDesignChangeYN", "");
                    sqlParameter.Add("CorrDesignChangeContents", "");
                    sqlParameter.Add("CloseYN", cboCloseYN_InGroupBox.SelectedValue.ToString());
                    sqlParameter.Add("DvlDefectYN", cboDvlYN_InGroupBox.SelectedValue.ToString());
                    sqlParameter.Add("ReasonImput", txtReasonImput.Text);

                    sqlParameter.Add("AttPath1", "");
                    sqlParameter.Add("AttFile1", "");
                    sqlParameter.Add("AttPath2", "");
                    sqlParameter.Add("AttFile2", "");
                    sqlParameter.Add("AttPath3", "");
                    sqlParameter.Add("AttFile3", "");
                    sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_Qul_iInspectCustom";
                    pro1.OutputUseYN = "Y";
                    pro1.OutputName = "DefectID";
                    pro1.OutputLength = "10";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);

                    List<KeyValue> list_Result = new List<KeyValue>();
                    list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter,"C");
                    string sGetDefectID = string.Empty;

                    if (list_Result[0].key.ToLower() == "success")
                    {
                        list_Result.RemoveAt(0);
                        for (int i = 0; i < list_Result.Count; i++)
                        {
                            KeyValue kv = list_Result[i];
                            if (kv.key == "DefectID")
                            {
                                sGetDefectID = kv.value;
                            }
                        }

                        bool AttachYesNo = false;

                        if (txtAttFile1.Text != string.Empty || txtAttFile2.Text != string.Empty || txtAttFile3.Text != string.Empty)       //첨부파일 1
                        {
                            if (FTP_Save_File(listFtpFile, sGetDefectID))
                            {
                                if (!txtAttFile1.Text.Equals(string.Empty)) { txtAttPath1.Text = "/ImageData/CustomDefect/" + sGetDefectID; }
                                if (!txtAttFile2.Text.Equals(string.Empty)) { txtAttPath2.Text = "/ImageData/CustomDefect/" + sGetDefectID; }
                                if (!txtAttFile3.Text.Equals(string.Empty)) { txtAttPath3.Text = "/ImageData/CustomDefect/" + sGetDefectID; }

                                AttachYesNo = true;
                            }
                            else
                            { MessageBox.Show("데이터 저장이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }

                            if (AttachYesNo == true) { AttachFileUpdate(sGetDefectID); }      //첨부문서 정보 DB 업데이트.
                        }
                    }
                    else
                    {
                        MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                        //return false;
                    }
                }

                else if (TagNUM == "2")         // 수정 저장입니다.
                {
                    // 수정 저장 update.
                    var ViewReceiver = dgdCustomDefect.SelectedItem as Win_Qul_CustomDefect_U_View;

                    string OccurDate = string.Empty;
                    if (chkOccurDate.IsChecked == true) { OccurDate = dtpOccurDate.Text.Substring(0, 10).Replace("-", ""); }

                    string ReplyReqDate = string.Empty;
                    if (chkReplyReqDate.IsChecked == true) { ReplyReqDate = dtpReplyReqDate.Text.Substring(0, 10).Replace("-", ""); }

                    string NotifyDate = string.Empty;
                    if (chkNotifyDate.IsChecked == true) { NotifyDate = dtpNotifyDate.Text.Substring(0, 10).Replace("-", ""); }

                    string ReplyDate = string.Empty;
                    if (chkReplyDate.IsChecked == true) { ReplyDate = dtpReplyDate.Text.Substring(0, 10).Replace("-", ""); }

                    double DefectQty = 0;
                    double ReDefectQty = 0;
                    if (txtDefectQty.Text != string.Empty) { DefectQty = Convert.ToDouble(txtDefectQty.Text); }  //(인트 체크완료)
                    if (txtReDefectQty.Text != string.Empty) { ReDefectQty = Convert.ToDouble(txtReDefectQty.Text); }  //(인트 체크완료)

                    string DefectSymtomCode = string.Empty;
                    if (txtDefectSymptom.Tag != null)
                    {
                        DefectSymtomCode = txtDefectSymptom.Tag.ToString();
                    }

                    string DefectReasonCode = string.Empty;
                    if (cboDefectReason.SelectedIndex != -1)
                    {
                        DefectReasonCode = cboDefectReason.SelectedValue.ToString();
                    }

                    string ImportantGrade = string.Empty;
                    if (cboImportantGrade.SelectedIndex != -1)
                    {
                        ImportantGrade = cboImportantGrade.SelectedValue.ToString();
                    }

                    string CriticalGrade = string.Empty;
                    if (cboCriticalGrade.SelectedIndex != -1)
                    {
                        CriticalGrade = cboCriticalGrade.SelectedValue.ToString();
                    }

                    string AcptDate = string.Empty;
                    if (chkAcptDate.IsChecked == true) { AcptDate = dtpAcptDate.Text.Substring(0, 10).Replace("-", ""); }

                    string DefectRespectDate = string.Empty;
                    if (chkDefectRespectDate.IsChecked == true) { DefectRespectDate = dtpDefectRespectDate.Text.Substring(0, 10).Replace("-", ""); }

                    string CorrExpectDate = string.Empty;
                    if (chkCorrExpectDate.IsChecked == true) { CorrExpectDate = dtpCorrExpectDate.Text.Substring(0, 10).Replace("-", ""); }

                    string CorrCompDate = string.Empty;
                    if (chkCorrCompDate.IsChecked == true) { CorrCompDate = dtpCorrCompDate.Text.Substring(0, 10).Replace("-", ""); }

                    string CorrRespectDate = string.Empty;
                    if (chkCorrRespectDate.IsChecked == true) { CorrRespectDate = dtpCorrRespectDate.Text.Substring(0, 10).Replace("-", ""); }

                    string CorrProdFromDate = string.Empty;
                    string CorrProdToDate = string.Empty;
                    if (chkCorrProdDate.IsChecked == true)
                    {
                        CorrProdFromDate = dtpCorrProdFromDate.Text.Substring(0, 10).Replace("-", "");
                        CorrProdToDate = dtpCorrProdToDate.Text.Substring(0, 10).Replace("-", "");
                    }

                    string DefectRespectSymtomCode = string.Empty;
                    if (txtDefectRespectSymptom.Tag != null)
                    {
                        DefectRespectSymtomCode = txtDefectRespectSymptom.Tag.ToString();
                    }

                    string DefectRespectReasonCode = string.Empty;
                    if (cboDefectRespectReason.SelectedIndex != -1)
                    {
                        DefectRespectReasonCode = cboDefectRespectReason.SelectedValue.ToString();
                    }



                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("DefectID", txtDefectID.Text);
                    sqlParameter.Add("DefectNo", txtDefectNo_InGroupBox.Text);
                    sqlParameter.Add("CustomID", txtCustomer_InGroupBox.Tag.ToString());        //필수기입 체크완료
                    sqlParameter.Add("ArticleID", txtArticle_InGroupBox.Tag.ToString());        //필수기입 체크완료
                    sqlParameter.Add("BuyerModelID", "");

                    sqlParameter.Add("OccurDate", OccurDate);
                    sqlParameter.Add("ReplyReqDate", ReplyReqDate);
                    sqlParameter.Add("NotifyDate", NotifyDate);
                    sqlParameter.Add("ReplyDate", ReplyDate);
                    sqlParameter.Add("DefectCheckMan", txtDefectCheckMan.Text);

                    sqlParameter.Add("NotifyMan", txtNotifyMan.Text);
                    sqlParameter.Add("DefectQty", DefectQty);          //numeric 
                    sqlParameter.Add("ReDefectQty", ReDefectQty);      //numeric 
                    sqlParameter.Add("DefectSymtomCode", DefectSymtomCode);
                    sqlParameter.Add("DefectReasonCode", DefectReasonCode);

                    sqlParameter.Add("ImportantGrade", ImportantGrade);
                    sqlParameter.Add("OccurProcess", txtOccurProcess.Text);
                    sqlParameter.Add("OccurMachine", txtOccurMachine.Text);
                    sqlParameter.Add("CriticalGrade", CriticalGrade);
                    sqlParameter.Add("DefectContents", txtDefectContent_InGroupBox.Text);

                    sqlParameter.Add("AcptDate", AcptDate);
                    sqlParameter.Add("AcptMan", txtAcptMan.Text);
                    sqlParameter.Add("DefectRespectDate", DefectRespectDate);
                    sqlParameter.Add("DefectRespectMan", txtDefectRespectMan.Text);
                    sqlParameter.Add("DefectRespectContents", txtDefectRespectContent.Text);

                    sqlParameter.Add("CorrExpectDate", CorrExpectDate);
                    sqlParameter.Add("CorrCompDate", CorrCompDate);
                    sqlParameter.Add("CorrRespectDate", CorrRespectDate);
                    sqlParameter.Add("CorrRespectMan", txtCorrRespectMan.Text);
                    sqlParameter.Add("CorrProdFromDate", CorrProdFromDate);

                    sqlParameter.Add("CorrProdToDate", CorrProdToDate);
                    sqlParameter.Add("DefectRespectSymtomCode", DefectRespectSymtomCode);
                    sqlParameter.Add("DefectRespectReasonCode", DefectRespectReasonCode);
                    sqlParameter.Add("CorrContents", txtCorrContent.Text);
                    sqlParameter.Add("CorrEOChangeYN", "");
                    sqlParameter.Add("Corr4MChangeYN", "");

                    sqlParameter.Add("CorrDesignChangeYN", "");
                    sqlParameter.Add("CorrDesignChangeContents", "");
                    sqlParameter.Add("CloseYN", cboCloseYN_InGroupBox.SelectedValue.ToString());
                    sqlParameter.Add("DvlDefectYN", cboDvlYN_InGroupBox.SelectedValue.ToString());
                    sqlParameter.Add("ReasonImput", txtReasonImput.Text);

                    sqlParameter.Add("AttPath1", "");
                    sqlParameter.Add("AttFile1", "");
                    sqlParameter.Add("AttPath2", "");
                    sqlParameter.Add("AttFile2", "");
                    sqlParameter.Add("AttPath3", "");
                    sqlParameter.Add("AttFile3", "");
                    sqlParameter.Add("UserID", MainWindow.CurrentUser);


                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_Qul_uInspectCustom";
                    pro1.OutputUseYN = "N";
                    pro1.OutputName = "DefectID";
                    pro1.OutputLength = "10";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"C");
                    if (Confirm[0] == "success")
                    {
                        bool AttachYesNo = false;

                        if (txtAttFile1.Text != string.Empty || txtAttFile2.Text != string.Empty || txtAttFile3.Text != string.Empty)       //첨부파일 1
                        {
                            if (FTP_Save_File(listFtpFile, txtDefectID.Text))
                            {
                                if (!txtAttFile1.Text.Equals(string.Empty)) { txtAttPath1.Text = "/ImageData/CustomDefect/" + txtDefectID.Text; }
                                if (!txtAttFile2.Text.Equals(string.Empty)) { txtAttPath2.Text = "/ImageData/CustomDefect/" + txtDefectID.Text; }
                                if (!txtAttFile3.Text.Equals(string.Empty)) { txtAttPath3.Text = "/ImageData/CustomDefect/" + txtDefectID.Text; }

                                AttachYesNo = true;
                            }
                            else
                            { MessageBox.Show("데이터 수정이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }

                            if (AttachYesNo == true) { AttachFileUpdate(txtDefectID.Text); }      //첨부문서 정보 DB 업데이트.
                        }

                    }
                    else
                    {
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                        //return false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        //삭제.
        private void DeleteData()
        {
            try
            {

                string DefectID = txtDefectID.Text;
                if (DefectID == "")
                {
                    MessageBox.Show("삭제대상이 정확하지 않습니다. 불량ID를 확인해 주세요.");
                    return;
                }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("DefectID", DefectID);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Qul_dInspectCustom", sqlParameter, "D");
                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("이상발생, 관리자에게 문의하세요.");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }


        // 1) 첨부문서가 있을경우, 2) FTP에 정상적으로 업로드가 완료된 경우.  >> DB에 정보 업데이트 
        private void AttachFileUpdate(string ID)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

            sqlParameter.Clear();
            sqlParameter.Add("DefectID", ID);

            sqlParameter.Add("AttPath1", txtAttPath1.Text);
            sqlParameter.Add("AttFile1", txtAttFile1.Text);
            sqlParameter.Add("AttPath2", txtAttPath2.Text);
            sqlParameter.Add("AttFile2", txtAttFile2.Text);
            sqlParameter.Add("AttPath3", txtAttPath3.Text);
            sqlParameter.Add("AttFile3", txtAttFile3.Text);
            sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

            string[] result = DataStore.Instance.ExecuteProcedure("xp_Qul_uInspectCustom_Ftp", sqlParameter, false);
            if (!result[0].Equals("success"))
            {
                MessageBox.Show("이상발생, 관리자에게 문의하세요");
            }
        }

        #endregion


        // 닫기 버튼 클릭.
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            int i = 0;
            foreach (MenuViewModel mvm in MainWindow.mMenulist)
            {
                if (mvm.subProgramID.ToString().Contains("MDI"))
                {
                    if (this.ToString().Equals((mvm.subProgramID as MdiChild).Content.ToString()))
                    {
                        (MainWindow.mMenulist[i].subProgramID as MdiChild).Close();
                    }
                }
                i++;
            }
        }


        #region 엑셀
        // 엑셀.
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            //if (dgdCustomDefect.Items.Count < 1)
            //{
            //    MessageBox.Show("먼저 검색해 주세요.");
            //    return;
            //}

            DataTable dt = null;
            string Name = string.Empty;
            Lib lib2 = new Lib();

            string[] lst = new string[2];
            lst[0] = "고객불량리스트";
            lst[1] = dgdCustomDefect.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdCustomDefect.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib2.DataGridToDTinHidden(dgdCustomDefect);
                    //dt = lib.DataGridToDTinHidden(dgdCustomDefect);
                    else
                        dt = lib2.DataGirdToDataTable(dgdCustomDefect);
                    //dt = lib.DataGirdToDataTable(dgdCustomDefect);

                    Name = dgdCustomDefect.Name;

                    if (lib2.GenerateExcel(dt, Name))
                    {
                        lib2.excel.Visible = true;
                        lib2.ReleaseExcelObject(lib2.excel);
                    }
                    else
                        return;
                }
                else
                {
                    if (dt != null)
                    {
                        dt.Clear();
                    }
                }
            }
            lib2 = null;
        }



        #endregion


        #region FTP // 파일첨부 관련 (등록, 내려받기, 삭제))

        //FTP_ 파일 등록하기.
        private void btnFileEnroll_Click(object sender, RoutedEventArgs e)
        {
            // (버튼)sender 마다 tag를 달자.
            string ClickPoint = ((Button)sender).Tag.ToString();
            string[] strTemp = null;
            Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();

            OFdlg.DefaultExt = ".xls";
            //OFdlg.Filter = "Office Files | *.doc; *.xls; *.pptx; *.pdf | All Files|*.*";
            OFdlg.Filter = MainWindow.OFdlg_Filter;


            Nullable<bool> result = OFdlg.ShowDialog();
            if (result == true)
            {
                if (ClickPoint == "1") { FullPath1 = OFdlg.FileName; }  //긴 경로(FULL 사이즈)
                if (ClickPoint == "2") { FullPath2 = OFdlg.FileName; }
                if (ClickPoint == "3") { FullPath3 = OFdlg.FileName; }

                string AttachFileName = OFdlg.SafeFileName;  //명.
                string AttachFilePath = string.Empty;       // 경로

                if (ClickPoint == "1") { AttachFilePath = FullPath1.Replace(AttachFileName, ""); }
                if (ClickPoint == "2") { AttachFilePath = FullPath2.Replace(AttachFileName, ""); }
                if (ClickPoint == "3") { AttachFilePath = FullPath3.Replace(AttachFileName, ""); }

                StreamReader sr = new StreamReader(OFdlg.FileName);
                long File_size = sr.BaseStream.Length;
                if (sr.BaseStream.Length > (2048 * 1000))
                {
                    // 업로드 파일 사이즈범위 초과
                    MessageBox.Show("밀시트의 파일사이즈가 2M byte를 초과하였습니다.");
                    sr.Close();
                    return;
                }
                if (ClickPoint == "1")
                {
                    txtAttFile1.Text = AttachFileName;
                    txtAttPath1.Text = AttachFilePath.ToString();
                }
                else if (ClickPoint == "2")
                {
                    txtAttFile2.Text = AttachFileName;
                    txtAttPath2.Text = AttachFilePath.ToString();
                }
                else if (ClickPoint == "3")
                {
                    txtAttFile3.Text = AttachFileName;
                    txtAttPath3.Text = AttachFilePath.ToString();
                }

                strTemp = new string[] { AttachFileName, AttachFilePath.ToString() };
                listFtpFile.Add(strTemp);

            }
        }

        // 파일 저장하기.
        private bool FTP_Save_File(List<string[]> listStrArrayFileInfo, string MakeFolderName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

            List<string[]> UpdateFilesInfo = new List<string[]>();
            string[] fileListSimple;
            string[] fileListDetail = null;
            fileListSimple = _ftp.directoryListSimple("", Encoding.Default);

            // 기존 폴더 확인작업.
            bool MakeFolder = false;
            MakeFolder = FolderInfoAndFlag(fileListSimple, MakeFolderName.Trim());

            if (MakeFolder == false)        // 같은 아이를 찾지 못한경우,
            {
                //MIL 폴더에 InspectionID로 저장
                if (_ftp.createDirectory(MakeFolderName.Trim()) == false)
                {
                    MessageBox.Show("업로드를 위한 폴더를 생성할 수 없습니다.");
                    return false;
                }
            }
            else
            {
                fileListDetail = _ftp.directoryListSimple(MakeFolderName.Trim(), Encoding.Default);
            }
            for (int i = 0; i < listStrArrayFileInfo.Count; i++)
            {
                bool flag = true;

                if (fileListDetail != null)
                {
                    foreach (string compare in fileListDetail)
                    {
                        if (compare.Equals(listStrArrayFileInfo[i][0]))
                        {
                            flag = false;
                            break;
                        }
                    }
                }

                if (flag)
                {
                    listStrArrayFileInfo[i][0] = MakeFolderName.Trim() + "/" + listStrArrayFileInfo[i][0];
                    UpdateFilesInfo.Add(listStrArrayFileInfo[i]);
                }
            }
            if (UpdateFilesInfo.Count > 0)
            {
                if (!_ftp.UploadTempFilesToFTP(UpdateFilesInfo))
                {
                    MessageBox.Show("파일업로드에 실패하였습니다.");
                    return false;
                }
            }
            return true;
        }




        // 파일 내려받기.
        private void btnFileDown_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 다운로드 하시겠습니까?", "다운로드 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                //버튼 태그값.
                string ClickPoint = ((Button)sender).Tag.ToString();

                if ((ClickPoint == "1") && (txtAttPath1.Text == string.Empty))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }
                if ((ClickPoint == "2") && (txtAttPath2.Text == string.Empty))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }
                if ((ClickPoint == "3") && (txtAttPath3.Text == string.Empty))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }

                var ViewReceiver = dgdCustomDefect.SelectedItem as Win_Qul_CustomDefect_U_View;
                if (ViewReceiver != null)
                {
                    if (ClickPoint == "1")
                    {
                        FTP_DownLoadFile(ViewReceiver.AttPath1, ViewReceiver.DefectID, ViewReceiver.AttFile1);
                    }
                    else if (ClickPoint == "2")
                    {
                        FTP_DownLoadFile(ViewReceiver.AttPath2, ViewReceiver.DefectID, ViewReceiver.AttFile2);
                    }
                    else if (ClickPoint == "3")
                    {
                        FTP_DownLoadFile(ViewReceiver.AttPath3, ViewReceiver.DefectID, ViewReceiver.AttFile3);
                    }
                }
            }
        }

        //다운로드
        private void FTP_DownLoadFile(string Path, string FolderName, string ImageName)
        {
            try
            {
                _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

                string[] fileListSimple;
                string[] fileListDetail;

                fileListSimple = _ftp.directoryListSimple("", Encoding.UTF8);

                bool ExistFile = false;

                ExistFile = FolderInfoAndFlag(fileListSimple, FolderName);

                if (ExistFile)
                {
                    ExistFile = false;
                    fileListDetail = _ftp.directoryListSimple(FolderName, Encoding.UTF8);

                    ExistFile = FileInfoAndFlag(fileListDetail, ImageName);

                    if (ExistFile)
                    {
                        string str_remotepath = string.Empty;
                        string str_localpath = string.Empty;

                        str_remotepath = FTP_ADDRESS + '/' + FolderName + '/' + ImageName;
                        str_localpath = LOCAL_DOWN_PATH + "\\" + ImageName;

                        DirectoryInfo DI = new DirectoryInfo(LOCAL_DOWN_PATH);
                        if (DI.Exists)
                        {
                            DI.Create();
                        }

                        FileInfo file = new FileInfo(str_localpath);
                        if (file.Exists)
                        {
                            file.Delete();
                        }

                        _ftp.download(str_remotepath.Substring(str_remotepath.Substring
                            (0, str_remotepath.LastIndexOf("/")).LastIndexOf("/")), str_localpath);

                        ProcessStartInfo proc = new ProcessStartInfo(str_localpath);
                        proc.UseShellExecute = true;
                        Process.Start(proc);
                    }
                    else
                    {
                        MessageBox.Show("파일을 찾을 수 없습니다.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }


        // 파일 삭제하기.
        private void btnFileDel_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 삭제 하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                string ClickPoint = ((Button)sender).Tag.ToString();

                if ((ClickPoint == "1") && (txtAttPath1.Text != string.Empty))
                {
                    txtAttFile1.Text = string.Empty;
                    txtAttPath1.Text = string.Empty;
                }
                if ((ClickPoint == "2") && (txtAttPath2.Text != string.Empty))
                {
                    txtAttFile2.Text = string.Empty;
                    txtAttPath2.Text = string.Empty;
                }
                if ((ClickPoint == "3") && (txtAttPath3.Text != string.Empty))
                {
                    txtAttFile3.Text = string.Empty;
                    txtAttPath3.Text = string.Empty;
                }
            }
        }


        /// <summary>
        /// 해당영역에 폴더가 있는지 확인
        /// </summary>
        bool FolderInfoAndFlag(string[] strFolderList, string FolderName)
        {
            bool flag = false;
            foreach (string FolderList in strFolderList)
            {
                if (FolderList == FolderName)
                {
                    flag = true;
                    break;
                }
            }
            return flag;
        }

        /// <summary>
        /// 해당영역에 파일 있는지 확인
        /// </summary>
        bool FileInfoAndFlag(string[] strFileList, string FileName)
        {
            bool flag = false;
            foreach (string FileList in strFileList)
            {
                if (FileList == FileName)
                {
                    flag = true;
                    break;
                }
            }
            return flag;
        }

        #endregion


        #region 텍스트박스 엔터 키 이벤트


        private void txtDefectNo_InGroupBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtCustomer_InGroupBox.Focus();
            }
        }
        private void txtCustomer_InGroupBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnCustomer_InGroupBox_Click(null, null);
                txtArticle_InGroupBox.Focus();
            }
        }
        private void txtArticle_InGroupBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnArticle_InGroupBox_Click(null, null);
            }
        }
        private void dtpOccurDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                dtpOccurDate.IsDropDownOpen = true;
            }
        }
        private void chkReplyReqDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                chkNotifyDate.Focus();
            }
        }
        private void chkReplyReqDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                dtpReplyReqDate.Focus();
                dtpReplyReqDate.IsDropDownOpen = true;
            }
        }
        private void chkNotifyDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                chkReplyDate.Focus();
            }
        }
        private void chkNotifyDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                dtpNotifyDate.Focus();
                dtpNotifyDate.IsDropDownOpen = true;
            }
        }
        private void chkReplyDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtDefectCheckMan.Focus();
            }
        }
        private void chkReplyDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                dtpReplyDate.Focus();
                dtpReplyDate.IsDropDownOpen = true;
            }
        }
        private void txtDefectSymptom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnDefectSymptom_Click(null, null);
                cboDefectReason.IsDropDownOpen = true;
            }
        }
        private void cboDefectReason_DropDownClosed(object sender, EventArgs e)
        {
            lib.SendK(Key.Tab, this);
            cboImportantGrade.IsDropDownOpen = true;
        }
        private void cboImportantGrade_DropDownClosed(object sender, EventArgs e)
        {
            lib.SendK(Key.Tab, this);
            cboCriticalGrade.IsDropDownOpen = true;
        }
        private void cboCriticalGrade_DropDownClosed(object sender, EventArgs e)
        {
            lib.SendK(Key.Tab, this);
        }
        private void chkDefectRespectDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtAcptMan.Focus();
            }
        }
        private void chkDefectRespectDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                dtpDefectRespectDate.Focus();
                dtpDefectRespectDate.IsDropDownOpen = true;
            }
        }
        private void chkCorrCompDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                chkCorrRespectDate.Focus();
            }
        }
        private void chkCorrCompDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                dtpCorrCompDate.Focus();
                dtpCorrCompDate.IsDropDownOpen = true;
            }
        }
        private void chkCorrRespectDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtCorrRespectMan.Focus();
            }
        }
        private void chkCorrRespectDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                dtpCorrRespectDate.Focus();
                dtpCorrRespectDate.IsDropDownOpen = true;
            }
        }
        private void chkCorrProdDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtDefectRespectSymptom.Focus();
            }
        }
        private void chkCorrProdDate_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                dtpCorrProdFromDate.Focus();
                dtpCorrProdFromDate.IsDropDownOpen = true;
            }
        }
        private void dtpCorrProdFromDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            dtpCorrProdToDate.Focus();
            dtpCorrProdToDate.IsDropDownOpen = true;
        }
        private void txtDefectRespectSymptom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnDefectRespectSymptom_Click(null, null);
                cboDefectRespectReason.IsDropDownOpen = true;
            }
        }
        private void cboDefectRespectReason_DropDownClosed(object sender, EventArgs e)
        {
            lib.SendK(Key.Tab, this);
        }
        private void dtpOccurDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            chkReplyReqDate.Focus();
        }

        // 엔터 키를 통한 탭 인덱스 키 이동.
        private void EnterMove_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                lib.SendK(Key.Tab, this);
            }
        }

        // 엔터 키를 통한 _ 달력 닫기.
        private void EnterMove_CalendarClosed(object sender, RoutedEventArgs e)
        {
            lib.SendK(Key.Tab, this);
        }





        #endregion

        private void DataGrid_SizeChange(object sender, SizeChangedEventArgs e)
        {
            DataGrid dgs = sender as DataGrid;
            if (dgs.ColumnHeaderHeight == 0)
            {
                dgs.ColumnHeaderHeight = 1;
            }
            double a = e.NewSize.Height / 100;
            double b = e.PreviousSize.Height / 100;
            double c = a / b;

            if (c != double.PositiveInfinity && c != 0 && double.IsNaN(c) == false)
            {
                dgs.ColumnHeaderHeight = dgs.ColumnHeaderHeight * c;
                dgs.FontSize = dgs.FontSize * c;
            }
        }

        private void rbnOccurDay_Click(object sender, RoutedEventArgs e)
        {
            rbnOccurDay.IsChecked = true;
            rbnReceiptDay.IsChecked = false;
            rbnActionDay.IsChecked = false;
        }
        private void rbnReceiptDay_Click(object sender, RoutedEventArgs e)
        {
            rbnOccurDay.IsChecked = false;
            rbnReceiptDay.IsChecked = true;
            rbnActionDay.IsChecked = false;
        }

        private void rbnActionDay_Click(object sender, RoutedEventArgs e)
        {
            rbnOccurDay.IsChecked = false;
            rbnReceiptDay.IsChecked = false;
            rbnActionDay.IsChecked = true;
        }

        private void rbnChoice1_Click(object sender, RoutedEventArgs e)
        {
            rbnChoice1.IsChecked = true;
            rbnChoice2.IsChecked = false;
        }

        private void rbnChoice2_Click(object sender, RoutedEventArgs e)
        {
            rbnChoice1.IsChecked = false;
            rbnChoice2.IsChecked = true;
        }
    }

    class Win_Qul_CustomDefect_U_View : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }

        // CustomDefect 조회 값.    
        public string DefectID { get; set; }
        public string DefectNo { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string ArticleID { get; set; }

        public string Sabun { get; set; }

        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string BuyerModelID { get; set; }
        public string BuyerModelName { get; set; }
        public string OccurDate { get; set; }

        public string ReplyReqDate { get; set; }
        public string NotifyDate { get; set; }
        public string ReplyDate { get; set; }
        public string DefectCheckMan { get; set; }
        public string NotifyMan { get; set; }

        public string DefectQty { get; set; }
        public string ReDefectQty { get; set; }
        public string DefectSymtomCode { get; set; }
        public string DefectSymtom { get; set; }
        public string DefectReasonCode { get; set; }

        public string DefectReason { get; set; }
        public string ImportantGrade { get; set; }
        public string ImportantGradeName { get; set; }
        public string OccurProcess { get; set; }
        public string OccurMachine { get; set; }

        public string CriticalGrade { get; set; }
        public string CriticalGradeName { get; set; }
        public string DefectContents { get; set; }
        public string ReasonImput { get; set; }
        public string AcptDate { get; set; }

        public string AcptMan { get; set; }
        public string DefectRespectDate { get; set; }
        public string DefectRespectMan { get; set; }
        public string DefectRespectContents { get; set; }
        public string CorrExpectDate { get; set; }

        public string CorrCompDate { get; set; }
        public string CorrRespectDate { get; set; }
        public string CorrRespectMan { get; set; }
        public string CorrProdFromDate { get; set; }
        public string CorrProdToDate { get; set; }

        public string DefectRespectSymtomCode { get; set; }
        public string DefectRespectSymtom { get; set; }
        public string DefectRespectReasonCode { get; set; }
        public string DefectRespectReason { get; set; }
        public string CorrContents { get; set; }

        public string CorrEOChangeYN { get; set; }
        public string Corr4MChangeYN { get; set; }
        public string CorrDesignChangeYN { get; set; }
        public string CorrDesignChangeContents { get; set; }
        public string CloseYN { get; set; }

        public string AttPath1 { get; set; }
        public string AttFile1 { get; set; }
        public string AttPath2 { get; set; }
        public string AttFile2 { get; set; }
        public string AttPath3 { get; set; }

        public string AttFile3 { get; set; }
        public string CreateUserID { get; set; }

        // 개발/양산 구별자
        public string DvlYN { get; set; }

        public Win_Qul_CustomDefect_U_View Copy()
        {
            return (Win_Qul_CustomDefect_U_View)this.MemberwiseClone();
        }


    }

}
