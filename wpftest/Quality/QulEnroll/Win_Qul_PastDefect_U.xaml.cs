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
using System.Windows.Media.Imaging;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_PastDefect_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_PastDefect_U : UserControl
    {
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        // 추가저장인지 / 수정저장인지 구별하는 용도입니다.
        string ButtonTag = string.Empty;

        int Wh_Ar_SelectedLastIndex = 0;        // 그리드 마지막 선택 줄 임시저장 그릇


        // FTP 활용모음.
        List<string[]> listFtpFile = new List<string[]>();
        private FTP_EX _ftp = null;

        string FullPath1 = string.Empty;
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/PastDefect";
        //string FTP_ADDRESS = "ftp://HKserver:210/ImageData/PastDefect";
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/PastDefect";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";


        public Win_Qul_PastDefect_U()
        {
            InitializeComponent();
        }

        // 첫 로드시.
        private void Win_Qul_PastDefect_U_Loaded(object sender, RoutedEventArgs e)
        {
            First_Step();
            ComboBoxSetting();
        }


        #region 첫 스텝 // 날짜용 버튼 // 조회용 체크박스 세팅 
        // 첫 스텝
        private void First_Step()
        {
            chkSearchDay.IsChecked = true;
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            txtCustomer.IsEnabled = false;
            btnCustomer.IsEnabled = false;
            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
            txtCarModel.IsEnabled = false;
            btnCarModel.IsEnabled = false;
            txtDefectSubject.IsEnabled = false;
            txtDefectSymptom.IsEnabled = false;
            btnDefectSymptom.IsEnabled = false;
            cboCloseYN.IsEnabled = false;

            rbnChoice1.IsChecked = true;

            grbPastDefectBox.IsEnabled = false;

            grbContentNImageBox.IsEnabled = true;


            txtComments.IsEnabled = false;
            txtDefectImageFile.IsEnabled = false;
            txtDefectImagePath.IsEnabled = false;
            btnFileEnroll.IsEnabled = false;
            btnFileDel.IsEnabled = false;
            btnFileSee.IsEnabled = false;

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

        //고객사
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
        //고객사
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
        //차종
        private void chkCarModel_Click(object sender, RoutedEventArgs e)
        {
            if (chkCarModel.IsChecked == true)
            {
                txtCarModel.IsEnabled = true;
                txtCarModel.Focus();
                btnCarModel.IsEnabled = true;
            }
            else
            {
                txtCarModel.IsEnabled = false;
                btnCarModel.IsEnabled = false;
            }
        }
        //차종
        private void chkCarModel_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkCarModel.IsChecked == true)
            {
                chkCarModel.IsChecked = false;
                txtCarModel.IsEnabled = false;
                btnCarModel.IsEnabled = false;
            }
            else
            {
                chkCarModel.IsChecked = true;
                txtCarModel.IsEnabled = true;
                btnCarModel.IsEnabled = true;
                txtCarModel.Focus();
            }
        }
        //불량제목
        private void chkDefectSubject_Click(object sender, RoutedEventArgs e)
        {
            if (chkDefectSubject.IsChecked == true)
            {
                txtDefectSubject.IsEnabled = true;
                txtDefectSubject.Focus();
            }
            else
            {
                txtDefectSubject.IsEnabled = false;
            }
        }
        //불량제목
        private void chkDefectSubject_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkDefectSubject.IsChecked == true)
            {
                chkDefectSubject.IsChecked = false;
                txtDefectSubject.IsEnabled = false;
            }
            else
            {
                chkDefectSubject.IsChecked = true;
                txtDefectSubject.IsEnabled = true;
                txtDefectSubject.Focus();
            }
        }
        //불량유형
        private void chkDefectSymptom_Click(object sender, RoutedEventArgs e)
        {
            if (chkDefectSymptom.IsChecked == true)
            {
                txtDefectSymptom.IsEnabled = true;
                txtDefectSymptom.Focus();
                btnDefectSymptom.IsEnabled = true;
            }
            else
            {
                txtDefectSymptom.IsEnabled = false;
                btnDefectSymptom.IsEnabled = false;
            }
        }
        //불량유형
        private void chkDefectSymptom_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkDefectSymptom.IsChecked == true)
            {
                chkDefectSymptom.IsChecked = false;
                txtDefectSymptom.IsEnabled = false;
                btnDefectSymptom.IsEnabled = false;
            }
            else
            {
                chkDefectSymptom.IsChecked = true;
                txtDefectSymptom.IsEnabled = true;
                txtDefectSymptom.Focus();
                btnDefectSymptom.IsEnabled = true;
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

        #endregion


        #region 콤보박스 세팅

        //콤보박스 세팅
        private void ComboBoxSetting()
        {
            cboCloseYN.Items.Clear();
            cboCloseYN_InGroupBox.Items.Clear();
            cboDefectReason.Items.Clear();
            cboOccurStep.Items.Clear();
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
            this.cboCloseYN_InGroupBox.SelectedIndex = -1;

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

            /////////////////////////////////////////////////////

            ObservableCollection<CodeView> cbDefectReason = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "QULDFRSN", "Y", "", "");
            ObservableCollection<CodeView> cbOccurStep = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "QULSTEP", "Y", "", "");

            this.cboDefectReason.ItemsSource = cbDefectReason;
            this.cboDefectReason.DisplayMemberPath = "code_name";
            this.cboDefectReason.SelectedValuePath = "code_id";
            this.cboDefectReason.SelectedIndex = -1;

            this.cboOccurStep.ItemsSource = cbOccurStep;
            this.cboOccurStep.DisplayMemberPath = "code_name";
            this.cboOccurStep.SelectedValuePath = "code_id";
            this.cboOccurStep.SelectedIndex = -1;

        }

        #endregion


        #region 플러스파인더

        // 플러스파인더 _ 거래처 찾기.
        private void btnCustomer_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomer, 0, "");
        }
        // 플러스파인더 _ 그룹박스 내 거래처 찾기.
        private void btnCustomer_InGroupBox_Click(object sender, RoutedEventArgs e)
        {
            //if ((txtArticle_InGroupBox.Tag != null) && (txtArticle_InGroupBox.Text.Length > 0))
            //{
            //    pf.ReturnCode(txtCustomer_InGroupBox, 65, txtArticle_InGroupBox.Tag.ToString());
            //}
            //else { pf.ReturnCode(txtCustomer_InGroupBox, 0, ""); }

            pf.ReturnCode(txtCustomer_InGroupBox, 0, "");

            lib.SendK(Key.Tab, this);
        }


        // 플러스파인더 _ 품명 찾기
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticle, 84, txtArticle.Text);
        }

        // 키다운 _ 품명 찾기(품번으로 변경요청, 2020.03.23, 장가빈)
        private void TxtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticle, 84, txtArticle.Text);
            }
        }

        // 플러스파인더 _ 그룹박스 내 품명 찾기.
        private void btnArticle_InGroupBox_Click(object sender, RoutedEventArgs e)
        {
            //if ((txtCustomer_InGroupBox.Tag != null) && (txtCustomer_InGroupBox.Text.Length > 0))
            //{
            //    pf.ReturnCode(txtArticle_InGroupBox, 64, txtCustomer_InGroupBox.Tag.ToString());
            //}
            //else
            //{
            //    pf.ReturnCode(txtArticle_InGroupBox, 1, "");
            //}

            pf.ReturnCode(txtArticle_InGroupBox, 1, "");

            if (txtArticle_InGroupBox.Text.Length > 0 && txtArticle_InGroupBox.Tag != null)
            {
                string ArticleID = txtArticle_InGroupBox.Tag.ToString();
                txtArticleID_InGroupBox.Text = ArticleID;
                Article_InGroupBox_OtherSearch(ArticleID);  // 품명정보 바탕으로 품번 자동으로 찾아 뿌리기.
            }
            lib.SendK(Key.Tab, this);
        }


        // 플러스파인더 _ 차종 찾기.
        private void btnCarModel_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCarModel, 28, "");
        }
        // 플러스파인더 _ 그룹박스 내 차종 찾기.
        private void btnBuyerModel_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtBuyerModel, 28, "");
            lib.SendK(Key.Tab, this);
        }


        // 플러스파인더 _ 불량유형 찾기.
        private void btnDefectSymptom_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtDefectSymptom, 3, "");
        }
        // 플러스파인더 _ 그룹박스 내 불량유형 찾기.
        private void btnDefectSymptom_InGroupBox_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtDefectSymptom_InGroupBox, 3, "");
            lib.SendK(Key.Tab, this);
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

            grbPastDefectBox.IsEnabled = false;

            //그룹박스 내부 컨텐츠 일부 활성화 오픈.
            grbContentNImageBox.IsEnabled = true;
            txtComments.IsEnabled = false;
            txtDefectImageFile.IsEnabled = false;
            txtDefectImagePath.IsEnabled = false;
            btnFileEnroll.IsEnabled = false;
            btnFileDel.IsEnabled = false;
            btnFileSee.IsEnabled = false;

            txtPastDefectID.IsEnabled = true;

            dgdPastDefect.IsHitTestVisible = true; //메인그리드 사용가능.
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

            grbPastDefectBox.IsEnabled = true;

            grbContentNImageBox.IsEnabled = true;
            txtComments.IsEnabled = true;
            txtDefectImageFile.IsEnabled = true;
            txtDefectImagePath.IsEnabled = true;
            btnFileEnroll.IsEnabled = true;
            btnFileDel.IsEnabled = true;
            btnFileSee.IsEnabled = false;

            txtPastDefectID.IsReadOnly = true;

            dgdPastDefect.IsHitTestVisible = false; //메인그리드 못 건드리게.

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

            if (dgdPastDefect.Items.Count > 0)
            {
                dgdPastDefect.SelectedIndex = selectedIndex;
            }
        }

        private void FillGrid()
        {
            if (dgdPastDefect.Items.Count > 0)
            {
                dgdPastDefect.Items.Clear();
            }


            try
            {
                int nchkDate = 0;
                string FromDate = string.Empty;
                string ToDate = string.Empty;
                if (chkSearchDay.IsChecked == true)
                {
                    nchkDate = 1;
                    //FromDate = dtpFromDate.ToString().Substring(0, 10).Replace("-", "");
                    //ToDate = dtpToDate.ToString().Substring(0, 10).Replace("-", "");
                    FromDate = dtpFromDate.SelectedDate.Value.ToString("yyyyMMdd");
                    ToDate = dtpToDate.SelectedDate.Value.ToString("yyyyMMdd");
                }
                string sDvlYN = "Y";
                if (rbnChoice2.IsChecked == true)
                {
                    sDvlYN = "N";
                }

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

                int nchkArticleID = 0;
                string ArticleID = string.Empty;
                string Article = string.Empty;
                if (chkArticle.IsChecked == true)
                {
                    if (txtArticle.Tag == null)
                    {
                        txtArticle.Tag = "";
                        if (txtArticle.Text.Length > 0)
                        {
                            nchkArticleID = 2;
                            Article = txtArticle.Text;
                        }
                    }
                    else
                    {
                        nchkArticleID = 1;
                        ArticleID = txtArticle.Tag.ToString();
                    }
                }

                int nchkBuyerModel = 0;
                string BuyerModel = string.Empty;
                string BuyerModelName = string.Empty;

                if (chkCarModel.IsChecked == true)
                {
                    if (txtCarModel.Tag == null)
                    {
                        txtCarModel.Tag = "";
                        if (txtCarModel.Text.Length > 0)
                        {
                            nchkBuyerModel = 2;
                            BuyerModelName = txtCarModel.Text;
                        }
                    }
                    else
                    {
                        nchkBuyerModel = 1;
                        BuyerModel = txtCarModel.Tag.ToString();
                    }
                }

                int nchkSubject = 0;
                string Subject = string.Empty;
                if (chkDefectSubject.IsChecked == true)
                {
                    nchkSubject = 1;
                    Subject = txtDefectSubject.Text;
                }

                int nchkCloseYN = 0;
                string CloseYN = string.Empty;
                if (chkCloseYN.IsChecked == true)
                {
                    nchkCloseYN = 1;
                    CloseYN = cboCloseYN.SelectedValue.ToString();
                }

                int nchkSymtom = 0;
                string DefectSymtomcode = string.Empty;
                string DefectSymptomName = string.Empty;

                if (chkDefectSymptom.IsChecked == true)
                {
                    if (txtDefectSymptom.Tag == null)
                    {
                        txtDefectSymptom.Tag = "";
                        if (txtDefectSymptom.Text.Length > 0)
                        {
                            nchkSymtom = 2;
                            DefectSymptomName = txtDefectSymptom.Text;
                        }
                    }
                    else
                    {
                        nchkSymtom = 1;
                        DefectSymtomcode = txtDefectSymptom.Tag.ToString();
                    }
                }


                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("nchkDate", nchkDate);               //int
                sqlParameter.Add("FromDate", FromDate);
                sqlParameter.Add("ToDate", ToDate);
                sqlParameter.Add("sDvlYN", sDvlYN);
                sqlParameter.Add("nchkCustomID", nchkCustomID);           //int

                sqlParameter.Add("CustomID", CustomID);
                //sqlParameter.Add("Custom", Custom);
                sqlParameter.Add("nchkArticleID", 0); // nchkArticleID);          //int
                sqlParameter.Add("ArticleID", ""); // ArticleID);
                //sqlParameter.Add("Article", Article);

                sqlParameter.Add("nchkBuyerModel", nchkBuyerModel);         //int
                sqlParameter.Add("BuyerModel", BuyerModel);
                //sqlParameter.Add("BuyerModelName", BuyerModelName);
                sqlParameter.Add("nchkSubject", nchkSubject);            //int
                sqlParameter.Add("Subject", Subject);

                sqlParameter.Add("nchkCloseYN", nchkCloseYN);            //int
                sqlParameter.Add("CloseYN", CloseYN);
                sqlParameter.Add("nchkSymtom", nchkSymtom);             //int
                sqlParameter.Add("DefectSymtomcode", DefectSymtomcode);
                //sqlParameter.Add("DefectSymptomName", DefectSymptomName);
                sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sPastDefect", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        //조회결과가 있다면,
                        //dgdPastDefect.Items.Clear();
                        int i = 1;
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow item in drc)
                        {
                            var Win_Qul_PastDefect_U_Insert = new Win_Qul_PastDefect_U_View()
                            {
                                NUM = i.ToString(),
                                PastDefectID = item["PastDefectID"].ToString(),                 //불량id
                                PastDefectSubject = item["PastDefectSubject"].ToString(),       //불량제목
                                CustomID = item["CustomID"].ToString(),
                                KCustom = item["KCustom"].ToString(),                           //고객사
                                ArticleID = item["ArticleID"].ToString(),

                                Article = item["Article"].ToString(),                           //품명
                                BuyerArticleNo = item["BuyerArticleNo"].ToString(),             //품번
                                BuyerModelID = item["BuyerModelID"].ToString(),
                                BuyerModel = item["BuyerModel"].ToString(),                     //차종
                                OccurDate = item["OccurDate"].ToString(),                       //발생일

                                DefectSymtomCode = item["DefectSymtomCode"].ToString(),
                                DefectSymtom = item["DefectSymtom"].ToString(),                 //불량유형
                                DefectReasonCode = item["DefectReasonCode"].ToString(),
                                DefectReason = item["DefectReason"].ToString(),                 //불량원인
                                OccurStep = item["OccurStep"].ToString(),

                                OccurStepName = item["OccurStepName"].ToString(),               //불량발생단계
                                EffectValidation = item["EffectValidation"].ToString(),         //유효성
                                CloseYN = item["CloseYN"].ToString(),                           //종결
                                Comments = item["Comments"].ToString(),
                                QPoint = item["QPoint"].ToString(),

                                QPointPath = item["QPointPath"].ToString(),
                                QPointYN = item["QPointYN"].ToString(),
                                DvlDefectYN = item["DvlDefectYN"].ToString(),                   // 개발 OR 양산

                            };

                            if (!Lib.Instance.CheckNull(Win_Qul_PastDefect_U_Insert.OccurDate).Equals(string.Empty))
                            {
                                Win_Qul_PastDefect_U_Insert.OccurDate = Lib.Instance.StrDateTimeBar(Win_Qul_PastDefect_U_Insert.OccurDate);
                            }

                            dgdPastDefect.Items.Add(Win_Qul_PastDefect_U_Insert);

                            i++;
                        }
                        tbkIndexCount.Text = "▶검색결과 : " + (i - 1) + "건";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        #endregion


        #region 메인그리드 연동 _ SHOW DATA

        // 그리드 셀 체인지 이벤트 _ show data.
        private void dgdPastDefect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                //this.DataContext = dgdPastDefect.SelectedItem as Win_Qul_PastDefect_U_View;
                var ViewReceiver = dgdPastDefect.SelectedItem as Win_Qul_PastDefect_U_View;

                if (ViewReceiver != null)
                {
                    txtDefectSymptom_InGroupBox.Tag = ViewReceiver.DefectSymtomCode;
                    txtArticle_InGroupBox.Tag = ViewReceiver.ArticleID;
                    txtCustomer_InGroupBox.Tag = ViewReceiver.CustomID;


                    Win_Qul_PastDefect_U_View ProCopy = ViewReceiver.Copy();
                    this.DataContext = ProCopy;


                    if ((ViewReceiver.OccurDate != string.Empty) && (ViewReceiver.OccurDate != null) && (lib.IsNullOrWhiteSpace(ViewReceiver.OccurDate) != true))   //발생일
                    {
                        dtpOccurDate.Text = DateTime.ParseExact(ViewReceiver.OccurDate, "yyyy-MM-dd", null).ToString("yyyy-MM-dd");
                    }
                    else { dtpOccurDate.Text = string.Empty; }

                    if (ViewReceiver.KCustom != "") { txtCustomer_InGroupBox.Tag = ViewReceiver.CustomID; }
                    if (ViewReceiver.Article != "") { txtArticle_InGroupBox.Tag = ViewReceiver.ArticleID; }
                    if (ViewReceiver.DefectSymtom != "") { txtDefectSymptom_InGroupBox.Tag = ViewReceiver.DefectSymtomCode; }
                    if (ViewReceiver.BuyerModelID != "") { txtBuyerModel.Tag = ViewReceiver.BuyerModelID; }

                    if (ViewReceiver.QPointYN == "Y") { chkQPointYN.IsChecked = true; }
                    else if (ViewReceiver.QPointYN == "N") { chkQPointYN.IsChecked = false; }

                    if ((ViewReceiver.QPointPath != string.Empty) && (ViewReceiver.QPointPath != null))
                    {
                        btnFileSee.IsEnabled = true;
                        if (!(SetImage(ViewReceiver.QPoint, ViewReceiver.PastDefectID)))
                        {
                            btnFileSee.IsEnabled = false;
                            imgDefectImage.Source = null;
                        }
                    }
                    else
                    {
                        btnFileSee.IsEnabled = false;
                        imgDefectImage.Source = null;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #endregion


        #region (추가 , 수정 , 삭제 , 저장, 취소) 버튼 이벤트 모음

        // 추가 버튼 클릭 시.
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //1. 그룹박스 데이터 클리어.
                grbBoxDataClear();

                //2. 공통 버튼이벤트
                PublicEnableFalse();
                EventLabel.Content = "자료 입력(추가) 중..";

                dtpOccurDate.Text = DateTime.Now.ToString("yyyy-MM-dd");            //발생일에 대한 기본세팅.(오늘자)
                cboDvlYN_InGroupBox.SelectedIndex = 0;                              //개발/양산 구별 콤보박스 기본세팅.(빈값으로 못두도록.)

                ButtonTag = ((Button)sender).Tag.ToString();

                if (dgdPastDefect.Items.Count > 0)
                {
                    Wh_Ar_SelectedLastIndex = dgdPastDefect.SelectedIndex;
                }
                else
                {
                    Wh_Ar_SelectedLastIndex = 0;
                }
                txtPastDefectSubject.Focus();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        // 수정 버튼 클릭 시.
        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. 수정할 자격은 있는거야? 조회? 데이터 선택??
                if (dgdPastDefect.Items.Count < 1)
                {
                    MessageBox.Show("먼저 검색해 주세요.");
                    return;
                }
                var OBJ = dgdPastDefect.SelectedItem as Win_Qul_PastDefect_U_View;
                if (OBJ == null)
                {
                    MessageBox.Show("수정할 항목이 정확히 선택되지 않았습니다.");
                    return;
                }

                // 2.공통 버튼이벤트
                PublicEnableFalse();
                EventLabel.Content = "자료 입력(수정) 중..";
                Wh_Ar_SelectedLastIndex = dgdPastDefect.SelectedIndex;

                ButtonTag = ((Button)sender).Tag.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        // 삭제 버튼 클릭 시.
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. 삭제할 자격은 있는거야? 조회? 데이터 선택??
                if (dgdPastDefect.Items.Count < 1)
                {
                    MessageBox.Show("먼저 검색해 주세요.");
                    return;
                }
                var OBJ = dgdPastDefect.SelectedItem as Win_Qul_PastDefect_U_View;
                if (OBJ == null)
                {
                    MessageBox.Show("삭제할 항목이 정확히 선택되지 않았습니다.");
                    return;
                }
                MessageBoxResult msgresult = MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
                if (msgresult == MessageBoxResult.Yes)
                {
                    if (dgdPastDefect.Items.Count > 0 && dgdPastDefect.SelectedItem != null)
                    {
                        Wh_Ar_SelectedLastIndex = dgdPastDefect.SelectedIndex;
                    }

                    // 2.  삭제용
                    DeleteData();

                    // 접속 경로
                    _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
                    string[] fileListSimple;
                    fileListSimple = _ftp.directoryListSimple("", Encoding.Default);

                    bool delFtp = FolderInfoAndFlag(fileListSimple, OBJ.PastDefectID);
                    if (delFtp)
                        _ftp.removeDir(OBJ.PastDefectID);

                    dgdPastDefect.Refresh();

                    // 3. 화면정리.
                    Wh_Ar_SelectedLastIndex -= 1;
                    re_Search(Wh_Ar_SelectedLastIndex);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        //저장 버튼 클릭 시.
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. 데이터 기입체크.(항목을 제대로 모두 똑바로 넣고 저장버튼을 누르는 거야??) 
                if (grbEnroll_DataCheck() == false) { return; }

                // 2. 저장.
                SaveData(ButtonTag);

                //공통 버튼이벤트
                PublicEnableTrue();
                //그룹박스 데이터 클리어
                grbBoxDataClear();

                if (ButtonTag == "1")     //1. 추가 > 저장했다면,
                {
                    if (dgdPastDefect.Items.Count > 0)
                    {
                        re_Search(dgdPastDefect.Items.Count - 1);
                        dgdPastDefect.Focus();
                    }
                    else
                    { re_Search(0); }
                }
                else        //2. 수정 > 저장했다면,
                {
                    re_Search(Wh_Ar_SelectedLastIndex);
                    dgdPastDefect.Focus();
                }

                ButtonTag = string.Empty;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        #region 저장 전, 그룹박스 데이터 기입체크
        private bool grbEnroll_DataCheck()
        {
            if (lib.IsNullOrWhiteSpace(txtPastDefectSubject.Text) == true)
            {
                MessageBox.Show("불량제목은 반드시 입력하셔야 합니다.");
                return false;
            }
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
            if (txtDefectSymptom_InGroupBox.Tag == null)
            {
                MessageBox.Show("불량유형은 반드시 플러스파인더를 통해 입력하셔야 합니다.");
                return false;
            }
            if (cboCloseYN_InGroupBox.SelectedIndex < 0)
            {
                MessageBox.Show("종결여부는 반드시 선택하셔야 합니다.");
                return false;
            }

            return true;
        }

        #endregion


        // 취소 버튼 클릭 시.
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //공통 버튼이벤트
                PublicEnableTrue();

                //그룹박스 데이터 클리어
                grbBoxDataClear();

                if (ButtonTag == "1") // 1. 추가하다가 취소했다면,
                {
                    if (dgdPastDefect.Items.Count > 0)
                    {
                        re_Search(Wh_Ar_SelectedLastIndex);
                        dgdPastDefect.Focus();
                    }
                    else
                    { re_Search(0); }
                }
                else        //2. 수정하다가 취소했다면
                {
                    re_Search(Wh_Ar_SelectedLastIndex);
                    dgdPastDefect.Focus();
                }

                ButtonTag = string.Empty;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #endregion


        #region 그룹박스 데이터 클리어
        // 그룹박스 데이터 클리어 하기.
        private void grbBoxDataClear()
        {
            txtPastDefectID.Text = string.Empty;
            txtPastDefectSubject.Text = string.Empty;
            txtCustomer_InGroupBox.Text = string.Empty;
            txtCustomer_InGroupBox.Tag = null;
            txtArticle_InGroupBox.Text = string.Empty;
            txtArticle_InGroupBox.Tag = null;
            txtArticleID_InGroupBox.Text = string.Empty;
            txtBuyerArticleNo.Text = string.Empty;
            txtBuyerModel.Text = string.Empty;
            txtBuyerModel.Tag = null;
            dtpOccurDate.Text = string.Empty;
            txtDefectSymptom_InGroupBox.Text = string.Empty;
            txtDefectSymptom_InGroupBox.Tag = null;
            cboDefectReason.SelectedIndex = -1;
            cboOccurStep.SelectedIndex = -1;
            txtEffectValidation.Text = string.Empty;
            cboCloseYN_InGroupBox.SelectedIndex = -1;
            chkQPointYN.IsChecked = false;
            cboDvlYN_InGroupBox.SelectedIndex = -1;

            txtComments.Text = string.Empty;
            txtDefectImageFile.Text = string.Empty;
            txtDefectImagePath.Text = string.Empty;
            imgDefectImage.Source = null;
        }

        #endregion


        #region CRUD // 각종 프로시저 모음

        //저장.
        private void SaveData(string TagNUM)
        {
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (TagNUM == "1")      // 신규추가입니다.
                {
                    // 신규추가 저장 insert.
                    string DvlDefectYN = cboDvlYN_InGroupBox.SelectedValue.ToString();
                    string PastDefectSubject = txtPastDefectSubject.Text;
                    string CustomID = txtCustomer_InGroupBox.Tag.ToString();
                    string ArticleID = txtArticle_InGroupBox.Tag.ToString();

                    string BuyerModelID = string.Empty;
                    if (txtBuyerModel.Tag != null) { BuyerModelID = txtBuyerModel.Tag.ToString(); }

                    string OccurDate = "";

                    if (!dtpOccurDate.Text.Trim().Equals(""))
                    {
                        OccurDate = dtpOccurDate.Text.Substring(0, 10).Replace("-", "");
                    }
                    string DefectSymtomCode = txtDefectSymptom_InGroupBox.Tag.ToString();

                    string DefectReasonCode = string.Empty;
                    if (cboDefectReason.SelectedIndex >= 0) { DefectReasonCode = cboDefectReason.SelectedValue.ToString(); }

                    string OccurStep = string.Empty;
                    if (cboOccurStep.SelectedIndex >= 0) { OccurStep = cboOccurStep.SelectedValue.ToString(); }

                    string EffectValidation = txtEffectValidation.Text;
                    string CloseYN = cboCloseYN_InGroupBox.SelectedValue.ToString();
                    string Comments = txtComments.Text;

                    string QPointYN = "N";
                    if (chkQPointYN.IsChecked == true) { QPointYN = "Y"; }

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("PastDefectID", "");                           // output, 자동생성.
                    sqlParameter.Add("DvlDefectYN", DvlDefectYN);                   // 개발 / 양산 여부
                    sqlParameter.Add("PastDefectSubject", PastDefectSubject);       // 제목.(필수입력 체크완료)  
                    sqlParameter.Add("CustomID", CustomID);                         // 고객사. (필수입력 체크완료)
                    sqlParameter.Add("ArticleID", ArticleID);                       // 품명. (필수입력 체크완료)

                    sqlParameter.Add("BuyerModelID", BuyerModelID);                 // 차종.
                    sqlParameter.Add("OccurDate", OccurDate);                       // 발생일.
                    sqlParameter.Add("DefectSymtomCode", DefectSymtomCode);         // 불량유형 (필수입력 체크완료)
                    sqlParameter.Add("DefectReasonCode", DefectReasonCode);         // 불량원인
                    sqlParameter.Add("OccurStep", OccurStep);                       // 단계

                    sqlParameter.Add("EffectValidation", EffectValidation);         // 유효성
                    sqlParameter.Add("CloseYN", CloseYN);                           // 종결 (필수입력 체크완료)
                    sqlParameter.Add("Comments", Comments);                         // 발생내역
                    sqlParameter.Add("QPoint", "");
                    sqlParameter.Add("QPointpath", "");

                    sqlParameter.Add("QPointYN", QPointYN);                         // 여부.
                    sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_Qul_iPastDefect";
                    pro1.OutputUseYN = "Y";
                    pro1.OutputName = "PastDefectID";
                    pro1.OutputLength = "10";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);

                    List<KeyValue> list_Result = new List<KeyValue>();
                    list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                    string sGetPastDefectID = string.Empty;

                    if (list_Result[0].key.ToLower() == "success")
                    {
                        list_Result.RemoveAt(0);
                        for (int i = 0; i < list_Result.Count; i++)
                        {
                            KeyValue kv = list_Result[i];
                            if (kv.key == "PastDefectID")
                            {
                                sGetPastDefectID = kv.value;
                            }
                        }
                        bool AttachYesNo = false;
                        if (txtDefectImagePath.Text != string.Empty)       //첨부파일 1
                        {
                            if (FTP_Save_File(listFtpFile, sGetPastDefectID))
                            {
                                txtDefectImagePath.Text = "/ImageData/PastDefect/" + sGetPastDefectID;
                                AttachYesNo = true;
                            }
                            else
                            { MessageBox.Show("데이터 저장이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }
                        }
                        if (AttachYesNo == true) { AttachFileUpdate(sGetPastDefectID); }      //첨부문서 정보 DB 업데이트.
                    }
                    else
                    {
                        MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                        return;
                    }
                }


                else if (TagNUM == "2")         // 수정 저장입니다.
                {
                    // 수정 저장 update.
                    var ViewReceiver = dgdPastDefect.SelectedItem as Win_Qul_PastDefect_U_View;

                    string DvlDefectYN = cboDvlYN_InGroupBox.SelectedValue.ToString();
                    string PastDefectSubject = txtPastDefectSubject.Text;
                    string CustomID = txtCustomer_InGroupBox.Tag.ToString();
                    string ArticleID = txtArticle_InGroupBox.Tag.ToString();

                    string BuyerModelID = string.Empty;
                    if (txtBuyerModel.Tag != null) { BuyerModelID = txtBuyerModel.Tag.ToString(); }

                    string OccurDate = "";

                    if (!dtpOccurDate.Text.Trim().Equals(""))
                    {
                        OccurDate = dtpOccurDate.Text.Substring(0, 10).Replace("-", "");
                    }

                    string DefectSymtomCode = txtDefectSymptom_InGroupBox.Tag.ToString();

                    string DefectReasonCode = string.Empty;
                    if (cboDefectReason.SelectedIndex >= 0) { DefectReasonCode = cboDefectReason.SelectedValue.ToString(); }

                    string OccurStep = string.Empty;
                    if (cboOccurStep.SelectedIndex >= 0) { OccurStep = cboOccurStep.SelectedValue.ToString(); }

                    string EffectValidation = txtEffectValidation.Text;
                    string CloseYN = cboCloseYN_InGroupBox.SelectedValue.ToString();
                    string Comments = txtComments.Text;

                    string QPointYN = string.Empty;
                    if (chkQPointYN.IsChecked == true) { QPointYN = "Y"; }
                    else { QPointYN = "N"; }

                    string QPoint_upgrade_yn = string.Empty;
                    if (txtDefectImagePath.Text == string.Empty)
                    {
                        // 첨부파일 경로가 깨끗하다면, ""으로 업그레이드.
                        // 경로가 깨끗하지 않다면, ViewReceiver와의 체크가 이루어질때까지 대기. 즉 업그레이드 파라미터 항목에서 제외.
                        QPoint_upgrade_yn = "Y";
                    }
                    else
                    { QPoint_upgrade_yn = "N"; }

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("PastDefectID", txtPastDefectID.Text);         // p key.
                    sqlParameter.Add("DvlDefectYN", DvlDefectYN);                   // 개발 / 양산 여부
                    sqlParameter.Add("PastDefectSubject", PastDefectSubject);       // 제목.(필수입력 체크완료)  
                    sqlParameter.Add("CustomID", CustomID);                         // 고객사. (필수입력 체크완료)
                    sqlParameter.Add("ArticleID", ArticleID);                       // 품명. (필수입력 체크완료)

                    sqlParameter.Add("BuyerModelID", BuyerModelID);                 // 차종.
                    sqlParameter.Add("OccurDate", OccurDate);                       // 발생일.
                    sqlParameter.Add("DefectSymtomCode", DefectSymtomCode);         // 불량유형 (필수입력 체크완료)
                    sqlParameter.Add("DefectReasonCode", DefectReasonCode);         // 불량원인
                    sqlParameter.Add("OccurStep", OccurStep);                       // 단계

                    sqlParameter.Add("EffectValidation", EffectValidation);         // 유효성
                    sqlParameter.Add("CloseYN", CloseYN);                           // 종결 (필수입력 체크완료)
                    sqlParameter.Add("Comments", Comments);                         // 발생내역

                    //sqlParameter.Add("QPoint_upgrade_yn", QPoint_upgrade_yn);
                    sqlParameter.Add("QPoint", "");
                    sqlParameter.Add("QPointpath", "");
                    sqlParameter.Add("QPointYN", QPointYN);                         // 여부.

                    sqlParameter.Add("UserID", MainWindow.CurrentUser);

                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_Qul_uPastDefect";
                    pro1.OutputUseYN = "N";
                    pro1.OutputName = "PastDefectID";
                    pro1.OutputLength = "10";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                    if (Confirm[0] == "success")
                    {
                        bool AttachYesNo = false;
                        if (txtDefectImagePath.Text != string.Empty)       //첨부파일1 > DB 업로드 조건은 통과
                        {
                            if (txtDefectImagePath.Text != ViewReceiver.QPointPath)   // 기존 저장된 경로랑 새로 들어온 경로랑 같지 않을때만,
                            {
                                if (FTP_Save_File(listFtpFile, txtPastDefectID.Text))
                                {
                                    txtDefectImagePath.Text = "/ImageData/PastDefect/" + txtPastDefectID.Text;
                                    AttachYesNo = true;
                                }
                                else
                                { MessageBox.Show("데이터 수정이 완료되었지만, 첨부문서 등록에 실패하였습니다."); }
                            }
                        }

                        if (AttachYesNo == true) { AttachFileUpdate(txtPastDefectID.Text); }      //첨부문서 정보 DB 업데이트.
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
                MessageBox.Show(ex.ToString());
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
                string PastDefectID = txtPastDefectID.Text;
                if (PastDefectID == "")
                {
                    MessageBox.Show("삭제대상이 정확하지 않습니다. 불량ID를 확인해 주세요.");
                    return;
                }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("PastDefectID", PastDefectID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Qul_dPastDefect", sqlParameter, false);
                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("이상발생, 관리자에게 문의하세요.");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }


        // 품명정보 바탕으로 품번 자동으로 찾아 뿌리기.
        private void Article_InGroupBox_OtherSearch(string ArticleID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ArticleID", ArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleData", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        txtBuyerArticleNo.Text = string.Empty;
                        return;
                    }
                    else
                    {
                        txtBuyerArticleNo.Text = dt.Rows[0]["BuyerArticleNo"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }


        // 1) 첨부문서가 있을경우, 2) FTP에 정상적으로 업로드가 완료된 경우.  >> DB에 정보 업데이트 
        private void AttachFileUpdate(string ID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                sqlParameter.Clear();
                sqlParameter.Add("PastDefectID", ID);

                sqlParameter.Add("AttPath1", txtDefectImagePath.Text);
                sqlParameter.Add("AttFile1", txtDefectImageFile.Text);

                sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Qul_uPastDefect_Ftp", sqlParameter, false);
                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("이상발생, 관리자에게 문의하세요");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        #endregion


        // 닫기.
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
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

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdPastDefect.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            Lib lib2 = new Lib();
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "메인그리드";
            lst[1] = dgdPastDefect.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdPastDefect.Name))
                {
                    //MessageBox.Show("대분류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib2.DataGridToDTinHidden(dgdPastDefect);
                    else
                        dt = lib2.DataGirdToDataTable(dgdPastDefect);

                    Name = dgdPastDefect.Name;

                    if (lib2.GenerateExcel(dt, Name))
                    {
                        lib2.excel.Visible = true;
                        lib2.ReleaseExcelObject(lib2.excel);
                    }
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


        #region FTP

        // FTP. 파일 첨부, 등록하기.
        private void btnFileEnroll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();

                OFdlg.DefaultExt = ".jpg";
                OFdlg.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png | All Files|*.*";

                Nullable<bool> result = OFdlg.ShowDialog();
                if (result == true)
                {
                    FullPath1 = OFdlg.FileName;  //긴 경로(FULL 사이즈)

                    string AttachFileName = OFdlg.SafeFileName;  //명.
                    string AttachFilePath = string.Empty;       // 경로

                    AttachFilePath = FullPath1.Replace(AttachFileName, "");

                    StreamReader sr = new StreamReader(OFdlg.FileName);
                    long File_size = sr.BaseStream.Length;
                    if (sr.BaseStream.Length > (2048 * 1000))
                    {
                        // 업로드 파일 사이즈범위 초과
                        MessageBox.Show("밀시트의 파일사이즈가 2M byte를 초과하였습니다.");
                        sr.Close();
                        return;
                    }

                    txtDefectImageFile.Text = AttachFileName;
                    txtDefectImagePath.Text = AttachFilePath.ToString();

                    listFtpFile.Add(new string[] { AttachFileName, AttachFilePath.ToString() });

                    //첨부 이미지 보기.
                    ShowingImage(FullPath1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        // 첨부등록 이미지 바로보기.
        private void ShowingImage(string source)
        {
            try
            {
                if (txtDefectImagePath.Text != string.Empty) // <내 컴퓨터> 첨부파일 경로가 있으니,
                {
                    var ImageSource = new Uri(source);
                    imgDefectImage.Source = new BitmapImage(ImageSource);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private bool SetImage(string ImageName, string FolderName)
        {
            bool ExistFile = false;
            BitmapImage bit = null;
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            if (_ftp == null) { return false; }

            string[] fileListDetail;
            //fileListDetail = _ftp.directoryListSimple(FolderName, Encoding.Default);
            fileListDetail = _ftp.directoryListSimple(FolderName, Encoding.UTF8);
            ExistFile = FileInfoAndFlag(fileListDetail, ImageName);
            if (ExistFile)
            {
                bit = _ftp.DrawingImageByByte(FTP_ADDRESS + '/' + FolderName + '/' + ImageName + "");
                imgDefectImage.Source = bit;
                return true;
            }

            return false;
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
            MakeFolder = FolderInfoAndFlag(fileListSimple, MakeFolderName);

            if (MakeFolder == false)        // 같은 아이를 찾지 못한경우,
            {
                //MIL 폴더에 InspectionID로 저장
                if (_ftp.createDirectory(MakeFolderName) == false)
                {
                    MessageBox.Show("업로드를 위한 폴더를 생성할 수 없습니다.");
                    return false;
                }
            }
            else
            {
                fileListDetail = _ftp.directoryListSimple(MakeFolderName, Encoding.Default);
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
                    listStrArrayFileInfo[i][0] = MakeFolderName + "/" + listStrArrayFileInfo[i][0];
                    UpdateFilesInfo.Add(listStrArrayFileInfo[i]);
                }
            }
            if (!_ftp.UploadTempFilesToFTP(UpdateFilesInfo))
            {
                MessageBox.Show("파일업로드에 실패하였습니다.");
                return false;
            }
            return true;
        }



        // ftp 파일 삭제하기.
        private void btnFileDel_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 삭제 하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                if (txtDefectImagePath.Text != string.Empty)
                {
                    txtDefectImageFile.Text = string.Empty;
                    txtDefectImagePath.Text = string.Empty;
                    imgDefectImage.Source = null;
                }
            }
        }


        // 파일 내려받기.
        private void btnFileSee_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 다운로드 하시겠습니까?", "다운로드 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                var ViewReceiver = dgdPastDefect.SelectedItem as Win_Qul_PastDefect_U_View;

                if (ViewReceiver != null && !ViewReceiver.QPointPath.Equals(""))
                {
                    FTP_DownLoadFile(ViewReceiver.QPointPath, ViewReceiver.PastDefectID, ViewReceiver.QPoint);
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


        private void txtPastDefectSubject_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                lib.SendK(Key.Tab, this);
                dtpOccurDate.IsDropDownOpen = true;
            }
        }
        private void dtpOccurDate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            lib.SendK(Key.Tab, this);
            cboDvlYN_InGroupBox.IsDropDownOpen = true;
        }
        private void cboDvlYN_InGroupBox_DropDownClosed(object sender, EventArgs e)
        {
            lib.SendK(Key.Tab, this);
        }
        private void txtCustomer_InGroupBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnCustomer_InGroupBox_Click(null, null);
            }
        }
        private void txtArticle_InGroupBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnArticle_InGroupBox_Click(null, null);
            }
        }
        private void txtBuyerModel_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnBuyerModel_Click(null, null);
            }
        }
        private void txtDefectSymptom_InGroupBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnDefectSymptom_InGroupBox_Click(null, null);
                cboDefectReason.IsDropDownOpen = true;
            }
        }
        private void cboDefectReason_DropDownClosed(object sender, EventArgs e)
        {
            lib.SendK(Key.Tab, this);
            cboOccurStep.IsDropDownOpen = true;
        }
        private void cboOccurStep_DropDownClosed(object sender, EventArgs e)
        {
            lib.SendK(Key.Tab, this);
        }

        // 엔터 키를 통한 탭 인덱스 키 이동.
        private void EnterMove_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                lib.SendK(Key.Tab, this);
            }
        }

        private void txtEffectValidation_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                lib.SendK(Key.Tab, this);
                cboCloseYN_InGroupBox.IsDropDownOpen = true;
            }
        }

        private void cboCloseYN_InGroupBox_DropDownClosed(object sender, EventArgs e)
        {
            lib.SendK(Key.Tab, this);
        }

        private void txtDefect_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btnDefectSymptom_Click(null, null);
            }
        }

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


    class Win_Qul_PastDefect_U_View : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }


        // 순번
        public string NUM { get; set; }

        // PastDefect 조회 값.    
        public string PastDefectID { get; set; }
        public string PastDefectSubject { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string ArticleID { get; set; }

        public string Sabun { get; set; }

        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string BuyerModelID { get; set; }
        public string BuyerModel { get; set; }
        public string OccurDate { get; set; }

        public string DefectSymtomCode { get; set; }
        public string DefectSymtom { get; set; }
        public string DefectReasonCode { get; set; }
        public string DefectReason { get; set; }
        public string OccurStep { get; set; }

        public string OccurStepName { get; set; }
        public string EffectValidation { get; set; }
        public string CloseYN { get; set; }
        public string Comments { get; set; }
        public string QPoint { get; set; }

        public string QPointPath { get; set; }
        public string QPointYN { get; set; }
        public string DvlDefectYN { get; set; }



        public Win_Qul_PastDefect_U_View Copy()
        {
            return (Win_Qul_PastDefect_U_View)this.MemberwiseClone();
        }
    }


}
