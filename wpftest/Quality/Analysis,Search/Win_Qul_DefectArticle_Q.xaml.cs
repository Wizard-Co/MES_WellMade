using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_DefectRepair_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_DefectArticle_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        //제품별 점유율에 사용
        Win_Qul_DefectArticle_Q_ModelOccupy_CodeView WinModelName = new Win_Qul_DefectArticle_Q_ModelOccupy_CodeView();
        //유형별 점유율에 사용
        Win_Qul_DefectArticle_Q_DefectType_CodeView WinTypeName = new Win_Qul_DefectArticle_Q_DefectType_CodeView();
        //작업자별
        Win_Qul_DefectArticle_Q_Worker_CodeView WinWorkerName = new Win_Qul_DefectArticle_Q_Worker_CodeView();

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        //이건 뭐...
        string M1 = "       -       ";
        string M2 = "       -       ";
        string M3 = "       -       ";
        string M4 = "       -       ";
        string M5 = "       -       ";
        string M6 = "       -       ";
        string M7 = "       -       ";
        string M8 = "       -       ";
        string M9 = "       -       ";
        string M10 = "       -       ";
        string M11 = "       -       ";
        string M12 = "       -       ";
        string M13 = "       -       ";


        string strLastDay = string.Empty;

        //Image 변수 선언
        System.Windows.Controls.Image ImageData = new System.Windows.Controls.Image();


        public Win_Qul_DefectArticle_Q()
        {
            InitializeComponent();
        }

        SeriesCollection SeriesCollection;
        public string[] X_Linelbl { get; set; }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            lib.UiLoading(sender);
            btnThisMonth_Click(null, null);
            CreateDataGridRowsColumns();
            SetComboBox();
            lblchartMonth.Content = "2. " + dtpSDate.SelectedDate.Value.ToString().Replace("-", "").Substring(4, 2) + "월 불량유형";
            cboProductGrpID.IsEnabled = false;
        }


        private void SetComboBox()
        {
            // 제품군
            ObservableCollection<CodeView> ovcProductGrp = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CMPRDGRPID", "Y", "");
            this.cboProductGrpID.ItemsSource = ovcProductGrp;
            this.cboProductGrpID.DisplayMemberPath = "code_name";
            this.cboProductGrpID.SelectedValuePath = "code_id";

            ObservableCollection<CodeView> oveOrderForm = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "QULSTEP", "Y", "", "");
            cboOccurStepSrh.ItemsSource = oveOrderForm;
            cboOccurStepSrh.DisplayMemberPath = "code_name";
            cboOccurStepSrh.SelectedValuePath = "code_id";

            if (cboOccurStepSrh.Items.Count > 0)
                cboOccurStepSrh.SelectedIndex = 0;
        }


        #region 날짜 관련

        // 이전년도
        private void btnLastYear_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = lib.BringLastYearDatetimeContinue(dtpSDate.SelectedDate.Value)[0];
        }

        //금년
        private void btnThisYear_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = lib.BringThisYearDatetimeFormat()[0];
        }

        //전월
        private void BtnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = lib.BringLastMonthDatetimeList()[0];
        }

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
        }

        #endregion

        #region 체크박스 action

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

        //품명 체크 이벤트
        private void ChkArticle_Checked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = true;
            btnArticle.IsEnabled = true;
        }

        //품명 체크 해제 이벤트
        private void ChkArticle_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
        }


        #region 플러스파인더

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
        #endregion

        #endregion


        // row header column 자동생성.
        private void CreateDataGridRowsColumns()
        {
            string[] RowHeaderName = new string[4];
            RowHeaderName[0] = "목표PPM";

            if (rbnInspoint0.IsChecked == true)
            {
                RowHeaderName[1] = "전체수량";
            }
            else if (rbnInspoint1.IsChecked == true)
            {
                RowHeaderName[1] = "입고수량";
            }
            else if (rbnInspoint3.IsChecked == true)
            {
                RowHeaderName[1] = "생산수량";
            }
            else if (rbnInspoint4.IsChecked == true)
            {
                RowHeaderName[1] = "검사수량";
            }
            else if (rbnInspoint5.IsChecked == true)
            {
                RowHeaderName[1] = "출하수량";
            }
            RowHeaderName[2] = "불량수량";
            RowHeaderName[3] = "불량PPM";

            for (int i = 0; i < 4; i++)
            {
                var Win_Qul_DefectArticle_Q_DefectCount_Insert = new Win_Qul_DefectArticle_Q_DefectCount_CodeView()
                {
                    DefectCount_RowHeaderColumns = RowHeaderName[i]
                };
                dgdDefectArticle_DefectCount.Items.Add(Win_Qul_DefectArticle_Q_DefectCount_Insert);
            }

            RowHeaderName = new string[2];
            RowHeaderName[0] = "품번";  //BuyerArticleNo
            RowHeaderName[1] = "불량수량";

            for (int i = 0; i < 2; i++)
            {
                var Win_Qul_DefectArticle_Q_ModelOccupy_Insert = new Win_Qul_DefectArticle_Q_ModelOccupy_CodeView()
                {
                    ModelOccupy_RowHeaderColumns = RowHeaderName[i]
                };
                dgdDefectArticle_ModelOccupy.Items.Add(Win_Qul_DefectArticle_Q_ModelOccupy_Insert);
            }

            RowHeaderName = new string[2];
            RowHeaderName[0] = "유형";
            RowHeaderName[1] = "불량수량";

            for (int i = 0; i < 2; i++)
            {
                var Win_Qul_DefectArticle_Q_DefectType_Insert = new Win_Qul_DefectArticle_Q_DefectType_CodeView()
                {
                    DefectType_RowHeaderColumns = RowHeaderName[i]
                };
                dgdDefectArticle_DefectType.Items.Add(Win_Qul_DefectArticle_Q_DefectType_Insert);
            }

            RowHeaderName = new string[2];
            RowHeaderName[0] = "작업자";
            RowHeaderName[1] = "불량수량";

            for (int i = 0; i < 2; i++)
            {
                var Win_Qul_DefectArticle_Q_Worker_Insert = new Win_Qul_DefectArticle_Q_Worker_CodeView()
                {
                    Worker_RowHeaderColumns = RowHeaderName[i]
                };
                dgdDefectArticle_Worker.Items.Add(Win_Qul_DefectArticle_Q_Worker_Insert);
            }
        }



        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByForm(this.GetType().Name, "R");

            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                //로직
                re_Search(0);

            }), System.Windows.Threading.DispatcherPriority.Background);


            Dispatcher.BeginInvoke(new Action(() =>
            {
                btnSearch.IsEnabled = true;
            }), System.Windows.Threading.DispatcherPriority.Background);
        }


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
                        break;
                    }
                }
                i++;
            }
        }

        private void chkDefectOccurStep_Checked(object sender, RoutedEventArgs e)
        {
            //cboOccurStepSrh.IsEnabled = true;
        }

        private void chkDefectOccurStep_Unchecked(object sender, RoutedEventArgs e)
        {
            //cboOccurStepSrh.IsEnabled = false;
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            string[] lst = new string[8];

            lst[0] = "불량건수";
            lst[1] = "점유모델";
            lst[2] = "불량유형";
            lst[3] = "작업자";
            lst[4] = dgdDefectArticle_DefectCount.Name;
            lst[5] = dgdDefectArticle_ModelOccupy.Name;
            lst[6] = dgdDefectArticle_DefectType.Name;
            lst[7] = dgdDefectArticle_Worker.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
            ExpExc.ShowDialog();

            DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");

            if (ExpExc.DialogResult.HasValue)
            {
                string[] ExcelRowHeaderName = new string[4];
                if (ExpExc.choice.Equals(dgdDefectArticle_DefectCount.Name))
                {
                    ExcelRowHeaderName[0] = "목표PPM";

                    if (cboOccurStepSrh.SelectedValue.ToString().Equals("0"))
                    {
                        ExcelRowHeaderName[1] = "전체수량";
                    }
                    else if (cboOccurStepSrh.SelectedValue.ToString().Equals("1"))
                    {
                        ExcelRowHeaderName[1] = "입고수량";
                    }
                    else if (cboOccurStepSrh.SelectedValue.ToString().Equals("2"))
                    {
                        ExcelRowHeaderName[1] = "생산수량";
                    }
                    else if (cboOccurStepSrh.SelectedValue.ToString().Equals("3"))
                    {
                        ExcelRowHeaderName[1] = "검사수량";
                    }
                    else if (cboOccurStepSrh.SelectedValue.ToString().Equals("6"))
                    {
                        ExcelRowHeaderName[1] = "출하수량";
                    }
                    ExcelRowHeaderName[2] = "불량수량";
                    ExcelRowHeaderName[3] = "불량PPM";

                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdDefectArticle_DefectCount);
                    else
                        dt = lib.DataGirdToDataTable(dgdDefectArticle_DefectCount);

                    Name = dgdDefectArticle_DefectCount.Name;
                    if (lib.HeaderAddGenerateExcel(dt, Name, ExcelRowHeaderName))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
                }

                else if (ExpExc.choice.Equals(dgdDefectArticle_ModelOccupy.Name))
                {
                    string[] ExcelRowHeaderName2 = new string[2];
                    ExcelRowHeaderName2[0] = "품번";  //BuyerArticleNo
                    ExcelRowHeaderName2[1] = "불량수량";

                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdDefectArticle_ModelOccupy);
                    else
                        dt = lib.DataGirdToDataTable(dgdDefectArticle_ModelOccupy);

                    Name = dgdDefectArticle_ModelOccupy.Name;

                    if (lib.HeaderAddGenerateExcel(dt, Name, ExcelRowHeaderName2))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
                }

                else if (ExpExc.choice.Equals(dgdDefectArticle_DefectType.Name))
                {
                    string[] ExcelRowHeaderName3 = new string[2];
                    ExcelRowHeaderName3[0] = "유형";
                    ExcelRowHeaderName3[1] = "불량수량";

                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdDefectArticle_DefectType);
                    else
                        dt = lib.DataGirdToDataTable(dgdDefectArticle_DefectType);

                    Name = dgdDefectArticle_DefectType.Name;

                    if (lib.HeaderAddGenerateExcel(dt, Name, ExcelRowHeaderName3))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
                }

                else if (ExpExc.choice.Equals(dgdDefectArticle_Worker.Name))
                {
                    string[] ExcelRowHeaderName4 = new string[2];
                    ExcelRowHeaderName4[0] = "작업자";
                    ExcelRowHeaderName4[1] = "불량수량";

                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdDefectArticle_Worker);
                    else
                        dt = lib.DataGirdToDataTable(dgdDefectArticle_Worker);

                    Name = dgdDefectArticle_Worker.Name;

                    if (lib.HeaderAddGenerateExcel(dt, Name, ExcelRowHeaderName4))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
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
        }

        //월별불량지수 데이터 그리드와 그래프 값
        private void re_Search(int selectIndex)
        {
            //검색 누를 때 re_Search 타니까 이 때 아이템들 비워주기

            if (dtpSDate.SelectedDate == null)
            {
                MessageBox.Show("년도를 정확히 선택해 주세요. 필수선택입니다.");
                //검색 다 되면 활성화
                btnSearch.IsEnabled = true;
                return;
            }

            try
            {

                //조회할 때 선택된 조회 날짜를 반영하여 "00월 불량유형"이 들어가도록 
                lblchartMonth.Content = "2. " + dtpSDate.SelectedDate.Value.ToString().Replace("-", "").Substring(4, 2) + "월 불량유형";

                //불량지수 그래프 비워주기
                if (lvcDayChart.Series != null && lvcDayChart.Series.Count > 0)
                {
                    lvcDayChart.Series.Clear();
                }

                //불량건수 dgd 비워주기
                if (dgdDefectArticle_DefectCount.Items.Count > 0)
                {
                    dgdDefectArticle_DefectCount.Items.Clear();
                }

                //제품별 dgd 비워주기
                if (dgdDefectArticle_ModelOccupy.Items.Count > 0)
                {
                    dgdDefectArticle_ModelOccupy.Items.Clear();
                }

                //제품별 원그래프 비워주기
                if (lvcProductPieChart.Series != null && lvcProductPieChart.Series.Count > 0)
                {
                    lvcProductPieChart.Series.Clear();
                }

                //유형별 dgd 비워주기
                if (dgdDefectArticle_DefectType.Items.Count > 0)
                {
                    dgdDefectArticle_DefectType.Items.Clear();
                }

                //유형별 원그래프 비워주기
                if (lvcTypePieChart.Series != null && lvcTypePieChart.Series.Count > 0)
                {
                    lvcTypePieChart.Series.Clear();
                }

                //hidden dgd 비워주기
                if (PieChartProductValue.Items.Count > 0)
                {
                    PieChartProductValue.Items.Clear();
                }

                //hidden dgd 비워주기
                if (PieChartTypeValue.Items.Count > 0)
                {
                    PieChartTypeValue.Items.Clear();
                }

                //유형별 dgd 비워주기
                if (dgdDefectArticle_Worker.Items.Count > 0)
                {
                    dgdDefectArticle_Worker.Items.Clear();
                }

                //RowHeader값 다시 읽어주기
                CreateDataGridRowsColumns();

                string InsPoint = string.Empty;  //불량 발생 시점

                if (cboOccurStepSrh.SelectedValue.ToString().Equals("0")) { InsPoint = ""; } //전체
                else if (cboOccurStepSrh.SelectedValue.ToString().Equals("2")) { InsPoint = "3"; } //생산
                else if (cboOccurStepSrh.SelectedValue.ToString().Equals("3")) { InsPoint = "2"; } //자주검사
                else if (cboOccurStepSrh.SelectedValue.ToString().Equals("4")) { InsPoint = "0"; } //
                else if (cboOccurStepSrh.SelectedValue.ToString().Equals("6")) { InsPoint = "6"; } //
                else { InsPoint = cboOccurStepSrh.SelectedValue.ToString(); }

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

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("YYYY", chkDate.IsChecked == true ? dtpSDate.Text.Substring(0, 4) : "");
                sqlParameter.Add("MM", chkDate.IsChecked == true ? dtpSDate.Text.Substring(5, 2) : "");
                sqlParameter.Add("chkProdGroupID", chkProductGroup.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProdGroupID", chkProductGroup.IsChecked == true ? cboProductGrpID.SelectedValue.ToString() : "");
                sqlParameter.Add("chkInspectPointID", !InsPoint.Equals("") ? 1 : 0); // 전체가 아니면 1로 
                sqlParameter.Add("InspectPointID", InsPoint); //빈값 : 전체, 1: 입고, 2: 자주검사, 3: 생산, 4: 최종검사, 5:출하

                sqlParameter.Add("chkArticleID", chkArticleID);
                sqlParameter.Add("ArticleID", ArticleID);
                sqlParameter.Add("BuyerArticleNo", chkArticleNo.IsChecked == true ? txtArticleNo.Text : "");
                sqlParameter.Add("BuyerArticleNme", chkArticle.IsChecked == true ? txtArticle.Text : "");


                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectArticle_DefectCount", sqlParameter, false);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        return;
                    }
                    else
                    {
                        DataRow drHeader = dt.Rows[4];
                        X_Linelbl = new[] {Convert.ToString(Convert.ToInt32(drHeader["M1"]))+"월", Convert.ToString(Convert.ToInt32(drHeader["M2"]))+"월",
                            Convert.ToString(Convert.ToInt32(drHeader["M3"]))+"월", Convert.ToString(Convert.ToInt32(drHeader["M4"]))+"월",
                            Convert.ToString(Convert.ToInt32(drHeader["M5"]))+"월", Convert.ToString(Convert.ToInt32(drHeader["M6"]))+"월",
                            Convert.ToString(Convert.ToInt32(drHeader["M7"]))+"월", Convert.ToString(Convert.ToInt32(drHeader["M8"]))+"월",
                            Convert.ToString(Convert.ToInt32(drHeader["M9"]))+"월", Convert.ToString(Convert.ToInt32(drHeader["M10"]))+"월",
                            Convert.ToString(Convert.ToInt32(drHeader["M11"]))+"월", Convert.ToString(Convert.ToInt32(drHeader["M12"]))+"월"};

                        //헤더변경
                        for (int i = 0; i < 12; i++)
                        {
                            dgdDefectArticle_DefectCount.Columns[i].Header = Convert.ToString(Convert.ToInt32(drHeader["M" + (i + 1) + ""])) + "월";
                        }

                        lbl_1Month.Text = Convert.ToString(Convert.ToInt32(drHeader["M1"])) + "월";
                        lbl_2Month.Text = Convert.ToString(Convert.ToInt32(drHeader["M2"])) + "월";
                        lbl_3Month.Text = Convert.ToString(Convert.ToInt32(drHeader["M3"])) + "월";
                        lbl_4Month.Text = Convert.ToString(Convert.ToInt32(drHeader["M4"])) + "월";
                        lbl_5Month.Text = Convert.ToString(Convert.ToInt32(drHeader["M5"])) + "월";
                        lbl_6Month.Text = Convert.ToString(Convert.ToInt32(drHeader["M6"])) + "월";
                        lbl_7Month.Text = Convert.ToString(Convert.ToInt32(drHeader["M7"])) + "월";
                        lbl_8Month.Text = Convert.ToString(Convert.ToInt32(drHeader["M8"])) + "월";
                        lbl_9Month.Text = Convert.ToString(Convert.ToInt32(drHeader["M9"])) + "월";
                        lbl_10Month.Text = Convert.ToString(Convert.ToInt32(drHeader["M10"])) + "월";
                        lbl_11Month.Text = Convert.ToString(Convert.ToInt32(drHeader["M11"])) + "월";
                        lbl_12Month.Text = Convert.ToString(Convert.ToInt32(drHeader["M12"])) + "월";

                        for (int j = 0; j < 4; j++)
                        {
                            DataRow dr = dt.Rows[j];
                            DataGridRow dgr = lib.GetRow(j, dgdDefectArticle_DefectCount);
                            var ViewReceiver = dgr.Item as Win_Qul_DefectArticle_Q_DefectCount_CodeView;

                            double M1 = 0;
                            double M2 = 0;
                            double M3 = 0;
                            double M4 = 0;
                            double M5 = 0;
                            double M6 = 0;
                            double M7 = 0;
                            double M8 = 0;
                            double M9 = 0;
                            double M10 = 0;
                            double M11 = 0;
                            double M12 = 0;
                            double M13 = 0;

                            double.TryParse(dr["M1"].ToString(), out M1);
                            double.TryParse(dr["M2"].ToString(), out M2);
                            double.TryParse(dr["M3"].ToString(), out M3);
                            double.TryParse(dr["M4"].ToString(), out M4);
                            double.TryParse(dr["M5"].ToString(), out M5);
                            double.TryParse(dr["M6"].ToString(), out M6);
                            double.TryParse(dr["M7"].ToString(), out M7);
                            double.TryParse(dr["M8"].ToString(), out M8);
                            double.TryParse(dr["M9"].ToString(), out M9);
                            double.TryParse(dr["M10"].ToString(), out M10);
                            double.TryParse(dr["M11"].ToString(), out M11);
                            double.TryParse(dr["M12"].ToString(), out M12);
                            double.TryParse(dr["M13"].ToString(), out M13);


                            if (j == 3)
                            {
                                ViewReceiver.M1 = stringFormatN2(M1);
                                ViewReceiver.M2 = stringFormatN2(M2);
                                ViewReceiver.M3 = stringFormatN2(M3);
                                ViewReceiver.M4 = stringFormatN2(M4);
                                ViewReceiver.M5 = stringFormatN2(M5);
                                ViewReceiver.M6 = stringFormatN2(M6);
                                ViewReceiver.M7 = stringFormatN2(M7);
                                ViewReceiver.M8 = stringFormatN2(M8);
                                ViewReceiver.M9 = stringFormatN2(M9);
                                ViewReceiver.M10 = stringFormatN2(M10);
                                ViewReceiver.M11 = stringFormatN2(M11);
                                ViewReceiver.M12 = stringFormatN2(M12);
                                ViewReceiver.M13 = stringFormatN2(M13);
                            }
                            else
                            {
                                ViewReceiver.M1 = stringFormatN0(M1);
                                ViewReceiver.M2 = stringFormatN0(M2);
                                ViewReceiver.M3 = stringFormatN0(M3);
                                ViewReceiver.M4 = stringFormatN0(M4);
                                ViewReceiver.M5 = stringFormatN0(M5);
                                ViewReceiver.M6 = stringFormatN0(M6);
                                ViewReceiver.M7 = stringFormatN0(M7);
                                ViewReceiver.M8 = stringFormatN0(M8);
                                ViewReceiver.M9 = stringFormatN0(M9);
                                ViewReceiver.M10 = stringFormatN0(M10);
                                ViewReceiver.M11 = stringFormatN0(M11);
                                ViewReceiver.M12 = stringFormatN0(M12);
                                ViewReceiver.M13 = stringFormatN0(M13);
                            }



                            if (j == 0)     // 목표, 그래프 그려야 한다.
                            {
                                SeriesCollection = new SeriesCollection
                            {
                                new LineSeries
                                {
                                    Title = "목표",
                                    Values = new ChartValues<double>
                                    {
                                        M1, M2, M3, M4, M5, M6,
                                        M7, M8, M9, M10, M11, M12
                                    }
                                }
                            };
                            }
                            else if (j == 3)        // 불량율,ㅡ 그래프 그려야 한다.
                            {
                                SeriesCollection.Add(new LineSeries
                                {
                                    Title = "불량율",
                                    Values = new ChartValues<double>
                                {
                                    M1, M2, M3, M4, M5, M6,
                                    M7, M8, M9, M10, M11, M12
                                }
                                });
                            }
                        }



                        lvcDayChart.Series = SeriesCollection;
                        DataContext = this;

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            FillGridModel();
            FillGridType();
            FillGridWorker();

        }


        private void testModel()
        {
            WinModelName = new Win_Qul_DefectArticle_Q_ModelOccupy_CodeView
            {
                DefectQty = "96",
                DefectCount = "3",
                GroupingName = "32076 7270",
                SumDefectQty = "137",
                SumDefectCount = "11",

                DefectQtyRate = "70.00",
                DefectCountRate = "27.00",
            };
            PieChartProductValue.Items.Add(WinModelName);

            WinModelName = new Win_Qul_DefectArticle_Q_ModelOccupy_CodeView
            {
                DefectQty = "34",
                DefectCount = "3",
                GroupingName = "32076 7311",
                SumDefectQty = "137",
                SumDefectCount = "11",

                DefectQtyRate = "25.00",
                DefectCountRate = "27.00",
            };
            PieChartProductValue.Items.Add(WinModelName);

            WinModelName = new Win_Qul_DefectArticle_Q_ModelOccupy_CodeView
            {
                DefectQty = "5",
                DefectCount = "4",
                GroupingName = "32060-7262",
                SumDefectQty = "137",
                SumDefectCount = "11",

                DefectQtyRate = "4.00",
                DefectCountRate = "36.00",
            };
            PieChartProductValue.Items.Add(WinModelName);

            WinModelName = new Win_Qul_DefectArticle_Q_ModelOccupy_CodeView
            {
                DefectQty = "2",
                DefectCount = "1",
                GroupingName = "25471-2T000",
                SumDefectQty = "137",
                SumDefectCount = "11",

                DefectQtyRate = "1.00",
                DefectCountRate = "9.00",
            };
            PieChartProductValue.Items.Add(WinModelName);



            WinTypeName = new Win_Qul_DefectArticle_Q_DefectType_CodeView
            {
                DefectQty = "136",
                DefectCount = "10",
                GroupingName = "찍힘",
                SumDefectQty = "137",
                SumDefectCount = "11",

                DefectQtyRate = "99.00",
                DefectCountRate = "91.00",
            };
            PieChartTypeValue.Items.Add(WinTypeName);

            WinTypeName = new Win_Qul_DefectArticle_Q_DefectType_CodeView
            {
                DefectQty = "1",
                DefectCount = "1",
                GroupingName = "면취량",
                SumDefectQty = "137",
                SumDefectCount = "11",

                DefectQtyRate = "1.00",
                DefectCountRate = "9.00",
            };
            PieChartTypeValue.Items.Add(WinTypeName);
        }

        //제품별 불량 dgdDefectArticle_ModelOccupy 데이터그리드 조회
        private void FillGridModel()
        {
            string InsPoint = string.Empty;  //불량 발생 시점

            if (cboOccurStepSrh.SelectedValue.ToString().Equals("0")) { InsPoint = ""; } //전체
            else if (cboOccurStepSrh.SelectedValue.ToString().Equals("2")) { InsPoint = "3"; }
            else if (cboOccurStepSrh.SelectedValue.ToString().Equals("3")) { InsPoint = "2"; }
            //else if (cboOccurStepSrh.SelectedValue.ToString().Equals("4")) { InsPoint = "0"; } //
            else if (cboOccurStepSrh.SelectedValue.ToString().Equals("4")) { InsPoint = "3"; } //
            else if (cboOccurStepSrh.SelectedValue.ToString().Equals("6")) { InsPoint = "5"; } //
            else { InsPoint = cboOccurStepSrh.SelectedValue.ToString(); }

            //헤더변경
            for (int i = 0; i < 12; i++)
            {
                dgdDefectArticle_ModelOccupy.Columns[i].Header = (i + 1) + "위";
            }

            //마지막행은 13위부터의 합계
            dgdDefectArticle_ModelOccupy.Columns[12].Header = "기타";

            try
            {
                DataSet ds = null;

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

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("YYYYMM", chkDate.IsChecked == true ? dtpSDate.SelectedDate.ToString().Replace("-", "").Substring(0, 6) : "");
                sqlParameter.Add("chkProdGroupID", chkProductGroup.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProdGroupID", chkProductGroup.IsChecked == true ? cboProductGrpID.SelectedValue.ToString() : "");
                sqlParameter.Add("chkInspectPointID", 1);   // 뭐든 체크는 되어있을 거니까
                sqlParameter.Add("InspectPointID", InsPoint);
                sqlParameter.Add("chkArticleID", chkArticleID);
                sqlParameter.Add("ArticleID", ArticleID);
                sqlParameter.Add("sGrouping", "3"); //varchar(1) 이라서 "3" ,  1 : 유형 , 3 : 제품
                sqlParameter.Add("BuyerArticleNme", chkArticle.IsChecked == true ? txtArticle.Text : "");
                sqlParameter.Add("BuyerArticleNo", chkArticleNo.IsChecked == true ? txtArticleNo.Text : "");

                ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectArticle_DefectCountsub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다. (점유제품)");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        //원그래프를 시작해보자
                        PieData pd = new PieData();

                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            DataRow dr = dt.Rows[j];
                            WinModelName = new Win_Qul_DefectArticle_Q_ModelOccupy_CodeView
                            {
                                DefectQty = dr["DefectQty"].ToString(),
                                DefectCount = dr["DefectCount"].ToString(),
                                GroupingName = dr["GroupingName"].ToString(),
                                SumDefectQty = dr["SumDefectQty"].ToString(),
                                SumDefectCount = dr["SumDefectCount"].ToString(),

                                DefectQtyRate = dr["DefectQtyRate"].ToString(),
                                DefectCountRate = dr["DefectCountRate"].ToString(),
                            };

                            PieChartProductValue.Items.Add(WinModelName);

                            DataGridRow dgr = lib.GetRow(0, dgdDefectArticle_ModelOccupy);
                            var ModelName = dgr.Item as Win_Qul_DefectArticle_Q_ModelOccupy_CodeView;

                            // ------- 원그래프 값 설정 -------
                            double value = 0.0;
                            if (WinModelName.DefectQtyRate == string.Empty)
                                WinModelName.DefectQtyRate = "0";

                            value = Convert.ToDouble(WinModelName.DefectQtyRate);
                            pd.AddSlice(WinModelName.GroupingName, value);
                            // --------------------------------

                            if (j < dt.Rows.Count)
                            {
                                if (j == 0)
                                {
                                    ModelName.M1 = M1;
                                    ModelName.M2 = M2;
                                    ModelName.M3 = M3;
                                    ModelName.M4 = M4;
                                    ModelName.M5 = M5;
                                    ModelName.M6 = M6;
                                    ModelName.M7 = M7;
                                    ModelName.M8 = M8;
                                    ModelName.M9 = M9;
                                    ModelName.M10 = M10;
                                    ModelName.M11 = M11;
                                    ModelName.M12 = M12;
                                    ModelName.M13 = M13;

                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M1 = WinModelName.GroupingName.ToString(); }
                                }
                                if (j == 1)
                                {
                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M2 = WinModelName.GroupingName.ToString(); }
                                }
                                if (j == 2)
                                {
                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M3 = WinModelName.GroupingName.ToString(); }
                                }
                                if (j == 3)
                                {
                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M4 = WinModelName.GroupingName.ToString(); }
                                }
                                if (j == 4)
                                {
                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M5 = WinModelName.GroupingName.ToString(); }
                                }
                                if (j == 5)
                                {
                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M6 = WinModelName.GroupingName.ToString(); }
                                }
                                if (j == 6)
                                {
                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M7 = WinModelName.GroupingName.ToString(); }
                                }
                                if (j == 7)
                                {
                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M8 = WinModelName.GroupingName.ToString(); }
                                }
                                if (j == 8)
                                {
                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M9 = WinModelName.GroupingName.ToString(); }
                                }
                                if (j == 9)
                                {
                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M10 = WinModelName.GroupingName.ToString(); }
                                }
                                if (j == 10)
                                {
                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M11 = WinModelName.GroupingName.ToString(); }
                                }
                                if (j == 11)
                                {
                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M12 = WinModelName.GroupingName.ToString(); }
                                }
                                if (j == 12)
                                {
                                    if (!WinModelName.GroupingName.ToString().Equals(""))
                                    { ModelName.M13 = WinModelName.GroupingName.ToString(); }
                                }
                            }
                        }

                        // ------- 원그래프 값 설정 -------
                        if (dt.Rows.Count > 0)
                        {
                            foreach (var n in pd.Slice)
                            {
                                var pieSeries = new PieSeries
                                {
                                    Title = n.Key,
                                    Values = new ChartValues<double> { n.Value },
                                    Fill = new SolidColorBrush(lvcProductPieChart.GetNextDefaultColor())
                                };

                                //계열, 범례, 데이터레이블을 보여주라!!
                                pieSeries.DataLabels = true;
                                lvcProductPieChart.Series.Add(pieSeries);

                                //조각 설정
                                var ChartDataGrid = new ChartGrid()
                                {
                                    ColorName = n.Key,
                                    FillColor = pieSeries.Fill.ToString(),
                                    Percentage = n.Value.ToString()
                                };

                                pieSeries.LabelPoint = Point => pieSeries.Title + ", " + ChartDataGrid.Percentage + "%";
                            }
                        }
                        // --------------------------------

                        //불량수량을 넣을 수 있을까
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            DataRow dr = dt.Rows[i];
                            var WinModelQty = new Win_Qul_DefectArticle_Q_ModelOccupy_CodeView
                            {
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                DefectCount = stringFormatN0(dr["DefectCount"]),
                                GroupingName = dr["GroupingName"].ToString(),
                                SumDefectQty = stringFormatN0(dr["SumDefectQty"]),
                                SumDefectCount = stringFormatN0(dr["SumDefectCount"]),

                                DefectQtyRate = dr["DefectQtyRate"].ToString(),
                                DefectCountRate = dr["DefectCountRate"].ToString(),
                            };

                            DataGridRow dgr = lib.GetRow(1, dgdDefectArticle_ModelOccupy);
                            var ModelQty = dgr.Item as Win_Qul_DefectArticle_Q_ModelOccupy_CodeView;


                            if (i < dt.Rows.Count)
                            {
                                if (i == 0)
                                {

                                    ModelQty.M1 = M1;
                                    ModelQty.M2 = M2;
                                    ModelQty.M3 = M3;
                                    ModelQty.M4 = M4;
                                    ModelQty.M5 = M5;
                                    ModelQty.M6 = M6;
                                    ModelQty.M7 = M7;
                                    ModelQty.M8 = M8;
                                    ModelQty.M9 = M9;
                                    ModelQty.M10 = M10;
                                    ModelQty.M11 = M11;
                                    ModelQty.M12 = M12;
                                    ModelQty.M13 = M13;

                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M1 = WinModelQty.DefectQty.ToString(); }
                                }
                                if (i == 1)
                                {
                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M2 = WinModelQty.DefectQty.ToString(); }
                                }
                                if (i == 2)
                                {
                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M3 = WinModelQty.DefectQty.ToString(); }
                                }
                                if (i == 3)
                                {
                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M4 = WinModelQty.DefectQty.ToString(); }
                                }
                                if (i == 4)
                                {
                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M5 = WinModelQty.DefectQty.ToString(); }
                                }
                                if (i == 5)
                                {
                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M6 = WinModelQty.DefectQty.ToString(); }
                                }
                                if (i == 6)
                                {
                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M7 = WinModelQty.DefectQty.ToString(); }
                                }
                                if (i == 7)
                                {
                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M8 = WinModelQty.DefectQty.ToString(); }
                                }
                                if (i == 8)
                                {
                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M9 = WinModelQty.DefectQty.ToString(); }
                                }
                                if (i == 9)
                                {
                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M10 = WinModelQty.DefectQty.ToString(); }
                                }
                                if (i == 10)
                                {
                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M11 = WinModelQty.DefectQty.ToString(); }
                                }
                                if (i == 11)
                                {
                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M12 = WinModelQty.DefectQty.ToString(); }
                                }
                                if (i == 12)
                                {
                                    if (!WinModelQty.DefectQty.ToString().Equals(""))
                                    { ModelQty.M13 = WinModelQty.DefectQty.ToString(); }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //종료겠지..
            }

            //원그래프소환
            //FillChartModel();
        }


        //제품별 점유율을 구현할 수 있을까.  //FillGridModel 먼저 조회하고 fillchartModel 실행해서 넣어줄 것
        private void FillChartModel()
        {
            try
            {
                //원그래프를 시작해보자
                PieData pd = new PieData();

                int TargetCount = PieChartProductValue.Items.Count;
                for (int i = 0; i < TargetCount; i++)
                {
                    DataGridRow dgr = lib.GetRow(i, PieChartProductValue);
                    var ViewReceiver = dgr.Item as Win_Qul_DefectArticle_Q_ModelOccupy_CodeView;

                    double value = 0.0;
                    if (ViewReceiver.DefectQtyRate == string.Empty)
                        ViewReceiver.DefectQtyRate = "0";

                    value = Convert.ToDouble(ViewReceiver.DefectQtyRate);
                    pd.AddSlice(ViewReceiver.GroupingName, value);
                }

                if (TargetCount > 0)
                {
                    foreach (var n in pd.Slice)
                    {
                        var pieSeries = new PieSeries
                        {
                            Title = n.Key,
                            Values = new ChartValues<double> { n.Value },
                            Fill = new SolidColorBrush(lvcProductPieChart.GetNextDefaultColor())
                        };


                        //계열, 범례, 데이터레이블을 보여주라!!
                        pieSeries.DataLabels = true;

                        lvcProductPieChart.Series.Add(pieSeries);


                        //조각 설정
                        var ChartDataGrid = new ChartGrid()
                        {
                            ColorName = n.Key,
                            FillColor = pieSeries.Fill.ToString(),
                            Percentage = n.Value.ToString()
                        };

                        pieSeries.LabelPoint = Point => pieSeries.Title + ", " + ChartDataGrid.Percentage + "%";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //종료겠지..
            }
        }


        //유형별 불량 dgdDefectArticle_DefectType 데이터그리드 조회
        private void FillGridType()
        {
            string InsPoint = string.Empty;  //불량 발생 시점

            if (cboOccurStepSrh.SelectedValue.ToString().Equals("0")) { InsPoint = ""; } //전체
            else if (cboOccurStepSrh.SelectedValue.ToString().Equals("2")) { InsPoint = "3"; }
            else if (cboOccurStepSrh.SelectedValue.ToString().Equals("3")) { InsPoint = "2"; }
            //else if (cboOccurStepSrh.SelectedValue.ToString().Equals("4")) { InsPoint = "0"; } //
            else if (cboOccurStepSrh.SelectedValue.ToString().Equals("4")) { InsPoint = "3"; } //
            else if (cboOccurStepSrh.SelectedValue.ToString().Equals("6")) { InsPoint = "5"; } //
            else { InsPoint = cboOccurStepSrh.SelectedValue.ToString(); }

            //헤더변경
            for (int i = 0; i < 12; i++)
            {
                dgdDefectArticle_DefectType.Columns[i].Header = (i + 1) + "위";
            }

            //마지막행은 13위부터의 합계
            dgdDefectArticle_DefectType.Columns[12].Header = "기타";

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

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

                sqlParameter.Clear();
                sqlParameter.Add("YYYYMM", chkDate.IsChecked == true ? dtpSDate.SelectedDate.ToString().Replace("-", "").Substring(0, 6) : "");
                sqlParameter.Add("chkProdGroupID", chkProductGroup.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProdGroupID", chkProductGroup.IsChecked == true ? cboProductGrpID.SelectedValue.ToString() : "");
                sqlParameter.Add("chkInspectPointID", 1);   // 뭐든 체크는 되어있을 거니까
                sqlParameter.Add("InspectPointID", InsPoint);
                sqlParameter.Add("chkArticleID", chkArticleID);
                sqlParameter.Add("ArticleID", ArticleID);
                sqlParameter.Add("sGrouping", "1"); //varchar(1) 이라서 "1" ,  1 : 유형 , 3 : 제품
                sqlParameter.Add("BuyerArticleNme", chkArticle.IsChecked == true ? txtArticle.Text : "");
                sqlParameter.Add("BuyerArticleNo", chkArticleNo.IsChecked == true ? txtArticleNo.Text : "");

                ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectArticle_DefectCountsub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다. (불량유형)");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        //원그래프를 시작해보자
                        PieData pd = new PieData();

                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            DataRow dr = dt.Rows[j];
                            WinTypeName = new Win_Qul_DefectArticle_Q_DefectType_CodeView
                            {
                                DefectQty = dr["DefectQty"].ToString(),
                                DefectCount = dr["DefectCount"].ToString(),
                                GroupingName = dr["GroupingName"].ToString(),
                                SumDefectQty = dr["SumDefectQty"].ToString(),
                                SumDefectCount = dr["SumDefectCount"].ToString(),

                                DefectQtyRate = dr["DefectQtyRate"].ToString(),
                                DefectCountRate = dr["DefectCountRate"].ToString(),
                            };

                            PieChartTypeValue.Items.Add(WinTypeName);

                            DataGridRow dgr = lib.GetRow(0, dgdDefectArticle_DefectType);
                            var TypeName = dgr.Item as Win_Qul_DefectArticle_Q_DefectType_CodeView;

                            // ---------- 원그래프 값 설정 ----------
                            double value = 0.0;
                            if (WinTypeName.DefectQtyRate == string.Empty)
                                WinTypeName.DefectQtyRate = "0";

                            value = Convert.ToDouble(WinTypeName.DefectQtyRate);
                            pd.AddSlice(WinTypeName.GroupingName, value);
                            // -----------------------------------

                            if (j < dt.Rows.Count)
                            {
                                if (j == 0)
                                {
                                    TypeName.M1 = M1;
                                    TypeName.M2 = M2;
                                    TypeName.M3 = M3;
                                    TypeName.M4 = M4;
                                    TypeName.M5 = M5;
                                    TypeName.M6 = M6;
                                    TypeName.M7 = M7;
                                    TypeName.M8 = M8;
                                    TypeName.M9 = M9;
                                    TypeName.M10 = M10;
                                    TypeName.M11 = M11;
                                    TypeName.M12 = M12;
                                    TypeName.M13 = M13;


                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M1 = WinTypeName.GroupingName.ToString(); }
                                }
                                if (j == 1)
                                {
                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M2 = WinTypeName.GroupingName.ToString(); }
                                }
                                if (j == 2)
                                {
                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M3 = WinTypeName.GroupingName.ToString(); }
                                }
                                if (j == 3)
                                {
                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M4 = WinTypeName.GroupingName.ToString(); }
                                }
                                if (j == 4)
                                {
                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M5 = WinTypeName.GroupingName.ToString(); }
                                }
                                if (j == 5)
                                {
                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M6 = WinTypeName.GroupingName.ToString(); }
                                }
                                if (j == 6)
                                {
                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M7 = WinTypeName.GroupingName.ToString(); }
                                }
                                if (j == 7)
                                {
                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M8 = WinTypeName.GroupingName.ToString(); }
                                }
                                if (j == 8)
                                {
                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M9 = WinTypeName.GroupingName.ToString(); }
                                }
                                if (j == 9)
                                {
                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M10 = WinTypeName.GroupingName.ToString(); }
                                }
                                if (j == 10)
                                {
                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M11 = WinTypeName.GroupingName.ToString(); }
                                }
                                if (j == 11)
                                {
                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M12 = WinTypeName.GroupingName.ToString(); }
                                }
                                if (j == 12)
                                {
                                    if (!WinTypeName.GroupingName.ToString().Equals(""))
                                    { TypeName.M13 = WinTypeName.GroupingName.ToString(); }
                                }
                            }
                        }

                        // ---------- 원그래프 값 설정 ----------
                        if (dt.Rows.Count > 0)
                        {
                            foreach (var n in pd.Slice)
                            {
                                var pieSeries = new PieSeries
                                {
                                    Title = n.Key,
                                    Values = new ChartValues<double> { n.Value },
                                    Fill = new SolidColorBrush(lvcTypePieChart.GetNextDefaultColor())
                                };

                                //계열, 범례, 데이터레이블을 보여주라!!
                                pieSeries.DataLabels = true;
                                pieSeries.LabelPoint = Point => pieSeries.Title + ", " + n.Value + "%";

                                lvcTypePieChart.Series.Add(pieSeries);

                                //조각 설정
                                var ChartDataGrid = new ChartGrid()
                                {
                                    ColorName = n.Key,
                                    FillColor = pieSeries.Fill.ToString(),
                                    Percentage = n.Value.ToString()
                                };
                            }
                        }
                        // -----------------------------------

                        //불량수량을 넣을 수 있을까
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            DataRow dr = dt.Rows[i];
                            var WinTypeQty = new Win_Qul_DefectArticle_Q_DefectType_CodeView
                            {
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                DefectCount = stringFormatN0(dr["DefectCount"]),
                                GroupingName = dr["GroupingName"].ToString(),
                                SumDefectQty = stringFormatN0(dr["SumDefectQty"]),
                                SumDefectCount = stringFormatN0(dr["SumDefectCount"]),

                                DefectQtyRate = dr["DefectQtyRate"].ToString(),
                                DefectCountRate = dr["DefectCountRate"].ToString(),
                            };

                            DataGridRow dgr = lib.GetRow(1, dgdDefectArticle_DefectType);
                            var TypeQty = dgr.Item as Win_Qul_DefectArticle_Q_DefectType_CodeView;


                            if (i < dt.Rows.Count)
                            {
                                if (i == 0)
                                {

                                    TypeQty.M1 = M1;
                                    TypeQty.M2 = M2;
                                    TypeQty.M3 = M3;
                                    TypeQty.M4 = M4;
                                    TypeQty.M5 = M5;
                                    TypeQty.M6 = M6;
                                    TypeQty.M7 = M7;
                                    TypeQty.M8 = M8;
                                    TypeQty.M9 = M9;
                                    TypeQty.M10 = M10;
                                    TypeQty.M11 = M11;
                                    TypeQty.M12 = M12;
                                    TypeQty.M13 = M13;

                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M1 = WinTypeQty.DefectQty.ToString(); }
                                }
                                if (i == 1)
                                {
                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M2 = WinTypeQty.DefectQty.ToString(); }
                                }
                                if (i == 2)
                                {
                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M3 = WinTypeQty.DefectQty.ToString(); }
                                }
                                if (i == 3)
                                {
                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M4 = WinTypeQty.DefectQty.ToString(); }
                                }
                                if (i == 4)
                                {
                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M5 = WinTypeQty.DefectQty.ToString(); }
                                }
                                if (i == 5)
                                {
                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M6 = WinTypeQty.DefectQty.ToString(); }
                                }
                                if (i == 6)
                                {
                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M7 = WinTypeQty.DefectQty.ToString(); }
                                }
                                if (i == 7)
                                {
                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M8 = WinTypeQty.DefectQty.ToString(); }
                                }
                                if (i == 8)
                                {
                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M9 = WinTypeQty.DefectQty.ToString(); }
                                }
                                if (i == 9)
                                {
                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M10 = WinTypeQty.DefectQty.ToString(); }
                                }
                                if (i == 10)
                                {
                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M11 = WinTypeQty.DefectQty.ToString(); }
                                }
                                if (i == 11)
                                {
                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M12 = WinTypeQty.DefectQty.ToString(); }
                                }
                                if (i == 12)
                                {
                                    if (!WinTypeQty.DefectQty.ToString().Equals(""))
                                    { TypeQty.M13 = WinTypeQty.DefectQty.ToString(); }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //종료겠지..
            }

            //원그래프소환
            //FillChartType();
        }


        //유형별 점유율을 구현할 수 있을까.  //FillGridType 먼저 조회하고 fillchartType 실행해서 넣어줄 것
        private void FillChartType()
        {
            try
            {
                //원그래프를 시작해보자
                PieData pd = new PieData();

                int TargetCount = PieChartTypeValue.Items.Count;
                for (int i = 0; i < TargetCount; i++)
                {
                    DataGridRow dgr = lib.GetRow(i, PieChartTypeValue);
                    var ViewReceiver = dgr.Item as Win_Qul_DefectArticle_Q_DefectType_CodeView;

                    double value = 0.0;
                    if (ViewReceiver.DefectQtyRate == string.Empty)
                        ViewReceiver.DefectQtyRate = "0";

                    value = Convert.ToDouble(ViewReceiver.DefectQtyRate);
                    pd.AddSlice(ViewReceiver.GroupingName, value);
                }

                if (TargetCount > 0)
                {
                    foreach (var n in pd.Slice)
                    {
                        var pieSeries = new PieSeries
                        {
                            Title = n.Key,
                            Values = new ChartValues<double> { n.Value },
                            Fill = new SolidColorBrush(lvcTypePieChart.GetNextDefaultColor())
                        };

                        //계열, 범례, 데이터레이블을 보여주라!!
                        pieSeries.DataLabels = true;

                        pieSeries.LabelPoint = Point => pieSeries.Title + ", " + n.Value + "%";



                        lvcTypePieChart.Series.Add(pieSeries);


                        //조각 설정
                        var ChartDataGrid = new ChartGrid()
                        {
                            ColorName = n.Key,
                            FillColor = pieSeries.Fill.ToString(),
                            Percentage = n.Value.ToString()
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //종료겠지..
            }
        }

        //작업자 불량 dgdDefectArticle_Worker 데이터그리드 조회
        private void FillGridWorker()
        {
            string InsPoint = string.Empty;  //불량 발생 시점

            if (cboOccurStepSrh.SelectedValue.ToString().Equals("0")) { InsPoint = ""; } //전체
            else if (cboOccurStepSrh.SelectedValue.ToString().Equals("2")) { InsPoint = "3"; }
            else if (cboOccurStepSrh.SelectedValue.ToString().Equals("3")) { InsPoint = "2"; }
            //else if (cboOccurStepSrh.SelectedValue.ToString().Equals("4")) { InsPoint = "0"; } //
            else if (cboOccurStepSrh.SelectedValue.ToString().Equals("4")) { InsPoint = "3"; } //
            else if (cboOccurStepSrh.SelectedValue.ToString().Equals("6")) { InsPoint = "5"; } //
            else { InsPoint = cboOccurStepSrh.SelectedValue.ToString(); }

            //헤더변경
            for (int i = 0; i < 12; i++)
            {
                dgdDefectArticle_Worker.Columns[i].Header = (i + 1) + "위";
            }

            //마지막행은 13위부터의 합계
            dgdDefectArticle_Worker.Columns[12].Header = "기타";

            try
            {
                DataSet ds = null;

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

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("YYYYMM", chkDate.IsChecked == true ? dtpSDate.SelectedDate.ToString().Replace("-", "").Substring(0, 6) : "");
                sqlParameter.Add("chkProdGroupID", chkProductGroup.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ProdGroupID", chkProductGroup.IsChecked == true ? cboProductGrpID.SelectedValue.ToString() : "");
                sqlParameter.Add("chkInspectPointID", 1);   // 뭐든 체크는 되어있을 거니까
                sqlParameter.Add("InspectPointID", InsPoint);

                sqlParameter.Add("chkArticleID", chkArticleID);
                sqlParameter.Add("ArticleID", ArticleID);
                sqlParameter.Add("sGrouping", "4"); //varchar(1) 이라서 "3" ,  1 : 유형 , 3 : 제품
                sqlParameter.Add("BuyerArticleNme", chkArticle.IsChecked == true ? txtArticle.Text : "");
                sqlParameter.Add("BuyerArticleNo", chkArticleNo.IsChecked == true ? txtArticleNo.Text : "");


                ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsWorker_DefectCountsub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다. (점유제품)");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            DataRow dr = dt.Rows[j];
                            WinWorkerName = new Win_Qul_DefectArticle_Q_Worker_CodeView
                            {
                                DefectQty = dr["DefectQty"].ToString(),
                                DefectCount = dr["DefectCount"].ToString(),
                                GroupingName = dr["GroupingName"].ToString(),
                                SumDefectQty = dr["SumDefectQty"].ToString(),
                                SumDefectCount = dr["SumDefectCount"].ToString(),

                                DefectQtyRate = dr["DefectQtyRate"].ToString(),
                                DefectCountRate = dr["DefectCountRate"].ToString(),
                            };

                            //PieChartProductValue.Items.Add(WinWorkerName);

                            DataGridRow dgr = lib.GetRow(0, dgdDefectArticle_Worker);
                            var WorkerName = dgr.Item as Win_Qul_DefectArticle_Q_Worker_CodeView;

                            if (j < dt.Rows.Count)
                            {
                                if (j == 0)
                                {
                                    WorkerName.M1 = M1;
                                    WorkerName.M2 = M2;
                                    WorkerName.M3 = M3;
                                    WorkerName.M4 = M4;
                                    WorkerName.M5 = M5;
                                    WorkerName.M6 = M6;
                                    WorkerName.M7 = M7;
                                    WorkerName.M8 = M8;
                                    WorkerName.M9 = M9;
                                    WorkerName.M10 = M10;
                                    WorkerName.M11 = M11;
                                    WorkerName.M12 = M12;
                                    WorkerName.M13 = M13;

                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M1 = WinWorkerName.GroupingName.ToString(); }
                                }
                                if (j == 1)
                                {
                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M2 = WinWorkerName.GroupingName.ToString(); }
                                }
                                if (j == 2)
                                {
                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M3 = WinWorkerName.GroupingName.ToString(); }
                                }
                                if (j == 3)
                                {
                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M4 = WinWorkerName.GroupingName.ToString(); }
                                }
                                if (j == 4)
                                {
                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M5 = WinWorkerName.GroupingName.ToString(); }
                                }
                                if (j == 5)
                                {
                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M6 = WinWorkerName.GroupingName.ToString(); }
                                }
                                if (j == 6)
                                {
                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M7 = WinWorkerName.GroupingName.ToString(); }
                                }
                                if (j == 7)
                                {
                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M8 = WinWorkerName.GroupingName.ToString(); }
                                }
                                if (j == 8)
                                {
                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M9 = WinWorkerName.GroupingName.ToString(); }
                                }
                                if (j == 9)
                                {
                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M10 = WinWorkerName.GroupingName.ToString(); }
                                }
                                if (j == 10)
                                {
                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M11 = WinWorkerName.GroupingName.ToString(); }
                                }
                                if (j == 11)
                                {
                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M12 = WinWorkerName.GroupingName.ToString(); }
                                }
                                if (j == 12)
                                {
                                    if (!WinWorkerName.GroupingName.ToString().Equals(""))
                                    { WorkerName.M13 = WinWorkerName.GroupingName.ToString(); }
                                }
                            }
                        }

                        //불량수량을 넣을 수 있을까
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            DataRow dr = dt.Rows[i];
                            var WinWorkerQty = new Win_Qul_DefectArticle_Q_Worker_CodeView
                            {
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                DefectCount = stringFormatN0(dr["DefectCount"]),
                                GroupingName = dr["GroupingName"].ToString(),
                                SumDefectQty = stringFormatN0(dr["SumDefectQty"]),
                                SumDefectCount = stringFormatN0(dr["SumDefectCount"]),

                                DefectQtyRate = dr["DefectQtyRate"].ToString(),
                                DefectCountRate = dr["DefectCountRate"].ToString(),
                            };

                            DataGridRow dgr = lib.GetRow(1, dgdDefectArticle_Worker);
                            var WorkerQty = dgr.Item as Win_Qul_DefectArticle_Q_Worker_CodeView;


                            if (i < dt.Rows.Count)
                            {
                                if (i == 0)
                                {

                                    WorkerQty.M1 = M1;
                                    WorkerQty.M2 = M2;
                                    WorkerQty.M3 = M3;
                                    WorkerQty.M4 = M4;
                                    WorkerQty.M5 = M5;
                                    WorkerQty.M6 = M6;
                                    WorkerQty.M7 = M7;
                                    WorkerQty.M8 = M8;
                                    WorkerQty.M9 = M9;
                                    WorkerQty.M10 = M10;
                                    WorkerQty.M11 = M11;
                                    WorkerQty.M12 = M12;
                                    WorkerQty.M13 = M13;

                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M1 = WinWorkerQty.DefectQty.ToString(); }
                                }
                                if (i == 1)
                                {
                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M2 = WinWorkerQty.DefectQty.ToString(); }
                                }
                                if (i == 2)
                                {
                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M3 = WinWorkerQty.DefectQty.ToString(); }
                                }
                                if (i == 3)
                                {
                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M4 = WinWorkerQty.DefectQty.ToString(); }
                                }
                                if (i == 4)
                                {
                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M5 = WinWorkerQty.DefectQty.ToString(); }
                                }
                                if (i == 5)
                                {
                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M6 = WinWorkerQty.DefectQty.ToString(); }
                                }
                                if (i == 6)
                                {
                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M7 = WinWorkerQty.DefectQty.ToString(); }
                                }
                                if (i == 7)
                                {
                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M8 = WinWorkerQty.DefectQty.ToString(); }
                                }
                                if (i == 8)
                                {
                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M9 = WinWorkerQty.DefectQty.ToString(); }
                                }
                                if (i == 9)
                                {
                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M10 = WinWorkerQty.DefectQty.ToString(); }
                                }
                                if (i == 10)
                                {
                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M11 = WinWorkerQty.DefectQty.ToString(); }
                                }
                                if (i == 11)
                                {
                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M12 = WinWorkerQty.DefectQty.ToString(); }
                                }
                                if (i == 12)
                                {
                                    if (!WinWorkerQty.DefectQty.ToString().Equals(""))
                                    { WorkerQty.M13 = WinWorkerQty.DefectQty.ToString(); }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //종료겠지..
            }

        }

        //주 모니터 대신 어플리케이션 활성화 영역을 캡쳐
        private void BtnCapture_Click(object sender, RoutedEventArgs e)
        {
            FrameworkElement targetElement = this;
            lib.ScreenCapture(targetElement);
        }

        //제품그룹 체크
        private void ChkProductGroup_Checked(object sender, RoutedEventArgs e)
        {
            cboProductGrpID.IsEnabled = true;
        }

        //제품그룹 체크해제
        private void ChkProductGroup_Unchecked(object sender, RoutedEventArgs e)
        {
            cboProductGrpID.IsEnabled = false;
        }

        //제품그룹 라벨 클릭
        private void LblProductGroup_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkProductGroup.IsChecked == true) { chkProductGroup.IsChecked = false; }
            else { chkProductGroup.IsChecked = true; }
        }

        // 천자리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천자리 콤마, 소수점 두 자리까지
        private string stringFormatN2(object obj)
        {
            return string.Format("{0:N2}", obj);
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
    }



    public class Win_Qul_DefectArticle_Q_DefectCount_CodeView : BaseView
    {
        // 그리드 로우컬럼 헤더._불량지수(건수)
        public string DefectCount_RowHeaderColumns { get; set; }

        public string M1 { get; set; }
        public string M2 { get; set; }
        public string M3 { get; set; }
        public string M4 { get; set; }
        public string M5 { get; set; }
        public string M6 { get; set; }
        public string M7 { get; set; }
        public string M8 { get; set; }
        public string M9 { get; set; }
        public string M10 { get; set; }
        public string M11 { get; set; }
        public string M12 { get; set; }
        public string M13 { get; set; }



    }

    public class Win_Qul_DefectArticle_Q_ModelOccupy_CodeView : BaseView
    {
        // 그리드 로우컬럼 헤더._점유모델(점유제품)
        public string ModelOccupy_RowHeaderColumns { get; set; }


        public string DefectQty { get; set; }          //불량 갯수
        public string DefectCount { get; set; }        //불량 발생 수
        public string GroupingName { get; set; }    //제품명 나오는데....
        public string SumDefectQty { get; set; }       //불량 수 합계
        public string SumDefectCount { get; set; }     //불량 발생 합계

        public string DefectQtyRate { get; set; }   //불량률
        public string DefectCountRate { get; set; } //불량발생률

        public string M1 { get; set; }
        public string M2 { get; set; }
        public string M3 { get; set; }
        public string M4 { get; set; }
        public string M5 { get; set; }
        public string M6 { get; set; }
        public string M7 { get; set; }
        public string M8 { get; set; }
        public string M9 { get; set; }
        public string M10 { get; set; }
        public string M11 { get; set; }
        public string M12 { get; set; }
        public string M13 { get; set; }



    }

    public class Win_Qul_DefectArticle_Q_DefectType_CodeView : BaseView
    {
        // 그리드 로우컬럼 헤더._불량유형
        public string DefectType_RowHeaderColumns { get; set; }

        public string DefectQty { get; set; }          //불량 갯수
        public string DefectCount { get; set; }        //불량 발생 수
        public string GroupingName { get; set; }    //제품명 나오는데....
        public string SumDefectQty { get; set; }       //불량 수 합계
        public string SumDefectCount { get; set; }     //불량 발생 합계

        public string DefectQtyRate { get; set; }   //불량률
        public string DefectCountRate { get; set; } //불량발생률

        public string M1 { get; set; }
        public string M2 { get; set; }
        public string M3 { get; set; }
        public string M4 { get; set; }
        public string M5 { get; set; }
        public string M6 { get; set; }
        public string M7 { get; set; }
        public string M8 { get; set; }
        public string M9 { get; set; }
        public string M10 { get; set; }
        public string M11 { get; set; }
        public string M12 { get; set; }
        public string M13 { get; set; }
    }


    public class Win_Qul_DefectArticle_Q_Sum_Daily_CodeView : BaseView
    {
        public int Num { get; set; }
        public string InspectDate { get; set; }
        public string GroupingName { get; set; }
        public string DefectQty { get; set; }
        public string RepairQty { get; set; }
        public string RepairRate { get; set; }
    }

    public class ChartGrid : BaseView
    {
        public string FillColor { get; set; }
        public string ColorName { get; set; }
        public string Percentage { get; set; }
    }

    public class Win_Qul_DefectArticle_Q_Worker_CodeView : BaseView
    {
        // 그리드 로우컬럼 헤더._점유모델(점유제품)
        public string Worker_RowHeaderColumns { get; set; }

        public string DefectQty { get; set; }          //불량 갯수
        public string DefectCount { get; set; }        //불량 발생 수
        public string GroupingName { get; set; }    //제품명 나오는데....
        public string SumDefectQty { get; set; }       //불량 수 합계
        public string SumDefectCount { get; set; }     //불량 발생 합계

        public string DefectQtyRate { get; set; }   //불량률
        public string DefectCountRate { get; set; } //불량발생률

        public string M1 { get; set; }
        public string M2 { get; set; }
        public string M3 { get; set; }
        public string M4 { get; set; }
        public string M5 { get; set; }
        public string M6 { get; set; }
        public string M7 { get; set; }
        public string M8 { get; set; }
        public string M9 { get; set; }
        public string M10 { get; set; }
        public string M11 { get; set; }
        public string M12 { get; set; }
        public string M13 { get; set; }



    }
}
