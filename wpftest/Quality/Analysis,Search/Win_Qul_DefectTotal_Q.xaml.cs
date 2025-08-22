using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_sts_DefectTotal_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_DefectTotal_Q : UserControl
    {
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();


        public Win_Qul_DefectTotal_Q()
        {
            InitializeComponent();
        }


        // 첫 로드시.
        private void Win_Qul_sts_DefectTotal_Q_Loaded(object sender, RoutedEventArgs e)
        {
            First_Step();
            CreateDataGridColumns();
        }


        #region 첫단계 / 날짜버튼 세팅 / 조회용 체크박스 세팅
        // 첫 단계
        private void First_Step()
        {
            chkYearDate.IsChecked = true;
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM");

            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
        }

        // 이전년도
        private void btnLastYear_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = lib.BringLastYearDatetimeContinue(dtpFromDate.SelectedDate.Value)[0];
            dtpToDate.SelectedDate = lib.BringLastYearDatetimeContinue(dtpFromDate.SelectedDate.Value)[1];
        }

        // 금년도 당해
        private void btnThisYear_Click(object sender, RoutedEventArgs e)
        {
            string[] receiver = lib.BringThisYearDatetime();

            dtpFromDate.SelectedDate = Convert.ToDateTime(receiver[0]);
            dtpToDate.SelectedDate = Convert.ToDateTime(receiver[1]);
        }


        // 검사일자
        private void chkYearDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkYearDate.IsChecked == true)
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
        private void chkYearDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkYearDate.IsChecked == true)
            {
                chkYearDate.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkYearDate.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }
        // 품명
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
        // 품명
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

        #endregion


        // 플러스 파인더 _ 품명.(품번으로 변경요청, 2020.03.24, 장가빈)
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticle, 81, txtArticle.Text);
        }
        // 품명 _ 키다운 (품번으로 변경요청, 2020.03.24, 장가빈)
        private void TxtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticle, 81, txtArticle.Text);
            }
        }


        #region 데이터 그리드 로우컬럼 자동생성.
        // 그리드 로우컬럼 자동생성.
        private void CreateDataGridColumns()
        {
            string[] RowHeaderName = new string[9];
            RowHeaderName[0] = "전체";
            RowHeaderName[1] = "인수목표";
            RowHeaderName[2] = "인수실적";
            RowHeaderName[3] = "공정목표";
            RowHeaderName[4] = "공정실적";
            RowHeaderName[5] = "출하목표";
            RowHeaderName[6] = "출하실적";
            RowHeaderName[7] = "고객목표";
            RowHeaderName[8] = "고객실적";

            for (int i = 0; i < 9; i++)
            {
                var Win_Qul_sts_DefectTotal_Q_Insert = new Win_Qul_DefectTotal_Q_View()
                {
                    TOTAL_RowHeaderColumns = RowHeaderName[i]
                };
                dgdDefectTotal_TotalGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_Insert);
            }

            RowHeaderName = new string[4];
            RowHeaderName[0] = "작업수량";
            RowHeaderName[1] = "불량수량";
            RowHeaderName[2] = "불량PPM";
            RowHeaderName[3] = "목표PPM";

            for (int i = 0; i < 4; i++)
            {
                var Win_Qul_sts_DefectTotal_Q_Insert = new Win_Qul_DefectTotal_Q_View()
                {
                    INSU_RowHeaderColumns = RowHeaderName[i]
                };
                dgdDefectTotal_INSU_MonthGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_Insert);
                dgdDefectTotal_PROC_MonthGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_Insert);
                dgdDefectTotal_SHIP_MonthGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_Insert);
                dgdDefectTotal_CUST_MonthGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_Insert);
            }
        }

        #endregion


        #region 검색버튼 클릭.
        // 검색버튼 클릭.
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                TabItem NowTI = tabconGrid.SelectedItem as TabItem;
                string sYYYY = dtpFromDate.Text.Substring(0, 4);
                string SDate = sYYYY + "01";
                string EDate = sYYYY + "12";

                int ChkArticleID = 0;
                string ArticleID = string.Empty;


                if (NowTI.Header.ToString() == "전체")
                {
                    // for문 돌면서 각 단계별 금년목표 + 월별 목표 값 구하기.
                    for (int i = 1; i < 8; i = i + 2)
                    {
                        FillGrid_Total_MonthlyGoal(i, sYYYY);
                    }
                    // 그리드 나머지.
                    FillGrid_Total_Monthly();
                }

                else if (NowTI.Header.ToString() == "인수")
                {
                    // 월 종합 그리드 + 상단 막대그래프.
                    FillGrid_InSu__Monthly(SDate, EDate, ChkArticleID, ArticleID);

                    string sMM_I = cboINSU_Month.SelectedValue.ToString();
                    if (sMM_I.Length == 1) { sMM_I = "0" + sMM_I; }

                    FillGrid_InSu_Symptom(sYYYY, sMM_I, ChkArticleID, ArticleID);
                }
                else if (NowTI.Header.ToString() == "자주/공정")
                {
                    // 월 종합 그리드 + 상단 막대그래프.
                    FillGrid_Proc_Monthly(SDate, EDate, ChkArticleID, ArticleID);

                    string sMM_P = cboPROC_Month.SelectedValue.ToString();
                    if (sMM_P.Length == 1) { sMM_P = "0" + sMM_P; }

                    FillGrid_Proc_Symptom(sYYYY, sMM_P, ChkArticleID, ArticleID);
                }
                else if (NowTI.Header.ToString() == "출하")
                {
                    // 월 종합 그리드 + 상단 막대그래프.
                    FillGrid_Ship_Monthly(SDate, EDate, ChkArticleID, ArticleID);

                    string sMM_S = cboSHIP_Month.SelectedValue.ToString();
                    if (sMM_S.Length == 1) { sMM_S = "0" + sMM_S; }

                    FillGrid_Ship_Symptom(sYYYY, sMM_S, ChkArticleID, ArticleID);
                }
                else if (NowTI.Header.ToString() == "고객")
                {
                    // 월 종합 그리드 + 상단 막대그래프.
                    FillGrid_Cust_Monthly(SDate, EDate, ChkArticleID, ArticleID);

                    string sMM_C = cboCUST_Month.SelectedValue.ToString();
                    if (sMM_C.Length == 1) { sMM_C = "0" + sMM_C; }

                    FillGrid_Cust_Symptom(sYYYY, sMM_C, ChkArticleID, ArticleID);
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);



        }

        #endregion


        #region 전체 탭 조회 프로시저/

        // 단계별 금년목표 + 월별 목표 값 구하기.
        private void FillGrid_Total_MonthlyGoal(int iValue, string sYYYY)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("sInspectGubun", iValue.ToString());
            sqlParameter.Add("sYYYY", sYYYY);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sDefectGoal", sqlParameter, false);

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
                    DataGridRow dgr = lib.GetRow(iValue, dgdDefectTotal_TotalGrid);  //1행, 3행, 5행, 7행에만 업데이트
                    var ViewReceiver = dgr.Item as Win_Qul_DefectTotal_Q_View;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        DataRow dr = dt.Rows[i];

                        ViewReceiver.TOTAL_AvgDefectGoal = stringFormat2(dr["AvgDefectGoal"]);
                        ViewReceiver.TOTAL_DefectGoal = stringFormat2(dr["AvgDefectGoal"]);
                        if (i == 0) { ViewReceiver.TOTAL_MM1 = dr["DefectGoal"].ToString(); }
                        if (i == 1) { ViewReceiver.TOTAL_MM2 = dr["DefectGoal"].ToString(); }
                        if (i == 2) { ViewReceiver.TOTAL_MM3 = dr["DefectGoal"].ToString(); }
                        if (i == 3) { ViewReceiver.TOTAL_MM4 = dr["DefectGoal"].ToString(); }
                        if (i == 4) { ViewReceiver.TOTAL_MM5 = dr["DefectGoal"].ToString(); }
                        if (i == 5) { ViewReceiver.TOTAL_MM6 = dr["DefectGoal"].ToString(); }
                        if (i == 6) { ViewReceiver.TOTAL_MM7 = dr["DefectGoal"].ToString(); }
                        if (i == 7) { ViewReceiver.TOTAL_MM8 = dr["DefectGoal"].ToString(); }
                        if (i == 8) { ViewReceiver.TOTAL_MM9 = dr["DefectGoal"].ToString(); }
                        if (i == 9) { ViewReceiver.TOTAL_MM10 = dr["DefectGoal"].ToString(); }
                        if (i == 10) { ViewReceiver.TOTAL_MM11 = dr["DefectGoal"].ToString(); }
                        if (i == 11) { ViewReceiver.TOTAL_MM12 = dr["DefectGoal"].ToString(); }

                    }
                }
            }

            DataStore.Instance.CloseConnection();
        }


        // 그리드 나머지 조회하기.
        private void FillGrid_Total_Monthly()
        {
            int nchkDate = 0;
            if (chkYearDate.IsChecked == true) { nchkDate = 1; }

            string StartMonth = dtpFromDate.Text.Substring(0, 4) + "01";
            string EndMonth = dtpToDate.Text.Substring(0, 4) + "12";

            int ChkArticleID = 0;              //품명.
            string ArticleID = string.Empty;
            string Article = string.Empty;

            //if (chkArticle.IsChecked == true)
            //{
            //    if (txtArticle.Tag == null)
            //    {
            //        txtArticle.Tag = "";
            //        if (txtArticle.Text.Length > 0)
            //        {
            //            ChkArticleID = 2;
            //            Article = txtArticle.Text;
            //        }
            //    }
            //    else
            //    {
            //        ChkArticleID = 1;
            //        ArticleID = txtArticle.Tag.ToString();
            //    }
            //}


            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("nchkDate", nchkDate);               //int
            sqlParameter.Add("StartMonth", StartMonth);
            sqlParameter.Add("EndMonth", EndMonth);
            sqlParameter.Add("nchkDefectStep", 0);         //int
            sqlParameter.Add("DefectStep", "");

            sqlParameter.Add("nchkCustom", 0);             //int
            sqlParameter.Add("CustomID", "");

            sqlParameter.Add("nchkArticleID", 0); // ChkArticleID);          //int
            sqlParameter.Add("ArticleID", ""); // ArticleID);
            sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : ""); // ArticleID);
            //sqlParameter.Add("Article", Article);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectTotal_Monthly", sqlParameter, false);
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
                    int x;

                    for (int j = 1; j < dt.Rows.Count; j++)
                    {
                        DataRow dr = dt.Rows[j];
                        if (dr["SortStep"].ToString() == "0")
                        {
                            x = 0;
                        }
                        else
                        {
                            x = Convert.ToInt32(dr["SortStep"].ToString()) + 1;
                        }

                        DataGridRow dgr = lib.GetRow(x, dgdDefectTotal_TotalGrid);      // x행 별로 업데이트.
                        var ViewReceiver = dgr.Item as Win_Qul_DefectTotal_Q_View;

                        ViewReceiver.TOTAL_PreDefectGoal = dr["PreYearDefectRate"].ToString();  //전년
                        ViewReceiver.TOTAL_DefectGoal = stringFormat2(dr["YearGoalDefectRate"]);    //금년
                        ViewReceiver.TOTAL_AvgDefectGoal = stringFormat2(dr["TAvgDefectRate"]);    //금년

                        ViewReceiver.TOTAL_MM1 = stringFormat2(dr["DefectRate1"]);
                        ViewReceiver.TOTAL_MM2 = stringFormat2(dr["DefectRate2"]);
                        ViewReceiver.TOTAL_MM3 = stringFormat2(dr["DefectRate3"]);
                        ViewReceiver.TOTAL_MM4 = stringFormat2(dr["DefectRate4"]);
                        ViewReceiver.TOTAL_MM5 = stringFormat2(dr["DefectRate5"]);
                        ViewReceiver.TOTAL_MM6 = stringFormat2(dr["DefectRate6"]);
                        ViewReceiver.TOTAL_MM7 = stringFormat2(dr["DefectRate7"]);
                        ViewReceiver.TOTAL_MM8 = stringFormat2(dr["DefectRate8"]);
                        ViewReceiver.TOTAL_MM9 = stringFormat2(dr["DefectRate9"]);
                        ViewReceiver.TOTAL_MM10 = stringFormat2(dr["DefectRate10"]);
                        ViewReceiver.TOTAL_MM11 = stringFormat2(dr["DefectRate11"]);
                        ViewReceiver.TOTAL_MM12 = stringFormat2(dr["DefectRate12"]);

                        double GoalLine = 0;
                        double mm1 = 0;
                        double mm2 = 0;
                        double mm3 = 0;
                        double mm4 = 0;
                        double mm5 = 0;
                        double mm6 = 0;
                        double mm7 = 0;
                        double mm8 = 0;
                        double mm9 = 0;
                        double mm10 = 0;
                        double mm11 = 0;
                        double mm12 = 0;


                        if (x == 1)         //전체다. >> 그래프 그려야 한다.
                        {
                            GoalLine = Convert.ToDouble(ViewReceiver.TOTAL_DefectGoal);

                            mm1 = Convert.ToDouble(ViewReceiver.TOTAL_MM1);
                            mm2 = Convert.ToDouble(ViewReceiver.TOTAL_MM2);
                            mm3 = Convert.ToDouble(ViewReceiver.TOTAL_MM3);
                            mm4 = Convert.ToDouble(ViewReceiver.TOTAL_MM4);
                            mm5 = Convert.ToDouble(ViewReceiver.TOTAL_MM5);
                            mm6 = Convert.ToDouble(ViewReceiver.TOTAL_MM6);
                            mm7 = Convert.ToDouble(ViewReceiver.TOTAL_MM7);
                            mm8 = Convert.ToDouble(ViewReceiver.TOTAL_MM8);
                            mm9 = Convert.ToDouble(ViewReceiver.TOTAL_MM9);
                            mm10 = Convert.ToDouble(ViewReceiver.TOTAL_MM10);
                            mm11 = Convert.ToDouble(ViewReceiver.TOTAL_MM11);
                            mm12 = Convert.ToDouble(ViewReceiver.TOTAL_MM12);
                        }
                        SeriesCollection SeriesCollection = new SeriesCollection
                            {
                                new ColumnSeries
                                {
                                    Title = "전체 불량율",
                                    Values = new ChartValues<double>
                                    {

                                        mm1, mm2, mm3, mm4, mm5, mm6,
                                        mm7, mm8, mm9, mm10, mm11, mm12

                                    }

                                },
                                new LineSeries
                                {
                                    Title = "금년 목표선",
                                    Values = new ChartValues<double>
                                    {

                                        GoalLine,GoalLine,GoalLine,GoalLine,GoalLine,GoalLine,
                                        GoalLine,GoalLine,GoalLine,GoalLine,GoalLine,GoalLine
                                    }

                                }
                            };
                        lvcTotalChart.Series = SeriesCollection;

                    }
                }
            }

            DataStore.Instance.CloseConnection();
        }

        #endregion


        #region 인수 탭 월 종합 그리드 조회

        private void FillGrid_InSu__Monthly(string SDate, string EDate, int ChkArticleID, string ArticleID)
        {

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("nchkDate", 1);                    //int
            sqlParameter.Add("StartYYYYMM", SDate);
            sqlParameter.Add("EndYYYYMM", EDate);
            sqlParameter.Add("nchkDefectStep", 1);              //int  (인수 > 무조건 1)
            sqlParameter.Add("DefectStep", "1");

            sqlParameter.Add("nchkCustom", 0);                  //int (존재불가)
            sqlParameter.Add("CustomID", "");
            sqlParameter.Add("Custom", "");

            sqlParameter.Add("nchkArticleID", 0); // ChkArticleID);    //int
            sqlParameter.Add("ArticleID", ""); // ArticleID);
            sqlParameter.Add("Article", "");
            sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectTotal_Detail_Monthly", sqlParameter, false);

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
                    DataRow dr = dt.Rows[0];
                    for (int i = 0; i < 4; i++)
                    {
                        DataGridRow dgr = lib.GetRow(i, dgdDefectTotal_INSU_MonthGrid);
                        var ViewReceiver = dgr.Item as Win_Qul_DefectTotal_Q_View;

                        if (i == 0)             //첫번째 행 (작업수량)
                        {
                            ViewReceiver.INSU_PreDefectGoal = stringFormat2(dr["PreYearProdQty"]);
                            ViewReceiver.INSU_DefectGoal = "";
                            ViewReceiver.INSU_MM1 = stringFormat2(dr["ProdQty1"]);
                            ViewReceiver.INSU_MM2 = stringFormat2(dr["ProdQty2"]);
                            ViewReceiver.INSU_MM3 = stringFormat2(dr["ProdQty3"]);
                            ViewReceiver.INSU_MM4 = stringFormat2(dr["ProdQty4"]);
                            ViewReceiver.INSU_MM5 = stringFormat2(dr["ProdQty5"]);
                            ViewReceiver.INSU_MM6 = stringFormat2(dr["ProdQty6"]);
                            ViewReceiver.INSU_MM7 = stringFormat2(dr["ProdQty7"]);
                            ViewReceiver.INSU_MM8 = stringFormat2(dr["ProdQty8"]);
                            ViewReceiver.INSU_MM9 = stringFormat2(dr["ProdQty9"]);
                            ViewReceiver.INSU_MM10 = stringFormat2(dr["ProdQty10"]);
                            ViewReceiver.INSU_MM11 = stringFormat2(dr["ProdQty11"]);
                            ViewReceiver.INSU_MM12 = stringFormat2(dr["ProdQty12"]);
                            ViewReceiver.INSU_TotalDefectGoal = stringFormat2(dr["TProdQty"]);
                        }
                        else if (i == 1)            //두번째 행 (불량수량)
                        {
                            ViewReceiver.INSU_PreDefectGoal = stringFormat2(dr["PreYearDefectQty"]);
                            ViewReceiver.INSU_DefectGoal = "";
                            ViewReceiver.INSU_MM1 = stringFormat2(dr["DefectQty1"]);
                            ViewReceiver.INSU_MM2 = stringFormat2(dr["DefectQty2"]);
                            ViewReceiver.INSU_MM3 = stringFormat2(dr["DefectQty3"]);
                            ViewReceiver.INSU_MM4 = stringFormat2(dr["DefectQty4"]);
                            ViewReceiver.INSU_MM5 = stringFormat2(dr["DefectQty5"]);
                            ViewReceiver.INSU_MM6 = stringFormat2(dr["DefectQty6"]);
                            ViewReceiver.INSU_MM7 = stringFormat2(dr["DefectQty7"]);
                            ViewReceiver.INSU_MM8 = stringFormat2(dr["DefectQty8"]);
                            ViewReceiver.INSU_MM9 = stringFormat2(dr["DefectQty9"]);
                            ViewReceiver.INSU_MM10 = stringFormat2(dr["DefectQty10"]);
                            ViewReceiver.INSU_MM11 = stringFormat2(dr["DefectQty11"]);
                            ViewReceiver.INSU_MM12 = stringFormat2(dr["DefectQty12"]);
                            ViewReceiver.INSU_TotalDefectGoal = stringFormat2(dr["TDefectQty"]);
                        }
                        else if (i == 2)            //세번째 행 (불량율 ppm) + 그래프.
                        {
                            ViewReceiver.INSU_PreDefectGoal = stringFormat2(dr["PreYearDefectRate"]);
                            ViewReceiver.INSU_DefectGoal = "";
                            ViewReceiver.INSU_MM1 = stringFormat2(dr["DefectRate1"]);
                            ViewReceiver.INSU_MM2 = stringFormat2(dr["DefectRate2"]);
                            ViewReceiver.INSU_MM3 = stringFormat2(dr["DefectRate3"]);
                            ViewReceiver.INSU_MM4 = stringFormat2(dr["DefectRate4"]);
                            ViewReceiver.INSU_MM5 = stringFormat2(dr["DefectRate5"]);
                            ViewReceiver.INSU_MM6 = stringFormat2(dr["DefectRate6"]);
                            ViewReceiver.INSU_MM7 = stringFormat2(dr["DefectRate7"]);
                            ViewReceiver.INSU_MM8 = stringFormat2(dr["DefectRate8"]);
                            ViewReceiver.INSU_MM9 = stringFormat2(dr["DefectRate9"]);
                            ViewReceiver.INSU_MM10 = stringFormat2(dr["DefectRate10"]);
                            ViewReceiver.INSU_MM11 = stringFormat2(dr["DefectRate11"]);
                            ViewReceiver.INSU_MM12 = stringFormat2(dr["DefectRate12"]);
                            ViewReceiver.INSU_TotalDefectGoal = stringFormat2(dr["TDefectRate"]);

                            double pregoal = 0;
                            double thisgoal = 0;
                            double mm1 = 0;
                            double mm2 = 0;
                            double mm3 = 0;
                            double mm4 = 0;
                            double mm5 = 0;
                            double mm6 = 0;
                            double mm7 = 0;
                            double mm8 = 0;
                            double mm9 = 0;
                            double mm10 = 0;
                            double mm11 = 0;
                            double mm12 = 0;
                            double total = 0;

                            pregoal = Convert.ToDouble(ViewReceiver.INSU_PreDefectGoal);
                            mm1 = Convert.ToDouble(ViewReceiver.INSU_MM1);
                            mm2 = Convert.ToDouble(ViewReceiver.INSU_MM2);
                            mm3 = Convert.ToDouble(ViewReceiver.INSU_MM3);
                            mm4 = Convert.ToDouble(ViewReceiver.INSU_MM4);
                            mm5 = Convert.ToDouble(ViewReceiver.INSU_MM5);
                            mm6 = Convert.ToDouble(ViewReceiver.INSU_MM6);
                            mm7 = Convert.ToDouble(ViewReceiver.INSU_MM7);
                            mm8 = Convert.ToDouble(ViewReceiver.INSU_MM8);
                            mm9 = Convert.ToDouble(ViewReceiver.INSU_MM9);
                            mm10 = Convert.ToDouble(ViewReceiver.INSU_MM10);
                            mm11 = Convert.ToDouble(ViewReceiver.INSU_MM11);
                            mm12 = Convert.ToDouble(ViewReceiver.INSU_MM12);
                            total = Convert.ToDouble(ViewReceiver.INSU_TotalDefectGoal);

                            SeriesCollection SeriesCollection = new SeriesCollection
                            {
                                new ColumnSeries
                                {
                                    Title = "불량율 종합현황",
                                    Values = new ChartValues<double>
                                    {
                                        pregoal, thisgoal,
                                        mm1, mm2, mm3, mm4, mm5, mm6,
                                        mm7, mm8, mm9, mm10, mm11, mm12, total
                                    }

                                },
                                new LineSeries
                                {
                                    Title = "작년 기준선",
                                    Values = new ChartValues<double>
                                    {
                                        0,0,
                                        pregoal, pregoal, pregoal, pregoal, pregoal, pregoal,
                                        pregoal, pregoal, pregoal, pregoal, pregoal, pregoal, 0
                                    }
                                }
                            };
                            lvcINSUChart.Series = SeriesCollection;

                        }
                        else if (i == 3)            //네번째 행 (목표 ppm)
                        {
                            ViewReceiver.INSU_PreDefectGoal = "";
                            ViewReceiver.INSU_DefectGoal = stringFormat2(dr["YearGoalDefectRate"]);
                            ViewReceiver.INSU_MM1 = stringFormat2(dr["YearGoal1"]);
                            ViewReceiver.INSU_MM2 = stringFormat2(dr["YearGoal2"]);
                            ViewReceiver.INSU_MM3 = stringFormat2(dr["YearGoal3"]);
                            ViewReceiver.INSU_MM4 = stringFormat2(dr["YearGoal4"]);
                            ViewReceiver.INSU_MM5 = stringFormat2(dr["YearGoal5"]);
                            ViewReceiver.INSU_MM6 = stringFormat2(dr["YearGoal6"]);
                            ViewReceiver.INSU_MM7 = stringFormat2(dr["YearGoal7"]);
                            ViewReceiver.INSU_MM8 = stringFormat2(dr["YearGoal8"]);
                            ViewReceiver.INSU_MM9 = stringFormat2(dr["YearGoal9"]);
                            ViewReceiver.INSU_MM10 = stringFormat2(dr["YearGoal10"]);
                            ViewReceiver.INSU_MM11 = stringFormat2(dr["YearGoal11"]);
                            ViewReceiver.INSU_MM12 = stringFormat2(dr["YearGoal12"]);
                            ViewReceiver.INSU_TotalDefectGoal = stringFormat2(dr["TotalYearGoal"]);
                        }
                    }
                }

            }

            DataStore.Instance.CloseConnection();
        }

        #endregion

        #region 인수 탭 유형별 불량현황 그리드 조회
        // 인수 탭 유형별 불량현황 3개월치 묶음 조회
        private void FillGrid_InSu_Symptom(string sYYYY, string sMM, int ChkArticleID, string ArticleID)
        {

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("nchkDate", 1);                    //int
            sqlParameter.Add("YYYY", sYYYY);
            sqlParameter.Add("MM", sMM);
            sqlParameter.Add("nchkDefectStep", "1");            //int  (인수 > 무조건 1)
            sqlParameter.Add("DefectStep", "1");

            sqlParameter.Add("nchkCustom", 0);                  //int (존재불가)
            sqlParameter.Add("CustomID", "");

            sqlParameter.Add("nchkArticleID", 0); // ChkArticleID);    //int
            sqlParameter.Add("ArticleID", ""); // ArticleID);
            //sqlParameter.Add("Article", "");
            sqlParameter.Add("sGrouping", "1");
            sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectTotal_Detail_Symptom", sqlParameter, false);

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
                    // 원형 그래프.
                    PieData pd = new PieData();

                    double TotalQtyMinus2 = 0;
                    double TotalQtyMinus1 = 0;
                    double TotalQty = 0;

                    dgdDefectTotal_INSU_SymptomGrid.Items.Clear();
                    DataRowCollection drc = dt.Rows;
                    foreach (DataRow item in drc)
                    {
                        if (item["GroupingName"].ToString() == string.Empty)
                        {
                            continue;
                        }
                        else
                        {
                            var Win_Qul_sts_DefectTotal_Q_Insert = new Win_Qul_DefectTotal_Q_View()
                            {
                                INSU_GroupingName = item["GroupingName"].ToString(),

                                INSU_Minus2qty = item["DefectQty1"].ToString(),
                                INSU_Minus2rate = item["DefectRate1"].ToString(),
                                INSU_Minus1qty = item["DefectQty2"].ToString(),
                                INSU_Minus1rate = item["DefectRate2"].ToString(),
                                INSU_MMqty = item["DefectQty3"].ToString(),
                                INSU_MMrate = item["DefectRate3"].ToString()
                            };
                            dgdDefectTotal_INSU_SymptomGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_Insert);

                            TotalQtyMinus2 = TotalQtyMinus2 + Convert.ToDouble(item["DefectQty1"].ToString());
                            TotalQtyMinus1 = TotalQtyMinus1 + Convert.ToDouble(item["DefectQty2"].ToString());
                            TotalQty = TotalQty + Convert.ToDouble(item["DefectQty3"].ToString());

                            double value = 0;
                            value = Convert.ToDouble(item["DefectRate3"].ToString());
                            pd.AddSlice(item["GroupingName"].ToString(), value);

                        }
                    }

                    // 마지막 총계 줄 업데이트
                    var Win_Qul_sts_DefectTotal_Q_lasttotal_Insert = new Win_Qul_DefectTotal_Q_View()
                    {
                        INSU_GroupingName = "누적합계",

                        INSU_Minus2qty = TotalQtyMinus2.ToString(),
                        INSU_Minus2rate = "100",
                        INSU_Minus1qty = TotalQtyMinus1.ToString(),
                        INSU_Minus1rate = "100",
                        INSU_MMqty = TotalQty.ToString(),
                        INSU_MMrate = "100"
                    };
                    dgdDefectTotal_INSU_SymptomGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_lasttotal_Insert);

                    foreach (var n in pd.Slice)
                    {
                        INSU_PieChart.Series.Add(new PieSeries
                        {
                            Title = n.Key,
                            Values = new ChartValues<double> { n.Value }
                        });
                    }
                }
            }

            DataStore.Instance.CloseConnection();
        }

        #endregion



        #region 자주/공정 탭 월 종합 그리드 조회
        // 자주/공정 탭 월 종합 그리드 조회
        private void FillGrid_Proc_Monthly(string SDate, string EDate, int ChkArticleID, string ArticleID)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("nchkDate", 1);                    //int
            sqlParameter.Add("StartYYYYMM", SDate);
            sqlParameter.Add("EndYYYYMM", EDate);
            sqlParameter.Add("nchkDefectStep", 1);              //int  (자주공정 > 무조건 3)
            sqlParameter.Add("DefectStep", "3");

            sqlParameter.Add("nchkCustom", 0);                  //int (존재불가)
            sqlParameter.Add("CustomID", "");
            sqlParameter.Add("Custom", "");

            sqlParameter.Add("nchkArticleID", 0);// ChkArticleID);    //int
            sqlParameter.Add("ArticleID", ""); // ArticleID);
            sqlParameter.Add("Article", "");
            sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");


            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectTotal_Detail_Monthly", sqlParameter, false);

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
                    DataRow dr = dt.Rows[0];
                    for (int i = 0; i < 4; i++)
                    {
                        DataGridRow dgr = lib.GetRow(i, dgdDefectTotal_PROC_MonthGrid);
                        var ViewReceiver = dgr.Item as Win_Qul_DefectTotal_Q_View;

                        if (i == 0)             //첫번째 행 (작업수량)
                        {
                            ViewReceiver.PROC_PreDefectGoal = stringFormat2(dr["PreYearProdQty"]);
                            ViewReceiver.PROC_DefectGoal = "";
                            ViewReceiver.PROC_MM1 = stringFormat2(dr["ProdQty1"]);
                            ViewReceiver.PROC_MM2 = stringFormat2(dr["ProdQty2"]);
                            ViewReceiver.PROC_MM3 = stringFormat2(dr["ProdQty3"]);
                            ViewReceiver.PROC_MM4 = stringFormat2(dr["ProdQty4"]);
                            ViewReceiver.PROC_MM5 = stringFormat2(dr["ProdQty5"]);
                            ViewReceiver.PROC_MM6 = stringFormat2(dr["ProdQty6"]);
                            ViewReceiver.PROC_MM7 = stringFormat2(dr["ProdQty7"]);
                            ViewReceiver.PROC_MM8 = stringFormat2(dr["ProdQty8"]);
                            ViewReceiver.PROC_MM9 = stringFormat2(dr["ProdQty9"]);
                            ViewReceiver.PROC_MM10 = stringFormat2(dr["ProdQty10"]);
                            ViewReceiver.PROC_MM11 = stringFormat2(dr["ProdQty11"]);
                            ViewReceiver.PROC_MM12 = stringFormat2(dr["ProdQty12"]);
                            ViewReceiver.PROC_TotalDefectGoal = stringFormat2(dr["TProdQty"]);
                        }
                        else if (i == 1)            //두번째 행 (불량수량)
                        {
                            ViewReceiver.PROC_PreDefectGoal = stringFormat2(dr["PreYearDefectQty"]);
                            ViewReceiver.PROC_DefectGoal = "";
                            ViewReceiver.PROC_MM1 = stringFormat2(dr["DefectQty1"]);
                            ViewReceiver.PROC_MM2 = stringFormat2(dr["DefectQty2"]);
                            ViewReceiver.PROC_MM3 = stringFormat2(dr["DefectQty3"]);
                            ViewReceiver.PROC_MM4 = stringFormat2(dr["DefectQty4"]);
                            ViewReceiver.PROC_MM5 = stringFormat2(dr["DefectQty5"]);
                            ViewReceiver.PROC_MM6 = stringFormat2(dr["DefectQty6"]);
                            ViewReceiver.PROC_MM7 = stringFormat2(dr["DefectQty7"]);
                            ViewReceiver.PROC_MM8 = stringFormat2(dr["DefectQty8"]);
                            ViewReceiver.PROC_MM9 = stringFormat2(dr["DefectQty9"]);
                            ViewReceiver.PROC_MM10 = stringFormat2(dr["DefectQty10"]);
                            ViewReceiver.PROC_MM11 = stringFormat2(dr["DefectQty11"]);
                            ViewReceiver.PROC_MM12 = stringFormat2(dr["DefectQty12"]);
                            ViewReceiver.PROC_TotalDefectGoal = stringFormat2(dr["TDefectQty"]);
                        }
                        else if (i == 2)            //세번째 행 (불량율 ppm) + 그래프.
                        {
                            ViewReceiver.PROC_PreDefectGoal = stringFormat2(dr["PreYearDefectRate"]);
                            ViewReceiver.PROC_DefectGoal = "";
                            ViewReceiver.PROC_MM1 = stringFormat2(dr["DefectRate1"]);
                            ViewReceiver.PROC_MM2 = stringFormat2(dr["DefectRate2"]);
                            ViewReceiver.PROC_MM3 = stringFormat2(dr["DefectRate3"]);
                            ViewReceiver.PROC_MM4 = stringFormat2(dr["DefectRate4"]);
                            ViewReceiver.PROC_MM5 = stringFormat2(dr["DefectRate5"]);
                            ViewReceiver.PROC_MM6 = stringFormat2(dr["DefectRate6"]);
                            ViewReceiver.PROC_MM7 = stringFormat2(dr["DefectRate7"]);
                            ViewReceiver.PROC_MM8 = stringFormat2(dr["DefectRate8"]);
                            ViewReceiver.PROC_MM9 = stringFormat2(dr["DefectRate9"]);
                            ViewReceiver.PROC_MM10 = stringFormat2(dr["DefectRate10"]);
                            ViewReceiver.PROC_MM11 = stringFormat2(dr["DefectRate11"]);
                            ViewReceiver.PROC_MM12 = stringFormat2(dr["DefectRate12"]);
                            ViewReceiver.PROC_TotalDefectGoal = stringFormat2(dr["TDefectRate"]);

                            double pregoal = 0;
                            double thisgoal = 0;
                            double mm1 = 0;
                            double mm2 = 0;
                            double mm3 = 0;
                            double mm4 = 0;
                            double mm5 = 0;
                            double mm6 = 0;
                            double mm7 = 0;
                            double mm8 = 0;
                            double mm9 = 0;
                            double mm10 = 0;
                            double mm11 = 0;
                            double mm12 = 0;
                            double total = 0;

                            pregoal = Convert.ToDouble(ViewReceiver.PROC_PreDefectGoal);
                            mm1 = Convert.ToDouble(ViewReceiver.PROC_MM1);
                            mm2 = Convert.ToDouble(ViewReceiver.PROC_MM2);
                            mm3 = Convert.ToDouble(ViewReceiver.PROC_MM3);
                            mm4 = Convert.ToDouble(ViewReceiver.PROC_MM4);
                            mm5 = Convert.ToDouble(ViewReceiver.PROC_MM5);
                            mm6 = Convert.ToDouble(ViewReceiver.PROC_MM6);
                            mm7 = Convert.ToDouble(ViewReceiver.PROC_MM7);
                            mm8 = Convert.ToDouble(ViewReceiver.PROC_MM8);
                            mm9 = Convert.ToDouble(ViewReceiver.PROC_MM9);
                            mm10 = Convert.ToDouble(ViewReceiver.PROC_MM10);
                            mm11 = Convert.ToDouble(ViewReceiver.PROC_MM11);
                            mm12 = Convert.ToDouble(ViewReceiver.PROC_MM12);
                            total = Convert.ToDouble(ViewReceiver.PROC_TotalDefectGoal);

                            SeriesCollection SeriesCollection = new SeriesCollection
                            {
                                new ColumnSeries
                                {
                                    Title = "불량율 종합현황",
                                    Values = new ChartValues<double>
                                    {
                                        pregoal, thisgoal,
                                        mm1, mm2, mm3, mm4, mm5, mm6,
                                        mm7, mm8, mm9, mm10, mm11, mm12, total
                                    }
                                },
                                new LineSeries
                                {
                                    Title = "작년 기준선",
                                    Values = new ChartValues<double>
                                    {
                                        0,0,
                                        pregoal, pregoal, pregoal, pregoal, pregoal, pregoal,
                                        pregoal, pregoal, pregoal, pregoal, pregoal, pregoal, 0
                                    }
                                }
                            };
                            lvcPROCChart.Series = SeriesCollection;


                        }
                        else if (i == 3)            //네번째 행 (목표 ppm)
                        {
                            ViewReceiver.PROC_PreDefectGoal = "";
                            ViewReceiver.PROC_DefectGoal = stringFormat2(dr["YearGoalDefectRate"]);
                            ViewReceiver.PROC_MM1 = stringFormat2(dr["YearGoal1"]);
                            ViewReceiver.PROC_MM2 = stringFormat2(dr["YearGoal2"]);
                            ViewReceiver.PROC_MM3 = stringFormat2(dr["YearGoal3"]);
                            ViewReceiver.PROC_MM4 = stringFormat2(dr["YearGoal4"]);
                            ViewReceiver.PROC_MM5 = stringFormat2(dr["YearGoal5"]);
                            ViewReceiver.PROC_MM6 = stringFormat2(dr["YearGoal6"]);
                            ViewReceiver.PROC_MM7 = stringFormat2(dr["YearGoal7"]);
                            ViewReceiver.PROC_MM8 = stringFormat2(dr["YearGoal8"]);
                            ViewReceiver.PROC_MM9 = stringFormat2(dr["YearGoal9"]);
                            ViewReceiver.PROC_MM10 = stringFormat2(dr["YearGoal10"]);
                            ViewReceiver.PROC_MM11 = stringFormat2(dr["YearGoal11"]);
                            ViewReceiver.PROC_MM12 = stringFormat2(dr["YearGoal12"]);
                            ViewReceiver.PROC_TotalDefectGoal = stringFormat2(dr["TotalYearGoal"]);
                        }
                    }
                }

            }

            DataStore.Instance.CloseConnection();
        }

        #endregion

        #region 자주/공정 탬 유형별 불량현황 그리드 조회
        // 자주/공정 탭 유형별 불량현황 3개월치 묶음 조회
        private void FillGrid_Proc_Symptom(string sYYYY, string sMM, int ChkArticleID, string ArticleID)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("nchkDate", 1);                    //int
            sqlParameter.Add("YYYY", sYYYY);
            sqlParameter.Add("MM", sMM);
            sqlParameter.Add("nchkDefectStep", "1");            //int  (자주/공정 > 무조건 3)
            sqlParameter.Add("DefectStep", "3");

            sqlParameter.Add("nchkCustom", 0);                  //int (존재불가)
            sqlParameter.Add("CustomID", "");

            sqlParameter.Add("nchkArticleID", 0); // ChkArticleID);    //int
            sqlParameter.Add("ArticleID", ""); // ArticleID);
            //sqlParameter.Add("Article", "");
            sqlParameter.Add("sGrouping", "1");
            sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectTotal_Detail_Symptom", sqlParameter, false);

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
                    // 원형 그래프.
                    PieData pd = new PieData();

                    double TotalQtyMinus2 = 0;
                    double TotalQtyMinus1 = 0;
                    double TotalQty = 0;

                    dgdDefectTotal_PROC_SymptomGrid.Items.Clear();
                    DataRowCollection drc = dt.Rows;
                    foreach (DataRow item in drc)
                    {
                        if (item["GroupingName"].ToString() == string.Empty)
                        {
                            continue;
                        }
                        else
                        {
                            var Win_Qul_sts_DefectTotal_Q_Insert = new Win_Qul_DefectTotal_Q_View()
                            {
                                PROC_GroupingName = item["GroupingName"].ToString(),

                                PROC_Minus2qty = item["DefectQty1"].ToString(),
                                PROC_Minus2rate = item["DefectRate1"].ToString(),
                                PROC_Minus1qty = item["DefectQty2"].ToString(),
                                PROC_Minus1rate = item["DefectRate2"].ToString(),
                                PROC_MMqty = item["DefectQty3"].ToString(),
                                PROC_MMrate = item["DefectRate3"].ToString()
                            };
                            dgdDefectTotal_PROC_SymptomGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_Insert);

                            TotalQtyMinus2 = TotalQtyMinus2 + Convert.ToDouble(item["DefectQty1"].ToString());
                            TotalQtyMinus1 = TotalQtyMinus1 + Convert.ToDouble(item["DefectQty2"].ToString());
                            TotalQty = TotalQty + Convert.ToDouble(item["DefectQty3"].ToString());

                            double value = 0;
                            value = Convert.ToDouble(item["DefectRate3"].ToString());
                            pd.AddSlice(item["GroupingName"].ToString(), value);

                        }
                    }

                    // 마지막 총계 줄 업데이트
                    var Win_Qul_sts_DefectTotal_Q_lasttotal_Insert = new Win_Qul_DefectTotal_Q_View()
                    {
                        PROC_GroupingName = "누적합계",

                        PROC_Minus2qty = TotalQtyMinus2.ToString(),
                        PROC_Minus2rate = "100",
                        PROC_Minus1qty = TotalQtyMinus1.ToString(),
                        PROC_Minus1rate = "100",
                        PROC_MMqty = TotalQty.ToString(),
                        PROC_MMrate = "100"
                    };
                    dgdDefectTotal_PROC_SymptomGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_lasttotal_Insert);

                    foreach (var n in pd.Slice)
                    {
                        PROC_PieChart.Series.Add(new PieSeries
                        {
                            Title = n.Key,
                            Values = new ChartValues<double> { n.Value }
                        });
                    }
                }
            }

            DataStore.Instance.CloseConnection();
        }

        #endregion



        #region 출하 탭 월 종합 그리드 조회
        // 출하 탭 월 종합 그리드 조회
        private void FillGrid_Ship_Monthly(string SDate, string EDate, int ChkArticleID, string ArticleID)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("nchkDate", 1);                    //int
            sqlParameter.Add("StartYYYYMM", SDate);
            sqlParameter.Add("EndYYYYMM", EDate);
            sqlParameter.Add("nchkDefectStep", 1);              //int  (출하 > 무조건 5)
            sqlParameter.Add("DefectStep", "5");

            sqlParameter.Add("nchkCustom", 0);                  //int (존재불가)
            sqlParameter.Add("CustomID", "");
            sqlParameter.Add("Custom", "");

            sqlParameter.Add("nchkArticleID", 0); // ChkArticleID);    //int
            sqlParameter.Add("ArticleID", ""); // ArticleID);
            sqlParameter.Add("Article", "");
            sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectTotal_Detail_Monthly", sqlParameter, false);

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
                    DataRow dr = dt.Rows[0];
                    for (int i = 0; i < 4; i++)
                    {
                        DataGridRow dgr = lib.GetRow(i, dgdDefectTotal_SHIP_MonthGrid);
                        var ViewReceiver = dgr.Item as Win_Qul_DefectTotal_Q_View;

                        if (i == 0)             //첫번째 행 (작업수량)
                        {
                            ViewReceiver.SHIP_PreDefectGoal = stringFormat2(dr["PreYearProdQty"]);
                            ViewReceiver.SHIP_DefectGoal = "";
                            ViewReceiver.SHIP_MM1 = stringFormat2(dr["ProdQty1"]);
                            ViewReceiver.SHIP_MM2 = stringFormat2(dr["ProdQty2"]);
                            ViewReceiver.SHIP_MM3 = stringFormat2(dr["ProdQty3"]);
                            ViewReceiver.SHIP_MM4 = stringFormat2(dr["ProdQty4"]);
                            ViewReceiver.SHIP_MM5 = stringFormat2(dr["ProdQty5"]);
                            ViewReceiver.SHIP_MM6 = stringFormat2(dr["ProdQty6"]);
                            ViewReceiver.SHIP_MM7 = stringFormat2(dr["ProdQty7"]);
                            ViewReceiver.SHIP_MM8 = stringFormat2(dr["ProdQty8"]);
                            ViewReceiver.SHIP_MM9 = stringFormat2(dr["ProdQty9"]);
                            ViewReceiver.SHIP_MM10 = stringFormat2(dr["ProdQty10"]);
                            ViewReceiver.SHIP_MM11 = stringFormat2(dr["ProdQty11"]);
                            ViewReceiver.SHIP_MM12 = stringFormat2(dr["ProdQty12"]);
                            ViewReceiver.SHIP_TotalDefectGoal = stringFormat2(dr["TProdQty"]);
                        }
                        else if (i == 1)            //두번째 행 (불량수량)
                        {
                            ViewReceiver.SHIP_PreDefectGoal = stringFormat2(dr["PreYearDefectQty"]);
                            ViewReceiver.SHIP_DefectGoal = "";
                            ViewReceiver.SHIP_MM1 = stringFormat2(dr["DefectQty1"]);
                            ViewReceiver.SHIP_MM2 = stringFormat2(dr["DefectQty2"]);
                            ViewReceiver.SHIP_MM3 = stringFormat2(dr["DefectQty3"]);
                            ViewReceiver.SHIP_MM4 = stringFormat2(dr["DefectQty4"]);
                            ViewReceiver.SHIP_MM5 = stringFormat2(dr["DefectQty5"]);
                            ViewReceiver.SHIP_MM6 = stringFormat2(dr["DefectQty6"]);
                            ViewReceiver.SHIP_MM7 = stringFormat2(dr["DefectQty7"]);
                            ViewReceiver.SHIP_MM8 = stringFormat2(dr["DefectQty8"]);
                            ViewReceiver.SHIP_MM9 = stringFormat2(dr["DefectQty9"]);
                            ViewReceiver.SHIP_MM10 = stringFormat2(dr["DefectQty10"]);
                            ViewReceiver.SHIP_MM11 = stringFormat2(dr["DefectQty11"]);
                            ViewReceiver.SHIP_MM12 = stringFormat2(dr["DefectQty12"]);
                            ViewReceiver.SHIP_TotalDefectGoal = stringFormat2(dr["TDefectQty"]);
                        }
                        else if (i == 2)            //세번째 행 (불량율 ppm) + 그래프.
                        {
                            ViewReceiver.SHIP_PreDefectGoal = stringFormat2(dr["PreYearDefectRate"]);
                            ViewReceiver.SHIP_DefectGoal = "";
                            ViewReceiver.SHIP_MM1 = stringFormat2(dr["DefectRate1"]);
                            ViewReceiver.SHIP_MM2 = stringFormat2(dr["DefectRate2"]);
                            ViewReceiver.SHIP_MM3 = stringFormat2(dr["DefectRate3"]);
                            ViewReceiver.SHIP_MM4 = stringFormat2(dr["DefectRate4"]);
                            ViewReceiver.SHIP_MM5 = stringFormat2(dr["DefectRate5"]);
                            ViewReceiver.SHIP_MM6 = stringFormat2(dr["DefectRate6"]);
                            ViewReceiver.SHIP_MM7 = stringFormat2(dr["DefectRate7"]);
                            ViewReceiver.SHIP_MM8 = stringFormat2(dr["DefectRate8"]);
                            ViewReceiver.SHIP_MM9 = stringFormat2(dr["DefectRate9"]);
                            ViewReceiver.SHIP_MM10 = stringFormat2(dr["DefectRate10"]);
                            ViewReceiver.SHIP_MM11 = stringFormat2(dr["DefectRate11"]);
                            ViewReceiver.SHIP_MM12 = stringFormat2(dr["DefectRate12"]);
                            ViewReceiver.SHIP_TotalDefectGoal = stringFormat2(dr["TDefectRate"]);

                            double pregoal = 0;
                            double thisgoal = 0;
                            double mm1 = 0;
                            double mm2 = 0;
                            double mm3 = 0;
                            double mm4 = 0;
                            double mm5 = 0;
                            double mm6 = 0;
                            double mm7 = 0;
                            double mm8 = 0;
                            double mm9 = 0;
                            double mm10 = 0;
                            double mm11 = 0;
                            double mm12 = 0;
                            double total = 0;

                            pregoal = Convert.ToDouble(ViewReceiver.SHIP_PreDefectGoal);
                            mm1 = Convert.ToDouble(ViewReceiver.SHIP_MM1);
                            mm2 = Convert.ToDouble(ViewReceiver.SHIP_MM2);
                            mm3 = Convert.ToDouble(ViewReceiver.SHIP_MM3);
                            mm4 = Convert.ToDouble(ViewReceiver.SHIP_MM4);
                            mm5 = Convert.ToDouble(ViewReceiver.SHIP_MM5);
                            mm6 = Convert.ToDouble(ViewReceiver.SHIP_MM6);
                            mm7 = Convert.ToDouble(ViewReceiver.SHIP_MM7);
                            mm8 = Convert.ToDouble(ViewReceiver.SHIP_MM8);
                            mm9 = Convert.ToDouble(ViewReceiver.SHIP_MM9);
                            mm10 = Convert.ToDouble(ViewReceiver.SHIP_MM10);
                            mm11 = Convert.ToDouble(ViewReceiver.SHIP_MM11);
                            mm12 = Convert.ToDouble(ViewReceiver.SHIP_MM12);
                            total = Convert.ToDouble(ViewReceiver.SHIP_TotalDefectGoal);

                            SeriesCollection SeriesCollection = new SeriesCollection
                            {
                                new ColumnSeries
                                {
                                    Title = "불량율 종합현황",
                                    Values = new ChartValues<double>
                                    {
                                        pregoal, thisgoal,
                                        mm1, mm2, mm3, mm4, mm5, mm6,
                                        mm7, mm8, mm9, mm10, mm11, mm12, total
                                    }
                                },
                                new LineSeries
                                {
                                    Title = "작년 기준선",
                                    Values = new ChartValues<double>
                                    {
                                        0,0,
                                        pregoal, pregoal, pregoal, pregoal, pregoal, pregoal,
                                        pregoal, pregoal, pregoal, pregoal, pregoal, pregoal, 0
                                    }
                                }
                            };
                            lvcSHIPChart.Series = SeriesCollection;


                        }
                        else if (i == 3)            //네번째 행 (목표 ppm)
                        {
                            ViewReceiver.SHIP_PreDefectGoal = "";
                            ViewReceiver.SHIP_DefectGoal = stringFormat2(dr["YearGoalDefectRate"]);
                            ViewReceiver.SHIP_MM1 = stringFormat2(dr["YearGoal1"]);
                            ViewReceiver.SHIP_MM2 = stringFormat2(dr["YearGoal2"]);
                            ViewReceiver.SHIP_MM3 = stringFormat2(dr["YearGoal3"]);
                            ViewReceiver.SHIP_MM4 = stringFormat2(dr["YearGoal4"]);
                            ViewReceiver.SHIP_MM5 = stringFormat2(dr["YearGoal5"]);
                            ViewReceiver.SHIP_MM6 = stringFormat2(dr["YearGoal6"]);
                            ViewReceiver.SHIP_MM7 = stringFormat2(dr["YearGoal7"]);
                            ViewReceiver.SHIP_MM8 = stringFormat2(dr["YearGoal8"]);
                            ViewReceiver.SHIP_MM9 = stringFormat2(dr["YearGoal9"]);
                            ViewReceiver.SHIP_MM10 = stringFormat2(dr["YearGoal10"]);
                            ViewReceiver.SHIP_MM11 = stringFormat2(dr["YearGoal11"]);
                            ViewReceiver.SHIP_MM12 = stringFormat2(dr["YearGoal12"]);
                            ViewReceiver.SHIP_TotalDefectGoal = stringFormat2(dr["TotalYearGoal"]);
                        }
                    }
                }
            }

            DataStore.Instance.CloseConnection();
        }

        #endregion

        #region 출하 탭 유형별 불량현황 그리드 조회
        // 출하 탭 유형별 불량현황 3개월치 묶음 조회
        private void FillGrid_Ship_Symptom(string sYYYY, string sMM, int ChkArticleID, string ArticleID)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("nchkDate", 1);                    //int
            sqlParameter.Add("YYYY", sYYYY);
            sqlParameter.Add("MM", sMM);
            sqlParameter.Add("nchkDefectStep", "1");            //int  (출하 > 무조건 5)
            sqlParameter.Add("DefectStep", "5");

            sqlParameter.Add("nchkCustom", 0);                  //int (존재불가)
            sqlParameter.Add("CustomID", "");

            sqlParameter.Add("nchkArticleID", 0);// ChkArticleID);    //int
            sqlParameter.Add("ArticleID", ""); // ArticleID);
            //sqlParameter.Add("Article", "");
            sqlParameter.Add("sGrouping", "1");
            sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");


            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectTotal_Detail_Symptom", sqlParameter, false);

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
                    // 원형 그래프.
                    PieData pd = new PieData();

                    double TotalQtyMinus2 = 0;
                    double TotalQtyMinus1 = 0;
                    double TotalQty = 0;

                    dgdDefectTotal_SHIP_SymptomGrid.Items.Clear();
                    DataRowCollection drc = dt.Rows;
                    foreach (DataRow item in drc)
                    {
                        if (item["GroupingName"].ToString() == string.Empty)
                        {
                            continue;
                        }
                        else
                        {
                            var Win_Qul_sts_DefectTotal_Q_Insert = new Win_Qul_DefectTotal_Q_View()
                            {
                                SHIP_GroupingName = item["GroupingName"].ToString(),

                                SHIP_Minus2qty = item["DefectQty1"].ToString(),
                                SHIP_Minus2rate = stringFormat2(item["DefectRate1"]),
                                SHIP_Minus1qty = item["DefectQty2"].ToString(),
                                SHIP_Minus1rate = stringFormat2(item["DefectRate2"]),
                                SHIP_MMqty = item["DefectQty3"].ToString(),
                                SHIP_MMrate = stringFormat2(item["DefectRate3"])
                            };
                            dgdDefectTotal_SHIP_SymptomGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_Insert);

                            TotalQtyMinus2 = TotalQtyMinus2 + Convert.ToDouble(item["DefectQty1"].ToString());
                            TotalQtyMinus1 = TotalQtyMinus1 + Convert.ToDouble(item["DefectQty2"].ToString());
                            TotalQty = TotalQty + Convert.ToDouble(item["DefectQty3"].ToString());

                            double value = 0;
                            value = Convert.ToDouble(item["DefectRate3"].ToString());
                            pd.AddSlice(item["GroupingName"].ToString(), value);

                        }
                    }

                    // 마지막 총계 줄 업데이트
                    var Win_Qul_sts_DefectTotal_Q_lasttotal_Insert = new Win_Qul_DefectTotal_Q_View()
                    {
                        SHIP_GroupingName = "누적합계",

                        SHIP_Minus2qty = TotalQtyMinus2.ToString(),
                        SHIP_Minus2rate = "100",
                        SHIP_Minus1qty = TotalQtyMinus1.ToString(),
                        SHIP_Minus1rate = "100",
                        SHIP_MMqty = TotalQty.ToString(),
                        SHIP_MMrate = "100"
                    };
                    dgdDefectTotal_SHIP_SymptomGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_lasttotal_Insert);


                    foreach (var n in pd.Slice)
                    {
                        SHIP_PieChart.Series.Add(new PieSeries
                        {
                            Title = n.Key,
                            Values = new ChartValues<double> { n.Value }
                        });
                    }

                }
            }

            DataStore.Instance.CloseConnection();
        }

        #endregion



        #region 고객 탭 월 종합 그리드 조회
        // 고객 탭 월 종합 그리드 조회
        private void FillGrid_Cust_Monthly(string SDate, string EDate, int ChkArticleID, string ArticleID)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("nchkDate", 1);                    //int
            sqlParameter.Add("StartYYYYMM", SDate);
            sqlParameter.Add("EndYYYYMM", EDate);
            sqlParameter.Add("nchkDefectStep", 1);              //int  (고객 > 무조건 7)
            sqlParameter.Add("DefectStep", "7");

            sqlParameter.Add("nchkCustom", 0);                  //int (존재불가)
            sqlParameter.Add("CustomID", "");
            sqlParameter.Add("Custom", "");

            sqlParameter.Add("nchkArticleID", 0); // ChkArticleID);    //int
            sqlParameter.Add("ArticleID", ""); // ArticleID);
            sqlParameter.Add("Article", "");
            sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectTotal_Detail_Monthly", sqlParameter, false);

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
                    DataRow dr = dt.Rows[0];
                    for (int i = 0; i < 4; i++)
                    {
                        DataGridRow dgr = lib.GetRow(i, dgdDefectTotal_CUST_MonthGrid);
                        var ViewReceiver = dgr.Item as Win_Qul_DefectTotal_Q_View;

                        if (i == 0)             //첫번째 행 (작업수량)
                        {
                            ViewReceiver.CUST_PreDefectGoal = dr["PreYearProdQty"].ToString();
                            ViewReceiver.CUST_DefectGoal = "";
                            ViewReceiver.CUST_MM1 = dr["ProdQty1"].ToString();
                            ViewReceiver.CUST_MM2 = dr["ProdQty2"].ToString();
                            ViewReceiver.CUST_MM3 = dr["ProdQty3"].ToString();
                            ViewReceiver.CUST_MM4 = dr["ProdQty4"].ToString();
                            ViewReceiver.CUST_MM5 = dr["ProdQty5"].ToString();
                            ViewReceiver.CUST_MM6 = dr["ProdQty6"].ToString();
                            ViewReceiver.CUST_MM7 = dr["ProdQty7"].ToString();
                            ViewReceiver.CUST_MM8 = dr["ProdQty8"].ToString();
                            ViewReceiver.CUST_MM9 = dr["ProdQty9"].ToString();
                            ViewReceiver.CUST_MM10 = dr["ProdQty10"].ToString();
                            ViewReceiver.CUST_MM11 = dr["ProdQty11"].ToString();
                            ViewReceiver.CUST_MM12 = dr["ProdQty12"].ToString();
                            ViewReceiver.CUST_TotalDefectGoal = dr["TProdQty"].ToString();
                        }
                        else if (i == 1)            //두번째 행 (불량수량)
                        {
                            ViewReceiver.CUST_PreDefectGoal = dr["PreYearDefectQty"].ToString();
                            ViewReceiver.CUST_DefectGoal = "";
                            ViewReceiver.CUST_MM1 = dr["DefectQty1"].ToString();
                            ViewReceiver.CUST_MM2 = dr["DefectQty2"].ToString();
                            ViewReceiver.CUST_MM3 = dr["DefectQty3"].ToString();
                            ViewReceiver.CUST_MM4 = dr["DefectQty4"].ToString();
                            ViewReceiver.CUST_MM5 = dr["DefectQty5"].ToString();
                            ViewReceiver.CUST_MM6 = dr["DefectQty6"].ToString();
                            ViewReceiver.CUST_MM7 = dr["DefectQty7"].ToString();
                            ViewReceiver.CUST_MM8 = dr["DefectQty8"].ToString();
                            ViewReceiver.CUST_MM9 = dr["DefectQty9"].ToString();
                            ViewReceiver.CUST_MM10 = dr["DefectQty10"].ToString();
                            ViewReceiver.CUST_MM11 = dr["DefectQty11"].ToString();
                            ViewReceiver.CUST_MM12 = dr["DefectQty12"].ToString();
                            ViewReceiver.CUST_TotalDefectGoal = dr["TDefectQty"].ToString();
                        }
                        else if (i == 2)            //세번째 행 (불량율 ppm) + 그래프.
                        {
                            ViewReceiver.CUST_PreDefectGoal = dr["PreYearDefectRate"].ToString();
                            ViewReceiver.CUST_DefectGoal = "";
                            ViewReceiver.CUST_MM1 = dr["DefectRate1"].ToString();
                            ViewReceiver.CUST_MM2 = dr["DefectRate2"].ToString();
                            ViewReceiver.CUST_MM3 = dr["DefectRate3"].ToString();
                            ViewReceiver.CUST_MM4 = dr["DefectRate4"].ToString();
                            ViewReceiver.CUST_MM5 = dr["DefectRate5"].ToString();
                            ViewReceiver.CUST_MM6 = dr["DefectRate6"].ToString();
                            ViewReceiver.CUST_MM7 = dr["DefectRate7"].ToString();
                            ViewReceiver.CUST_MM8 = dr["DefectRate8"].ToString();
                            ViewReceiver.CUST_MM9 = dr["DefectRate9"].ToString();
                            ViewReceiver.CUST_MM10 = dr["DefectRate10"].ToString();
                            ViewReceiver.CUST_MM11 = dr["DefectRate11"].ToString();
                            ViewReceiver.CUST_MM12 = dr["DefectRate12"].ToString();
                            ViewReceiver.CUST_TotalDefectGoal = dr["TDefectRate"].ToString();

                            double pregoal = 0;
                            double thisgoal = 0;
                            double mm1 = 0;
                            double mm2 = 0;
                            double mm3 = 0;
                            double mm4 = 0;
                            double mm5 = 0;
                            double mm6 = 0;
                            double mm7 = 0;
                            double mm8 = 0;
                            double mm9 = 0;
                            double mm10 = 0;
                            double mm11 = 0;
                            double mm12 = 0;
                            double total = 0;

                            pregoal = Convert.ToDouble(ViewReceiver.CUST_PreDefectGoal);
                            mm1 = Convert.ToDouble(ViewReceiver.CUST_MM1);
                            mm2 = Convert.ToDouble(ViewReceiver.CUST_MM2);
                            mm3 = Convert.ToDouble(ViewReceiver.CUST_MM3);
                            mm4 = Convert.ToDouble(ViewReceiver.CUST_MM4);
                            mm5 = Convert.ToDouble(ViewReceiver.CUST_MM5);
                            mm6 = Convert.ToDouble(ViewReceiver.CUST_MM6);
                            mm7 = Convert.ToDouble(ViewReceiver.CUST_MM7);
                            mm8 = Convert.ToDouble(ViewReceiver.CUST_MM8);
                            mm9 = Convert.ToDouble(ViewReceiver.CUST_MM9);
                            mm10 = Convert.ToDouble(ViewReceiver.CUST_MM10);
                            mm11 = Convert.ToDouble(ViewReceiver.CUST_MM11);
                            mm12 = Convert.ToDouble(ViewReceiver.CUST_MM12);
                            total = Convert.ToDouble(ViewReceiver.CUST_TotalDefectGoal);

                            SeriesCollection SeriesCollection = new SeriesCollection
                            {
                                new ColumnSeries
                                {
                                    Title = "불량율 종합현황",
                                    Values = new ChartValues<double>
                                    {
                                        pregoal, thisgoal,
                                        mm1, mm2, mm3, mm4, mm5, mm6,
                                        mm7, mm8, mm9, mm10, mm11, mm12, total
                                    }
                                },
                                new LineSeries
                                {
                                    Title = "작년 기준선",
                                    Values = new ChartValues<double>
                                    {
                                        0,0,
                                        pregoal, pregoal, pregoal, pregoal, pregoal, pregoal,
                                        pregoal, pregoal, pregoal, pregoal, pregoal, pregoal, 0
                                    }
                                }
                            };
                            lvcCUSTChart.Series = SeriesCollection;


                        }
                        else if (i == 3)            //네번째 행 (목표 ppm)
                        {
                            ViewReceiver.CUST_PreDefectGoal = "";
                            ViewReceiver.CUST_DefectGoal = dr["YearGoalDefectRate"].ToString();
                            ViewReceiver.CUST_MM1 = dr["YearGoal1"].ToString();
                            ViewReceiver.CUST_MM2 = dr["YearGoal2"].ToString();
                            ViewReceiver.CUST_MM3 = dr["YearGoal3"].ToString();
                            ViewReceiver.CUST_MM4 = dr["YearGoal4"].ToString();
                            ViewReceiver.CUST_MM5 = dr["YearGoal5"].ToString();
                            ViewReceiver.CUST_MM6 = dr["YearGoal6"].ToString();
                            ViewReceiver.CUST_MM7 = dr["YearGoal7"].ToString();
                            ViewReceiver.CUST_MM8 = dr["YearGoal8"].ToString();
                            ViewReceiver.CUST_MM9 = dr["YearGoal9"].ToString();
                            ViewReceiver.CUST_MM10 = dr["YearGoal10"].ToString();
                            ViewReceiver.CUST_MM11 = dr["YearGoal11"].ToString();
                            ViewReceiver.CUST_MM12 = dr["YearGoal12"].ToString();
                            ViewReceiver.CUST_TotalDefectGoal = dr["TotalYearGoal"].ToString();
                        }
                    }
                }
            }

            DataStore.Instance.CloseConnection();
        }

        #endregion

        #region 고객 탭 유형별 불량현황 그리드 조회
        // 고객 탭 유형별 불량현황 3개월치 묶음 조회
        private void FillGrid_Cust_Symptom(string sYYYY, string sMM, int ChkArticleID, string ArticleID)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("nchkDate", 1);                    //int
            sqlParameter.Add("YYYY", sYYYY);
            sqlParameter.Add("MM", sMM);
            sqlParameter.Add("nchkDefectStep", "1");            //int  (고객 > 무조건 7)
            sqlParameter.Add("DefectStep", "7");

            sqlParameter.Add("nchkCustom", 0);                  //int (존재불가)
            sqlParameter.Add("CustomID", "");

            sqlParameter.Add("nchkArticleID", 0); // ChkArticleID);    //int
            sqlParameter.Add("ArticleID", ""); // ArticleID);
            //sqlParameter.Add("Article", "");
            sqlParameter.Add("sGrouping", "1");
            sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectTotal_Detail_Symptom", sqlParameter, false);

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
                    // 원형 그래프.
                    PieData pd = new PieData();

                    double TotalQtyMinus2 = 0;
                    double TotalQtyMinus1 = 0;
                    double TotalQty = 0;

                    dgdDefectTotal_CUST_SymptomGrid.Items.Clear();
                    DataRowCollection drc = dt.Rows;
                    foreach (DataRow item in drc)
                    {
                        if (item["GroupingName"].ToString() == string.Empty)
                        {
                            continue;
                        }
                        else
                        {
                            var Win_Qul_sts_DefectTotal_Q_Insert = new Win_Qul_DefectTotal_Q_View()
                            {
                                CUST_GroupingName = item["GroupingName"].ToString(),

                                CUST_Minus2qty = item["DefectQty1"].ToString(),
                                CUST_Minus2rate = item["DefectRate1"].ToString(),
                                CUST_Minus1qty = item["DefectQty2"].ToString(),
                                CUST_Minus1rate = item["DefectRate2"].ToString(),
                                CUST_MMqty = item["DefectQty3"].ToString(),
                                CUST_MMrate = item["DefectRate3"].ToString()
                            };
                            dgdDefectTotal_CUST_SymptomGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_Insert);

                            TotalQtyMinus2 = TotalQtyMinus2 + Convert.ToDouble(item["DefectQty1"].ToString());
                            TotalQtyMinus1 = TotalQtyMinus1 + Convert.ToDouble(item["DefectQty2"].ToString());
                            TotalQty = TotalQty + Convert.ToDouble(item["DefectQty3"].ToString());

                            double value = 0;
                            value = Convert.ToDouble(item["DefectRate3"].ToString());
                            pd.AddSlice(item["GroupingName"].ToString(), value);

                        }
                    }

                    // 마지막 총계 줄 업데이트
                    var Win_Qul_sts_DefectTotal_Q_lasttotal_Insert = new Win_Qul_DefectTotal_Q_View()
                    {
                        CUST_GroupingName = "누적합계",

                        CUST_Minus2qty = TotalQtyMinus2.ToString(),
                        CUST_Minus2rate = "100",
                        CUST_Minus1qty = TotalQtyMinus1.ToString(),
                        CUST_Minus1rate = "100",
                        CUST_MMqty = TotalQty.ToString(),
                        CUST_MMrate = "100"
                    };
                    dgdDefectTotal_CUST_SymptomGrid.Items.Add(Win_Qul_sts_DefectTotal_Q_lasttotal_Insert);

                    foreach (var n in pd.Slice)
                    {
                        CUST_PieChart.Series.Add(new PieSeries
                        {
                            Title = n.Key,
                            Values = new ChartValues<double> { n.Value }
                        });
                    }

                }
            }

            DataStore.Instance.CloseConnection();
        }

        #endregion



        #region 각 상세 탭 월 콤보박스 변화시키기

        // 날짜 변화에 맞추어 월 콤보박스 변화시키기 <각 상세 탭 기준.>
        private void dtpToDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            cbo_Month_Changed();
        }


        // 각 디테일 탭 하단 > 월 콤보박스 변화시키기.
        private void cbo_Month_Changed()
        {
            int yyyy = Convert.ToInt32(dtpToDate.Text.Substring(0, 4)); //선택년도
            int zzzz = Convert.ToInt32(DateTime.Now.ToString("yyyy-MM").Substring(0, 4));  //현재 년도
            int thismm = Convert.ToInt32(DateTime.Now.ToString("yyyy-MM").Substring(5, 2));

            cboINSU_Month.ItemsSource = null;
            cboPROC_Month.ItemsSource = null;
            cboSHIP_Month.ItemsSource = null;
            cboCUST_Month.ItemsSource = null;

            DataTable dt = new DataTable();
            dt.Columns.Add("value");
            dt.Columns.Add("display");

            DataRow row0 = dt.NewRow();
            row0["value"] = "1";
            row0["display"] = "1";
            DataRow row1 = dt.NewRow();
            row1["value"] = "2";
            row1["display"] = "2";
            DataRow row2 = dt.NewRow();
            row2["value"] = "3";
            row2["display"] = "3";
            DataRow row3 = dt.NewRow();
            row3["value"] = "4";
            row3["display"] = "4";
            DataRow row4 = dt.NewRow();
            row4["value"] = "5";
            row4["display"] = "5";
            DataRow row5 = dt.NewRow();
            row5["value"] = "6";
            row5["display"] = "6";
            DataRow row6 = dt.NewRow();
            row6["value"] = "7";
            row6["display"] = "7";
            DataRow row7 = dt.NewRow();
            row7["value"] = "8";
            row7["display"] = "8";
            DataRow row8 = dt.NewRow();
            row8["value"] = "9";
            row8["display"] = "9";
            DataRow row9 = dt.NewRow();
            row9["value"] = "10";
            row9["display"] = "10";
            DataRow row10 = dt.NewRow();
            row10["value"] = "11";
            row10["display"] = "11";
            DataRow row11 = dt.NewRow();
            row11["value"] = "12";
            row11["display"] = "12";

            if (yyyy < zzzz)    // 현재년도가 아닌 과거년도를 선택한 경우,
            {
                dt.Rows.Add(row0);
                dt.Rows.Add(row1);
                dt.Rows.Add(row2);
                dt.Rows.Add(row3);
                dt.Rows.Add(row4);
                dt.Rows.Add(row5);
                dt.Rows.Add(row6);
                dt.Rows.Add(row7);
                dt.Rows.Add(row8);
                dt.Rows.Add(row9);
                dt.Rows.Add(row10);
                dt.Rows.Add(row11);

                this.cboINSU_Month.ItemsSource = dt.DefaultView;
                this.cboINSU_Month.DisplayMemberPath = "display";
                this.cboINSU_Month.SelectedValuePath = "value";
                this.cboINSU_Month.SelectedIndex = 11;

                this.cboPROC_Month.ItemsSource = dt.DefaultView;
                this.cboPROC_Month.DisplayMemberPath = "display";
                this.cboPROC_Month.SelectedValuePath = "value";
                this.cboPROC_Month.SelectedIndex = 11;

                this.cboSHIP_Month.ItemsSource = dt.DefaultView;
                this.cboSHIP_Month.DisplayMemberPath = "display";
                this.cboSHIP_Month.SelectedValuePath = "value";
                this.cboSHIP_Month.SelectedIndex = 11;

                this.cboCUST_Month.ItemsSource = dt.DefaultView;
                this.cboCUST_Month.DisplayMemberPath = "display";
                this.cboCUST_Month.SelectedValuePath = "value";
                this.cboCUST_Month.SelectedIndex = 11;

            }
            else        //같은 년도다.
            {
                dt.Rows.Add(row0);
                if (thismm >= 2) { dt.Rows.Add(row1); }
                if (thismm >= 3) { dt.Rows.Add(row2); }
                if (thismm >= 4) { dt.Rows.Add(row3); }
                if (thismm >= 5) { dt.Rows.Add(row4); }
                if (thismm >= 6) { dt.Rows.Add(row5); }
                if (thismm >= 7) { dt.Rows.Add(row6); }
                if (thismm >= 8) { dt.Rows.Add(row7); }
                if (thismm >= 9) { dt.Rows.Add(row8); }
                if (thismm >= 10) { dt.Rows.Add(row9); }
                if (thismm >= 11) { dt.Rows.Add(row10); }
                if (thismm >= 12) { dt.Rows.Add(row11); }


                this.cboINSU_Month.ItemsSource = dt.DefaultView;
                this.cboINSU_Month.DisplayMemberPath = "display";
                this.cboINSU_Month.SelectedValuePath = "value";
                this.cboINSU_Month.SelectedIndex = thismm - 1;

                this.cboPROC_Month.ItemsSource = dt.DefaultView;
                this.cboPROC_Month.DisplayMemberPath = "display";
                this.cboPROC_Month.SelectedValuePath = "value";
                this.cboPROC_Month.SelectedIndex = thismm - 1;

                this.cboSHIP_Month.ItemsSource = dt.DefaultView;
                this.cboSHIP_Month.DisplayMemberPath = "display";
                this.cboSHIP_Month.SelectedValuePath = "value";
                this.cboSHIP_Month.SelectedIndex = thismm - 1;

                this.cboCUST_Month.ItemsSource = dt.DefaultView;
                this.cboCUST_Month.DisplayMemberPath = "display";
                this.cboCUST_Month.SelectedValuePath = "value";
                this.cboCUST_Month.SelectedIndex = thismm - 1;

            }
        }

        #endregion

        #region 인수탭 월 콤보박스 셀렉션 체인지 이벤트
        // 인수 _ 콤보박스 셀렉션 체인지 
        private void cboINSU_Month_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //1. 데이터 그리드 컬럼명칭 변경시키기.
            if (cboINSU_Month.SelectedValue != null)
            {
                string MM = cboINSU_Month.SelectedValue.ToString();

                int imm = Convert.ToInt32(MM);

                int minus1 = imm - 1;
                if (minus1 == 0) { minus1 = 12; }

                int minus2 = minus1 - 1;
                if (minus2 == 0) { minus2 = 12; }


                ColumnHeaderChange chc = new ColumnHeaderChange()
                {
                    column1 = (minus2.ToString()) + "월 수량",
                    column2 = (minus2.ToString()) + "월 불량율",
                    column3 = (minus1.ToString()) + "월 수량",
                    column4 = (minus1.ToString()) + "월 불량율",
                    column5 = (imm.ToString()) + "월 수량",
                    column6 = (imm.ToString()) + "월 불량율"

                };
                this.DataContext = chc;
            }
        }

        #endregion

        #region 자주/공정탭 월 콤보박스 셀렉션 체인지 이벤트
        // 자주/공정 _ 콤보박스 셀렉션 체인지 
        private void cboPROC_Month_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //1. 데이터 그리드 컬럼명칭 변경시키기.
            if (cboPROC_Month.SelectedValue != null)
            {
                string MM = cboPROC_Month.SelectedValue.ToString();

                int imm = Convert.ToInt32(MM);

                int minus1 = imm - 1;
                if (minus1 == 0) { minus1 = 12; }

                int minus2 = minus1 - 1;
                if (minus2 == 0) { minus2 = 12; }


                ColumnHeaderChange chc = new ColumnHeaderChange()
                {
                    column11 = (minus2.ToString()) + "월 수량",
                    column12 = (minus2.ToString()) + "월 불량율",
                    column13 = (minus1.ToString()) + "월 수량",
                    column14 = (minus1.ToString()) + "월 불량율",
                    column15 = (imm.ToString()) + "월 수량",
                    column16 = (imm.ToString()) + "월 불량율"

                };
                this.DataContext = chc;
            }
        }

        #endregion

        #region 출하탭 월 콤보박스 셀렉션 체인지 이벤트
        // 출하 _ 콤보박스 셀렉션 체인지 
        private void cboSHIP_Month_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //1. 데이터 그리드 컬럼명칭 변경시키기.
            if (cboSHIP_Month.SelectedValue != null)
            {
                string MM = cboSHIP_Month.SelectedValue.ToString();

                int imm = Convert.ToInt32(MM);

                int minus1 = imm - 1;
                if (minus1 == 0) { minus1 = 12; }

                int minus2 = minus1 - 1;
                if (minus2 == 0) { minus2 = 12; }


                ColumnHeaderChange chc = new ColumnHeaderChange()
                {
                    column21 = (minus2.ToString()) + "월 수량",
                    column22 = (minus2.ToString()) + "월 불량율",
                    column23 = (minus1.ToString()) + "월 수량",
                    column24 = (minus1.ToString()) + "월 불량율",
                    column25 = (imm.ToString()) + "월 수량",
                    column26 = (imm.ToString()) + "월 불량율"

                };
                this.DataContext = chc;
            }
        }

        #endregion

        #region 고객탭 월 콤보박스 셀렉션 체인지 이벤트
        // 고객 _ 콤보박스 셀렉션 체인지 
        private void cboCUST_Month_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //1. 데이터 그리드 컬럼명칭 변경시키기.
            if (cboCUST_Month.SelectedValue != null)
            {
                string MM = cboCUST_Month.SelectedValue.ToString();

                int imm = Convert.ToInt32(MM);

                int minus1 = imm - 1;
                if (minus1 == 0) { minus1 = 12; }

                int minus2 = minus1 - 1;
                if (minus2 == 0) { minus2 = 12; }


                ColumnHeaderChange chc = new ColumnHeaderChange()
                {
                    column31 = (minus2.ToString()) + "월 수량",
                    column32 = (minus2.ToString()) + "월 불량율",
                    column33 = (minus1.ToString()) + "월 수량",
                    column34 = (minus1.ToString()) + "월 불량율",
                    column35 = (imm.ToString()) + "월 수량",
                    column36 = (imm.ToString()) + "월 불량율"

                };
                this.DataContext = chc;
            }
        }





        #endregion



        // 닫기 버튼 클릭.
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
                        break;
                    }
                }
                i++;
            }
        }


        #region 엑셀 클릭 시. // 탭별로 분리.
        //엑셀. (탭별 분리.)
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            string sNowTI = (tabconGrid.SelectedItem as TabItem).Header as string;

            int IsExcelOK = 0;
            int LittlePoint = 0;

            Lib lib2 = new Lib();
            DataTable dt = null;
            switch (sNowTI)
            {
                case "전체":
                    if (dgdDefectTotal_TotalGrid.Items.Count < 1)
                    {
                        MessageBox.Show("전체 탭에 대한 검색자료가 없습니다.");
                        return;
                    }

                    string[] lst = new string[2];
                    lst[0] = "전체 종합그리드";
                    lst[1] = dgdDefectTotal_TotalGrid.Name;

                    ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
                    ExpExc.ShowDialog();

                    if (ExpExc.DialogResult.HasValue)
                    {
                        if (ExpExc.choice.Equals(dgdDefectTotal_TotalGrid.Name))
                        {
                            //MessageBox.Show("전체 종합");
                            if (ExpExc.Check.Equals("Y"))
                                dt = lib.DataGridToDTinHidden(dgdDefectTotal_TotalGrid);
                            else
                                dt = lib.DataGirdToDataTable(dgdDefectTotal_TotalGrid);

                            Name = dgdDefectTotal_TotalGrid.Name;
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
                    lib = null;
                    break;


                case "인수":
                    if ((dgdDefectTotal_INSU_MonthGrid.Items.Count < 1)
                        && (dgdDefectTotal_INSU_SymptomGrid.Items.Count < 1))
                    {
                        MessageBox.Show("인수 탭에 대한 검색자료가 없습니다.");
                        return;
                    }
                    if (dgdDefectTotal_INSU_MonthGrid.Items.Count > 0)
                    {
                        IsExcelOK = IsExcelOK + 2;
                    }
                    if (dgdDefectTotal_INSU_SymptomGrid.Items.Count > 0)
                    {
                        IsExcelOK = IsExcelOK + 2;
                    }

                    string[] lst2 = new string[IsExcelOK];       // 리스트 갯수 결정하고 난 후에,

                    if (dgdDefectTotal_INSU_MonthGrid.Items.Count > 0)
                    {
                        lst2[0] = "인수 월별 그리드";
                        lst2[IsExcelOK / 2] = dgdDefectTotal_INSU_MonthGrid.Name;
                        LittlePoint++;
                    }
                    if (dgdDefectTotal_INSU_SymptomGrid.Items.Count > 0)
                    {
                        lst2[0 + LittlePoint] = "인수 유형별 불량그리드";
                        lst2[IsExcelOK / 2 + LittlePoint] = dgdDefectTotal_INSU_SymptomGrid.Name;
                        LittlePoint++;
                    }

                    ExportExcelxaml ExpExc2 = new ExportExcelxaml(lst2);
                    ExpExc2.ShowDialog();

                    if (ExpExc2.DialogResult.HasValue)
                    {
                        if (ExpExc2.choice.Equals(dgdDefectTotal_INSU_MonthGrid.Name))
                        {
                            //MessageBox.Show("인수 월별 그리드");
                            if (ExpExc2.Check.Equals("Y"))
                                dt = lib2.DataGridToDTinHidden(dgdDefectTotal_INSU_MonthGrid);
                            else
                                dt = lib2.DataGirdToDataTable(dgdDefectTotal_INSU_MonthGrid);

                            Name = dgdDefectTotal_INSU_MonthGrid.Name;
                            if (lib2.GenerateExcel(dt, Name))
                            {
                                lib2.excel.Visible = true;
                                lib2.ReleaseExcelObject(lib2.excel);
                            }
                        }
                        else if (ExpExc2.choice.Equals(dgdDefectTotal_INSU_SymptomGrid.Name))
                        {
                            //MessageBox.Show("인수 유형별 그리드");
                            if (ExpExc2.Check.Equals("Y"))
                                dt = lib2.DataGridToDTinHidden(dgdDefectTotal_INSU_SymptomGrid);
                            else
                                dt = lib2.DataGirdToDataTable(dgdDefectTotal_INSU_SymptomGrid);
                            Name = dgdDefectTotal_INSU_SymptomGrid.Name;
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
                    break;

                case "자주/공정":
                    if ((dgdDefectTotal_PROC_MonthGrid.Items.Count < 1)
                        && (dgdDefectTotal_PROC_SymptomGrid.Items.Count < 1))
                    {
                        MessageBox.Show("자주/공정 탭에 대한 검색자료가 없습니다.");
                        return;
                    }
                    if (dgdDefectTotal_PROC_MonthGrid.Items.Count > 0)
                    {
                        IsExcelOK = IsExcelOK + 2;
                    }
                    if (dgdDefectTotal_PROC_SymptomGrid.Items.Count > 0)
                    {
                        IsExcelOK = IsExcelOK + 2;
                    }

                    string[] lst3 = new string[IsExcelOK];       // 리스트 갯수 결정하고 난 후에,

                    if (dgdDefectTotal_PROC_MonthGrid.Items.Count > 0)
                    {
                        lst3[0] = "자주/공정 월별 그리드";
                        lst3[IsExcelOK / 2] = dgdDefectTotal_PROC_MonthGrid.Name;
                        LittlePoint++;
                    }
                    if (dgdDefectTotal_PROC_SymptomGrid.Items.Count > 0)
                    {
                        lst3[0 + LittlePoint] = "자주/공정 유형별 불량그리드";
                        lst3[IsExcelOK / 2 + LittlePoint] = dgdDefectTotal_PROC_SymptomGrid.Name;
                        LittlePoint++;
                    }

                    ExportExcelxaml ExpExc3 = new ExportExcelxaml(lst3);
                    ExpExc3.ShowDialog();

                    if (ExpExc3.DialogResult.HasValue)
                    {
                        if (ExpExc3.choice.Equals(dgdDefectTotal_PROC_MonthGrid.Name))
                        {
                            //MessageBox.Show("자주/공정 월별 그리드");
                            if (ExpExc3.Check.Equals("Y"))
                                dt = lib2.DataGridToDTinHidden(dgdDefectTotal_PROC_MonthGrid);
                            else
                                dt = lib2.DataGirdToDataTable(dgdDefectTotal_PROC_MonthGrid);

                            Name = dgdDefectTotal_PROC_MonthGrid.Name;
                            if (lib2.GenerateExcel(dt, Name))
                            {
                                lib2.excel.Visible = true;
                                lib2.ReleaseExcelObject(lib2.excel);
                            }
                        }
                        else if (ExpExc3.choice.Equals(dgdDefectTotal_PROC_SymptomGrid.Name))
                        {
                            //MessageBox.Show("자주/공정 유형별 그리드");
                            if (ExpExc3.Check.Equals("Y"))
                                dt = lib2.DataGridToDTinHidden(dgdDefectTotal_PROC_SymptomGrid);
                            else
                                dt = lib2.DataGirdToDataTable(dgdDefectTotal_PROC_SymptomGrid);
                            Name = dgdDefectTotal_PROC_SymptomGrid.Name;
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
                    break;

                case "출하":
                    if ((dgdDefectTotal_SHIP_MonthGrid.Items.Count < 1)
                        && (dgdDefectTotal_SHIP_SymptomGrid.Items.Count < 1))
                    {
                        MessageBox.Show("출하 탭에 대한 검색자료가 없습니다.");
                        return;
                    }
                    if (dgdDefectTotal_SHIP_MonthGrid.Items.Count > 0)
                    {
                        IsExcelOK = IsExcelOK + 2;
                    }
                    if (dgdDefectTotal_SHIP_SymptomGrid.Items.Count > 0)
                    {
                        IsExcelOK = IsExcelOK + 2;
                    }

                    string[] lst4 = new string[IsExcelOK];       // 리스트 갯수 결정하고 난 후에,

                    if (dgdDefectTotal_SHIP_MonthGrid.Items.Count > 0)
                    {
                        lst4[0] = "출하 월별 그리드";
                        lst4[IsExcelOK / 2] = dgdDefectTotal_SHIP_MonthGrid.Name;
                        LittlePoint++;
                    }
                    if (dgdDefectTotal_SHIP_SymptomGrid.Items.Count > 0)
                    {
                        lst4[0 + LittlePoint] = "출하 유형별 불량그리드";
                        lst4[IsExcelOK / 2 + LittlePoint] = dgdDefectTotal_SHIP_SymptomGrid.Name;
                        LittlePoint++;
                    }

                    ExportExcelxaml ExpExc4 = new ExportExcelxaml(lst4);
                    ExpExc4.ShowDialog();

                    if (ExpExc4.DialogResult.HasValue)
                    {
                        if (ExpExc4.choice.Equals(dgdDefectTotal_SHIP_MonthGrid.Name))
                        {
                            //MessageBox.Show("출하 월별 그리드");
                            if (ExpExc4.Check.Equals("Y"))
                                dt = lib2.DataGridToDTinHidden(dgdDefectTotal_SHIP_MonthGrid);
                            else
                                dt = lib2.DataGirdToDataTable(dgdDefectTotal_SHIP_MonthGrid);

                            Name = dgdDefectTotal_SHIP_MonthGrid.Name;
                            if (lib2.GenerateExcel(dt, Name))
                            {
                                lib2.excel.Visible = true;
                                lib2.ReleaseExcelObject(lib2.excel);
                            }
                        }
                        else if (ExpExc4.choice.Equals(dgdDefectTotal_SHIP_SymptomGrid.Name))
                        {
                            //MessageBox.Show("출하 유형별 그리드");
                            if (ExpExc4.Check.Equals("Y"))
                                dt = lib2.DataGridToDTinHidden(dgdDefectTotal_SHIP_SymptomGrid);
                            else
                                dt = lib2.DataGirdToDataTable(dgdDefectTotal_SHIP_SymptomGrid);
                            Name = dgdDefectTotal_SHIP_SymptomGrid.Name;
                            if (lib.GenerateExcel(dt, Name))
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
                    break;

                case "고객":
                    if ((dgdDefectTotal_CUST_MonthGrid.Items.Count < 1)
                        && (dgdDefectTotal_CUST_SymptomGrid.Items.Count < 1))
                    {
                        MessageBox.Show("고객 탭에 대한 검색자료가 없습니다.");
                        return;
                    }
                    if (dgdDefectTotal_CUST_MonthGrid.Items.Count > 0)
                    {
                        IsExcelOK = IsExcelOK + 2;
                    }
                    if (dgdDefectTotal_CUST_SymptomGrid.Items.Count > 0)
                    {
                        IsExcelOK = IsExcelOK + 2;
                    }

                    string[] lst5 = new string[IsExcelOK];       // 리스트 갯수 결정하고 난 후에,

                    if (dgdDefectTotal_CUST_MonthGrid.Items.Count > 0)
                    {
                        lst5[0] = "고객 월별 그리드";
                        lst5[IsExcelOK / 2] = dgdDefectTotal_CUST_MonthGrid.Name;
                        LittlePoint++;
                    }
                    if (dgdDefectTotal_CUST_SymptomGrid.Items.Count > 0)
                    {
                        lst5[0 + LittlePoint] = "고객 유형별 불량그리드";
                        lst5[IsExcelOK / 2 + LittlePoint] = dgdDefectTotal_CUST_SymptomGrid.Name;
                        LittlePoint++;
                    }

                    ExportExcelxaml ExpExc5 = new ExportExcelxaml(lst5);
                    ExpExc5.ShowDialog();

                    if (ExpExc5.DialogResult.HasValue)
                    {
                        if (ExpExc5.choice.Equals(dgdDefectTotal_CUST_MonthGrid.Name))
                        {
                            //MessageBox.Show("고객 월별 그리드");
                            if (ExpExc5.Check.Equals("Y"))
                                dt = lib2.DataGridToDTinHidden(dgdDefectTotal_CUST_MonthGrid);
                            else
                                dt = lib2.DataGirdToDataTable(dgdDefectTotal_CUST_MonthGrid);

                            Name = dgdDefectTotal_CUST_MonthGrid.Name;
                            if (lib2.GenerateExcel(dt, Name))
                            {
                                lib2.excel.Visible = true;
                                lib2.ReleaseExcelObject(lib2.excel);
                            }
                        }
                        else if (ExpExc5.choice.Equals(dgdDefectTotal_CUST_SymptomGrid.Name))
                        {
                            //MessageBox.Show("고객 유형별 그리드");
                            if (ExpExc5.Check.Equals("Y"))
                                dt = lib2.DataGridToDTinHidden(dgdDefectTotal_CUST_SymptomGrid);
                            else
                                dt = lib2.DataGirdToDataTable(dgdDefectTotal_CUST_SymptomGrid);
                            Name = dgdDefectTotal_CUST_SymptomGrid.Name;
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
                    break;


                default: return;
            }
        }



        #endregion


        private string stringFormat2(object sender)
        {
            return String.Format("{0:N2}", sender);
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


    public class ColumnHeaderChange
    {
        //그리드 월 기간 변경용 _인수
        public string column1 { get; set; }
        public string column2 { get; set; }
        public string column3 { get; set; }
        public string column4 { get; set; }
        public string column5 { get; set; }
        public string column6 { get; set; }

        //그리드 월 기간 변경용 _자주공정
        public string column11 { get; set; }
        public string column12 { get; set; }
        public string column13 { get; set; }
        public string column14 { get; set; }
        public string column15 { get; set; }
        public string column16 { get; set; }

        //그리드 월 기간 변경용 _출하
        public string column21 { get; set; }
        public string column22 { get; set; }
        public string column23 { get; set; }
        public string column24 { get; set; }
        public string column25 { get; set; }
        public string column26 { get; set; }

        //그리드 월 기간 변경용 _고객
        public string column31 { get; set; }
        public string column32 { get; set; }
        public string column33 { get; set; }
        public string column34 { get; set; }
        public string column35 { get; set; }
        public string column36 { get; set; }

    }

}
