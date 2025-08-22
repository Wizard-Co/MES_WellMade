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
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using WizMes_WellMade.PopUP;
using WPF.MDI;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_XBarR_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_XBarR_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        string RASpecMax = string.Empty;
        string RASpecMin = string.Empty;

        // X-BAR 차트용 전역변수.
        double X_chart_UCL = 0;
        double X_chart_CL = 0;
        double X_chart_LCL = 0;

        List<double> x_chart_val = new List<double>();

        // R 차트용 전역변수.
        double R_chart_UCL = 0;
        double R_chart_CL = 0;

        List<double> r_chart_val = new List<double>();

        //Image 변수 선언
        System.Windows.Controls.Image ImageData = new System.Windows.Controls.Image();

        public Win_Qul_XBarR_Q()
        {
            InitializeComponent();
        }

        private void Win_Qul_sts_XBarR_Q_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            First_Step();
        }

        #region 첫 단계 / 날짜버튼 세팅 / 조회용 체크박스 세팅 
        // 첫 단계
        private void First_Step()
        {
            chkMonthDate.IsChecked = true;
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM");

            dtpToDate.Visibility = Visibility.Hidden;           //날짜는 하나만 있으면 됨.

            tbnJaju.IsChecked = true;

            dtpFromDate.IsEnabled = true;
            dtpToDate.IsEnabled = true;

            txtCustomer.IsEnabled = false;
            btnCustomer.IsEnabled = false;
            txtArticle.IsEnabled = false;
            btnArticle.IsEnabled = false;
        }

        /// <summary>
        /// 수입
        /// </summary>
        private void TbnInCome_Click(object sender, RoutedEventArgs e)
        {
            ToggleButton toggleButton = sender as ToggleButton;
            if (toggleButton.IsChecked == true)
            {
                tbnOutCome.IsChecked = false;
                tbnProcessCycle.IsChecked = false;
                tbnJaju.IsChecked = false;
            }
            else
            {
                toggleButton.IsChecked = true;
            }
        }

        /// <summary>
        /// 출하
        /// </summary>
        private void TbnOutCome_Click(object sender, RoutedEventArgs e)
        {
            ToggleButton toggleButton = sender as ToggleButton;
            if (toggleButton.IsChecked == true)
            {
                tbnInCome.IsChecked = false;
                tbnProcessCycle.IsChecked = false;
                tbnJaju.IsChecked = false;
            }
            else
            {
                toggleButton.IsChecked = true;
            }
        }

        /// <summary>
        /// 공정순회
        /// </summary>
        private void TbnProcessCycle_Click(object sender, RoutedEventArgs e)
        {
            ToggleButton toggleButton = sender as ToggleButton;
            if (toggleButton.IsChecked == true)
            {
                tbnInCome.IsChecked = false;
                tbnOutCome.IsChecked = false;
                tbnJaju.IsChecked = false;
            }
            else
            {
                toggleButton.IsChecked = true;
            }
        }

        /// <summary>
        /// 자주
        /// </summary>
        private void TbnJaju_Click(object sender, RoutedEventArgs e)
        {
            ToggleButton toggleButton = sender as ToggleButton;
            if (toggleButton.IsChecked == true)
            {
                tbnInCome.IsChecked = false;
                tbnOutCome.IsChecked = false;
                tbnProcessCycle.IsChecked = false;
            }
            else
            {
                toggleButton.IsChecked = true;
            }
        }

        // 이번 달
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            string[] receiver = Lib.Instance.BringThisMonthDatetime();

            dtpFromDate.Text = receiver[0].Substring(0, 7);
            dtpToDate.Text = receiver[1].Substring(0, 7); ;
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

            dtpFromDate.Text = receiver[0].Substring(0, 7);
            dtpToDate.Text = receiver[1].Substring(0, 7);
        }

        // 마지막 반년.
        private void btnLastSixMonth_Click(object sender, RoutedEventArgs e)
        {
            string[] receiver = lib.BringLastSixMonthDateTime();

            dtpFromDate.Text = receiver[0].Substring(0, 7); ;
            dtpToDate.Text = receiver[1].Substring(0, 7); ;
        }

        // 검사일자
        private void chkMonthDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkMonthDate.IsChecked == true)
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
        private void chkMonthDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkMonthDate.IsChecked == true)
            {
                chkMonthDate.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkMonthDate.IsChecked = true;
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

        //품목코드
        private void chkArticleNo_Click(object sender, RoutedEventArgs e)
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

        //품목코드
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

        // 거래처
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
        // 거래처
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

        #endregion


        #region 플러스 파인더
        // 플러스파인더 _ 고객사 찾기.
        private void btnCustomer_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomer, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }


        // 플러스파인더 _ 품명 찾기.
        private void btnArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, 77, "");
        }
        // 플러스파인더 _ 품번 찾기.
        private void btnArticleNo_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, 76, "");
        }
        #endregion


        private void Allclear()
        {
            dgdXBar_std.Items.Clear();
            dgdXBar_DailySpread.Items.Clear();
            txtCP.Clear();
            txtCPK.Clear();
            txtCPL.Clear();
            txtCPU.Clear();
            txtMaxValue.Clear();
            txtMinValue.Clear();
            txtRCL.Clear();
            txtRLCL.Clear();
            txtRUCL.Clear();
            txtXbarCL.Clear();
            txtXbarLCL.Clear();
            txtXbarUCL.Clear();
            txtAverage.Clear();
            txtStandardDeviation.Clear();

            if (lvcXBarChart.Series != null && lvcRChart.Series != null)
            {
                lvcXBarChart.Series.Clear();
                lvcRChart.Series.Clear();
            }

        }

        #region 조회 / 조회용 프로시저
        // 검색버튼 클릭.
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                Allclear();
                FillGrid();

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);


        }

        private void FillGrid()
        {
            int nchkDate = 0;
            if (chkMonthDate.IsChecked == true) { nchkDate = 1; }

            string SDate = string.Empty;
            string EDate = string.Empty;
            SDate = dtpFromDate.ToString().Substring(0, 7).Replace("-", "") + "01";
            EDate = dtpFromDate.ToString().Substring(0, 7).Replace("-", "") + "31";

            string InspectPoint = string.Empty;
            if (tbnInCome.IsChecked == true) { InspectPoint = "1"; }
            else if (tbnProcessCycle.IsChecked == true) { InspectPoint = "3"; }
            else if (tbnOutCome.IsChecked == true) { InspectPoint = "5"; }
            else if (tbnJaju.IsChecked == true) { InspectPoint = "9"; } //자주


            int ChkCustomID = 0;                //거래처.
            string CustomID = string.Empty;
            string Custom = string.Empty;

            int ChkArticleID = 0;              //품명.
            string ArticleID = string.Empty;
            string Article = string.Empty;


            if (chkCustomer.IsChecked == true)
            {
                if (txtCustomer.Tag == null)
                {
                    txtCustomer.Tag = "";
                    if (txtCustomer.Text.Length > 0)
                    {
                        ChkCustomID = 2;
                        Custom = txtCustomer.Text;
                    }
                }
                else
                {
                    ChkCustomID = 1;
                    CustomID = txtCustomer.Tag.ToString();
                }
            }


            if (chkArticle.IsChecked == true)
            {
                //if (txtArticle.Tag == null)
                //{
                //    txtArticle.Tag = "";
                //    if (txtArticle.Text.Length > 0)
                //    {
                //        ChkArticleID = 2;
                //        Article = txtArticle.Text;
                //    }
                //}
                //else
                //{
                //    ChkArticleID = 1;
                //    ArticleID = txtArticle.Tag.ToString();
                //}
                if (txtArticle == null || txtArticle.Text.Equals(""))
                {

                }
                else
                {
                    //ChkArticleID = 1;
                    ArticleID = txtArticle.Text.ToString();
                }
            }


            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            //sqlParameter.Add("nchkDate", nchkDate);               //int         
            sqlParameter.Add("FromDate", chkMonthDate.IsChecked == true ? SDate : "");
            sqlParameter.Add("ToDate", chkMonthDate.IsChecked == true ? EDate : "");

            sqlParameter.Add("InspectPoint", InspectPoint);

            //sqlParameter.Add("nchkCustom", ChkCustomID);             //int
            sqlParameter.Add("CustomID", CustomID);
            //sqlParameter.Add("Custom", Custom);

            //sqlParameter.Add("nchkArticleID", ChkArticleID);          //int
            sqlParameter.Add("@BuyerArticleNo", ArticleID);
            //sqlParameter.Add("Article", Article);

            DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Qual_sSpc_std", sqlParameter, true, "R");

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = null;
                dt = ds.Tables[0];

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("SPC 항목 조회결과가 없습니다.");
                    return;
                }
                else
                {
                    dgdXBar_std.Items.Clear();
                    int i = 1;
                    DataRowCollection drc = dt.Rows;
                    foreach (DataRow item in drc)
                    {
                        var Win_Qul_sts_XBarR_Q_Insert = new Win_Qul_XBarR_Q_View()
                        {
                            STD_NUM = i.ToString(),

                            STD_Article = item["Article"].ToString(),
                            STD_ArticleID = item["ArticleID"].ToString(),
                            STD_BuyerArticleNo = item["BuyerArticleNo"].ToString(),
                            STD_EcoNo = item["EcoNo"].ToString(),
                            STD_InspectBasisID = item["InspectBasisID"].ToString(),
                            STD_insItemName = item["insItemName"].ToString(),

                            STD_SubSeq = item["SubSeq"].ToString()
                        };
                        dgdXBar_std.Items.Add(Win_Qul_sts_XBarR_Q_Insert);
                        i++;
                    }

                    tbkCount.Text = "▶ 검색결과 : " + (i - 1) + " 건";
                }
            }

            DataStore.Instance.CloseConnection();
        }

        #endregion


        #region 메인그리드 row enter. Show Data
        // 좌측 std 그리드 클릭시. 로우엔터 이벤트.
        private void dgdXBar_std_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var ViewReceiver = dgdXBar_std.SelectedItem as Win_Qul_XBarR_Q_View;

            if (ViewReceiver != null)
            {
                int nchkDate = 0;
                if (chkMonthDate.IsChecked == true) { nchkDate = 1; }

                string SDate = string.Empty;
                string EDate = string.Empty;
                SDate = dtpFromDate.ToString().Substring(0, 7).Replace("-", "") + "01";
                EDate = dtpFromDate.ToString().Substring(0, 7).Replace("-", "") + "31";

                string InspectPoint = string.Empty;
                if (tbnInCome.IsChecked == true) { InspectPoint = "1"; }
                else if (tbnProcessCycle.IsChecked == true) { InspectPoint = "3"; }
                else if (tbnOutCome.IsChecked == true) { InspectPoint = "5"; }
                else if (tbnJaju.IsChecked == true) { InspectPoint = "9"; }


                string InspectBasisID = ViewReceiver.STD_InspectBasisID;
                int Seq = Convert.ToInt32(ViewReceiver.STD_SubSeq);


                // Specification 공간 채울 데이터 구하기.
                FillGrid_Specification(nchkDate, SDate, EDate, InspectBasisID, Seq);
                //(최대값 / 최소값 구하기 용도...)


                //통계치 Summary 공간 채울 데이터 구하기.
                FillGrid_Summary(InspectPoint, nchkDate, SDate, InspectBasisID, Seq);


                // 차트 그리기.
                FillChart_Double();


                // 전역 값들 초기화
                RASpecMax = string.Empty;
                RASpecMin = string.Empty;

                X_chart_UCL = 0;
                X_chart_CL = 0;
                X_chart_LCL = 0;
                x_chart_val.Clear();

                R_chart_UCL = 0;
                R_chart_CL = 0;
                r_chart_val.Clear();
            }

        }

        #endregion



        private void FillGrid_Specification(int nchkDate, string FromDate, string ToDate,
                                            string InspectBasisID, int SubSeq)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            //sqlParameter.Add("nchkDate", nchkDate);
            sqlParameter.Add("FromDate", FromDate);
            sqlParameter.Add("ToDate", ToDate);
            sqlParameter.Add("InspectBasisID", InspectBasisID);
            sqlParameter.Add("SubSeq", SubSeq);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qual_sSpc_spec", sqlParameter, false);

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
                    RASpecMax = dr["InsRASpecMax"].ToString();
                    RASpecMin = dr["InsRASpecMin"].ToString();
                }

            }

            DataStore.Instance.CloseConnection();
        }


        #region 통계치 Summary
        //통계치 Summary 공간 채울 데이터 구하기.
        private void FillGrid_Summary(string InspectPoint, int nchkDate, string sMonth, string InspectBasisID,
                                        int InspectSubSeq)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("InspectPoint", InspectPoint);

            //sqlParameter.Add("nchkDate", nchkDate);                  //int
            sqlParameter.Add("sMonth", sMonth.Substring(0, 6));

            sqlParameter.Add("InspectBasisID", InspectBasisID);
            sqlParameter.Add("InspectSubSeq", InspectSubSeq);          //int

            sqlParameter.Add("CustomID", "");

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qual_sSpc_DailySpread", sqlParameter, false);

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
                    dgdXBar_DailySpread.Items.Clear();
                    DataRowCollection drc = dt.Rows;
                    foreach (DataRow item in drc)
                    {
                        if (item["cls"].ToString() == "01")         //  cls 가 1인 애들이 스프레드 시트에 뿌려진다.
                        {
                            var Win_Qul_sts_XBarR_Q_Insert = new Win_Qul_XBarR_Q_View()
                            {
                                DAYSPREAD_seq = "측정값" + item["seq"].ToString(),

                                DAYSPREAD_DD1 = item["InspectValue1"].ToString(),
                                DAYSPREAD_DD2 = item["InspectValue2"].ToString(),
                                DAYSPREAD_DD3 = item["InspectValue3"].ToString(),
                                DAYSPREAD_DD4 = item["InspectValue4"].ToString(),
                                DAYSPREAD_DD5 = item["InspectValue5"].ToString(),
                                DAYSPREAD_DD6 = item["InspectValue6"].ToString(),
                                DAYSPREAD_DD7 = item["InspectValue7"].ToString(),
                                DAYSPREAD_DD8 = item["InspectValue8"].ToString(),
                                DAYSPREAD_DD9 = item["InspectValue9"].ToString(),
                                DAYSPREAD_DD10 = item["InspectValue10"].ToString(),
                                DAYSPREAD_DD11 = item["InspectValue11"].ToString(),
                                DAYSPREAD_DD12 = item["InspectValue12"].ToString(),
                                DAYSPREAD_DD13 = item["InspectValue13"].ToString(),
                                DAYSPREAD_DD14 = item["InspectValue14"].ToString(),
                                DAYSPREAD_DD15 = item["InspectValue15"].ToString(),
                                DAYSPREAD_DD16 = item["InspectValue16"].ToString(),
                                DAYSPREAD_DD17 = item["InspectValue17"].ToString(),
                                DAYSPREAD_DD18 = item["InspectValue18"].ToString(),
                                DAYSPREAD_DD19 = item["InspectValue19"].ToString(),
                                DAYSPREAD_DD20 = item["InspectValue20"].ToString(),
                                DAYSPREAD_DD21 = item["InspectValue21"].ToString(),
                                DAYSPREAD_DD22 = item["InspectValue22"].ToString(),
                                DAYSPREAD_DD23 = item["InspectValue23"].ToString(),
                                DAYSPREAD_DD24 = item["InspectValue24"].ToString(),
                                DAYSPREAD_DD25 = item["InspectValue25"].ToString(),
                                DAYSPREAD_DD26 = item["InspectValue26"].ToString(),
                                DAYSPREAD_DD27 = item["InspectValue27"].ToString(),
                                DAYSPREAD_DD28 = item["InspectValue28"].ToString(),
                                DAYSPREAD_DD29 = item["InspectValue29"].ToString(),
                                DAYSPREAD_DD30 = item["InspectValue30"].ToString(),
                                DAYSPREAD_DD31 = item["InspectValue31"].ToString(),

                            };
                            dgdXBar_DailySpread.Items.Add(Win_Qul_sts_XBarR_Q_Insert);
                        }


                        // X_BAR CHART 그리기.
                        else if (item["cls"].ToString() == "02")
                        {
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue1"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue2"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue3"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue4"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue5"]));

                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue6"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue7"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue8"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue9"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue10"]));

                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue11"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue12"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue13"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue14"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue15"]));

                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue16"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue17"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue18"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue19"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue20"]));

                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue21"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue22"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue23"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue24"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue25"]));

                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue26"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue27"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue28"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue29"]));
                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue30"]));

                            x_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue31"]));
                        }


                        // R CHART 그리기.
                        else if (item["cls"].ToString() == "03")
                        {
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue1"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue2"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue3"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue4"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue5"]));

                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue6"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue7"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue8"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue9"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue10"]));

                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue11"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue12"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue13"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue14"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue15"]));

                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue16"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue17"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue18"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue19"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue20"]));

                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue21"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue22"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue23"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue24"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue25"]));

                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue26"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue27"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue28"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue29"]));
                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue30"]));

                            r_chart_val.Add(Lib.Instance.returnDouble(item["InspectValue31"]));
                        }


                        // X-Bar UCL
                        else if (item["cls"].ToString() == "04")
                        {
                            txtXbarUCL.Text = Lib.Instance.returnNumStringTargetNum(item["InspectValue1"].ToString(), 3);
                        }
                        // X-Bar CL
                        else if (item["cls"].ToString() == "05")
                        {
                            txtXbarCL.Text = Lib.Instance.returnNumStringTargetNum(item["InspectValue1"].ToString(), 3);
                        }
                        // X-Bar LCL
                        else if (item["cls"].ToString() == "06")
                        {
                            txtXbarLCL.Text = Lib.Instance.returnNumStringTargetNum(item["InspectValue1"].ToString(), 3);
                        }
                        // R UCL
                        else if (item["cls"].ToString() == "07")
                        {
                            txtRUCL.Text = Lib.Instance.returnNumStringTargetNum(item["InspectValue1"].ToString(), 3);
                        }
                        // R CL
                        else if (item["cls"].ToString() == "08")
                        {
                            txtRCL.Text = Lib.Instance.returnNumStringTargetNum(item["InspectValue1"].ToString(), 3);
                        }

                        // 최대값 텍스트박스.
                        else if (item["cls"].ToString() == "09")
                        {
                            txtMaxValue.Text = Lib.Instance.returnNumStringTargetNum(item["InspectValue1"].ToString(), 3);
                        }
                        // 최소값 텍스트박스.
                        else if (item["cls"].ToString() == "10")
                        {
                            txtMinValue.Text = Lib.Instance.returnNumStringTargetNum(item["InspectValue1"].ToString(), 3);
                        }
                        // 평균 텍스트박스.
                        else if (item["cls"].ToString() == "12")
                        {
                            txtAverage.Text = Lib.Instance.returnNumStringTargetNum(item["InspectValue1"].ToString(), 3);
                        }
                        // 표준편차 텍스트박스.
                        else if (item["cls"].ToString() == "13")
                        {
                            txtStandardDeviation.Text = Lib.Instance.returnNumStringTargetNum(item["InspectValue1"].ToString(), 3);
                        }

                    }
                    if (txtStandardDeviation.Text != "")
                    {
                        double dSTDev = Convert.ToDouble(txtStandardDeviation.Text);    // 표준편차 값.
                        double dAvr;
                        double dUSL;
                        double dLSL;

                        if (dSTDev == 0)
                        {
                            txtCPU.Text = "0";
                            txtCPL.Text = "0";
                            txtCP.Text = "0";

                        }
                        else
                        {
                            dAvr = Convert.ToDouble(txtAverage.Text);                // 평균.
                            dUSL = Convert.ToDouble(RASpecMax);           // 상한공차 값.
                            dLSL = Convert.ToDouble(RASpecMin);           // 하한공차 값.

                            //txtCPU.Text = Math.Round((dUSL - dAvr) / (3 * dSTDev)).ToString();
                            //txtCPL.Text = Math.Round((dAvr - dLSL) / (3 * dSTDev)).ToString();
                            //txtCP.Text = Math.Round((dUSL - dLSL) / (6 * dSTDev)).ToString();

                            //소수점 2째자리까지 나타나게 요청(3자리까지 보여주자...)_20190508 최준호
                            txtCPU.Text = Lib.Instance.returnNumStringTargetNum(((dUSL - dAvr) / (3 * dSTDev)).ToString(), 3);
                            txtCPL.Text = Lib.Instance.returnNumStringTargetNum(((dAvr - dLSL) / (3 * dSTDev)).ToString(), 3);
                            txtCP.Text = Lib.Instance.returnNumStringTargetNum(((dUSL - dLSL) / (6 * dSTDev)).ToString(), 3);
                        }
                        if (Convert.ToDouble(txtCPU.Text) > Convert.ToDouble(txtCPL.Text))
                        {
                            txtCPK.Text = txtCPL.Text;
                        }
                        else
                        {
                            txtCPK.Text = txtCPU.Text;
                        }
                    }
                    else
                    {
                        string caution = "경고";
                        string msg = "입력된 값이 없습니다!";

                        MessageBox.Show(msg, caution);
                    }
                }
            }

            DataStore.Instance.CloseConnection();
        }

        #endregion


        #region 차트
        // 차트 ㅡ그리기.
        private void FillChart_Double()
        {
            if (txtXbarUCL.Text != string.Empty) { X_chart_UCL = Convert.ToDouble(Lib.Instance.returnNumStringTargetNum(txtXbarUCL.Text, 3).Replace(",", "")); }
            if (txtXbarCL.Text != string.Empty) { X_chart_CL = Convert.ToDouble(Lib.Instance.returnNumStringTargetNum(txtXbarCL.Text, 3).Replace(",", "")); }
            if (txtXbarLCL.Text != string.Empty) { X_chart_LCL = Convert.ToDouble(Lib.Instance.returnNumStringTargetNum(txtXbarLCL.Text, 3).Replace(",", "")); }

            if (txtRUCL.Text != string.Empty) { R_chart_UCL = Convert.ToDouble(Lib.Instance.returnNumStringTargetNum(txtRUCL.Text, 3).Replace(",", "")); }
            if (txtRCL.Text != string.Empty) { R_chart_CL = Convert.ToDouble(Lib.Instance.returnNumStringTargetNum(txtRCL.Text, 3).Replace(",", "")); }

            ChartValues<double> XUCL_Charts = new ChartValues<double>();
            ChartValues<double> XCL_Charts = new ChartValues<double>();
            ChartValues<double> XLCL_Charts = new ChartValues<double>();
            ChartValues<double> XVal_Charts = new ChartValues<double>();
            List<string> XLabels = new List<string>();
            List<string> RLabels = new List<string>();

            int j = 0;
            for (int i = 0; i < x_chart_val.Count; i++)
            {
                if (x_chart_val[i] > 0.00)
                {
                    j++;
                    XUCL_Charts.Add(X_chart_UCL);
                    XCL_Charts.Add(X_chart_CL);
                    XLCL_Charts.Add(X_chart_LCL);
                    XVal_Charts.Add(x_chart_val[i]);
                    XLabels.Add("순" + j + " - " + (i + 1) + "일");
                }
            }

            ChartValues<double> RUCL_Charts = new ChartValues<double>();
            ChartValues<double> RCL_Charts = new ChartValues<double>();
            ChartValues<double> RVal_Charts = new ChartValues<double>();

            j = 0;
            for (int i = 0; i < r_chart_val.Count; i++)
            {
                if (r_chart_val[i] > 0.00)
                {
                    j++;
                    RUCL_Charts.Add(R_chart_UCL);
                    RCL_Charts.Add(R_chart_CL);
                    RVal_Charts.Add(r_chart_val[i]);
                    RLabels.Add("순" + j + " - " + (i + 1) + "일");
                }
            }

            SeriesCollection SeriesCollection = new SeriesCollection
            {
                new LineSeries
                {
                    Title = "UCL",
                    Values = XUCL_Charts
                },
                new LineSeries
                {
                    Title = "CL",
                    Values = XCL_Charts
                },
                new LineSeries
                {
                    Title = "LCL",
                    Values = XLCL_Charts
                },
                new LineSeries
                {
                    Title = "측정값",
                    Values = XVal_Charts
                }
            };
            lvcXBarChart.Series = SeriesCollection;
            lvcXBarChart.AxisX[0].Labels = XLabels;

            SeriesCollection SeriesCollection2 = new SeriesCollection
            {
                new LineSeries
                {
                    Title = "UCL",
                    Values = RUCL_Charts
                },
                new LineSeries
                {
                    Title = "CL",
                    Values = RCL_Charts
                },
                new LineSeries
                {
                    Title = "측정값",
                    Values = RVal_Charts
                }
            };
            lvcRChart.Series = SeriesCollection2;
            lvcRChart.AxisX[0].Labels = RLabels;

        }

        #endregion


        // 닫기.
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }



        #region 엑셀
        // 엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdXBar_std.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            DataTable dt = null;
            string Name = string.Empty;
            Lib lib = new Lib();

            string[] lst = new string[4];
            lst[0] = "메인그리드";
            lst[1] = "통계치그리드";
            lst[2] = dgdXBar_std.Name;
            lst[3] = dgdXBar_DailySpread.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                if (ExpExc.choice.Equals(dgdXBar_std.Name))
                {
                    //MessageBox.Show("대분류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdXBar_std);
                    else
                        dt = lib.DataGirdToDataTable(dgdXBar_std);

                    Name = dgdXBar_std.Name;

                    if (lib.GenerateExcel(dt, Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
                }
                else if (ExpExc.choice.Equals(dgdXBar_DailySpread.Name))
                {
                    //MessageBox.Show("정성류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdXBar_DailySpread);
                    else
                        dt = lib.DataGirdToDataTable(dgdXBar_DailySpread);
                    Name = dgdXBar_DailySpread.Name;

                    if (lib.GenerateExcel(dt, Name))
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
            lib = null;
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

        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, 77, "");
            }
        }

        private void txtArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, 76, "");
            }
        }

        //주 모니터 대신 어플리케이션 활성화 영역을 캡쳐
        private void BtnCapture_Click(object sender, RoutedEventArgs e)
        {
            FrameworkElement targetElement = this;
            lib.ScreenCapture(targetElement);
        }

        private void btnDetail_Click(object sender, RoutedEventArgs e)
        {
            var Detail = dgdXBar_std.SelectedItem as Win_Qul_XBarR_Q_View;

            if (Detail != null)
            {

            }
        }

        private void dgdXBar_std_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var Detail = dgdXBar_std.SelectedItem as Win_Qul_XBarR_Q_View;

            if (Detail != null)
            {
                string ItemCode = Detail.STD_ArticleID;
                string ItemName = Detail.STD_Article;
                string ChkDate = chkMonthDate.IsChecked.ToString();
                string SDate = dtpFromDate.SelectedDate.Value.ToString("yyyyMM") + "01";
                string EDate = Convert.ToDateTime(dtpFromDate.SelectedDate.Value.ToString("yyyy-MM") + "-01").AddMonths(1).AddDays(-1).ToString("yyyyMMdd");

                MainWindow.ChartDeatil.Clear();
                MainWindow.ChartDeatil.Add(ItemCode);
                MainWindow.ChartDeatil.Add(ItemName);
                MainWindow.ChartDeatil.Add(ChkDate);
                MainWindow.ChartDeatil.Add(SDate);
                MainWindow.ChartDeatil.Add(EDate);

                int i = 0;
                foreach (MenuViewModel mvm in MainWindow.mMenulist)
                {
                    if (mvm.Menu.Equals("품질 상세분석"))
                    {
                        break;
                    }
                    i++;
                }
                try
                {
                    if (MainWindow.MainMdiContainer.Children.Contains(MainWindow.mMenulist[i].subProgramID as MdiChild))
                    {
                        (MainWindow.mMenulist[i].subProgramID as MdiChild).Focus();
                    }
                    else
                    {
                        Type type = Type.GetType("WizMes_WellMade." + MainWindow.mMenulist[i].ProgramID.Trim(), true);
                        object uie = Activator.CreateInstance(type);

                        MainWindow.mMenulist[i].subProgramID = new MdiChild()
                        {
                            Title = "WellMade [" + MainWindow.mMenulist[i].MenuID.Trim() + "] " + MainWindow.mMenulist[i].Menu.Trim() +
                                    " (→" + MainWindow.mMenulist[i].ProgramID + ")",
                            Height = SystemParameters.PrimaryScreenHeight * 0.8,
                            MaxHeight = SystemParameters.PrimaryScreenHeight * 0.85,
                            Width = SystemParameters.WorkArea.Width * 0.85,
                            MaxWidth = SystemParameters.WorkArea.Width,
                            Content = uie as UIElement,
                            Tag = MainWindow.mMenulist[i]
                        };

                        Lib.Instance.AllMenuLogInsert(MainWindow.mMenulist[i].MenuID, MainWindow.mMenulist[i].Menu, MainWindow.mMenulist[i].subProgramID);
                        MainWindow.MainMdiContainer.Children.Add(MainWindow.mMenulist[i].subProgramID as MdiChild);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("해당 화면이 존재하지 않습니다.");
                }
            }
        }
    }

    class Win_Qul_XBarR_Q_View : BaseView
    {
        public Win_Qul_XBarR_Q_View()
        {
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }


        // x - bar_std. 리스트 그리드.
        public string STD_NUM { get; set; }

        public string STD_Article { get; set; }
        public string STD_ArticleID { get; set; }
        public string STD_BuyerArticleNo { get; set; }
        public string STD_EcoNo { get; set; }
        public string STD_InspectBasisID { get; set; }
        public string STD_insItemName { get; set; }

        public string STD_SubSeq { get; set; }



        // x - bar _ DailySpread 그리드.
        public string DAYSPREAD_seq { get; set; }

        public string DAYSPREAD_DD1 { get; set; }
        public string DAYSPREAD_DD2 { get; set; }
        public string DAYSPREAD_DD3 { get; set; }
        public string DAYSPREAD_DD4 { get; set; }
        public string DAYSPREAD_DD5 { get; set; }
        public string DAYSPREAD_DD6 { get; set; }
        public string DAYSPREAD_DD7 { get; set; }
        public string DAYSPREAD_DD8 { get; set; }
        public string DAYSPREAD_DD9 { get; set; }
        public string DAYSPREAD_DD10 { get; set; }
        public string DAYSPREAD_DD11 { get; set; }
        public string DAYSPREAD_DD12 { get; set; }
        public string DAYSPREAD_DD13 { get; set; }
        public string DAYSPREAD_DD14 { get; set; }
        public string DAYSPREAD_DD15 { get; set; }
        public string DAYSPREAD_DD16 { get; set; }
        public string DAYSPREAD_DD17 { get; set; }
        public string DAYSPREAD_DD18 { get; set; }
        public string DAYSPREAD_DD19 { get; set; }
        public string DAYSPREAD_DD20 { get; set; }
        public string DAYSPREAD_DD21 { get; set; }
        public string DAYSPREAD_DD22 { get; set; }
        public string DAYSPREAD_DD23 { get; set; }
        public string DAYSPREAD_DD24 { get; set; }
        public string DAYSPREAD_DD25 { get; set; }
        public string DAYSPREAD_DD26 { get; set; }
        public string DAYSPREAD_DD27 { get; set; }
        public string DAYSPREAD_DD28 { get; set; }
        public string DAYSPREAD_DD29 { get; set; }
        public string DAYSPREAD_DD30 { get; set; }
        public string DAYSPREAD_DD31 { get; set; }

    }
}
