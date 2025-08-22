using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Shapes;
using WizMes_WellMade.PopUP;
using WizMes_WellMade.PopUp;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_ord_InOutSum_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_InOutSum_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        // 그리드 셀렉트 도전(2018_08_09)
        int Clicked_row = 0;
        int Clicked_col = 0;
        List<Rectangle> PreRect = new List<Rectangle>();

        //전역변수는 이럴때 쓰는거 아니겠어??!!?
        private DataTable PeriodDataTable = null;
        private DataTable DaysDataTable = null;
        private DataTable MonthDataTable = null;
        private DataTable SpreadMonthDataTable = null;


        public Win_ord_InOutSum_Q()
        {
            InitializeComponent();
        }

        // 화면 첫 시작.
        private void Window_InOutTotalGrid_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            First_Step();
            ComboBoxSetting();
        }

        #region  첫 스텝 // 일자버튼 // 초기설정 // 조회용 체크박스 컨트롤 
        private void First_Step()
        {
            // 월별 가로집계 최근 3개월 지정하기.
            List<MonthChange> MC = new List<MonthChange>();
            MC.Add(new MonthChange()
            {
                H_MON1 = DateTime.Now.ToString("yyyy-MM"),
                H_MON2 = DateTime.Now.AddMonths(-1).ToString("yyyy-MM"),
                H_MON3 = DateTime.Now.AddMonths(-2).ToString("yyyy-MM"),
            });

            this.DataContext = MC;
            //////////////////////////////////////


            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            txtblMessage.Visibility = Visibility.Hidden;


        }

        // 어제.(전일)
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            //string[] receiver = lib.BringYesterdayDatetime();

            //dtpFromDate.Text = receiver[0];
            //dtpToDate.Text = receiver[1];

            if (dtpFromDate.SelectedDate != null)
            {
                dtpFromDate.SelectedDate = dtpFromDate.SelectedDate.Value.AddDays(-1);
                dtpToDate.SelectedDate = dtpFromDate.SelectedDate;
            }
            else
            {
                dtpFromDate.SelectedDate = DateTime.Today.AddDays(-1);
                dtpToDate.SelectedDate = DateTime.Today.AddDays(-1);
            }
        }
        // 오늘(금일)
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
        }
        // 지난 달(전월)
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            //string[] receiver = lib.BringLastMonthDatetime();

            //dtpFromDate.Text = receiver[0];
            //dtpToDate.Text = receiver[1];

            if (dtpFromDate.SelectedDate != null)
            {
                DateTime ThatMonth1 = dtpFromDate.SelectedDate.Value.AddDays(-(dtpFromDate.SelectedDate.Value.Day - 1)); // 선택한 일자 달의 1일!

                DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                dtpFromDate.SelectedDate = LastMonth1;
                dtpToDate.SelectedDate = LastMonth31;
            }

        }
        // 이번 달(금월)
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            string[] receiver = lib.BringThisMonthDatetime();

            dtpFromDate.Text = receiver[0];
            dtpToDate.Text = receiver[1];
        }


        #endregion

        #region 콤보박스 세팅
        // 콤보박스 세팅.
        private void ComboBoxSetting()
        {
            cboInOutGubunSrh.Items.Clear();
            cboInInspectGubunSrh.Items.Clear();

            string[] DirectCombo = new string[2];
            DirectCombo[0] = "Y";
            DirectCombo[1] = "합격";
            string[] DirectCombo1 = new string[2];
            DirectCombo1[0] = "N";
            DirectCombo1[1] = "불합격";

            List<string[]> DirectCombOList = new List<string[]>();
            DirectCombOList.Add(DirectCombo.ToArray());
            DirectCombOList.Add(DirectCombo1.ToArray());

            ObservableCollection<CodeView> cbInInspectGubunSrh = ComboBoxUtil.Instance.Direct_SetComboBox(DirectCombOList);

            DirectCombo = new string[2];
            DirectCombo[0] = "1";
            DirectCombo[1] = "입고";
            DirectCombo1 = new string[2];
            DirectCombo1[0] = "2";
            DirectCombo1[1] = "출고";

            DirectCombOList = new List<string[]>();
            DirectCombOList.Add(DirectCombo.ToArray());
            DirectCombOList.Add(DirectCombo1.ToArray());

            ObservableCollection<CodeView> cbInOutGubunSrh = ComboBoxUtil.Instance.Direct_SetComboBox(DirectCombOList);

            this.cboInOutGubunSrh.ItemsSource = cbInOutGubunSrh;
            this.cboInOutGubunSrh.DisplayMemberPath = "code_name";
            this.cboInOutGubunSrh.SelectedValuePath = "code_id";
            this.cboInOutGubunSrh.SelectedIndex = 0;

            this.cboInInspectGubunSrh.ItemsSource = cbInInspectGubunSrh;
            this.cboInInspectGubunSrh.DisplayMemberPath = "code_name";
            this.cboInInspectGubunSrh.SelectedValuePath = "code_id";
            this.cboInInspectGubunSrh.SelectedIndex = 0;

        }
        #endregion

        #region 플러스 파인더
        //플러스 파인더

        //거래처
        private void btnCustomIDSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomIDSrh, 0, "");
        }

        // 품명
        private void btnArticleIDSrh_click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticleIDSrh, 77, "");
        }

        #endregion


        // 검색(조회) 버튼 클릭
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if (tiMonth_H.IsSelected || tiMonth_V.IsSelected || lib.DatePickerCheck(dtpFromDate, dtpToDate, chkDateSrh) )
            {
                using (Loading ld = new Loading(beSearch))
                {
                    ld.ShowDialog();
                }
            }

        }

        private void beSearch()
        {
            //검색버튼 비활성화   
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                DataStore.Instance.InsertLogByForm(this.GetType().Name, "R");
                TabItem NowTI = tabconGrid.SelectedItem as TabItem;

                if (NowTI.Header.ToString() == "기간집계") { FillGrid_Period(); }
                else if (NowTI.Header.ToString() == "일일집계") { FillGrid_Day(); }
                else if (NowTI.Header.ToString() == "월별집계(세로)") { FillGrid_Month_V(); }
                else if (NowTI.Header.ToString() == "월별집계(가로)") { FillGrid_Month_H(); }

            }), System.Windows.Threading.DispatcherPriority.Background);      

            btnSearch.IsEnabled = true;
        }

        #region 기간집계 조회
        //기간집계 조회
        private void FillGrid_Period()
        {
            grdPeriod.Items.Clear();
            dgdPeriodTotal.Items.Clear();



            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", !lib.IsDatePickerNull(dtpFromDate) ? lib.ConvertDate(dtpFromDate) : "");
                sqlParameter.Add("EDate", !lib.IsDatePickerNull(dtpToDate) ? lib.ConvertDate(dtpToDate) : "");

                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag != null ? txtBuyerArticleNoSrh.Tag.ToString() : string.Empty : string.Empty);

                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkOrder", chkOrderIDSrh.IsChecked == true ? rbnOrderNOSrh.IsChecked == true ? 1 : 2 : 0);
                sqlParameter.Add("Order", chkOrderIDSrh.IsChecked == true ? !string.IsNullOrEmpty(txtOrderIDSrh.Text) ? txtOrderIDSrh.Text : "" : "");



                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sInOutwareSum_Period", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    PeriodDataTable = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        //grdMerge_Period.RowDefinitions.Clear();

                        int i = 0;
                        int totalOutRoll = 0;
                        int totalOutQty = 0;
                        int totalStuffRoll = 0;
                        int totalStuffQty = 0;
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var PeriodItem = new Win_ord_InOutSum_QView
                            {
                                P_NUM = i,
                                P_Gbn = dr["Gbn"].ToString(),
                                P_CustomName = dr["KCustom"].ToString(),
                                P_BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                P_Article = dr["Article"].ToString(),
                                P_Roll = stringFormatN0(dr["Roll"]),
                                P_Qty = stringFormatN0(dr["TotQty"]),
                                P_UnitClssName = dr["UnitClssName"].ToString(),
                                P_CustomRate = stringFormatN2(dr["CustomRate"]),

                            };

                            if (PeriodItem.P_Gbn.Equals("1"))
                            {
                                PeriodItem.P_Gbn = "입고";
                                totalStuffRoll += ConvertInt(PeriodItem.P_Roll);
                                totalStuffQty += ConvertInt(PeriodItem.P_Qty);

                                grdPeriod.Items.Add(PeriodItem);
                            }
                            else if (PeriodItem.P_Gbn.Equals("2"))
                            {
                                PeriodItem.P_Gbn = "출고";
                                totalOutRoll += ConvertInt(PeriodItem.P_Roll);
                                totalOutQty += ConvertInt(PeriodItem.P_Qty);

                                grdPeriod.Items.Add(PeriodItem);

                            }
                            else if (PeriodItem.P_Gbn.Equals("3"))
                            {
                                PeriodItem.P_Color1 = true;
                                PeriodItem.P_Gbn = string.Empty;
                                PeriodItem.P_Article = "거래처 계";
                                PeriodItem.P_CustomName = string.Empty;
                                grdPeriod.Items.Add(PeriodItem);
                            }
                            else if (PeriodItem.P_Gbn.Equals("4"))
                            {
                                PeriodItem.P_Color2 = true;
                                PeriodItem.P_Gbn = string.Empty;
                                PeriodItem.P_CustomName = string.Empty;
                                grdPeriod.Items.Add(PeriodItem);
                            }
                        }

                        if (grdPeriod.Items.Count > 0)
                        {
                            var PeriodTotal = new Win_ord_InOutSum_Total_QView
                            {
                                P_TotalOutRoll = stringFormatN0(totalOutRoll),
                                P_TotalOutQty = stringFormatN0(totalOutQty),
                                P_TotalStuffRoll = stringFormatN0(totalStuffRoll),
                                P_TotalStuffQty = stringFormatN0(totalStuffQty),
                            };

                            dgdPeriodTotal.Items.Add(PeriodTotal);
                        }


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

        #region 일일집계 조회
        //일일집계 조회
        private void FillGrid_Day()
        {

            grdMergeDays.Items.Clear();
            dgdDaysOutTotal.Items.Clear();

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", !lib.IsDatePickerNull(dtpFromDate) ? lib.ConvertDate(dtpFromDate) : "");
                sqlParameter.Add("EDate", !lib.IsDatePickerNull(dtpToDate) ? lib.ConvertDate(dtpToDate) : "");

                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag != null ? txtBuyerArticleNoSrh.Tag.ToString() : string.Empty : string.Empty);

                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkOrder", chkOrderIDSrh.IsChecked == true ? rbnOrderNOSrh.IsChecked == true ? 1 : 2 : 0);
                sqlParameter.Add("Order", chkOrderIDSrh.IsChecked == true ? !string.IsNullOrEmpty(txtOrderIDSrh.Text) ? txtOrderIDSrh.Text : "" : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sInOutwareSum_Day", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    DaysDataTable = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;


                        int i = 0;
                        int totalOutRoll = 0;
                        int totalOutQty = 0;
                        int totalStuffRoll = 0;
                        int totalStuffQty = 0;
                        int totalOutAmount = 0;
                        int totalOutVatAmount = 0;
                        int totalOutPrice = 0;
                        int totalStuffAmount = 0;
                        int totalStuffVatAmount = 0;
                        int totalStuffPrice = 0;
                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var DayItem = new Win_ord_InOutSum_QView
                            {
                                D_NUM = i,
                                D_IODate = lib.DateTypeHyphen(dr["IODate"].ToString()),
                                D_Gbn = dr["Gbn"].ToString(),
                                D_CustomName = dr["KCustom"].ToString(),
                                D_BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                D_Article = dr["Article"].ToString(),
                                D_Roll = stringFormatN0(dr["Roll"]),
                                D_Qty = stringFormatN0(dr["TotQty"]),
                                D_UnitClssName = dr["UnitClssName"].ToString(),
                                D_Amount = stringFormatN0(dr["Amount"]),
                                D_VatAmount = stringFormatN0(dr["VatAmount"]),
                                D_TotAmount = stringFormatN0(dr["TotalAmount"]),
                                D_CustomRate = stringFormatN2(dr["CustomRate"])

                            };

                            if (DayItem.D_Gbn.Equals("1"))
                            {
                                DayItem.D_Gbn = "입고";
                                totalStuffRoll += ConvertInt(DayItem.D_Roll);
                                totalStuffQty += ConvertInt(DayItem.D_Qty);
                                totalStuffAmount += ConvertInt(DayItem.D_Amount);
                                totalStuffVatAmount += ConvertInt(DayItem.D_VatAmount);
                                totalStuffPrice += ConvertInt(DayItem.D_TotAmount);
                                grdMergeDays.Items.Add(DayItem);
                            }
                            else if (DayItem.D_Gbn.Equals("2"))
                            {
                                DayItem.D_Gbn = "출고";
                                totalOutRoll += ConvertInt(DayItem.D_Roll);
                                totalOutQty += ConvertInt(DayItem.D_Qty);
                                totalOutAmount += ConvertInt(DayItem.D_Amount);
                                totalOutVatAmount += ConvertInt(DayItem.D_VatAmount);
                                totalOutPrice += ConvertInt(DayItem.D_TotAmount);
                                grdMergeDays.Items.Add(DayItem);

                            }
                            else if (DayItem.D_Gbn.Equals("3"))
                            {
                                DayItem.D_Color1 = true;
                                DayItem.D_Gbn = string.Empty;
                                grdMergeDays.Items.Add(DayItem);
                            }
                        }

                        if (grdMergeDays.Items.Count > 0)
                        {
                            var PeriodStuffTotal = new Win_ord_InOutSum_Total_QView
                            {
                                D_TotalStuffRoll = stringFormatN0(totalStuffRoll),
                                D_TotalStuffQty = stringFormatN0(totalStuffQty),
                                D_TotalStuffAmount = stringFormatN0(totalStuffAmount),
                                D_TotalStuffVatAmount = stringFormatN0(totalStuffVatAmount),
                                D_TotalStuffPrice = stringFormatN0(totalStuffPrice),
                            };


                            var PeriodOutTotal = new Win_ord_InOutSum_Total_QView
                            {
                                D_TotalOutRoll = stringFormatN0(totalOutRoll),
                                D_TotalOutQty = stringFormatN0(totalOutQty),
                                D_TotalOutAmount = stringFormatN0(totalOutAmount),
                                D_TotalOutVatAmount = stringFormatN0(totalOutVatAmount),
                                D_TotalOutPrice = stringFormatN0(totalOutPrice),
                            };

                            dgdDaysStuffTotal.Items.Add(PeriodStuffTotal);
                            dgdDaysOutTotal.Items.Add(PeriodOutTotal);
                        }

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

        #region 월별집계 (세로) 조회
        //월별집계 (세로) 조회
        private void FillGrid_Month_V()
        {

            grdMergeMonth_V.Items.Clear();


            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", !lib.IsDatePickerNull(dtpFromDate) ? lib.ConvertDate(dtpFromDate) : "");
                sqlParameter.Add("EDate", !lib.IsDatePickerNull(dtpToDate) ? lib.ConvertDate(dtpToDate) : "");

                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag != null ? txtBuyerArticleNoSrh.Tag.ToString() : string.Empty : string.Empty);

                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkOrder", chkOrderIDSrh.IsChecked == true ? rbnOrderNOSrh.IsChecked == true ? 1 : 2 : 0);
                sqlParameter.Add("Order", chkOrderIDSrh.IsChecked == true ? !string.IsNullOrEmpty(txtOrderIDSrh.Text) ? txtOrderIDSrh.Text : "" : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sInOutwareSum_Month", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    MonthDataTable = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        //grdMerge_Month_V.RowDefinitions.Clear();

                        DataRowCollection drc = dt.Rows;
                        int i = 0;
                        int totalOutRoll = 0;
                        int totalOutQty = 0;
                        int totalStuffRoll = 0;
                        int totalStuffQty = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++; ;
                            var MonthHItem = new Win_ord_InOutSum_QView
                            {
                                V_NUM = i,
                                V_IODate = lib.DateTypeHyphen(dr["IODate"].ToString()),
                                V_Gbn = dr["Gbn"].ToString(),
                                V_CustomName = dr["KCustom"].ToString(),
                                V_BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                V_Article = dr["Article"].ToString(),
                                V_Roll = stringFormatN0(dr["Roll"]),
                                V_Qty = stringFormatN0(dr["TotQty"]),
                                V_UnitClssName = dr["UnitClssName"].ToString(),
                                V_CustomRate = stringFormatN2(dr["CustomRate"])
                            };

                            if (MonthHItem.V_Gbn.Equals("1"))
                            {
                                MonthHItem.V_Gbn = "입고";
                                totalStuffRoll += ConvertInt(MonthHItem.V_Roll);
                                totalStuffQty += ConvertInt(MonthHItem.V_Qty);
                                grdMergeMonth_V.Items.Add(MonthHItem);
                            }
                            else if (MonthHItem.V_Gbn.Equals("2"))
                            {
                                MonthHItem.V_Gbn = "출고";
                                totalOutRoll += ConvertInt(MonthHItem.V_Roll);
                                totalOutQty += ConvertInt(MonthHItem.V_Qty);
                                grdMergeMonth_V.Items.Add(MonthHItem);

                            }
                            else if (MonthHItem.V_Gbn.Equals("3"))
                            {
                                MonthHItem.V_Color1 = true;
                                MonthHItem.V_Gbn = string.Empty;
                                MonthHItem.V_Article = "거래처 계";
                                MonthHItem.V_CustomName = string.Empty;
                                grdMergeMonth_V.Items.Add(MonthHItem);
                            }

                        }

                        if (grdMergeMonth_V.Items.Count > 0)
                        {
                            var MonthVtotal = new Win_ord_InOutSum_Total_QView
                            {
                                V_TotalOutRoll = stringFormatN0(totalOutRoll),
                                V_TotalOutQty = stringFormatN0(totalOutQty),
                                V_TotalStuffRoll = stringFormatN0(totalStuffRoll),
                                V_TotalStuffQty = stringFormatN0(totalStuffQty),
                            };

                            dgdMonthVtotal.Items.Add(MonthVtotal);
                        }

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

        #region 월별집계 최근 3개월 가로집계
        // 월별집계 (가로) (최근 3개월)
        private void FillGrid_Month_H()
        {

            try
            {

                grdMergeMonth_H.Items.Clear();
                dgdMonthHOutTotal.Items.Clear();
                dgdMonthHStuffTotal.Items.Clear();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", 1);
                sqlParameter.Add("SDate", !lib.IsDatePickerNull(dtpFromDate) ? lib.ConvertDate(dtpFromDate) : "");
                sqlParameter.Add("EDate", !lib.IsDatePickerNull(dtpToDate) ? lib.ConvertDate(dtpToDate) : "");

                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag != null ? txtBuyerArticleNoSrh.Tag.ToString() : string.Empty : string.Empty);

                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : "" : "");

                sqlParameter.Add("ChkOrder", chkOrderIDSrh.IsChecked == true ? rbnOrderNOSrh.IsChecked == true ? 1 : 2 : 0);
                sqlParameter.Add("Order", chkOrderIDSrh.IsChecked == true ? !string.IsNullOrEmpty(txtOrderIDSrh.Text) ? txtOrderIDSrh.Text : "" : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Outware_sInOutwareSum_MonthSpread3", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = null;
                    dt = ds.Tables[0];
                    SpreadMonthDataTable = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }
                    else
                    {
                        //grdMerge_Month_H.RowDefinitions.Clear();

                        DataRowCollection drc = dt.Rows;

                        int i = 0;
                        int totalOutRoll = 0;
                        int totalOutQty = 0;
                        int totalStuffRoll = 0;
                        int totalStuffQty = 0;

                        int totalBaseMonthOutRoll = 0;
                        int totalBaseMonthOutQty = 0;
                        int totalBaseMonthStuffRoll = 0;
                        int totalBaseMonthStuffQty = 0;

                        int totalAdd1MonthOutRoll = 0;
                        int totalAdd1MonthOutQty = 0;
                        int totalAdd1MonthStuffRoll = 0;
                        int totalAdd1MonthStuffQty = 0;

                        int totalAdd2MonthOutRoll = 0;
                        int totalAdd2MonthOutQty = 0;
                        int totalAdd2MonthStuffRoll = 0;
                        int totalAdd2MonthStuffQty = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var MonthHItem = new Win_ord_InOutSum_QView
                            {
                                H_NUM = i,
                                H_Gbn = dr["Gbn"].ToString(),
                                H_CustomName = dr["KCustom"].ToString(),
                                H_Article = dr["Article"].ToString(),
                                H_UnitClssName = dr["UnitClssName"].ToString(),

                                H_TotalMonthRoll = stringFormatN0(dr["TotalRoll"]),
                                H_TotalMonthQty = stringFormatN0(dr["TotalQty"]),
                                H_TotalMonthAmount = stringFormatN0(dr["TotalAmount"]),
                                H_BaseMonthRoll = stringFormatN0(dr["BaseMonthRoll"]),
                                H_BaseMonthQty = stringFormatN0(dr["BaseMonthQty"]),
                                H_BaseMonthAmount = stringFormatN0(dr["BaseMonthAmount"]),
                                H_Add1MonthRoll = stringFormatN0(dr["Add1MonthRoll"]),
                                H_Add1MonthQty = stringFormatN0(dr["Add1MonthQty"]),
                                H_Add1MonthAmount = stringFormatN0(dr["Add1MonthAmount"]),
                                H_Add2MonthRoll = stringFormatN0(dr["Add2MonthRoll"]),
                                H_Add2MonthQty = stringFormatN0(dr["Add2MonthQty"]),
                                H_Add2MonthAmount = stringFormatN0(dr["Add2MonthAmount"]),
                            };

                            if (MonthHItem.H_Gbn.Equals("1"))       //화면디자인이 출고가 먼저 나와야 하기에..
                            {
                                MonthHItem.H_Gbn = "출고";
                                totalOutRoll += ConvertInt(MonthHItem.H_TotalMonthRoll);
                                totalOutQty += ConvertInt(MonthHItem.H_TotalMonthQty);

                                totalBaseMonthOutRoll += ConvertInt(MonthHItem.H_BaseMonthRoll);
                                totalBaseMonthOutQty += ConvertInt(MonthHItem.H_BaseMonthQty);

                                totalAdd1MonthOutRoll += ConvertInt(MonthHItem.H_Add1MonthRoll);
                                totalAdd1MonthOutQty += ConvertInt(MonthHItem.H_Add1MonthQty);
                                totalAdd2MonthOutRoll += ConvertInt(MonthHItem.H_Add2MonthRoll);
                                totalAdd2MonthOutQty += ConvertInt(MonthHItem.H_Add2MonthQty);
                                grdMergeMonth_H.Items.Add(MonthHItem);
                            }
                            else if (MonthHItem.H_Gbn.Equals("2"))
                            {
                                MonthHItem.H_Gbn = "입고";
                                totalStuffRoll += ConvertInt(MonthHItem.H_TotalMonthRoll);
                                totalStuffQty += ConvertInt(MonthHItem.H_TotalMonthQty);

                                totalBaseMonthStuffRoll += ConvertInt(MonthHItem.H_BaseMonthRoll);
                                totalBaseMonthStuffQty += ConvertInt(MonthHItem.H_BaseMonthQty);

                                totalAdd1MonthStuffRoll += ConvertInt(MonthHItem.H_Add1MonthRoll);
                                totalAdd1MonthStuffQty += ConvertInt(MonthHItem.H_Add1MonthQty);
                                totalAdd2MonthStuffRoll += ConvertInt(MonthHItem.H_Add2MonthRoll);
                                totalAdd2MonthStuffQty += ConvertInt(MonthHItem.H_Add2MonthQty);
                                grdMergeMonth_H.Items.Add(MonthHItem);

                            }
                            else if (MonthHItem.H_Gbn.Equals("3"))
                            {
                                MonthHItem.H_Color1 = true;
                                MonthHItem.H_Gbn = string.Empty;
                                MonthHItem.H_CustomName = string.Empty;
                                grdMergeMonth_H.Items.Add(MonthHItem);
                            }
                            else if (MonthHItem.H_Gbn.Equals("4"))
                            {
                                MonthHItem.H_Color1 = true;
                                MonthHItem.H_Gbn = string.Empty;
                                MonthHItem.H_CustomName = string.Empty;
                                grdMergeMonth_H.Items.Add(MonthHItem);
                            }
                        }

                        if (grdMergeMonth_H.Items.Count > 0)
                        {
                            var MonthHTotalOut = new Win_ord_InOutSum_Total_QView
                            {
                                H_TotalOutRoll = stringFormatN0(totalOutRoll),
                                H_TotalOutQty = stringFormatN0(totalOutQty),

                                H_TotalBaseOutRoll = stringFormatN0(totalBaseMonthOutRoll),
                                H_TotalBaseOutQty = stringFormatN0(totalBaseMonthOutQty),

                                H_TotalAdd1OutRoll = stringFormatN0(totalAdd1MonthOutRoll),
                                H_TotalAdd1OutQty = stringFormatN0(totalAdd1MonthOutQty),

                                H_TotalAdd2OutRoll = stringFormatN0(totalAdd2MonthOutRoll),
                                H_TotalAdd2OutQty = stringFormatN0(totalAdd2MonthOutQty),

                            };

                            var MonthHTotalStuff = new Win_ord_InOutSum_Total_QView
                            {
                                H_TotalStuffRoll = stringFormatN0(totalStuffRoll),
                                H_TotalStuffQty = stringFormatN0(totalStuffQty),

                                H_TotalBaseStuffRoll = stringFormatN0(totalBaseMonthStuffRoll),
                                H_TotalBaseStuffQty = stringFormatN0(totalBaseMonthStuffQty),

                                H_TotalAdd1StuffRoll = stringFormatN0(totalAdd1MonthStuffRoll),
                                H_TotalAdd1StuffQty = stringFormatN0(totalAdd1MonthStuffRoll),

                                H_TotalAdd2StuffRoll = stringFormatN0(totalAdd2MonthStuffRoll),
                                H_TotalAdd2StuffQty = stringFormatN0(totalAdd2MonthStuffQty),
                            };

                            dgdMonthHOutTotal.Items.Add(MonthHTotalOut);
                            dgdMonthHStuffTotal.Items.Add(MonthHTotalStuff);
                        }
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


        #region 월별 가로집계 셀렉션 체인지 이벤트
        // 탬 컨트롤 셀렉션 체인지 이벤트.
        private void tabconGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string sNowTI = ((sender as TabControl).SelectedItem as TabItem).Header as string;

            switch (sNowTI)
            {
                case "기간집계":
                    txtblMessage.Visibility = Visibility.Hidden;
                    dtpFromDate.IsEnabled = true;
                    dtpToDate.IsEnabled = true;
                    break;
                case "일일집계":
                    txtblMessage.Visibility = Visibility.Hidden;
                    dtpFromDate.IsEnabled = true;
                    dtpToDate.IsEnabled = true;
                    break;
                case "월별집계(세로)":
                    txtblMessage.Visibility = Visibility.Hidden;
                    dtpFromDate.IsEnabled = true;
                    dtpToDate.IsEnabled = true;
                    break;
                case "월별집계(가로)":
                    txtblMessage.Visibility = Visibility.Visible;
                    dtpFromDate.IsEnabled = true;
                    dtpToDate.IsEnabled = true;
                    break;
                default: return;
            }
        }


        #endregion


        //닫기 버튼 클릭.
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


        #region 엑셀

        // 엑셀 버튼 클릭.
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            string sNowTI = (tabconGrid.SelectedItem as TabItem).Header as string;
            string Listname1 = string.Empty;
            string Listname2 = string.Empty;
            DataTable choicedt = null;
            Lib lib2 = new Lib();

            if (PeriodDataTable != null)
            {
                switch (sNowTI)
                {
                    case "기간집계":
                        if (PeriodDataTable.Rows.Count < 1)
                        {
                            MessageBox.Show("먼저 기간집계를 검색해 주세요.");
                            return;
                        }
                        Listname1 = "기간집계";
                        Listname2 = "PeriodData";
                        choicedt = PeriodDataTable;
                        break;
                    case "일일집계":
                        if (DaysDataTable.Rows.Count < 1)
                        {
                            MessageBox.Show("먼저 일일집계를 검색해 주세요.");
                            return;
                        }
                        Listname1 = "일일집계";
                        Listname2 = "DayData";
                        choicedt = DaysDataTable;
                        break;
                    case "월별집계(세로)":
                        if (MonthDataTable.Rows.Count < 1)
                        {
                            MessageBox.Show("먼저 월별(세로)집계를 검색해 주세요.");
                            return;
                        }
                        Listname1 = "월(세로)집계";
                        Listname2 = "MonthData";
                        choicedt = MonthDataTable;
                        break;
                    case "월별집계(가로)":
                        if (SpreadMonthDataTable.Rows.Count < 1)
                        {
                            MessageBox.Show("먼저 월별(가로)집계를 검색해 주세요.");
                            return;
                        }
                        Listname1 = "월(가로)집계";
                        Listname2 = "SpreadMonthData";
                        choicedt = SpreadMonthDataTable;
                        break;
                    default: return;
                }

                string Name = string.Empty;

                string[] lst = new string[4];
                lst[0] = Listname1;
                lst[2] = Listname2;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

                ExpExc.ShowDialog();

                // 어쨋든 머든 여기서 dt로 만들어서 주면 된다는 거네.
                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(Listname2))
                    {
                        Name = Listname2;
                        if (lib2.GenerateExcel(choicedt, Name))
                        {
                            DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                            lib2.excel.Visible = true;
                            lib2.ReleaseExcelObject(lib2.excel);
                        }
                    }
                    else
                    {
                        if (choicedt != null)
                        {
                            choicedt.Clear();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("엑설로 변환할 자료가 없습니다.");
            }

            lib2 = null;
        }





        #endregion





        private void txtCustomIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtCustomIDSrh, 0, "");
            }
        }

        private void txtArticleIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticleIDSrh, 77, "");
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


        private void txtBuyerArticleNoSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtBuyerArticleNoSrh, 76, "");
            }
        }

        private void btnBuyerArticleNoSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtBuyerArticleNoSrh, 76, "");
        }

        private void CommonControl_Click(object sender, MouseButtonEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private void CommonControl_Click(object sender, RoutedEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private void rbnOrderNOSrh_Click(object sender, RoutedEventArgs e)
        {
            tblOrderIDSrh.Text = "발주번호";
            //dtcOrderID.Visibility = Visibility.Hidden;
            //dtcOrderNO.Visibility = Visibility.Visible;
        }

        private void rbnOrderIDSrh_Click(object sender, RoutedEventArgs e)
        {
            tblOrderIDSrh.Text = "관리번호";
            //dtcOrderID.Visibility = Visibility.Visible;
            //dtcOrderNO.Visibility = Visibility.Hidden;
        }



        // 천 단위 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천 단위 콤마, 소수점 두자리
        private string stringFormatN2(object obj)
        {
            return string.Format("{0:N2}", obj);
        }

        // Int로 변환
        private int ConvertInt(string str)
        {
            int result = 0;
            int chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");

                if (Int32.TryParse(str, out chkInt) == true)
                {
                    result = Int32.Parse(str);
                }
            }

            return result;
        }

    }





    /// <summary>
    /// /////////////////////////////////////////////////////////////////////
    /// </summary>


    public class MonthChange
    {
        //SpreadMonth 월 기간 확인용
        public string H_MON1 { get; set; }
        public string H_MON2 { get; set; }
        public string H_MON3 { get; set; }

    }



    class Win_ord_InOutSum_QView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }

        // 조회 - 기간집계용 ( P_ (Period))
        public int P_NUM { get; set; }
        public string P_cls { get; set; }
        public string P_Gbn { get; set; }
        public string P_IODate { get; set; }
        public string P_CustomID { get; set; }
        public string P_CustomName { get; set; }

        public string P_Sabun { get; set; }

        public string P_BuyerArticleNo { get; set; }
        public string P_ArticleID { get; set; }
        public string P_Article { get; set; }
        public string P_Roll { get; set; }
        public string P_Qty { get; set; }
        public string P_UnitClss { get; set; }

        public string P_UnitClssName { get; set; }
        public string P_UnitPrice { get; set; }
        public string P_PriceClss { get; set; }
        public string P_PriceClssName { get; set; }
        public string P_Amount { get; set; }

        public string P_VatAmount { get; set; }
        public string P_TotAmount { get; set; }
        public string P_CustomRate { get; set; }
        public string P_CustomRateOrder { get; set; }
        public bool P_Color1 { get; set; } = false;
        public bool P_Color2 { get; set; } = false;



        // 조회 - 일별집계용 ( D_ (Day))
        public int D_NUM { get; set; }
        public string D_cls { get; set; }
        public string D_Gbn { get; set; }
        public string D_IODate { get; set; }
        public string D_CustomID { get; set; }
        public string D_CustomName { get; set; }

        public string D_BuyerArticleNo { get; set; }
        public string D_ArticleID { get; set; }
        public string D_Article { get; set; }
        public string D_Roll { get; set; }
        public string D_Qty { get; set; }
        public string D_UnitClss { get; set; }

        public string D_Sabun { get; set; }

        public string D_UnitClssName { get; set; }
        public string D_UnitPrice { get; set; }
        public string D_PriceClss { get; set; }
        public string D_PriceClssName { get; set; }
        public string D_Amount { get; set; }

        public string D_VatAmount { get; set; }
        public string D_TotAmount { get; set; }
        public string D_CustomRate { get; set; }
        public string D_CustomRateOrder { get; set; }
        public bool D_Color1 { get; set; } = false;
        public bool D_Color2 { get; set; } = false;


        // 조회 - 월별집계용 _V ( V_ (V_Month))
        public int V_NUM { get; set; }
        public string V_cls { get; set; }
        public string V_Gbn { get; set; }
        public string V_IODate { get; set; }
        public string V_CustomID { get; set; }
        public string V_CustomName { get; set; }

        public string V_BuyerArticleNo { get; set; }
        public string V_ArticleID { get; set; }
        public string V_Article { get; set; }
        public string V_Roll { get; set; }
        public string V_Qty { get; set; }
        public string V_UnitClss { get; set; }

        public string V_Sabun { get; set; }

        public string V_UnitClssName { get; set; }
        public string V_UnitPrice { get; set; }
        public string V_PriceClss { get; set; }
        public string V_PriceClssName { get; set; }
        public string V_Amount { get; set; }

        public string V_VatAmount { get; set; }
        public string V_TotAmount { get; set; }
        public string V_CustomRate { get; set; }
        public string V_CustomRateOrder { get; set; }
        public string V_RN { get; set; }

        public bool V_Color1 { get; set; } = false;
        public bool V_Color2 { get; set; } = false;



        // 조회 - 월별집계용 _H ( H_ (H_Month))
        public int H_NUM { get; set; }
        public string H_cls { get; set; }
        public string H_Gbn { get; set; }
        public string H_CustomID { get; set; }
        public string H_CustomName { get; set; }

        public string H_BuyerArticleNo { get; set; }
        public string H_ArticleID { get; set; }
        public string H_Article { get; set; }
        public string H_UnitClss { get; set; }
        public string H_UnitClssName { get; set; }
        public string H_UnitPrice { get; set; }

        public string H_Sabun { get; set; }

        public string H_PriceClss { get; set; }
        public string H_PriceClssName { get; set; }
        public string H_YYYYMM1 { get; set; }
        public string H_YYYYMM2 { get; set; }
        public string H_YYYYMM3 { get; set; }

        public string H_YYYYMM4 { get; set; }
        public string H_YYYYMM5 { get; set; }
        public string H_YYYYMM6 { get; set; }
        public string H_YYYYMM7 { get; set; }
        public string H_YYYYMM8 { get; set; }

        public string H_YYYYMM9 { get; set; }
        public string H_YYYYMM10 { get; set; }
        public string H_roll10 { get; set; }
        public string H_Qty10 { get; set; }
        public string H_Amount10 { get; set; }

        public string H_VatAmount10 { get; set; }
        public string H_YYYYMM11 { get; set; }
        public string H_roll11 { get; set; }
        public string H_Qty11 { get; set; }
        public string H_Amount11 { get; set; }

        public string H_VatAmount11 { get; set; }
        public string H_YYYYMM12 { get; set; }
        public string H_roll12 { get; set; }
        public string H_Qty12 { get; set; }
        public string H_Amount12 { get; set; }

        public string H_VatAmount12 { get; set; }
        public string H_YYYYMM13 { get; set; }
        public string H_roll13 { get; set; }
        public string H_Qty13 { get; set; }
        public string H_Amount13 { get; set; }

        public string H_VatAmount13 { get; set; }
        public string H_RN { get; set; }
        public string H_CustomRate { get; set; }
        public string H_CustomAmount { get; set; }
        public string H_AllTotalAmount { get; set; }

        public string H_TotalMonthQty { get; set; }
        public string H_TotalMonthRoll { get; set; }
        public string H_TotalMonthAmount { get; set; }
        public string H_BaseMonthQty { get; set; }
        public string H_BaseMonthRoll { get; set; }
        public string H_BaseMonthAmount { get; set; }
        public string H_Add1MonthQty { get; set; }
        public string H_Add1MonthRoll { get; set; }
        public string H_Add1MonthAmount { get; set; }
        public string H_Add2MonthQty { get; set; }
        public string H_Add2MonthRoll { get; set; }
        public string H_Add2MonthAmount { get; set; }
        public bool H_Color1 { get; set; } = false;
        public bool H_Color2 { get; set; } = false;


        public List<P_listmodel> P_listmodel { get; set; }
        public List<D_gbnmodel> D_gbnmodel { get; set; }
        public List<V_gbnmodel> V_gbnmodel { get; set; }
        public List<H_custommodel> H_custommodel { get; set; }


    }

    public class D_gbnmodel
    {
        public string D_Gbn { get; set; }
        public string D_YesColor { get; set; }

        public List<D_custommodel> D_custommodel { get; set; }
    }

    public class V_gbnmodel
    {
        public string V_Gbn { get; set; }
        public List<V_custommodel> V_custommodel { get; set; }
    }



    public class D_custommodel
    {
        public string D_CustomName { get; set; }
        public List<D_listmodel> D_listmodel { get; set; }
    }

    public class V_custommodel
    {
        public string V_CustomName { get; set; }
        public string V_YesColor { get; set; }
        public List<V_listmodel> V_listmodel { get; set; }
    }

    public class H_custommodel
    {
        public string H_CustomName { get; set; }
        public List<H_listmodel> H_listmodel { get; set; }
    }



    public class D_listmodel
    {
        public string D_ArticleID { get; set; }
        public string D_Article { get; set; }
        public string D_Roll { get; set; }
        public string D_Qty { get; set; }
        public string D_UnitClssName { get; set; }
        public string D_PriceClssName { get; set; }

        public string D_VatAmount { get; set; }
        public string D_TotAmount { get; set; }
        public string D_CustomRate { get; set; }

    }

    public class P_listmodel
    {
        public string P_ArticleID { get; set; }
        public string P_Article { get; set; }
        public string P_Roll { get; set; }
        public string P_Qty { get; set; }
        public string P_UnitClssName { get; set; }
        public string P_CustomRate { get; set; }

        public string P_YesColor { get; set; }

    }

    public class V_listmodel
    {
        public string V_ArticleID { get; set; }
        public string V_Article { get; set; }
        public string V_Roll { get; set; }
        public string V_Qty { get; set; }
        public string V_UnitClssName { get; set; }
        public string V_CustomRate { get; set; }
    }



    public class H_listmodel
    {
        public string H_ArticleID { get; set; }
        public string H_Article { get; set; }
        public string H_UnitClssName { get; set; }
        public string H_PriceClssName { get; set; }
        public string H_roll10 { get; set; }
        public string H_Qty10 { get; set; }
        public string H_CustomRate { get; set; }

        public string H_roll11 { get; set; }
        public string H_Qty11 { get; set; }
        public string H_roll12 { get; set; }
        public string H_Qty12 { get; set; }
        public string H_roll13 { get; set; }
        public string H_Qty13 { get; set; }

    }

    public class Win_ord_InOutSum_Total_QView : BaseView
    {
        public string P_TotalOutRoll { get; set; }
        public string P_TotalOutQty { get; set; }
        public string P_TotalStuffRoll { get; set; }
        public string P_TotalStuffQty { get; set; }

        public string D_TotalOutRoll { get; set; }
        public string D_TotalOutQty { get; set; }
        public string D_TotalOutAmount { get; set; }
        public string D_TotalOutVatAmount { get; set; }
        public string D_TotalOutPrice { get; set; }
        public string D_TotalStuffRoll { get; set; }
        public string D_TotalStuffQty { get; set; }
        public string D_TotalStuffAmount { get; set; }
        public string D_TotalStuffVatAmount { get; set; }
        public string D_TotalStuffPrice { get; set; }
        public bool D_Color1 { get; set; } = false;
        public bool D_Color2 { get; set; } = false;

        public string V_TotalOutRoll { get; set; }
        public string V_TotalOutQty { get; set; }
        public string V_TotalStuffRoll { get; set; }
        public string V_TotalStuffQty { get; set; }
        public bool V_Color1 { get; set; } = false;
        public bool V_Color2 { get; set; } = false;

        public string H_TotalOutRoll { get; set; }
        public string H_TotalOutQty { get; set; }
        public string H_TotalStuffRoll { get; set; }
        public string H_TotalStuffQty { get; set; }

        public string H_TotalBaseOutRoll { get; set; }
        public string H_TotalBaseOutQty { get; set; }
        public string H_TotalBaseStuffRoll { get; set; }
        public string H_TotalBaseStuffQty { get; set; }

        public string H_TotalAdd1OutRoll { get; set; }
        public string H_TotalAdd1OutQty { get; set; }
        public string H_TotalAdd1StuffRoll { get; set; }
        public string H_TotalAdd1StuffQty { get; set; }

        public string H_TotalAdd2OutRoll { get; set; }
        public string H_TotalAdd2OutQty { get; set; }
        public string H_TotalAdd2StuffRoll { get; set; }
        public string H_TotalAdd2StuffQty { get; set; }
        public bool H_Color1 { get; set; } = false;
        public bool H_Color2 { get; set; } = false;
    }


}

