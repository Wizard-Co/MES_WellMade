using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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
    /// Win_ord_OutWare_Product_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_OutWare_Product_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Win_ord_OutWare_Product_QView wopqv = new Win_ord_OutWare_Product_QView();
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        public Win_ord_OutWare_Product_Q()
        {
            InitializeComponent();
            this.DataContext = wopqv;
        }

        private void Window_OutwareProduct_Loaded(object sender, RoutedEventArgs e)
        {

            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            First_Step();
            ComboBoxSetting();
        }

        #region 시작 첫 스텝 // 날짜용 버튼 // ComboSetting // 조회용 체크박스 이벤트
        // 시작 첫 단추.
        private void First_Step()
        {
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            // 시작 지정 및 사용불가 설정.
            chkOutwareDay.IsChecked = true;

            // 폼 하단 안쓰는 버튼들 가리기.
            chkBuyCustom.Visibility = Visibility.Hidden;
            tbkInsertSheetNO.Visibility = Visibility.Hidden;
            txtBuyCustom.Visibility = Visibility.Hidden;
            txtInsertSheetNO.Visibility = Visibility.Hidden;
            btnBuyCustom.Visibility = Visibility.Hidden;

        }

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpToDate);

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
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpFromDate);

            dtpFromDate.SelectedDate = SearchDate[0];
            dtpToDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpToDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }



        // 콤보박스 두개 목록 불러오기.  (제품그룹, 출고구분)
        private void ComboBoxSetting()
        {
            //cboArticleGroup.Items.Clear();
            //cboOutClss.Items.Clear();

            //ObservableCollection<CodeView> cbArticleGroup = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            ////ObservableCollection<CodeView> cbOutClss = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "OCD", "Y", "", "");

            //this.cboArticleGroup.ItemsSource = cbArticleGroup;
            //this.cboArticleGroup.DisplayMemberPath = "code_name";
            //this.cboArticleGroup.SelectedValuePath = "code_id";
            //this.cboArticleGroup.SelectedIndex = 3;  //제품이보이게



            ////this.cboOutClss.ItemsSource = cbOutClss;
            ////this.cboOutClss.DisplayMemberPath = "code_id_plus_code_name";
            ////this.cboOutClss.SelectedValuePath = "code_id";
            ////this.cboOutClss.SelectedIndex = 0;

            //List<string> cbOutClss = new List<string>();
            //cbOutClss.Add("01.제품정상출고");
            //cbOutClss.Add("11.제품출고반품");
            //cbOutClss.Add("08.예외출고");
            //cbOutClss.Add("18.예외출고반품");

            //ObservableCollection<CodeView> cboOutClass = ComboBoxUtil.Instance.Direct_SetComboBox(cbOutClss);
            //this.cboOutClss.ItemsSource = cboOutClass;
            //this.cboOutClss.DisplayMemberPath = "code_name";
            //this.cboOutClss.SelectedValuePath = "code_id";
            //this.cboOutClss.SelectedIndex = 0;

        }

        //출고일자(날짜) 체크
        private void chkOutwareDay_Click(object sender, RoutedEventArgs e)
        {
            if (chkOutwareDay.IsChecked == true)
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
        //출고일자(날짜) 체크
        private void chkOutwareDay_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkOutwareDay.IsChecked == true)
            {
                chkOutwareDay.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkOutwareDay.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }

        private void rbnOrderNOSrh_Click(object sender, RoutedEventArgs e)
        {
            tblOrderIDSrh.Text = "발주번호";
            dtcOrderID.Visibility = Visibility.Hidden;
            dtcOrderNO.Visibility = Visibility.Visible;
        }

        private void rbnOrderIDSrh_Click(object sender, RoutedEventArgs e)
        {
            tblOrderIDSrh.Text = "관리번호";
            dtcOrderID.Visibility = Visibility.Visible;
            dtcOrderNO.Visibility = Visibility.Hidden;
        }

        #endregion


        #region 플러스 파인더
        //플러스 파인더





        #endregion


        #region 조회 // 조회용 프로시저

        // 검색버튼 클릭. (조회)
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            if (lib.DatePickerCheck(dtpFromDate, dtpToDate, chkOutwareDay))
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
                //로직
                FillGrid();
            }), System.Windows.Threading.DispatcherPriority.Background);

            btnSearch.IsEnabled = true;
        }

        private void FillGrid()
        {


            //출고구분
            //string outclssGBN = string.Empty;
            //dgdTotal.Items.Clear();

            //if (cboOutClss.SelectedIndex == 0) { outclssGBN = "01"; }     //제품정상출고
            //else if (cboOutClss.SelectedIndex == 1) { outclssGBN = "11"; } //제품출고반품
            //else if (cboOutClss.SelectedIndex == 2) { outclssGBN = "08"; } //예외출고
            //else if (cboOutClss.SelectedIndex == 3) { outclssGBN = "18"; } //예외출고반품


            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", chkOutwareDay.IsChecked == true ? 1 : 0);
                sqlParameter.Add("SDate", chkOutwareDay.IsChecked == true ? !lib.IsDatePickerNull(dtpFromDate) ? lib.ConvertDate(dtpFromDate) : string.Empty : string.Empty);
                sqlParameter.Add("EDate", chkOutwareDay.IsChecked == true ? !lib.IsDatePickerNull(dtpToDate) ? lib.ConvertDate(dtpToDate) : string.Empty : string.Empty);

                sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag != null ? txtCustomIDSrh.Tag.ToString() : string.Empty : string.Empty);

                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag != null ? txtBuyerArticleNoSrh.Tag.ToString() : string.Empty : string.Empty);

                sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag != null ? txtArticleIDSrh.Tag.ToString() : string.Empty : string.Empty);

                sqlParameter.Add("ChkOrderID", chkOrderIDSrh.IsChecked == true ? rbnOrderIDSrh.IsChecked == true ? 1 : 2 : 0);
                sqlParameter.Add("OrderID", txtOrderIDSrh.Text);

                //sqlParameter.Add("chkArticleGrpID", chkArticleGrpID);
                //sqlParameter.Add("sArticleGrpID", cboArticleGroup.SelectedValue.ToString());
                //sqlParameter.Add("sProductYN", "Y"); // 제품여부 Y인데 빈값넣으니까 됐어 왜지???

                //sqlParameter.Add("chkOutClss", int_chkOutClss);
                //sqlParameter.Add("OutClss", outclssGBN); //cboOutClss.SelectedValue.ToString()
                //sqlParameter.Add("nMainItem", interestitems);

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Outware_sOutwareProduct", sqlParameter, true, "R");

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
                        dgdOutware.Items.Clear();
                        dgdTotal.Items.Clear();

                        DataRowCollection drc = dt.Rows;

                        int i = 0;
                        int count = 0;
                        string totalOutQty = string.Empty;
                        string totalOutPrice = string.Empty;
                        foreach (DataRow dr in drc)
                        {

                            i++;
                            var OutWareInfo = new Win_ord_OutWare_Product_QView
                            {
                                NUM = i,
                                Depth = dr["Depth"].ToString(),
                                OutDate = lib.DateTypeHyphen(dr["OutDate"].ToString()),
                                KCustom = dr["KCustom"].ToString(),
                                OrderID = dr["OrderID"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                OutClssName = dr["OutClssName"].ToString(),
                                Model = dr["Model"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                LabelID = dr["LabelID"].ToString(),
                                OutQty = stringFormatN0(dr["OutQty"]),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                OutPrice = stringFormatN0(dr["OutPrice"]),
                                OutwareID = dr["OutWareID"].ToString(),
                                Remark = dr["remark"].ToString()

                            };


                            if (OutWareInfo.Depth == "0")
                            {
                                count++;
                                dgdOutware.Items.Add(OutWareInfo);
                            }
                            else if (OutWareInfo.Depth == "1")
                            {
                                OutWareInfo.Color1 = true;
                                OutWareInfo.OutDate = string.Empty;
                                dgdOutware.Items.Add(OutWareInfo);
                            }
                            else if (OutWareInfo.Depth == "2")
                            {
                                OutWareInfo.Color2 = true;
                                OutWareInfo.OutDate = string.Empty;
                                totalOutQty = OutWareInfo.OutQty;
                                totalOutPrice = OutWareInfo.OutPrice;
                                dgdOutware.Items.Add(OutWareInfo);
                            }
                        }

                        if (dgdOutware.Items.Count > 0)
                        {
                            var OutWareTotalInfo = new Win_ord_OutWare_Product_QView_Total
                            {
                                TotalCount = count,
                                TotalOutPrice = totalOutPrice,
                                TotalOutQty = totalOutQty,
                            };

                            dgdTotal.Items.Add(OutWareTotalInfo);
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


        #region 엑셀
        // 엑셀 버튼 클릭.
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdOutware.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            Lib lib2 = new Lib();
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "메인 그리드";
            lst[2] = dgdOutware.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdOutware.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    //MessageBox.Show("대분류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib2.DataGridToDTinHidden(dgdOutware);
                    else
                        dt = lib2.DataGirdToDataTable(dgdOutware);

                    Name = dgdOutware.Name;

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

        //닫기 버튼 클릭./
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


        //정렬.
        private void btnMultiSort_Click(object sender, RoutedEventArgs e)
        {
            PopUp.MultiLevelSort MLS = new PopUp.MultiLevelSort(dgdOutware);
            MLS.ShowDialog();

            if (MLS.DialogResult.HasValue)
            {
                string targetSortProperty = string.Empty;
                int targetColIndex;
                dgdOutware.Items.SortDescriptions.Clear();

                for (int x = 0; x < MLS.ColName.Count; x++)
                {
                    targetSortProperty = MLS.SortingProperty[x];
                    targetColIndex = MLS.ColIndex[x];
                    var targetCol = dgdOutware.Columns[targetColIndex];

                    if (targetSortProperty == "UP")
                    {
                        dgdOutware.Items.SortDescriptions.Add(new SortDescription(targetCol.SortMemberPath, ListSortDirection.Ascending));
                        targetCol.SortDirection = ListSortDirection.Ascending;
                    }
                    else
                    {
                        dgdOutware.Items.SortDescriptions.Add(new SortDescription(targetCol.SortMemberPath, ListSortDirection.Descending));
                        targetCol.SortDirection = ListSortDirection.Descending;
                    }
                }
                dgdOutware.Refresh();
            }
        }



        // 사용자 편의. 엔터키로 플러스파인더 호출.
        //검색조건 - 거래처 - 키다운
        private void txtCustomIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                pf.ReturnCode(txtCustomIDSrh, 0, "");
        }

        //검색조건 - 거래처 - 버튼
        private void btnCustomIDSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomIDSrh, 0, "");

        }

        //검색조건 - 품번 - 키다운
        private void txtBuyerArticleNoSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtBuyerArticleNoSrh, 7071, "");
            }
        }

        //검색조건 - 품번 - 버튼
        private void btnBuyerArticleNoSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtBuyerArticleNoSrh, 76, "");

        }

        //검색조건 - 품명 - 키다운
        private void txtArticleIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticleIDSrh, 77, "");
            }
        }
        //검색조건 - 품명 - 버튼
        private void btnArticleIDSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticleIDSrh, 77, "");
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

        #region 기타 메서드 모음



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

        // 데이터피커 포맷으로 변경
        private string DatePickerFormat(string str)
        {
            string result = "";

            if (str.Length == 8)
            {
                if (!str.Trim().Equals(""))
                {
                    result = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
                }
            }

            return result;
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

        // 소수로 변환 가능한지 체크 이벤트
        private bool CheckConvertDouble(string str)
        {
            bool flag = false;
            double chkDouble = 0;

            if (!str.Trim().Equals(""))
            {
                if (Double.TryParse(str, out chkDouble) == true)
                {
                    flag = true;
                }
            }

            return flag;
        }

        // 숫자로 변환 가능한지 체크 이벤트
        private bool CheckConvertInt(string str)
        {
            bool flag = false;
            int chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Trim().Replace(",", "");

                if (Int32.TryParse(str, out chkInt) == true)
                {
                    flag = true;
                }
            }

            return flag;
        }

        // 소수로 변환
        private double ConvertDouble(string str)
        {
            double result = 0;
            double chkDouble = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");

                if (Double.TryParse(str, out chkDouble) == true)
                {
                    result = Double.Parse(str);
                }
            }

            return result;
        }






        #endregion

        private void CommonControl_Click(object sender, RoutedEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private void CommonControl_Click(object sender, MouseButtonEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }


    }







    /// <summary>
    /// //////////////////////////////////////////////////////////////////////
    /// </summary>

    class Win_ord_OutWare_Product_QView : BaseView
    {
        //public override string ToString()
        //{
        //    return (this.ReportAllProperties());
        //}

        public ObservableCollection<CodeView> cboTrade { get; set; }

        // 조회 값.    
        public int NUM { get; set; }
        public string Depth { get; set; }
        public string OutwareID { get; set; }
        public string OutDate { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }

        public string OrderNo { get; set; }
        public string OrderID { get; set; }
        public string OutCustom { get; set; }
        public string Model { get; set; }

        public string BuyerArticleNo { get; set; }
        public string Article { get; set; }
        public string ArticleID { get; set; }
        public string Sabun { get; set; }

        public string WorkName { get; set; }

        public string OrderQty { get; set; }
        public string UnitClss { get; set; }
        public string UnitClssName { get; set; }
        public string LabelID { get; set; }
        public string LabelGubun { get; set; }

        public string FromLocName { get; set; }
        public string TOLocname { get; set; }
        public string OutClssName { get; set; }
        public string OutRoll { get; set; }
        public string OutQty { get; set; }

        public string UnitPrice { get; set; }
        public string OutPrice { get; set; }
        public string Amount { get; set; }
        public string VatAmount { get; set; }
        public string TotAmount { get; set; }
        public string Remark { get; set; }

        public bool Color1 { get; set; } = false;
        public bool Color2 { get; set; } = false;
        //순번


        //컬러 칠하기
        public string ColorGreen { get; set; }
        public string ColorRed { get; set; }

        public string LotID { get; set; }
    }

    public class Win_ord_OutWare_Product_QView_Total : BaseView
    {
        public int TotalCount { get; set; }
        public string TotalOutQty { get; set; }
        public string TotalOutPrice { get; set; }
    }
}
