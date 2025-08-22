using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /**************************************************************************************************
   '** System 명 : WizMes_WellMade
   '** Author    : Wizard
   '** 작성자    : 최준호
   '** 내용      : 금형/설비 부품 재고조회
   '** 생성일자  : 2018.10월~2019.01월 사이
   '** 변경일자  : 
   '**------------------------------------------------------------------------------------------------
   ''*************************************************************************************************
   ' 변경일자  , 변경자, 요청자    , 요구사항ID  , 요청 및 작업내용
   '**************************************************************************************************
   ' ex) 2015.11.09, 박진성, 오영      ,S_201510_AFT_03 , 월별집계(가로) 순서 변경 : 합계/10월/9월/8월 순으로
   ' 2019.07.09  최준호 , 최규한  부품명-> 예비품으로 , 출고량 -> 사용량(작업 중)
   '**************************************************************************************************/
    /// <summary>
    /// Win_prd_PartStock_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_PartStock_Q : UserControl
    {
        #region 변수 선언 및 로드
        Lib lib = new Lib();
        public Win_prd_PartStock_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            SetComboBox();

            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        private void SetComboBox()
        {
            List<string[]> lstForUse = new List<string[]>();
            string[] strForUse_1 = { "1", "공용" };
            string[] strForUse_2 = { "2", "설비" };
            string[] strForUse_3 = { "3", "Tool" };
            lstForUse.Add(strForUse_1);
            lstForUse.Add(strForUse_2);
            lstForUse.Add(strForUse_3);

            ObservableCollection<CodeView> ovcDvlYN = ComboBoxUtil.Instance.Direct_SetComboBox(lstForUse);
            this.cboForUseSrh.ItemsSource = ovcDvlYN;
            this.cboForUseSrh.DisplayMemberPath = "code_name";
            this.cboForUseSrh.SelectedValuePath = "code_id";
        }

        #endregion

        #region 날짜관련 이벤트

        //입출일자
        private void lblMcInOutDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMcInOutDate.IsChecked == true) { chkMcInOutDate.IsChecked = false; }
            else { chkMcInOutDate.IsChecked = true; }
        }

        //입출일자
        private void chkMcInOutDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        //입출일자
        private void chkMcInOutDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(dtpEDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(dtpSDate.SelectedDate.Value);

            dtpSDate.SelectedDate = SearchDate[0];
            dtpEDate.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }

        #endregion

        #region 상단 체크 이벤트

        //부품용도
        private void lblForUseSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkForUseSrh.IsChecked == true) { chkForUseSrh.IsChecked = false; }
            else { chkForUseSrh.IsChecked = true; }
        }

        //부품용도
        private void chkForUseSrh_Checked(object sender, RoutedEventArgs e)
        {
            cboForUseSrh.IsEnabled = true;
            cboForUseSrh.Focus();
        }

        //부품용도
        private void chkForUseSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            cboForUseSrh.IsEnabled = false;
        }

        //품명
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true) { chkArticleSrh.IsChecked = false; }
            else { chkArticleSrh.IsChecked = true; }
        }

        //품명
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = true;
            btnPfArticleSrh.IsEnabled = true;
            txtArticleSrh.Focus();
        }

        //품명
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticleSrh.IsEnabled = false;
            btnPfArticleSrh.IsEnabled = false;
        }

        //품명
        private void txtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticleSrh, 13, "");
            }
        }

        //품명
        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSrh, 13, "");
        }

        //거래처
        private void lblCustomSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustomSrh.IsChecked == true) { chkCustomSrh.IsChecked = false; }
            else { chkCustomSrh.IsChecked = true; }
        }

        //거래처
        private void chkCustomSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtCustomSrh.IsEnabled = true;
            btnPfCustomSrh.IsEnabled = true;
            txtCustomSrh.Focus();
        }

        //거래처
        private void chkCustomSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustomSrh.IsEnabled = false;
            btnPfCustomSrh.IsEnabled = false;
        }

        //거래처
        private void txtCustomSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        //거래처
        private void btnPfCustomSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustomSrh, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //적정재고량 미달건 조회
        private void tbkMissSafelyStock_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMissSafelyStock.IsChecked == true) { chkMissSafelyStock.IsChecked = false; }
            else { chkMissSafelyStock.IsChecked = true; }
        }

        #endregion

        #region 우측 상단 버튼 클릭

        //조회
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {

            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                if (dgdMcStock.Items.Count > 0)
                {
                    dgdMcStock.Items.Clear();
                }

                FillGrid();

                if (dgdMcStock.Items.Count > 0)
                {
                    dgdMcStock.SelectedIndex = 0;
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
            

        }

        //닫기
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

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] dgdStr = new string[2];
            dgdStr[0] = "설비부품 재고조회";
            dgdStr[1] = dgdMcStock.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMcStock.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMcStock);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdMcStock);

                    Name = dgdMcStock.Name;
                    if (Lib.Instance.GenerateExcel(dt, Name))
                        Lib.Instance.excel.Visible = true;
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
        }

        #endregion


        // 천단위 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        //실조회
        private void FillGrid()
        {
            try
            {
                dgdTotal.Items.Clear();
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nChkDate", chkMcInOutDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sSDate", chkMcInOutDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sEDate", chkMcInOutDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nChkCustom", chkCustomSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sCustomID", chkCustomSrh.IsChecked == true ? (txtCustomSrh.Tag != null ? txtCustomSrh.Tag.ToString() : txtCustomSrh.Text) : "");
                sqlParameter.Add("nChkArticleID", chkArticleSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sArticleID", chkArticleSrh.IsChecked == true ? (txtArticleSrh.Tag != null ? txtArticleSrh.Tag.ToString() : txtArticleSrh.Text) : "");
                sqlParameter.Add("sForUse", chkForUseSrh.IsChecked == true ? cboForUseSrh.SelectedValue.ToString() : "");
                sqlParameter.Add("sMissSafelyStockQty", chkMissSafelyStock.IsChecked == true ? "Y" : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_mc_sMcPartStockList", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        for (int i = 0; i < drc.Count; i++)
                        {
                            DataRow dr = drc[i];
                            var WinMcStock = new Win_prd_PartStock_Q_CodeView()
                            {
                                Num = (i + 1).ToString(),
                                cls = dr["cls"].ToString(),
                                McPartID = dr["McPartID"].ToString(),
                                MCPartName = dr["MCPartName"].ToString(),
                                InitStockRoll = dr["InitStockRoll"].ToString(),
                                InitStockQty = dr["InitStockQty"].ToString(),
                                StuffRoll = dr["StuffRoll"].ToString(),
                                StuffQty = dr["StuffQty"].ToString(),
                                OutRoll = dr["OutRoll"].ToString(),
                                OutQty = dr["OutQty"].ToString(),
                                StockQty = dr["StockQty"].ToString(),
                                UnitClss = dr["UnitClss"].ToString(),
                                UnitClssName = dr["UnitClssName"].ToString(),
                                NeedstockQty = dr["NeedstockQty"].ToString(),
                                ForUseName = dr["ForUseName"].ToString(),
                                ShortFall = dr["ShortFall"].ToString() //2021-04-26 과부족량을 추가하면서 생성
                            };

                            if (WinMcStock.UnitClss != null)
                            {
                                if (WinMcStock.InitStockQty != null && Lib.Instance.IsNumOrAnother(WinMcStock.InitStockQty))
                                {
                                    WinMcStock.InitStockQty = String.Format("{0:#,0}", double.Parse(WinMcStock.InitStockQty));
                                    //WinMcStock.InitStockQty = stringFormatN0(((int)(double.Parse(WinMcStock.InitStockQty))).ToString());
                                    //WinMcStock.InitStockQty = WinMcStock.UnitClss.Replace(" ", "").Equals("2") ? string.Format("{0:#,###.######0}", WinMcStock.InitStockQty) : string.Format("{0:#,##0}", WinMcStock.InitStockQty);
                                }

                                if (WinMcStock.StuffQty != null && Lib.Instance.IsNumOrAnother(WinMcStock.StuffQty))
                                {
                                    WinMcStock.StuffQty = String.Format("{0:#,0}", double.Parse(WinMcStock.StuffQty));
                                    //WinMcStock.StuffQty = stringFormatN0(((int)(double.Parse(WinMcStock.StuffQty))).ToString());
                                    //WinMcStock.StuffQty = WinMcStock.UnitClss.Replace(" ","").Equals("2") ? string.Format("{0:#,###.######0}", WinMcStock.StuffQty) : string.Format("{0:#,##0}", WinMcStock.StuffQty);
                                }

                                if (WinMcStock.OutQty != null && Lib.Instance.IsNumOrAnother(WinMcStock.OutQty))
                                {
                                    WinMcStock.OutQty = String.Format("{0:#,0}", double.Parse(WinMcStock.OutQty));
                                    //WinMcStock.OutQty = stringFormatN0(((int)(double.Parse(WinMcStock.OutQty))).ToString());
                                    //WinMcStock.OutQty = WinMcStock.UnitClss.Replace(" ", "").Equals("2") ? string.Format("{0:#,###.######0}", WinMcStock.OutQty) : string.Format("{0:#,##0}", WinMcStock.OutQty);
                                }

                                if (WinMcStock.StockQty != null && Lib.Instance.IsNumOrAnother(WinMcStock.StockQty))
                                {
                                    WinMcStock.StockQty = String.Format("{0:#,0}", double.Parse(WinMcStock.StockQty));
                                    //WinMcStock.StockQty = stringFormatN0(((int)(double.Parse(WinMcStock.StockQty))).ToString());
                                    //WinMcStock.StockQty = WinMcStock.UnitClss.Replace(" ", "").Equals("2") ? string.Format("{0:#,###.######0}", WinMcStock.StockQty) : string.Format("{0:#,##0}", WinMcStock.StockQty);
                                }

                                if (WinMcStock.NeedstockQty != null && Lib.Instance.IsNumOrAnother(WinMcStock.NeedstockQty))
                                {
                                    WinMcStock.NeedstockQty = String.Format("{0:#,0}", double.Parse(WinMcStock.NeedstockQty));
                                    //WinMcStock.NeedstockQty = stringFormatN0(((int)(double.Parse(WinMcStock.NeedstockQty))).ToString());
                                    //WinMcStock.NeedstockQty = WinMcStock.UnitClss.Replace(" ", "").Equals("2") ? string.Format("{0:#,###.######0}", WinMcStock.NeedstockQty) : string.Format("{0:#,##0}", WinMcStock.NeedstockQty);
                                }
                                //2021-04-28 단위 맞추기 추가
                                if (WinMcStock.ShortFall != null && Lib.Instance.IsNumOrAnother(WinMcStock.ShortFall))
                                {
                                    WinMcStock.ShortFall = String.Format("{0:#,0}", double.Parse(WinMcStock.ShortFall));
                                    //WinMcStock.ShortFall = stringFormatN0(((int)(double.Parse(WinMcStock.ShortFall))).ToString());
                                    //WinMcStock.NeedstockQty = WinMcStock.UnitClss.Replace(" ", "").Equals("2") ? string.Format("{0:#,###.######0}", WinMcStock.NeedstockQty) : string.Format("{0:#,##0}", WinMcStock.NeedstockQty);
                                }
                            }

                            if (WinMcStock.cls.Equals("3"))
                            {
                                WinMcStock.MCPartName = "품명계";
                                dgdMcStock.Items.Add(WinMcStock);
                            }
                            else if (WinMcStock.cls.Equals("4"))
                            {
                                WinMcStock.MCPartName = "총계";
                                WinMcStock.NeedstockQty = "";
                                dgdTotal.Items.Add(WinMcStock);
                            }
                            else
                            {
                                dgdMcStock.Items.Add(WinMcStock);

                            }

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
    }


    class Win_prd_PartStock_Q_CodeView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Num { get; set; }
        public string cls { get; set; }
        public string McPartID { get; set; }
        public string MCPartName { get; set; }
        public string InitStockRoll { get; set; }
        public string InitStockQty { get; set; }
        public string StuffRoll { get; set; }
        public string StuffQty { get; set; }
        public string OutRoll { get; set; }
        public string OutQty { get; set; }
        public string StockQty { get; set; }
        public string UnitClss { get; set; }
        public string UnitClssName { get; set; }
        public string NeedstockQty { get; set; }
        public string ForUse { get; set; }
        public string ForUseName { get; set; }
        public string ShortFall { get; set; } //2021-04-26 과부족량을 추가하면서 생성
    }
}
