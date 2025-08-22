using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_MIS_CustomArticleInSum_MM_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_MIS_CustomArticleInSum_MM_Q : UserControl
    {
        #region 변수 선언 및 로드

        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();
        Lib lib = new Lib();
        public Win_MIS_CustomArticleInSum_MM_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            chkDate.IsChecked = true;



        }

        #endregion

        #region 날짜 관련 이벤트

        //검색기간
        private void lblDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true) { chkDate.IsChecked = false; }
            else { chkDate.IsChecked = true; }
        }

        //검색기간
        private void chkDate_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;

        }

        //검색기간
        private void chkDate_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;

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
                Thread.Sleep(2000);

                //로직
                FillGrid();

                if (dgdmain.Items.Count > 0)
                {
                    dgdmain.SelectedIndex = 0;
                    this.DataContext = dgdmain.SelectedItem as Win_MIS_CustomArticleInSum_MM_Q_CodeView;
                }
                else
                {
                    MessageBox.Show("조회된 데이터가 없습니다.");
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

        //인쇄
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "월간 매입분석";
            lst[1] = dgdmain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdmain.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdmain);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdmain);

                    Name = dgdmain.Name;

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

        #region 검색조건
        //거래처
        private void lblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustom.IsChecked == true) { chkCustom.IsChecked = false; }
            else { chkCustom.IsChecked = true; }
        }

        //거래처
        private void chkCustom_Checked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = true;
            btnPfCustom.IsEnabled = true;
        }

        //거래처
        private void chkCustom_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = false;
            btnPfCustom.IsEnabled = false;
        }

        //거래처
        private void txtCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        //거래처
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //품번 플러스 파인더 
        private void BtnPfArticleNo_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyerArticleNo, 95, txtBuyerArticleNo.Text);
        }

        //품번 키다운
        private void TxtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtBuyerArticleNo, 95, txtBuyerArticleNo.Text);
            }
        }

        //품번
        private void lblBuyerArticleNo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkBuyerArticleNo.IsChecked == true) { chkBuyerArticleNo.IsChecked = false; }
            else { chkBuyerArticleNo.IsChecked = true; }
        }

        //품번
        private void chkBuyerArticleNo_Checked(object sender, RoutedEventArgs e)
        {
            txtBuyerArticleNo.IsEnabled = true;
            btnPfArticleNo.IsEnabled = true;
        }

        //품번
        private void chkBuyerArticleNo_Unchecked(object sender, RoutedEventArgs e)
        {
            txtBuyerArticleNo.IsEnabled = false;
            btnPfArticleNo.IsEnabled = false;
        }

        #endregion

        //실 조회
        private void FillGrid()
        {
            try
            {
                if (dgdmain.Items.Count > 0)
                {
                    dgdmain.Items.Clear();
                }


                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                sqlParameter.Add("EYYYYMM", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMM") : "");

                sqlParameter.Add("chkCustomID", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true ? txtCustom.Tag : "");
                sqlParameter.Add("chkBuyerArticleNo", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNo.IsChecked == true ? txtBuyerArticleNo.Text : "");



                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_MIS_sInwareMMSpread", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var DD = new Win_MIS_CustomArticleInSum_MM_Q_CodeView()
                            {

                                CustomID = dr["CustomName"].ToString(),
                                QTY00 = stringFormatN0(dr["QTY00"]),
                                AMOUNT00 = stringFormatN0(dr["AMOUNT00"]),
                                QTY01 = stringFormatN0(dr["QTY01"]),
                                AMOUNT01 = stringFormatN0(dr["AMOUNT01"]),
                                QTY02 = stringFormatN0(dr["QTY02"]),
                                AMOUNT02 = stringFormatN0(dr["AMOUNT02"]),
                                QTY03 = stringFormatN0(dr["QTY03"]),
                                AMOUNT03 = stringFormatN0(dr["AMOUNT03"]),
                                QTY04 = stringFormatN0(dr["QTY04"]),
                                AMOUNT04 = stringFormatN0(dr["AMOUNT04"]),
                                QTY05 = stringFormatN0(dr["QTY05"]),
                                AMOUNT05 = stringFormatN0(dr["AMOUNT05"]),
                                QTY06 = stringFormatN0(dr["QTY06"]),
                                AMOUNT06 = stringFormatN0(dr["AMOUNT06"]),
                                QTY07 = stringFormatN0(dr["QTY07"]),
                                AMOUNT07 = stringFormatN0(dr["AMOUNT07"]),
                                QTY08 = stringFormatN0(dr["QTY08"]),
                                AMOUNT08 = stringFormatN0(dr["AMOUNT08"]),
                                QTY09 = stringFormatN0(dr["QTY09"]),
                                AMOUNT09 = stringFormatN0(dr["AMOUNT09"]),
                                QTY10 = stringFormatN0(dr["QTY10"]),
                                AMOUNT10 = stringFormatN0(dr["AMOUNT10"]),
                                QTY11 = stringFormatN0(dr["QTY11"]),
                                AMOUNT11 = stringFormatN0(dr["AMOUNT11"]),
                                QTY12 = stringFormatN0(dr["QTY12"]),
                                AMOUNT12 = stringFormatN0(dr["AMOUNT12"])
                            };

                            if (DD.CustomID.Equals("합계"))
                            {
                                DD.TotalColor = "true";
                            }

                            DD.QTY00 = lib.returnNumStringZero(DD.QTY00);
                            DD.AMOUNT00 = lib.returnNumStringZero(DD.AMOUNT00);
                            DD.QTY01 = lib.returnNumStringZero(DD.QTY01);
                            DD.AMOUNT01 = lib.returnNumStringZero(DD.AMOUNT01);
                            DD.QTY02 = lib.returnNumStringZero(DD.QTY02);
                            DD.AMOUNT02 = lib.returnNumStringZero(DD.AMOUNT02);
                            DD.QTY03 = lib.returnNumStringZero(DD.QTY03);
                            DD.AMOUNT03 = lib.returnNumStringZero(DD.AMOUNT03);
                            DD.QTY04 = lib.returnNumStringZero(DD.QTY04);
                            DD.AMOUNT04 = lib.returnNumStringZero(DD.AMOUNT04);
                            DD.QTY05 = lib.returnNumStringZero(DD.QTY05);
                            DD.AMOUNT05 = lib.returnNumStringZero(DD.AMOUNT05);
                            DD.QTY06 = lib.returnNumStringZero(DD.QTY06);
                            DD.AMOUNT06 = lib.returnNumStringZero(DD.AMOUNT06);
                            DD.QTY07 = lib.returnNumStringZero(DD.QTY07);
                            DD.AMOUNT07 = lib.returnNumStringZero(DD.AMOUNT07);
                            DD.QTY08 = lib.returnNumStringZero(DD.QTY08);
                            DD.AMOUNT08 = lib.returnNumStringZero(DD.AMOUNT08);
                            DD.QTY09 = lib.returnNumStringZero(DD.QTY09);
                            DD.AMOUNT09 = lib.returnNumStringZero(DD.AMOUNT09);
                            DD.QTY10 = lib.returnNumStringZero(DD.QTY10);
                            DD.AMOUNT10 = lib.returnNumStringZero(DD.AMOUNT10);
                            DD.QTY11 = lib.returnNumStringZero(DD.QTY11);
                            DD.AMOUNT11 = lib.returnNumStringZero(DD.AMOUNT11);
                            DD.QTY12 = lib.returnNumStringZero(DD.QTY12);
                            DD.AMOUNT12 = lib.returnNumStringZero(DD.AMOUNT12);

                            if (i == 1)
                            {
                                date1.Content = DD.QTY01;
                                date2.Content = DD.QTY02;
                                date3.Content = DD.QTY03;
                                date4.Content = DD.QTY04;
                                date5.Content = DD.QTY05;
                                date6.Content = DD.QTY06;
                                date7.Content = DD.QTY07;
                                date8.Content = DD.QTY08;
                                date9.Content = DD.QTY09;
                                date10.Content = DD.QTY10;
                                date11.Content = DD.QTY11;
                                date12.Content = DD.QTY12;
                            }
                            else
                            {
                                dgdmain.Items.Add(DD);

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
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }
        #region 히든

        //미리보기(인쇄 하위버튼)
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            if (dgdmain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(true);
        }

        //바로인쇄(인쇄 하위버튼)
        private void menuRightPrint_Click(object sender, RoutedEventArgs e)
        {
            if (dgdmain.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            PrintWork(false);
        }

        //닫   기(인쇄 하위버튼)
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }

        private void PrintWork(bool preview_click)
        {

        }

        #endregion

        private void Grid_Scroll(object sender, System.Windows.Controls.Primitives.ScrollEventArgs e)
        {
            svrHeader.ScrollToHorizontalOffset(dgdscroll.HorizontalOffset);
            svrHeader.UpdateLayout();

        }

        private void Scroll_Scroll(object sender, System.Windows.Controls.Primitives.ScrollEventArgs e)
        {
            //dgdmain. = svrHeader.VerticalOffset;
            //dgdmain.UpdateLayout();
        }

        private void Dgdmain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }

    class Win_MIS_CustomArticleInSum_MM_Q_CodeView : BaseView
    {
        public string CustomID { get; internal set; }
        public string QTY00 { get; internal set; }
        public string AMOUNT00 { get; internal set; }
        public string QTY01 { get; internal set; }
        public string AMOUNT01 { get; internal set; }
        public string QTY02 { get; internal set; }
        public string AMOUNT02 { get; internal set; }
        public string QTY03 { get; internal set; }
        public string AMOUNT03 { get; internal set; }
        public string QTY04 { get; internal set; }
        public string AMOUNT04 { get; internal set; }
        public string QTY05 { get; internal set; }
        public string AMOUNT05 { get; internal set; }
        public string QTY06 { get; internal set; }
        public string AMOUNT06 { get; internal set; }
        public string QTY07 { get; internal set; }
        public string AMOUNT07 { get; internal set; }
        public string QTY08 { get; internal set; }
        public string AMOUNT08 { get; internal set; }
        public string QTY09 { get; internal set; }
        public string AMOUNT09 { get; internal set; }
        public string QTY10 { get; internal set; }
        public string AMOUNT10 { get; internal set; }
        public string QTY11 { get; internal set; }
        public string AMOUNT11 { get; internal set; }
        public string QTY12 { get; internal set; }
        public string AMOUNT12 { get; internal set; }

        public string TotalColor { get; internal set; }
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }


    }
}



