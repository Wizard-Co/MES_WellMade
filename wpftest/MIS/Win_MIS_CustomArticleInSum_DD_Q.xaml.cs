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
    /// Win_MIS_CustomArticleInSum_DD_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_MIS_CustomArticleInSum_DD_Q : UserControl
    {
        #region 변수 선언 및 로드

        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();
        Lib lib = new Lib();

        public Win_MIS_CustomArticleInSum_DD_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
            dtpSDate.SelectedDate = DateTime.Today;
            dtpEDate.SelectedDate = DateTime.Today;
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
            dtpEDate.IsEnabled = true;
        }

        //검색기간
        private void chkDate_Unchecked(object sender, RoutedEventArgs e)
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
                    this.DataContext = dgdmain.SelectedItem as Win_MIS_CustomArticleInSum_DD_Q_CodeView;
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
            lst[0] = "일일 매입분석 조회";
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

                sqlParameter.Add("SDate", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("chkCustomID", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true ? txtCustom.Tag : "");
                sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNo.IsChecked == true ? txtBuyerArticleNo.Text : "");
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_MIS_sInwareDDSpread", sqlParameter, false);


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


                            var DD = new Win_MIS_CustomArticleInSum_DD_Q_CodeView()
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
                                AMOUNT12 = stringFormatN0(dr["AMOUNT12"]),
                                QTY13 = stringFormatN0(dr["QTY13"]),
                                AMOUNT13 = stringFormatN0(dr["AMOUNT13"]),
                                QTY14 = stringFormatN0(dr["QTY14"]),
                                AMOUNT14 = stringFormatN0(dr["AMOUNT14"]),
                                QTY15 = stringFormatN0(dr["QTY15"]),
                                AMOUNT15 = stringFormatN0(dr["AMOUNT15"]),
                                QTY16 = stringFormatN0(dr["QTY16"]),
                                AMOUNT16 = stringFormatN0(dr["AMOUNT16"]),
                                QTY17 = stringFormatN0(dr["QTY17"]),
                                AMOUNT17 = stringFormatN0(dr["AMOUNT17"]),
                                QTY18 = stringFormatN0(dr["QTY18"]),
                                AMOUNT18 = stringFormatN0(dr["AMOUNT18"]),
                                QTY19 = stringFormatN0(dr["QTY19"]),
                                AMOUNT19 = stringFormatN0(dr["AMOUNT19"]),

                                QTY20 = stringFormatN0(dr["QTY20"]),
                                AMOUNT20 = stringFormatN0(dr["AMOUNT20"]),
                                QTY21 = stringFormatN0(dr["QTY21"]),
                                AMOUNT21 = stringFormatN0(dr["AMOUNT21"]),
                                QTY22 = stringFormatN0(dr["QTY22"]),
                                AMOUNT22 = stringFormatN0(dr["AMOUNT22"]),
                                QTY23 = stringFormatN0(dr["QTY23"]),
                                AMOUNT23 = stringFormatN0(dr["AMOUNT23"]),
                                QTY24 = stringFormatN0(dr["QTY24"]),
                                AMOUNT24 = stringFormatN0(dr["AMOUNT24"]),
                                QTY25 = stringFormatN0(dr["QTY25"]),
                                AMOUNT25 = stringFormatN0(dr["AMOUNT25"]),
                                QTY26 = stringFormatN0(dr["QTY26"]),
                                AMOUNT26 = stringFormatN0(dr["AMOUNT26"]),
                                QTY27 = stringFormatN0(dr["QTY27"]),
                                AMOUNT27 = stringFormatN0(dr["AMOUNT27"]),
                                QTY28 = stringFormatN0(dr["QTY28"]),
                                AMOUNT28 = stringFormatN0(dr["AMOUNT28"]),
                                QTY29 = stringFormatN0(dr["QTY29"]),
                                AMOUNT29 = stringFormatN0(dr["AMOUNT29"]),
                                QTY30 = stringFormatN0(dr["QTY30"]),
                                AMOUNT30 = stringFormatN0(dr["AMOUNT30"]),
                                QTY31 = stringFormatN0(dr["QTY31"]),
                                AMOUNT31 = stringFormatN0(dr["AMOUNT31"]),
                                QTY32 = stringFormatN0(dr["QTY32"]),
                                AMOUNT32 = stringFormatN0(dr["AMOUNT32"]),
                                QTY33 = stringFormatN0(dr["QTY33"]),
                                AMOUNT33 = stringFormatN0(dr["AMOUNT33"]),
                                QTY34 = stringFormatN0(dr["QTY34"]),
                                AMOUNT34 = stringFormatN0(dr["AMOUNT34"]),
                                QTY35 = stringFormatN0(dr["QTY35"]),
                                AMOUNT35 = stringFormatN0(dr["AMOUNT35"]),
                                QTY36 = stringFormatN0(dr["QTY36"]),
                                AMOUNT36 = stringFormatN0(dr["AMOUNT36"]),
                                QTY37 = stringFormatN0(dr["QTY37"]),
                                AMOUNT37 = stringFormatN0(dr["AMOUNT37"]),
                                QTY38 = stringFormatN0(dr["QTY38"]),
                                AMOUNT38 = stringFormatN0(dr["AMOUNT38"]),

                                QTY39 = stringFormatN0(dr["QTY39"]),
                                AMOUNT39 = stringFormatN0(dr["AMOUNT39"]),
                                QTY40 = stringFormatN0(dr["QTY40"]),
                                AMOUNT40 = stringFormatN0(dr["AMOUNT40"]),
                                QTY41 = stringFormatN0(dr["QTY41"]),
                                AMOUNT41 = stringFormatN0(dr["AMOUNT41"]),
                                QTY42 = stringFormatN0(dr["QTY42"]),
                                AMOUNT42 = stringFormatN0(dr["AMOUNT42"]),
                                QTY43 = stringFormatN0(dr["QTY43"]),
                                AMOUNT43 = stringFormatN0(dr["AMOUNT43"]),
                                QTY44 = stringFormatN0(dr["QTY44"]),
                                AMOUNT44 = stringFormatN0(dr["AMOUNT44"]),
                                QTY45 = stringFormatN0(dr["QTY45"]),
                                AMOUNT45 = stringFormatN0(dr["AMOUNT45"]),
                                QTY46 = stringFormatN0(dr["QTY46"]),
                                AMOUNT46 = stringFormatN0(dr["AMOUNT46"]),
                                QTY47 = stringFormatN0(dr["QTY47"]),
                                AMOUNT47 = stringFormatN0(dr["AMOUNT47"]),
                                QTY48 = stringFormatN0(dr["QTY48"]),
                                AMOUNT48 = stringFormatN0(dr["AMOUNT48"]),
                                QTY49 = stringFormatN0(dr["QTY49"]),
                                AMOUNT49 = stringFormatN0(dr["AMOUNT49"]),
                                QTY50 = stringFormatN0(dr["QTY50"]),
                                AMOUNT50 = stringFormatN0(dr["AMOUNT50"]),
                                QTY51 = stringFormatN0(dr["QTY51"]),
                                AMOUNT51 = stringFormatN0(dr["AMOUNT51"]),
                                QTY52 = stringFormatN0(dr["QTY52"]),
                                AMOUNT52 = stringFormatN0(dr["AMOUNT52"]),
                                QTY53 = stringFormatN0(dr["QTY53"]),
                                AMOUNT53 = stringFormatN0(dr["AMOUNT53"]),

                                QTY54 = stringFormatN0(dr["QTY54"]),
                                AMOUNT54 = stringFormatN0(dr["AMOUNT54"]),
                                QTY55 = stringFormatN0(dr["QTY55"]),
                                AMOUNT55 = stringFormatN0(dr["AMOUNT55"]),
                                QTY56 = stringFormatN0(dr["QTY56"]),
                                AMOUNT56 = stringFormatN0(dr["AMOUNT56"]),
                                QTY57 = stringFormatN0(dr["QTY57"]),
                                AMOUNT57 = stringFormatN0(dr["AMOUNT57"]),
                                QTY58 = stringFormatN0(dr["QTY58"]),
                                AMOUNT58 = stringFormatN0(dr["AMOUNT58"]),
                                QTY59 = stringFormatN0(dr["QTY59"]),
                                AMOUNT59 = stringFormatN0(dr["AMOUNT59"]),
                                QTY60 = stringFormatN0(dr["QTY60"]),
                                AMOUNT60 = stringFormatN0(dr["AMOUNT60"]),
                                QTY61 = stringFormatN0(dr["QTY61"]),
                                AMOUNT61 = stringFormatN0(dr["AMOUNT61"]),
                                QTY62 = stringFormatN0(dr["QTY62"]),
                                AMOUNT62 = stringFormatN0(dr["AMOUNT62"]),
                                QTY63 = stringFormatN0(dr["QTY63"]),
                                AMOUNT63 = stringFormatN0(dr["AMOUNT63"]),
                                QTY64 = stringFormatN0(dr["QTY64"]),
                                AMOUNT64 = stringFormatN0(dr["AMOUNT64"]),
                                QTY65 = stringFormatN0(dr["QTY65"]),
                                AMOUNT65 = stringFormatN0(dr["AMOUNT65"]),
                                QTY66 = stringFormatN0(dr["QTY66"]),
                                AMOUNT66 = stringFormatN0(dr["AMOUNT66"]),
                                QTY67 = stringFormatN0(dr["QTY67"]),
                                AMOUNT67 = stringFormatN0(dr["AMOUNT67"]),
                                QTY68 = stringFormatN0(dr["QTY68"]),
                                AMOUNT68 = stringFormatN0(dr["AMOUNT68"]),
                                QTY69 = stringFormatN0(dr["QTY69"]),
                                AMOUNT69 = stringFormatN0(dr["AMOUNT69"]),

                                QTY70 = stringFormatN0(dr["QTY70"]),
                                AMOUNT70 = stringFormatN0(dr["AMOUNT70"]),
                                QTY71 = stringFormatN0(dr["QTY71"]),
                                AMOUNT71 = stringFormatN0(dr["AMOUNT71"]),
                                QTY72 = stringFormatN0(dr["QTY72"]),
                                AMOUNT72 = stringFormatN0(dr["AMOUNT72"]),
                                QTY73 = stringFormatN0(dr["QTY73"]),
                                AMOUNT73 = stringFormatN0(dr["AMOUNT73"]),
                                QTY74 = stringFormatN0(dr["QTY74"]),
                                AMOUNT74 = stringFormatN0(dr["AMOUNT74"]),
                                QTY75 = stringFormatN0(dr["QTY75"]),
                                AMOUNT75 = stringFormatN0(dr["AMOUNT75"]),
                                QTY76 = stringFormatN0(dr["QTY76"]),
                                AMOUNT76 = stringFormatN0(dr["AMOUNT76"]),
                                QTY77 = stringFormatN0(dr["QTY77"]),
                                AMOUNT77 = stringFormatN0(dr["AMOUNT77"]),
                                QTY78 = stringFormatN0(dr["QTY78"]),
                                AMOUNT78 = stringFormatN0(dr["AMOUNT78"]),
                                QTY79 = stringFormatN0(dr["QTY79"]),
                                AMOUNT79 = stringFormatN0(dr["AMOUNT79"]),
                                QTY80 = stringFormatN0(dr["QTY80"]),
                                AMOUNT80 = stringFormatN0(dr["AMOUNT80"]),
                                QTY81 = stringFormatN0(dr["QTY81"]),
                                AMOUNT81 = stringFormatN0(dr["AMOUNT81"]),

                                QTY82 = stringFormatN0(dr["QTY82"]),
                                AMOUNT82 = stringFormatN0(dr["AMOUNT82"]),
                                QTY83 = stringFormatN0(dr["QTY83"]),
                                AMOUNT83 = stringFormatN0(dr["AMOUNT83"]),
                                QTY84 = stringFormatN0(dr["QTY84"]),
                                AMOUNT84 = stringFormatN0(dr["AMOUNT84"]),
                                QTY85 = stringFormatN0(dr["QTY85"]),
                                AMOUNT85 = stringFormatN0(dr["AMOUNT85"]),
                                QTY86 = stringFormatN0(dr["QTY86"]),
                                AMOUNT86 = stringFormatN0(dr["AMOUNT86"]),
                                QTY87 = stringFormatN0(dr["QTY87"]),
                                AMOUNT87 = stringFormatN0(dr["AMOUNT87"]),
                                QTY88 = stringFormatN0(dr["QTY88"]),
                                AMOUNT88 = stringFormatN0(dr["AMOUNT88"]),
                                QTY89 = stringFormatN0(dr["QTY89"]),
                                AMOUNT89 = stringFormatN0(dr["AMOUNT89"]),
                                QTY90 = stringFormatN0(dr["QTY90"]),
                                AMOUNT90 = stringFormatN0(dr["AMOUNT90"]),
                                QTY91 = stringFormatN0(dr["QTY91"]),
                                AMOUNT91 = stringFormatN0(dr["AMOUNT91"]),
                                QTY92 = stringFormatN0(dr["QTY92"]),
                                AMOUNT92 = stringFormatN0(dr["AMOUNT92"]),
                                QTY93 = stringFormatN0(dr["QTY93"]),
                                AMOUNT93 = stringFormatN0(dr["AMOUNT93"]),
                                QTY94 = stringFormatN0(dr["QTY94"]),
                                AMOUNT94 = stringFormatN0(dr["AMOUNT94"]),
                                QTY95 = stringFormatN0(dr["QTY95"]),
                                AMOUNT95 = stringFormatN0(dr["AMOUNT95"])
                            };



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
                            DD.QTY13 = lib.returnNumStringZero(DD.QTY13);
                            DD.AMOUNT13 = lib.returnNumStringZero(DD.AMOUNT13);
                            DD.QTY14 = lib.returnNumStringZero(DD.QTY14);
                            DD.AMOUNT14 = lib.returnNumStringZero(DD.AMOUNT14);
                            DD.QTY15 = lib.returnNumStringZero(DD.QTY15);
                            DD.AMOUNT15 = lib.returnNumStringZero(DD.AMOUNT15);
                            DD.QTY16 = lib.returnNumStringZero(DD.QTY16);
                            DD.AMOUNT16 = lib.returnNumStringZero(DD.AMOUNT16);
                            DD.QTY17 = lib.returnNumStringZero(DD.QTY17);
                            DD.AMOUNT17 = lib.returnNumStringZero(DD.AMOUNT17);
                            DD.QTY18 = lib.returnNumStringZero(DD.QTY18);
                            DD.AMOUNT18 = lib.returnNumStringZero(DD.AMOUNT18);
                            DD.QTY19 = lib.returnNumStringZero(DD.QTY19);
                            DD.AMOUNT19 = lib.returnNumStringZero(DD.AMOUNT19);
                            DD.QTY20 = lib.returnNumStringZero(DD.QTY20);
                            DD.AMOUNT20 = lib.returnNumStringZero(DD.AMOUNT20);
                            DD.QTY21 = lib.returnNumStringZero(DD.QTY21);
                            DD.AMOUNT21 = lib.returnNumStringZero(DD.AMOUNT21);
                            DD.QTY22 = lib.returnNumStringZero(DD.QTY22);
                            DD.AMOUNT22 = lib.returnNumStringZero(DD.AMOUNT22);
                            DD.QTY23 = lib.returnNumStringZero(DD.QTY23);
                            DD.AMOUNT23 = lib.returnNumStringZero(DD.AMOUNT23);
                            DD.QTY24 = lib.returnNumStringZero(DD.QTY24);
                            DD.AMOUNT24 = lib.returnNumStringZero(DD.AMOUNT24);
                            DD.QTY25 = lib.returnNumStringZero(DD.QTY25);
                            DD.AMOUNT25 = lib.returnNumStringZero(DD.AMOUNT25);
                            DD.QTY26 = lib.returnNumStringZero(DD.QTY26);
                            DD.AMOUNT26 = lib.returnNumStringZero(DD.AMOUNT26);
                            DD.QTY27 = lib.returnNumStringZero(DD.QTY27);
                            DD.AMOUNT27 = lib.returnNumStringZero(DD.AMOUNT27);
                            DD.QTY28 = lib.returnNumStringZero(DD.QTY28);
                            DD.AMOUNT28 = lib.returnNumStringZero(DD.AMOUNT28);
                            DD.QTY29 = lib.returnNumStringZero(DD.QTY29);
                            DD.AMOUNT29 = lib.returnNumStringZero(DD.AMOUNT29);
                            DD.QTY30 = lib.returnNumStringZero(DD.QTY30);
                            DD.AMOUNT30 = lib.returnNumStringZero(DD.AMOUNT30);
                            DD.QTY31 = lib.returnNumStringZero(DD.QTY31);
                            DD.AMOUNT31 = lib.returnNumStringZero(DD.AMOUNT31);
                            DD.QTY32 = lib.returnNumStringZero(DD.QTY32);
                            DD.AMOUNT32 = lib.returnNumStringZero(DD.AMOUNT32);
                            DD.QTY33 = lib.returnNumStringZero(DD.QTY33);
                            DD.AMOUNT33 = lib.returnNumStringZero(DD.AMOUNT33);
                            DD.QTY34 = lib.returnNumStringZero(DD.QTY34);
                            DD.AMOUNT34 = lib.returnNumStringZero(DD.AMOUNT34);
                            DD.QTY35 = lib.returnNumStringZero(DD.QTY35);
                            DD.AMOUNT35 = lib.returnNumStringZero(DD.AMOUNT35);
                            DD.QTY36 = lib.returnNumStringZero(DD.QTY36);
                            DD.AMOUNT36 = lib.returnNumStringZero(DD.AMOUNT36);
                            DD.QTY37 = lib.returnNumStringZero(DD.QTY37);
                            DD.AMOUNT37 = lib.returnNumStringZero(DD.AMOUNT37);
                            DD.QTY38 = lib.returnNumStringZero(DD.QTY38);
                            DD.AMOUNT38 = lib.returnNumStringZero(DD.AMOUNT38);
                            DD.QTY39 = lib.returnNumStringZero(DD.QTY39);
                            DD.AMOUNT39 = lib.returnNumStringZero(DD.AMOUNT39);
                            DD.QTY40 = lib.returnNumStringZero(DD.QTY40);
                            DD.AMOUNT40 = lib.returnNumStringZero(DD.AMOUNT40);
                            DD.QTY41 = lib.returnNumStringZero(DD.QTY41);
                            DD.AMOUNT41 = lib.returnNumStringZero(DD.AMOUNT41);
                            DD.QTY42 = lib.returnNumStringZero(DD.QTY42);
                            DD.AMOUNT42 = lib.returnNumStringZero(DD.AMOUNT42);
                            DD.QTY43 = lib.returnNumStringZero(DD.QTY43);
                            DD.AMOUNT43 = lib.returnNumStringZero(DD.AMOUNT43);
                            DD.QTY44 = lib.returnNumStringZero(DD.QTY44);
                            DD.AMOUNT44 = lib.returnNumStringZero(DD.AMOUNT44);
                            DD.QTY45 = lib.returnNumStringZero(DD.QTY45);
                            DD.AMOUNT45 = lib.returnNumStringZero(DD.AMOUNT45);
                            DD.QTY46 = lib.returnNumStringZero(DD.QTY46);
                            DD.AMOUNT46 = lib.returnNumStringZero(DD.AMOUNT46);
                            DD.QTY47 = lib.returnNumStringZero(DD.QTY47);
                            DD.AMOUNT47 = lib.returnNumStringZero(DD.AMOUNT47);
                            DD.QTY48 = lib.returnNumStringZero(DD.QTY48);
                            DD.AMOUNT48 = lib.returnNumStringZero(DD.AMOUNT48);
                            DD.QTY49 = lib.returnNumStringZero(DD.QTY49);
                            DD.AMOUNT49 = lib.returnNumStringZero(DD.AMOUNT49);
                            DD.QTY50 = lib.returnNumStringZero(DD.QTY50);
                            DD.AMOUNT50 = lib.returnNumStringZero(DD.AMOUNT50);
                            DD.QTY51 = lib.returnNumStringZero(DD.QTY51);
                            DD.AMOUNT51 = lib.returnNumStringZero(DD.AMOUNT51);
                            DD.QTY52 = lib.returnNumStringZero(DD.QTY52);
                            DD.AMOUNT52 = lib.returnNumStringZero(DD.AMOUNT52);
                            DD.QTY53 = lib.returnNumStringZero(DD.QTY53);
                            DD.AMOUNT53 = lib.returnNumStringZero(DD.AMOUNT53);
                            DD.QTY54 = lib.returnNumStringZero(DD.QTY54);
                            DD.AMOUNT54 = lib.returnNumStringZero(DD.AMOUNT54);
                            DD.QTY55 = lib.returnNumStringZero(DD.QTY55);
                            DD.AMOUNT55 = lib.returnNumStringZero(DD.AMOUNT55);
                            DD.QTY56 = lib.returnNumStringZero(DD.QTY56);
                            DD.AMOUNT56 = lib.returnNumStringZero(DD.AMOUNT56);
                            DD.QTY57 = lib.returnNumStringZero(DD.QTY57);
                            DD.AMOUNT57 = lib.returnNumStringZero(DD.AMOUNT57);
                            DD.QTY58 = lib.returnNumStringZero(DD.QTY58);
                            DD.AMOUNT58 = lib.returnNumStringZero(DD.AMOUNT58);
                            DD.QTY59 = lib.returnNumStringZero(DD.QTY59);
                            DD.AMOUNT59 = lib.returnNumStringZero(DD.AMOUNT59);
                            DD.QTY60 = lib.returnNumStringZero(DD.QTY60);
                            DD.AMOUNT60 = lib.returnNumStringZero(DD.AMOUNT60);
                            DD.QTY61 = lib.returnNumStringZero(DD.QTY61);
                            DD.AMOUNT61 = lib.returnNumStringZero(DD.AMOUNT61);
                            DD.QTY62 = lib.returnNumStringZero(DD.QTY62);
                            DD.AMOUNT62 = lib.returnNumStringZero(DD.AMOUNT62);
                            DD.QTY63 = lib.returnNumStringZero(DD.QTY63);
                            DD.AMOUNT63 = lib.returnNumStringZero(DD.AMOUNT63);
                            DD.QTY64 = lib.returnNumStringZero(DD.QTY64);
                            DD.AMOUNT64 = lib.returnNumStringZero(DD.AMOUNT64);
                            DD.QTY65 = lib.returnNumStringZero(DD.QTY65);
                            DD.AMOUNT65 = lib.returnNumStringZero(DD.AMOUNT65);
                            DD.QTY66 = lib.returnNumStringZero(DD.QTY66);
                            DD.AMOUNT66 = lib.returnNumStringZero(DD.AMOUNT66);
                            DD.QTY67 = lib.returnNumStringZero(DD.QTY67);
                            DD.AMOUNT67 = lib.returnNumStringZero(DD.AMOUNT67);
                            DD.QTY68 = lib.returnNumStringZero(DD.QTY68);
                            DD.AMOUNT68 = lib.returnNumStringZero(DD.AMOUNT68);
                            DD.QTY69 = lib.returnNumStringZero(DD.QTY69);
                            DD.AMOUNT69 = lib.returnNumStringZero(DD.AMOUNT69);
                            DD.QTY70 = lib.returnNumStringZero(DD.QTY70);
                            DD.AMOUNT70 = lib.returnNumStringZero(DD.AMOUNT70);
                            DD.QTY71 = lib.returnNumStringZero(DD.QTY71);
                            DD.AMOUNT71 = lib.returnNumStringZero(DD.AMOUNT71);
                            DD.QTY72 = lib.returnNumStringZero(DD.QTY72);
                            DD.AMOUNT72 = lib.returnNumStringZero(DD.AMOUNT72);
                            DD.QTY73 = lib.returnNumStringZero(DD.QTY73);
                            DD.AMOUNT73 = lib.returnNumStringZero(DD.AMOUNT73);
                            DD.QTY74 = lib.returnNumStringZero(DD.QTY74);
                            DD.AMOUNT74 = lib.returnNumStringZero(DD.AMOUNT74);
                            DD.QTY75 = lib.returnNumStringZero(DD.QTY75);
                            DD.AMOUNT75 = lib.returnNumStringZero(DD.AMOUNT75);
                            DD.QTY76 = lib.returnNumStringZero(DD.QTY76);
                            DD.AMOUNT76 = lib.returnNumStringZero(DD.AMOUNT76);
                            DD.QTY77 = lib.returnNumStringZero(DD.QTY77);
                            DD.AMOUNT77 = lib.returnNumStringZero(DD.AMOUNT77);
                            DD.QTY78 = lib.returnNumStringZero(DD.QTY78);
                            DD.AMOUNT78 = lib.returnNumStringZero(DD.AMOUNT78);
                            DD.QTY79 = lib.returnNumStringZero(DD.QTY79);
                            DD.AMOUNT79 = lib.returnNumStringZero(DD.AMOUNT79);
                            DD.QTY80 = lib.returnNumStringZero(DD.QTY80);
                            DD.AMOUNT80 = lib.returnNumStringZero(DD.AMOUNT80);
                            DD.QTY81 = lib.returnNumStringZero(DD.QTY81);
                            DD.AMOUNT81 = lib.returnNumStringZero(DD.AMOUNT81);
                            DD.QTY82 = lib.returnNumStringZero(DD.QTY82);
                            DD.AMOUNT82 = lib.returnNumStringZero(DD.AMOUNT82);
                            DD.QTY83 = lib.returnNumStringZero(DD.QTY83);
                            DD.AMOUNT83 = lib.returnNumStringZero(DD.AMOUNT83);
                            DD.QTY84 = lib.returnNumStringZero(DD.QTY84);
                            DD.AMOUNT84 = lib.returnNumStringZero(DD.AMOUNT84);
                            DD.QTY85 = lib.returnNumStringZero(DD.QTY85);
                            DD.AMOUNT85 = lib.returnNumStringZero(DD.AMOUNT85);
                            DD.QTY86 = lib.returnNumStringZero(DD.QTY86);
                            DD.AMOUNT86 = lib.returnNumStringZero(DD.AMOUNT86);
                            DD.QTY87 = lib.returnNumStringZero(DD.QTY87);
                            DD.AMOUNT87 = lib.returnNumStringZero(DD.AMOUNT87);
                            DD.QTY88 = lib.returnNumStringZero(DD.QTY88);
                            DD.AMOUNT88 = lib.returnNumStringZero(DD.AMOUNT88);
                            DD.QTY89 = lib.returnNumStringZero(DD.QTY89);
                            DD.AMOUNT89 = lib.returnNumStringZero(DD.AMOUNT89);
                            DD.QTY90 = lib.returnNumStringZero(DD.QTY90);
                            DD.AMOUNT90 = lib.returnNumStringZero(DD.AMOUNT90);
                            DD.QTY91 = lib.returnNumStringZero(DD.QTY91);
                            DD.AMOUNT91 = lib.returnNumStringZero(DD.AMOUNT91);
                            DD.QTY92 = lib.returnNumStringZero(DD.QTY92);
                            DD.AMOUNT92 = lib.returnNumStringZero(DD.AMOUNT92);
                            DD.QTY93 = lib.returnNumStringZero(DD.QTY93);
                            DD.AMOUNT93 = lib.returnNumStringZero(DD.AMOUNT93);
                            DD.QTY94 = lib.returnNumStringZero(DD.QTY94);
                            DD.AMOUNT94 = lib.returnNumStringZero(DD.AMOUNT94);
                            DD.QTY95 = lib.returnNumStringZero(DD.QTY95);
                            DD.AMOUNT95 = lib.returnNumStringZero(DD.AMOUNT95);
                            if (DD.CustomID.Equals("합계"))
                            {
                                DD.TotalColor = "true";
                            }
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
                                date13.Content = DD.QTY13;
                                date14.Content = DD.QTY14;
                                date15.Content = DD.QTY15;
                                date16.Content = DD.QTY16;
                                date17.Content = DD.QTY17;
                                date18.Content = DD.QTY18;
                                date19.Content = DD.QTY19;
                                date20.Content = DD.QTY20;

                                date21.Content = DD.QTY21;
                                date22.Content = DD.QTY22;
                                date23.Content = DD.QTY23;
                                date24.Content = DD.QTY24;
                                date25.Content = DD.QTY25;
                                date26.Content = DD.QTY26;
                                date27.Content = DD.QTY27;
                                date28.Content = DD.QTY28;
                                date29.Content = DD.QTY29;
                                date30.Content = DD.QTY30;

                                date31.Content = DD.QTY31;
                                date32.Content = DD.QTY32;
                                date33.Content = DD.QTY33;
                                date34.Content = DD.QTY34;
                                date35.Content = DD.QTY35;
                                date36.Content = DD.QTY36;
                                date37.Content = DD.QTY37;
                                date38.Content = DD.QTY38;
                                date39.Content = DD.QTY39;
                                date40.Content = DD.QTY40;

                                date41.Content = DD.QTY41;
                                date42.Content = DD.QTY42;
                                date43.Content = DD.QTY43;
                                date44.Content = DD.QTY44;
                                date45.Content = DD.QTY45;
                                date46.Content = DD.QTY46;
                                date47.Content = DD.QTY47;
                                date48.Content = DD.QTY48;
                                date49.Content = DD.QTY49;
                                date50.Content = DD.QTY50;

                                date51.Content = DD.QTY51;
                                date52.Content = DD.QTY52;
                                date53.Content = DD.QTY53;
                                date54.Content = DD.QTY54;
                                date55.Content = DD.QTY55;
                                date56.Content = DD.QTY56;
                                date57.Content = DD.QTY57;
                                date58.Content = DD.QTY58;
                                date59.Content = DD.QTY59;
                                date60.Content = DD.QTY60;

                                date61.Content = DD.QTY61;
                                date62.Content = DD.QTY62;
                                date63.Content = DD.QTY63;
                                date64.Content = DD.QTY64;
                                date65.Content = DD.QTY65;
                                date66.Content = DD.QTY66;
                                date67.Content = DD.QTY67;
                                date68.Content = DD.QTY68;
                                date69.Content = DD.QTY69;
                                date70.Content = DD.QTY70;

                                date71.Content = DD.QTY71;
                                date72.Content = DD.QTY72;
                                date73.Content = DD.QTY73;
                                date74.Content = DD.QTY74;
                                date75.Content = DD.QTY75;
                                date76.Content = DD.QTY76;
                                date77.Content = DD.QTY77;
                                date78.Content = DD.QTY78;
                                date79.Content = DD.QTY79;
                                date80.Content = DD.QTY80;

                                date81.Content = DD.QTY81;
                                date82.Content = DD.QTY82;
                                date83.Content = DD.QTY83;
                                date84.Content = DD.QTY84;
                                date85.Content = DD.QTY85;
                                date86.Content = DD.QTY86;
                                date87.Content = DD.QTY87;
                                date88.Content = DD.QTY88;
                                date89.Content = DD.QTY89;
                                date90.Content = DD.QTY90;

                                date91.Content = DD.QTY91;
                                date92.Content = DD.QTY92;
                                date93.Content = DD.QTY93;
                                date94.Content = DD.QTY94;
                                date95.Content = DD.QTY95;


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
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N1}", obj);
        }

        private void dgdmainColumHeaderClick(object sender, RoutedEventArgs e)
        {
            foreach (DataGridColumn column in dgdmain.Columns)
            {
                column.CanUserSort = false;
            }
        }

        private void Dgdmain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

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
    }



    class Win_MIS_CustomArticleInSum_DD_Q_CodeView : BaseView
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
        public string QTY13 { get; internal set; }
        public string AMOUNT13 { get; internal set; }
        public string QTY14 { get; internal set; }
        public string AMOUNT14 { get; internal set; }
        public string QTY15 { get; internal set; }
        public string AMOUNT15 { get; internal set; }
        public string QTY16 { get; internal set; }
        public string AMOUNT16 { get; internal set; }
        public string AMOUNT19 { get; internal set; }
        public string QTY19 { get; internal set; }
        public string AMOUNT18 { get; internal set; }
        public string QTY18 { get; internal set; }
        public string AMOUNT17 { get; internal set; }
        public string QTY17 { get; internal set; }
        public string QTY20 { get; internal set; }
        public string QTY21 { get; internal set; }
        public string AMOUNT20 { get; internal set; }
        public string AMOUNT21 { get; internal set; }
        public string QTY22 { get; internal set; }
        public string AMOUNT22 { get; internal set; }
        public string QTY23 { get; internal set; }
        public string AMOUNT23 { get; internal set; }
        public string QTY24 { get; internal set; }
        public string AMOUNT24 { get; internal set; }
        public string QTY25 { get; internal set; }
        public string AMOUNT25 { get; internal set; }
        public string QTY26 { get; internal set; }
        public string AMOUNT26 { get; internal set; }
        public string QTY27 { get; internal set; }
        public string AMOUNT27 { get; internal set; }
        public string QTY28 { get; internal set; }
        public string AMOUNT28 { get; internal set; }
        public string QTY29 { get; internal set; }
        public string AMOUNT29 { get; internal set; }
        public string QTY30 { get; internal set; }
        public string AMOUNT30 { get; internal set; }
        public string QTY31 { get; internal set; }
        public string QTY32 { get; internal set; }
        public string AMOUNT31 { get; internal set; }
        public string AMOUNT37 { get; internal set; }
        public string QTY37 { get; internal set; }
        public string AMOUNT36 { get; internal set; }
        public string QTY36 { get; internal set; }
        public string AMOUNT35 { get; internal set; }
        public string QTY35 { get; internal set; }
        public string AMOUNT34 { get; internal set; }
        public string AMOUNT33 { get; internal set; }
        public string QTY33 { get; internal set; }
        public string AMOUNT32 { get; internal set; }
        public string AMOUNT38 { get; internal set; }
        public string QTY38 { get; internal set; }
        public string QTY34 { get; internal set; }
        public string QTY39 { get; internal set; }
        public string AMOUNT39 { get; internal set; }
        public string QTY40 { get; internal set; }
        public string AMOUNT40 { get; internal set; }
        public string QTY41 { get; internal set; }
        public string AMOUNT41 { get; internal set; }
        public string QTY42 { get; internal set; }
        public string AMOUNT42 { get; internal set; }
        public string QTY44 { get; internal set; }
        public string AMOUNT43 { get; internal set; }
        public string QTY43 { get; internal set; }
        public string AMOUNT44 { get; internal set; }
        public string QTY45 { get; internal set; }
        public string AMOUNT45 { get; internal set; }
        public string QTY46 { get; internal set; }
        public string AMOUNT46 { get; internal set; }

        public string AMOUNT95 { get; internal set; }
        public string QTY95 { get; internal set; }
        public string AMOUNT94 { get; internal set; }
        public string QTY94 { get; internal set; }
        public string AMOUNT93 { get; internal set; }
        public string QTY93 { get; internal set; }
        public string AMOUNT92 { get; internal set; }
        public string QTY92 { get; internal set; }
        public string AMOUNT91 { get; internal set; }
        public string QTY91 { get; internal set; }
        public string QTY88 { get; internal set; }

        public string QTY47 { get; internal set; }
        public string AMOUNT47 { get; internal set; }
        public string QTY48 { get; internal set; }
        public string AMOUNT48 { get; internal set; }
        public string QTY49 { get; internal set; }
        public string AMOUNT49 { get; internal set; }
        public string QTY50 { get; internal set; }
        public string AMOUNT50 { get; internal set; }
        public string QTY51 { get; internal set; }
        public string AMOUNT51 { get; internal set; }
        public string QTY52 { get; internal set; }
        public string AMOUNT52 { get; internal set; }
        public string QTY53 { get; internal set; }
        public string AMOUNT53 { get; internal set; }
        public string QTY54 { get; internal set; }
        public string AMOUNT54 { get; internal set; }
        public string QTY55 { get; internal set; }
        public string AMOUNT55 { get; internal set; }
        public string QTY56 { get; internal set; }
        public string AMOUNT56 { get; internal set; }
        public string QTY57 { get; internal set; }
        public string AMOUNT57 { get; internal set; }
        public string QTY58 { get; internal set; }
        public string AMOUNT58 { get; internal set; }
        public string AMOUNT59 { get; internal set; }
        public string QTY59 { get; internal set; }
        public string QTY60 { get; internal set; }
        public string AMOUNT60 { get; internal set; }
        public string QTY61 { get; internal set; }
        public string AMOUNT61 { get; internal set; }
        public string QTY62 { get; internal set; }
        public string AMOUNT62 { get; internal set; }
        public string QTY63 { get; internal set; }
        public string AMOUNT63 { get; internal set; }
        public string QTY64 { get; internal set; }
        public string AMOUNT64 { get; internal set; }
        public string QTY65 { get; internal set; }
        public string AMOUNT65 { get; internal set; }
        public string QTY66 { get; internal set; }
        public string AMOUNT66 { get; internal set; }
        public string QTY67 { get; internal set; }
        public string AMOUNT67 { get; internal set; }
        public string QTY68 { get; internal set; }
        public string AMOUNT68 { get; internal set; }
        public string QTY69 { get; internal set; }
        public string AMOUNT69 { get; internal set; }
        public string QTY70 { get; internal set; }
        public string AMOUNT70 { get; internal set; }
        public string QTY71 { get; internal set; }
        public string AMOUNT90 { get; internal set; }
        public string QTY90 { get; internal set; }
        public string AMOUNT89 { get; internal set; }
        public string QTY89 { get; internal set; }
        public string AMOUNT88 { get; internal set; }
        public string AMOUNT87 { get; internal set; }
        public string QTY87 { get; internal set; }
        public string AMOUNT86 { get; internal set; }
        public string QTY86 { get; internal set; }
        public string AMOUNT85 { get; internal set; }
        public string QTY85 { get; internal set; }
        public string AMOUNT84 { get; internal set; }
        public string QTY84 { get; internal set; }
        public string AMOUNT83 { get; internal set; }
        public string QTY83 { get; internal set; }
        public string AMOUNT82 { get; internal set; }
        public string QTY82 { get; internal set; }
        public string AMOUNT81 { get; internal set; }
        public string QTY81 { get; internal set; }
        public string AMOUNT80 { get; internal set; }
        public string QTY80 { get; internal set; }
        public string AMOUNT79 { get; internal set; }
        public string QTY79 { get; internal set; }
        public string AMOUNT78 { get; internal set; }
        public string QTY78 { get; internal set; }
        public string AMOUNT77 { get; internal set; }
        public string QTY77 { get; internal set; }
        public string AMOUNT76 { get; internal set; }
        public string QTY76 { get; internal set; }
        public string AMOUNT75 { get; internal set; }
        public string QTY75 { get; internal set; }
        public string AMOUNT71 { get; internal set; }
        public string QTY72 { get; internal set; }
        public string AMOUNT72 { get; internal set; }
        public string QTY73 { get; internal set; }
        public string AMOUNT73 { get; internal set; }
        public string QTY74 { get; internal set; }
        public string AMOUNT74 { get; internal set; }
        public string TotalColor { get; set; }
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }


    }
}
