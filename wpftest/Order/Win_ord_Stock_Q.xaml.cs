using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;
using WizMes_WellMade.PopUp;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_ord_Stock_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_ord_Stock_Q : UserControl
    {
        public Win_ord_Stock_Q()
        {
            InitializeComponent();
        }
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        string stDate = string.Empty;
        string stTime = string.Empty;

        // 엑셀 활용 용도 (프린트)
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();
        DataTable DT;


        // 첫 로드시.
        private void Win_ord_Stock_Q_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            First_Step();
            ComboBoxSetting();
            //제품으로 고정
            cboArticleGroupSrh.SelectedIndex = 3;
        }

        #region 첫단계 / 날짜버튼 세팅 / 조회용 체크박스 세팅

        // 첫 단계
        private void First_Step()
        {
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            chkInOutDate.IsChecked = true;

            chkIn_NotApprovedIncloudSrh.IsChecked = true;
            chkAutoInOutItemsIncloudSrh.IsChecked = true;

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
            else
            {
                DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

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

        // 입출일자
        private void chkInOutDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkInOutDate.IsChecked == true)
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
        //입출일자
        private void chkInOutDate_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkInOutDate.IsChecked == true)
            {
                chkInOutDate.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkInOutDate.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }
  
     
      
      
  
      

        #endregion


        #region 콤보박스 세팅
        // 콤보박스 세팅.
        private void ComboBoxSetting()
        {
            cboArticleGroupSrh.Items.Clear();
            cboWareHouseSrh.Items.Clear();
            cboInGbnSrh.Items.Clear();
            cboOutGbnSrh.Items.Clear();
            cboSupplyTypeSrh.Items.Clear();

            ObservableCollection<CodeView> cbArticleGroup = ComboBoxUtil.Instance.Gf_DB_MT_sArticleGrp();
            ObservableCollection<CodeView> cbWareHouse = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "LOC", "Y", "", "");
            ObservableCollection<CodeView> cbInGbn = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "OCD", "Y", "", "");
            ObservableCollection<CodeView> cbSupplyType = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "CMMASPLTYPE", "Y", "", "");

            this.cboArticleGroupSrh.ItemsSource = cbArticleGroup;
            this.cboArticleGroupSrh.DisplayMemberPath = "code_name";
            this.cboArticleGroupSrh.SelectedValuePath = "code_id";
            this.cboArticleGroupSrh.SelectedIndex = 0;

            this.cboWareHouseSrh.ItemsSource = cbWareHouse;
            this.cboWareHouseSrh.DisplayMemberPath = "code_name";
            this.cboWareHouseSrh.SelectedValuePath = "code_id";
            this.cboWareHouseSrh.SelectedIndex = 0;

            this.cboInGbnSrh.ItemsSource = cbInGbn;
            this.cboInGbnSrh.DisplayMemberPath = "code_id_plus_code_name";
            this.cboInGbnSrh.SelectedValuePath = "code_id";
            this.cboInGbnSrh.SelectedIndex = 0;

            this.cboOutGbnSrh.ItemsSource = cbInGbn;
            this.cboOutGbnSrh.DisplayMemberPath = "code_id_plus_code_name";
            this.cboOutGbnSrh.SelectedValuePath = "code_id";
            this.cboOutGbnSrh.SelectedIndex = 0;

            this.cboSupplyTypeSrh.ItemsSource = cbSupplyType;
            this.cboSupplyTypeSrh.DisplayMemberPath = "code_name";
            this.cboSupplyTypeSrh.SelectedValuePath = "code_id";
            this.cboSupplyTypeSrh.SelectedIndex = 0;
        }

        #endregion


        #region 조회 , 조회용 프로시저 
        // 조회.
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            using (Loading ld = new Loading(beSearch))
            {
                ld.ShowDialog();
            }
        }

        private void beSearch()
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            if(lib.DatePickerCheck(dtpFromDate, dtpToDate, chkInOutDate))
            {

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    //로직
                    FillGrid();
                }), System.Windows.Threading.DispatcherPriority.Background);
            }

            btnSearch.IsEnabled = true;
        }

        private void FillGrid()
        {
            dgdStock.Items.Clear();
            dgdStockQTotal.Items.Clear();

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nChkDate", chkInOutDate.IsChecked == true ?  1: 0);
                sqlParameter.Add("sSDate", dtpFromDate.SelectedDate?.ToString("yyyyMMdd") ?? string.Empty);
                sqlParameter.Add("sEDate", dtpToDate.SelectedDate?.ToString("yyyyMMdd") ?? string.Empty);
                sqlParameter.Add("nChkCustom", chkCustomIDSrh.IsChecked == true ? 1:0);
                sqlParameter.Add("sCustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag?.ToString() ?? string.Empty : string.Empty);

                sqlParameter.Add("nChkArticleID", chkArticleIDSrh.IsChecked == true ? 1:0 );
                sqlParameter.Add("sArticleID", chkArticleIDSrh.IsChecked == true? txtArticleIDSrh.Tag?.ToString() ?? string.Empty : string.Empty); 
                sqlParameter.Add("nChkOrder", 0);               //화면내 조건없음 --추가시에 적용 -- orderID 또는 orderNo
                sqlParameter.Add("sOrder", string.Empty);       //화면내 조건없음 --추가시에 적용
                sqlParameter.Add("ArticleGrpID", cboArticleGroupSrh.SelectedValue?.ToString() ?? string.Empty); //제품구분

                sqlParameter.Add("sFromLocID", cboWareHouseSrh.SelectedValue?.ToString() ?? string.Empty); //창고
                sqlParameter.Add("sToLocID", string.Empty);                                                //창고위치
                sqlParameter.Add("nChkOutClss", chkOutGbnSrh.IsChecked == true ? 1:0);                     //출고구분
                sqlParameter.Add("sOutClss", chkOutGbnSrh.IsChecked == true ? cboOutGbnSrh.SelectedValue?.ToString() ?? string.Empty : string.Empty);
                sqlParameter.Add("nChkInClss", chkInGbnSrh.IsChecked == true ? 1:0);                        //입고구분
                sqlParameter.Add("sInClss", chkInGbnSrh.IsChecked == true ? cboInGbnSrh.SelectedValue?.ToString() ?? string.Empty : string.Empty);

                sqlParameter.Add("nChkReqID", 0);               //발주번호
                sqlParameter.Add("sReqID", string.Empty);       //발주번호
                sqlParameter.Add("incNotApprovalYN", chkIn_NotApprovedIncloudSrh.IsChecked == true ? 1:0);
                sqlParameter.Add("incAutoInOutYN", chkAutoInOutItemsIncloudSrh.IsChecked == true? 1:0);

                sqlParameter.Add("sArticleIDS", string.Empty);
                sqlParameter.Add("sMissSafelyStockQty", chkOptimumStockBelowSeeSrh.IsChecked == true ? "Y" : "");
                sqlParameter.Add("sProductYN", "Y");
                sqlParameter.Add("nMainItem", chkMainInterestItemsSeeSrh.IsChecked == true ? 1:0);
                sqlParameter.Add("nCustomItem", chkRegistItemsByCustomerSrh.IsChecked == true ? 1:0);

                sqlParameter.Add("nSupplyType", chkSupplyTypeSrh.IsChecked == true ? 1:0);
                sqlParameter.Add("sSupplyType", chkCustomIDSrh.IsChecked == true ? cboSupplyTypeSrh.SelectedValue?.ToString() ?? string.Empty : string.Empty);

                sqlParameter.Add("JaturiNoYN", ""); //이건 뭐하는 거지
                sqlParameter.Add("nBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0); //일단 빈값
                sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag?.ToString() ?? string.Empty : string.Empty);

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Subul_sStockList", sqlParameter, true, "R");
                DataTable dt = null;

                if (ds != null && ds.Tables.Count > 0)
                {
                    dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회결과가 없습니다.");
                        return;
                    }

                    else
                    {
                        int NUM = 1;
                        int totalInitStockQty = 0;
                        int totalStuffQty = 0;
                        int TotalOutQty = 0;
                        int TotalStockQty = 0;

                        DT = dt;
                     
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow item in drc)
                        {
                            if (((item["InitStockQty"].ToString().Split('.')[0].Trim() == "") &&
                                    (item["StuffQty"].ToString().Split('.')[0].Trim() == "") &&
                                    (item["OutQty"].ToString().Split('.')[0].Trim() == "") &&
                                    (item["StockQty"].ToString().Split('.')[0].Trim() == ""))
                                    ||
                                    (item["cls"].ToString() == "3"))
                            {
                                continue;
                            }

                            if ((item["cls"].ToString() != "3") && (item["cls"].ToString() != "4") &&
                                (Convert.ToInt32(item["StockQty"].ToString().Split('.')[0].Trim()) <
                                    Convert.ToInt32(item["NeedstockQty"].ToString().Split('.')[0].Trim())))
                            {

                                // 적정재고 미달건으로 빨간색 재고량에 빨간색 글자색을 입혀주어야 한다.
                                var Win_ord_Stock_Q_Insert_red = new Win_ord_Stock_Q_View()
                                {
                                    NUM = NUM.ToString(),

                                    cls = item["cls"].ToString(),
                                    BuyerArticleNo = item["BuyerArticleNo"].ToString(),
                                    ArticleID = item["ArticleID"].ToString(),
                                    Article = item["Article"].ToString(),
                                    LocID = item["LocID"].ToString(),
                                    LocName = item["LocName"].ToString(),

                                    InitStockRoll = item["InitStockRoll"].ToString(),
                                    InitStockQty = String.Format("{0:#,##0}", Convert.ToDouble(item["InitStockQty"].ToString())),
                                    StuffRoll = item["StuffRoll"].ToString(),
                                    StuffQty = String.Format("{0:#,##0}", Convert.ToDouble(item["StuffQty"].ToString().Split('.')[0].Trim())),
                                    OutRoll = item["OutRoll"].ToString(),

                                    OutQty = String.Format("{0:#,##0}", Convert.ToDouble(item["OutQty"].ToString().Split('.')[0].Trim())),
                                    StockQty = String.Format("{0:#,##0}", Convert.ToDouble(item["StockQty"].ToString())),
                                    UnitClss = item["UnitClss"].ToString(),
                                    UnitClssName = item["UnitClssName"].ToString(),
                                    NeedstockQty = String.Format("{0:#,##0}", Convert.ToDouble(item["NeedstockQty"].ToString().Split('.')[0].Trim())),

                                    OverQty = String.Format("{0:#,##0}", Convert.ToDouble(item["OverQty"].ToString())),
                                    StockRate = item["StockRate"].ToString().Split('.')[0].Trim(),
                                    FontRed = "true",
                                    ColorGreen = "false"

                                };
                                dgdStock.Items.Add(Win_ord_Stock_Q_Insert_red);
                                NUM++;

                            }

                            else if (item["cls"].ToString() == "4")
                            {
                                var Win_ord_Stock_Q_Insert = new Win_ord_Stock_Q_View()
                                {
                                    NUM = NUM.ToString(),

                                    cls = item["cls"].ToString(),
                                    BuyerArticleNo = "",
                                    ArticleID = "",
                                    Article = "총계",
                                    LocID = item["LocID"].ToString(),
                                    LocName = "",

                                    InitStockRoll = item["InitStockRoll"].ToString(),
                                    InitStockQty = String.Format("{0:#,##0}", Convert.ToDouble(item["InitStockQty"].ToString())),
                                    StuffRoll = item["StuffRoll"].ToString(),
                                    StuffQty = String.Format("{0:#,##0}", Convert.ToDouble(item["StuffQty"].ToString().Split('.')[0].Trim())),
                                    OutRoll = item["OutRoll"].ToString(),

                                    OutQty = String.Format("{0:#,##0}", Convert.ToDouble(item["OutQty"].ToString().Split('.')[0].Trim())),
                                    StockQty = String.Format("{0:#,##0}", Convert.ToDouble(item["StockQty"].ToString())),
                                    UnitClss = item["UnitClss"].ToString(),
                                    UnitClssName = item["UnitClssName"].ToString(),
                                    NeedstockQty = "",

                                    OverQty = String.Format("{0:#,##0}", Convert.ToDouble(item["OverQty"].ToString())),
                                    StockRate = item["StockRate"].ToString().Split('.')[0].Trim(),
                                    FontRed = "false",
                                    ColorGreen = "true"                                   

                                };

                                totalInitStockQty += lib.RemoveComma(item["initStockQty"].ToString(), 0);
                                totalStuffQty += lib.RemoveComma(item["StuffQty"].ToString(), 0);
                                TotalOutQty += lib.RemoveComma(item["OutQty"].ToString(), 0);
                                TotalStockQty += lib.RemoveComma(item["StockQty"].ToString(), 0);

                                dgdStock.Items.Add(Win_ord_Stock_Q_Insert);
                                NUM++;
                            }

                            else
                            {
                                var Win_ord_Stock_Q_Insert = new Win_ord_Stock_Q_View()
                                {
                                    NUM = NUM.ToString(),

                                    cls = item["cls"].ToString(),
                                    BuyerArticleNo = item["BuyerArticleNo"].ToString(),
                                    ArticleID = item["ArticleID"].ToString(),
                                    Article = item["Article"].ToString(),
                                    LocID = item["LocID"].ToString(),
                                    LocName = item["LocName"].ToString(),

                                    InitStockRoll = String.Format("{0:#,##0}", Convert.ToDouble(item["InitStockRoll"].ToString())),
                                    InitStockQty = String.Format("{0:#,##0}", Convert.ToDouble(item["InitStockQty"].ToString())),
                                    StuffRoll = String.Format("{0:#,##0}", Convert.ToDouble(item["StuffRoll"].ToString())),
                                    StuffQty = String.Format("{0:#,##0}", Convert.ToDouble(item["StuffQty"].ToString().Split('.')[0].Trim())),
                                    OutRoll = item["OutRoll"].ToString(),

                                    OutQty = String.Format("{0:#,##0}", Convert.ToDouble(item["OutQty"].ToString().Split('.')[0].Trim())),
                                    StockQty = String.Format("{0:#,##0}", Convert.ToDouble(item["StockQty"].ToString())),
                                    UnitClss = item["UnitClss"].ToString(),
                                    UnitClssName = item["UnitClssName"].ToString(),
                                    NeedstockQty = String.Format("{0:#,##0}", Convert.ToDouble(item["NeedstockQty"].ToString().Split('.')[0].Trim())),

                                    OverQty = String.Format("{0:#,##0}", Convert.ToDouble(item["OverQty"].ToString())),
                                    StockRate = item["StockRate"].ToString().Split('.')[0].Trim(),
                                    FontRed = "false",
                                    ColorGreen = "false"
                                };
                                dgdStock.Items.Add(Win_ord_Stock_Q_Insert);
                                NUM++;
                            }
                        }

                        if(dgdStock.Items.Count > 0)
                        {
                            var stockTotal = new Win_ord_Stock_Q_View_Total
                            {
                                TotalInitStockQty = stringFormatN0(totalInitStockQty),
                                TotalOutQty = stringFormatN0(TotalOutQty),
                                TotalStuffQty = stringFormatN0(totalStuffQty),
                                TotalStockQty = stringFormatN0(TotalStockQty),
                            };

                            dgdStockQTotal.Items.Add(stockTotal);
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


        // 닫기 버튼클릭.
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
        // 엑셀버튼 클릭
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdStock.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            Lib lib3 = new Lib();
            System.Data.DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "메인 그리드";
            lst[2] = dgdStock.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdStock.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    //MessageBox.Show("대분류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib3.DataGridToDTinHidden(dgdStock);
                    else
                        dt = lib3.DataGirdToDataTable(dgdStock);

                    Name = dgdStock.Name;

                    if (lib3.GenerateExcel(dt, Name))
                    {
                        lib3.excel.Visible = true;
                        lib3.ReleaseExcelObject(lib3.excel);
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

            lib3 = null;
        }



        #endregion


        #region 플러스 파인더 
        //플러스파인더 _ 거래처_클릭.
        private void btnCustomIDSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomIDSrh, 0, "");
        }
        //플러스파인더 _ 거래처_키다운.
        private void txtCustomIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            pf.ReturnCode(txtCustomIDSrh, 0, "");
        }

        //플러스파인더 _ 품명_클릭. (품번이 검색되도록 수정, 2020.03.18, 장가빈)
        private void btnArticleIDSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticleIDSrh, 77, txtArticleIDSrh.Text);
        }
        //플러스파인더 _ 품명_키다운
        private void txtArticleIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticleIDSrh, 77, txtArticleIDSrh.Text);
            }
        }

        #endregion


        #region 인쇄

        // 인쇄버튼 클릭
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = true;
            menu.IsOpen = true;
        }

        // 인쇄 - 미리보기 클릭.
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {
            if (dgdStock.Items.Count == 0)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }
            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            lib.Delay(1000);

            PrintWork(true);
            msg.Visibility = Visibility.Hidden;
        }
        // 인쇄 서브메뉴2. 바로인쇄
        private void menuRighPrint_Click(object sender, RoutedEventArgs e)
        {
            if (dgdStock.Items.Count == 0)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }
            DataStore.Instance.InsertLogByForm(this.GetType().Name, "P");
            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            lib.Delay(1000);

            PrintWork(false);
            msg.Visibility = Visibility.Hidden;
        }
        //인쇄 서브메뉴3. 그냥 닫기
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }


        // 실제 엑셀작업 스타트.
        private void PrintWork(bool previewYN)
        {
            Lib lib2 = new Lib();

            try
            {
                excelapp = new Microsoft.Office.Interop.Excel.Application();

                string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\재고조회(영업관리).xls";
                workbook = excelapp.Workbooks.Add(MyBookPath);
                worksheet = workbook.Sheets["Form"];

                if (chkInOutDate.IsChecked == true)             // 기간.
                {
                    string fyyyy = dtpFromDate.Text.Substring(0, 4) + "년";
                    string fmm = dtpFromDate.Text.Substring(5, 2) + "월";
                    string fdd = dtpFromDate.Text.Substring(8, 2) + "일";

                    string tyyyy = dtpToDate.Text.Substring(0, 4) + "년";
                    string tmm = dtpToDate.Text.Substring(5, 2) + "월";
                    string tdd = dtpToDate.Text.Substring(8, 2) + "일";

                    workrange = worksheet.get_Range("D4");//셀 범위 지정
                    workrange.Value2 = fyyyy + fmm + fdd + "~" + tyyyy + tmm + tdd;
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    workrange.Font.Size = 11;
                }

                if ((chkWareHouseSrh.IsChecked == true) && (cboWareHouseSrh.SelectedIndex != -1))     // 창고정보.
                {
                    workrange = worksheet.get_Range("D3");//셀 범위 지정
                    workrange.Value2 = ((WizMes_WellMade.CodeView)cboWareHouseSrh.SelectedItem).code_name.ToString();
                    workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    workrange.Font.Size = 11;
                }

                workrange = worksheet.get_Range("AE46");//셀 범위 지정
                workrange.Value2 = "주식회사 지엘에스";
                workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                workrange.Font.Size = 11;

                workrange = worksheet.get_Range("AH4", "AO4");//셀 범위 지정       //발행일자.
                workrange.Value2 = DateTime.Now.ToString("yyyy-MM-dd");
                workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                workrange.Font.Size = 11;

                /////////////////////////////////
                int Page = 0;
                int DataCount = 0;
                int copyLine = 0;

                copysheet = workbook.Sheets["Form"];
                pastesheet = workbook.Sheets["Print"];

                string str_article = string.Empty;
                string str_locname = string.Empty;
                string str_initstockqty = string.Empty;
                string str_stuffqty = string.Empty;
                string str_outqty = string.Empty;
                string str_stockqty = string.Empty;
                string str_unitclssname = string.Empty;
                string str_needstockqty = string.Empty;
                string str_overqty = string.Empty;
                string str_stockrate = string.Empty;

                while (DT.Rows.Count - 1 > DataCount)
                {
                    Page++;
                    if (Page != 1) { DataCount++; }           // +1. 
                    copyLine = (Page - 1) * 48;
                    copysheet.Select();
                    copysheet.UsedRange.Copy();
                    pastesheet.Select();
                    workrange = pastesheet.Cells[copyLine + 1, 1];
                    workrange.Select();
                    pastesheet.Paste();                 // 프린트 열에 page번째 항목 복사완료.


                    int j = 0;
                    for (int i = DataCount; i < DT.Rows.Count; i++)
                    {

                        if (((DT.Rows[i]["InitStockQty"].ToString().Split('.')[0].Trim() == "") &&
                                    (DT.Rows[i]["StuffQty"].ToString().Split('.')[0].Trim() == "") &&
                                    (DT.Rows[i]["OutQty"].ToString().Split('.')[0].Trim() == "") &&
                                    (DT.Rows[i]["StockQty"].ToString().Split('.')[0].Trim() == ""))
                                    ||
                                    (DT.Rows[i]["cls"].ToString() == "3"))
                        {
                            continue;
                        }

                        if (j == 39) { break; }
                        int insertline = copyLine + 7 + j;


                        str_article = DT.Rows[i]["Article"].ToString();
                        str_locname = DT.Rows[i]["LocName"].ToString();
                        str_initstockqty = DT.Rows[i]["InitStockQty"].ToString();
                        str_stuffqty = DT.Rows[i]["StuffQty"].ToString();
                        str_outqty = DT.Rows[i]["OutQty"].ToString();
                        str_stockqty = DT.Rows[i]["StockQty"].ToString();
                        str_unitclssname = DT.Rows[i]["UnitClssName"].ToString();
                        str_needstockqty = DT.Rows[i]["NeedstockQty"].ToString();
                        str_overqty = DT.Rows[i]["OverQty"].ToString();
                        str_stockrate = DT.Rows[i]["StockRate"].ToString();

                        if (str_article == "zzzzzz")
                        {
                            str_article = "총계";
                            str_locname = "";
                        }


                        workrange = pastesheet.get_Range("A" + insertline, "G" + insertline);    //품명
                        workrange.Value2 = str_article;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 11;

                        workrange = pastesheet.get_Range("H" + insertline, "K" + insertline);    //창고
                        workrange.Value2 = str_locname;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 11;

                        workrange = pastesheet.get_Range("L" + insertline, "O" + insertline);    //이월
                        workrange.Value2 = str_initstockqty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 11;

                        workrange = pastesheet.get_Range("P" + insertline, "S" + insertline);    //입고
                        workrange.Value2 = str_stuffqty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 11;

                        workrange = pastesheet.get_Range("T" + insertline, "W" + insertline);    //출고
                        workrange.Value2 = str_outqty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 11;

                        workrange = pastesheet.get_Range("X" + insertline, "AA" + insertline);    //재고
                        workrange.Value2 = str_stockqty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 11;

                        workrange = pastesheet.get_Range("AB" + insertline, "AC" + insertline);    //단위
                        workrange.Value2 = str_unitclssname;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 11;

                        workrange = pastesheet.get_Range("AD" + insertline, "AG" + insertline);    //적정재고량
                        workrange.Value2 = str_needstockqty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 11;

                        workrange = pastesheet.get_Range("AH" + insertline, "AK" + insertline);    //과부족
                        workrange.Value2 = str_overqty;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 11;

                        workrange = pastesheet.get_Range("AL" + insertline, "AO" + insertline);    //재고율
                        workrange.Value2 = str_stockrate;
                        workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        workrange.Font.Size = 11;

                        //라인 색깔
                        if (str_article == "총계")
                        {
                            workrange = pastesheet.get_Range("A" + insertline, "AO" + insertline);
                            workrange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                        }

                        DataCount = i;
                        j++;
                    }

                }

                pastesheet.PageSetup.Zoom = 96;

                excelapp.Visible = true;
                msg.Hide();
                if (previewYN == true)
                {
                    pastesheet.PrintPreview();
                }
                else
                {
                    pastesheet.PrintOutEx();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {

                //excelapp.Quit();
                lib2.ReleaseExcelObject(workbook);
                lib2.ReleaseExcelObject(worksheet);
                lib2.ReleaseExcelObject(pastesheet);
                lib2.ReleaseExcelObject(excelapp);
                lib2 = null;
            }


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

        //검색조건 -품번 - 클릭
        private void btnBuyerArticleNoSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtBuyerArticleNoSrh, 76, txtBuyerArticleNoSrh.Text);
        }

        //검색조건 - 품번 -키다운
        private void txtBuyerArticleNoSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtBuyerArticleNoSrh, 76, txtBuyerArticleNoSrh.Text);
            }
        }

        private void CommonControl_Click(object sender, MouseButtonEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private void CommonControl_Click(object sender, RoutedEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }
    }


    class Win_ord_Stock_Q_View : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }

        // 조회용
        public string NUM { get; set; }

        public string cls { get; set; }
        public string ArticleID { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Article { get; set; }
        public string Sabun { get; set; }

        public string LocID { get; set; }
        public string LocName { get; set; }

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

        public string OverQty { get; set; }
        public string StockRate { get; set; }

        public string FontRed { get; set; }
        public string ColorGreen { get; set; }


    }

    public class Win_ord_Stock_Q_View_Total : BaseView
    {
        public string TotalInitStockQty { get; set; }
        public string TotalStuffQty { get; set; }
        public string TotalOutQty { get; set; }
        public string TotalStockQty { get; set; }
    }
}
