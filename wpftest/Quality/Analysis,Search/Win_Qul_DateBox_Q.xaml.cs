using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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
    /// 
    /// </summary>
    public partial class Win_Qul_DateBox_Q : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Win_Qul_DateBox_QView WIDV = new Win_Qul_DateBox_QView();
        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        // 엑셀 활용 용도 (프린트)


        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();
        //(기다림 알림 메시지창)

        System.Data.DataTable DT;

        public Win_Qul_DateBox_Q()
        {
            InitializeComponent();
            this.DataContext = WIDV;
        }

        private void Window_InsDateBox_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            First_Step();
            ComboBoxSetting();
        }

        #region 시작 첫단계 // 콤보박스 세팅 // 조회용 각종 체크박스 활성화 // 일자버튼

        // 시작 첫 단계.
        private void First_Step()
        {
            chkInspectDay.IsChecked = true;
            dtpFromDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpToDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            dtpFromDate.IsEnabled = true;
            dtpToDate.IsEnabled = true;
            rbnManageNumberSrh.IsChecked = true;

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




        //출고일자(날짜) 체크
        private void chkInspectDay_Click(object sender, RoutedEventArgs e)
        {
            if (chkInspectDay.IsChecked == true)
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
        private void chkInspectDay_Click(object sender, MouseButtonEventArgs e)
        {
            if (chkInspectDay.IsChecked == true)
            {
                chkInspectDay.IsChecked = false;
                dtpFromDate.IsEnabled = false;
                dtpToDate.IsEnabled = false;
            }
            else
            {
                chkInspectDay.IsChecked = true;
                dtpFromDate.IsEnabled = true;
                dtpToDate.IsEnabled = true;
            }
        }

     

        private void rbnOrderNoSrh_Click(object sender, RoutedEventArgs e)
        {
            tblOrderID.Text = "Order NO";
        }

        private void rbnManageNumberSrh_Click(object sender, RoutedEventArgs e)
        {
            tblOrderID.Text = "관리번호";
        }

        // 콤보박스 목록 불러오기.
        private void ComboBoxSetting()
        {
            ObservableCollection<CodeView> cbFaultyGBN = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "DFGRP", "Y", "", "");

            this.cboDefectGrpIDSrh.ItemsSource = cbFaultyGBN;
            this.cboDefectGrpIDSrh.DisplayMemberPath = "code_name";
            this.cboDefectGrpIDSrh.SelectedValuePath = "code_id";
            this.cboDefectGrpIDSrh.SelectedIndex = 0;
        }

        #endregion


        #region 플러스 파인더
        //플러스 파인더

        // 거래처
        private void btnCustomIDSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtCustomIDSrh, 0, txtCustomIDSrh.Text);
        }

        private void btnArticleIDSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtArticleIDSrh, 77, txtArticleIDSrh.Text);
        }

        #endregion


        #region 조회 // 조회 프로시저

        // 검색. 조회버튼 클릭.
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            if(lib.DatePickerCheck(dtpFromDate, dtpToDate, chkInspectDay))
            {
                Dispatcher.BeginInvoke(new Action(() =>

                {
                    Thread.Sleep(2000);

                    //로직
                    FillGrid();

                }), System.Windows.Threading.DispatcherPriority.Background);



                Dispatcher.BeginInvoke(new Action(() =>

                {
                    btnSearch.IsEnabled = true;

                }), System.Windows.Threading.DispatcherPriority.Background);
            }

        }

        private void FillGrid()
        {

            dgdInspect.Items.Clear();
            dgdTotal.Items.Clear();


            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("ChkDate", chkInspectDay.IsChecked == true ? 1 : 0);
            sqlParameter.Add("sDate", chkInspectDay.IsChecked == true ?  dtpFromDate.SelectedDate?.ToString("yyyyMMdd") ?? string.Empty :string.Empty);
            sqlParameter.Add("eDate", chkInspectDay.IsChecked == true ? dtpToDate.SelectedDate?.ToString("yyyyMMdd") ?? string.Empty : string.Empty);

            sqlParameter.Add("ChkCustomID", chkCustomIDSrh.IsChecked ==true ? 1:0);
            sqlParameter.Add("CustomID", chkCustomIDSrh.IsChecked == true ? txtCustomIDSrh.Tag?.ToString() ?? string.Empty : string.Empty);

            sqlParameter.Add("ChkArticleID", chkArticleIDSrh.IsChecked == true ? 1 : 0);
            sqlParameter.Add("ArticleID", chkArticleIDSrh.IsChecked == true ? txtArticleIDSrh.Tag?.ToString() ?? string.Empty : string.Empty);

            sqlParameter.Add("ChkBuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? 1 : 0);
            sqlParameter.Add("BuyerArticleNo", chkBuyerArticleNoSrh.IsChecked == true ? txtBuyerArticleNoSrh.Tag?.ToString() ?? string.Empty : string.Empty);

            sqlParameter.Add("ChkDefectGrpID", chkDefectGrpIDSrh.IsChecked == true ? 1 : 0);
            sqlParameter.Add("DefectGrpID", chkDefectGrpIDSrh.IsChecked == true ? cboDefectGrpIDSrh.SelectedValue?.ToString() ?? string.Empty : string.Empty);

            sqlParameter.Add("ChkLabelID", chkCLabelSrh.IsChecked == true ? 1 : 0);
            sqlParameter.Add("LabelID", chkCLabelSrh.IsChecked == true ? txtCLabelSrh.Text : string.Empty);

            sqlParameter.Add("ChkBoxID", chkBLabelSrh.IsChecked == true ? 1 : 0);
            sqlParameter.Add("BoxID", chkBLabelSrh.IsChecked == true ? txtBLabelSrh.Text : string.Empty);

            sqlParameter.Add("ChkOrderID", chkOrderIDSrh.IsChecked != true ?  0 : rbnManageNumberSrh.IsChecked == true ? 1 : rbnOrderNoSrh.IsChecked == true ? 2 : 0);
            sqlParameter.Add("OrderID", txtOrderIDSrh.Text ?? string.Empty);

            DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Inspect_sInspectByBox", sqlParameter, true, "R");

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = null;
                dt = ds.Tables[0];
                DT = null;
                DT = dt;

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("조회결과가 없습니다.");
                    return;
                }
                else
                {
                    try
                    {
                        dgdInspect.Items.Clear();
                        DataRowCollection drc = dt.Rows;

                        int i = 0;
                        int rowCount = 0;
                        int totalRealQty = 0;
                        int totalCtrlQty = 0;
                        int totalDefectQty = 0;
                        foreach (DataRow item in drc)
                        {
                            i++;
                            var DefectInfo = new Win_Qul_DateBox_QView
                            {
                                num = i,
                                Gbn = item["Gbn"].ToString(),
                                ExamDate = lib.DateTypeHyphen(item["ExamDate"].ToString()),
                                KCustom = item["KCustom"].ToString(),
                                OrderID = item["OrderID"].ToString(),
                                OrderNo = item["OrderNo"].ToString(),
                                Model = item["Model"].ToString(),
                                Article = item["Article"].ToString(),
                                BuyerArticleNo = item["BuyerArticleNo"].ToString(),
                                OrderQty = stringFormatN0(item["OrderQty"]),
                                InBoxID = item["InBoxID"].ToString(),
                                LabelID = item["LabelID"].ToString(),
                                RealQty = stringFormatN0(item["RealQty"]),
                                CtrlQty = stringFormatN0(item["CtrlQty"]),
                                DefectQty = stringFormatN0(item["DefectQty"]),
                                UnitClssName = item["UnitClssName"].ToString(),
                                Name = item["Name"].ToString()

                            };

                            if (DefectInfo.Gbn.Equals("1"))
                            {
                                rowCount++;
                                dgdInspect.Items.Add(DefectInfo);
                            }
                            else if (DefectInfo.Gbn.Equals("2"))
                            {
                                DefectInfo.Color1 = true;
                                DefectInfo.OrderQty = string.Empty;
                                totalRealQty += lib.RemoveComma(item["RealQty"].ToString(),0);
                                totalCtrlQty += lib.RemoveComma(item["CtrlQty"].ToString(), 0);
                                totalDefectQty += lib.RemoveComma(item["DefectQty"].ToString(), 0);
                                dgdInspect.Items.Add(DefectInfo);
                            }                            
                        }

                        if(dgdInspect.Items.Count > 0)
                        {
                            var total = new Win_Qul_DateBox_QView_Total
                            {
                                TotalCount = rowCount,
                                TotalRealQty = totalRealQty,
                                TotalCtrlQty = totalCtrlQty,
                                TotalDefectQty = totalDefectQty
                            };

                            dgdTotal.Items.Add(total);
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
        }

        #endregion


        #region 엑셀

        // 엑셀버튼 클릭.
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdInspect.Items.Count < 1)
            {
                MessageBox.Show("먼저 검색해 주세요.");
                return;
            }

            Lib lib = new Lib();
            System.Data.DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[4];
            lst[0] = "메인 그리드";
            lst[2] = dgdInspect.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);

            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdInspect.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    //MessageBox.Show("대분류");
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdInspect);
                    else
                        dt = lib.DataGirdToDataTable(dgdInspect);

                    Name = dgdInspect.Name;
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

        }

        #endregion



        //닫기 기능.
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
                    }
                }
                i++;
            }
        }


        //정렬 버튼이벤트.
        private void btnMultiSort_Click(object sender, RoutedEventArgs e)
        {
            PopUp.MultiLevelSort MLS = new PopUp.MultiLevelSort(dgdInspect);
            MLS.ShowDialog();

            if (MLS.DialogResult.HasValue)
            {
                string targetSortProperty = string.Empty;
                int targetColIndex;
                dgdInspect.Items.SortDescriptions.Clear();

                for (int x = 0; x < MLS.ColName.Count; x++)
                {
                    targetSortProperty = MLS.SortingProperty[x];
                    targetColIndex = MLS.ColIndex[x];
                    var targetCol = dgdInspect.Columns[targetColIndex];

                    if (targetSortProperty == "UP")
                    {
                        dgdInspect.Items.SortDescriptions.Add(new SortDescription(targetCol.SortMemberPath, ListSortDirection.Ascending));
                        targetCol.SortDirection = ListSortDirection.Ascending;
                    }
                    else
                    {
                        dgdInspect.Items.SortDescriptions.Add(new SortDescription(targetCol.SortMemberPath, ListSortDirection.Descending));
                        targetCol.SortDirection = ListSortDirection.Descending;
                    }
                }
                dgdInspect.Refresh();
            }
        }

        private void txtArticleIDSrh_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtArticleIDSrh, 77, txtArticleIDSrh.Text);
            }
        }

        private void txtCustomIDSrh_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pf.ReturnCode(txtCustomIDSrh, 0, txtCustomIDSrh.Text);
            }   
        }

        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
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

      

        // 플러스파인더 _ 품번 찾기
        private void btnBuyerArticleNoSrh_Click(object sender, RoutedEventArgs e)
        {
            pf.ReturnCode(txtBuyerArticleNoSrh, 76, txtBuyerArticleNoSrh.Text);
        }

     

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
    }

    class Win_Qul_DateBox_QView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int num { get; set; }
        public string Gbn { get; set; }
        public string ExamDate { get; set; }
        public string KCustom { get; set; }
        public string OrderID { get; set; } 
        public string OrderNo { get; set; }
        public string Model { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; } 
        public string OrderQty { get; set; }
        public string InBoxID { get; set; }
        public string LabelID { get; set; }
        public string RealQty { get; set; }
        public string CtrlQty { get; set; }
        public string DefectQty { get; set; }
        public string UnitClssName { get; set; }
        public string Name { get; set; }
        public bool Color1 { get; set; } = false;
        public bool Color2 { get; set; } = false;

    }

    class Win_Qul_DateBox_QView_Total : BaseView
    {
        public int TotalCount { get; set; }
        public int TotalRealQty { get; set; }
        public int TotalCtrlQty { get; set; }
        public int TotalDefectQty { get; set; }
    }

}
