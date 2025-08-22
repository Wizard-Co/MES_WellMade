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
using System.Windows.Media.Imaging;
using WizMes_WellMade.PopUP;
using WPF.MDI;

/**************************************************************************************************
'** 프로그램명 : Win_Qul_DefectRepair_Q
'** 설명       : 불량발생 및 시정등록
'** 작성일자   : 2023.03.31
'** 작성자     : 장시영
'**------------------------------------------------------------------------------------------------
'**************************************************************************************************
' 변경일자  , 변경자, 요청자    , 요구사항ID      , 요청 및 작업내용
'**************************************************************************************************
' 2023.03.31, 장시영, 불량발생단계 중 전체는 제외
'**************************************************************************************************/

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_Qul_DefectRepair_Q.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_DefectRepair_Q : UserControl
    {
        DataTable DT_SYMPTOM = null;
        DataTable DT_REASON = null;
        DataTable DT_CUSTOM = null;
        Lib lib = new Lib();
        string strLastDay = string.Empty;

        //Image 변수 선언
        System.Windows.Controls.Image ImageData = new System.Windows.Controls.Image();

        public Win_Qul_DefectRepair_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);

            btnThisYear_Click(null, null);
            //cboOccurStepSrh.IsEnabled = false;
            SetComboBox();
        }

        private void SetComboBox()
        {
            ObservableCollection<CodeView> oveOrderForm = ComboBoxUtil.Instance.Gf_DB_CM_GetComCodeDataset(null, "QULSTEP", "Y", "", "");
            if (oveOrderForm.Count > 0)
                oveOrderForm.RemoveAt(0);

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
            dtpEDate.SelectedDate = lib.BringLastYearDatetimeContinue(dtpSDate.SelectedDate.Value)[1];
        }

        //금년
        private void btnThisYear_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisYearDatetimeFormat()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisYearDatetimeFormat()[1];
        }

        //최근6개월
        private void btnLastSixMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringLastSixMonthDateTimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringLastSixMonthDateTimeList()[1];
        }

        //금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        #endregion

        #region 체크박스 action

        //고객사
        private void lblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustom.IsChecked == true) { chkCustom.IsChecked = false; }
            else { chkCustom.IsChecked = true; }
        }

        //고객사
        private void chkCustom_Checked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = true;
            btnPfCustom.IsEnabled = true;
        }

        //고객사
        private void chkCustom_Unchecked(object sender, RoutedEventArgs e)
        {
            txtCustom.IsEnabled = false;
            btnPfCustom.IsEnabled = false;
        }

        //고객사
        private void txtCustom_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
            }
        }

        //고객사
        private void btnPfCustom_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtCustom, (int)Defind_CodeFind.DCF_CUSTOM, "");
        }

        //품명
        private void lblArticle_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticle.IsChecked == true) { chkArticle.IsChecked = false; }
            else { chkArticle.IsChecked = true; }
        }

        //품명
        private void chkArticle_Checked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = true;
            btnPfArticle.IsEnabled = true;
        }

        //품명
        private void chkArticle_Unchecked(object sender, RoutedEventArgs e)
        {
            txtArticle.IsEnabled = false;
            btnPfArticle.IsEnabled = false;
        }

        //품명(품번으로 변경요청, 2020.03.23, 장가빈) 
        private void txtArticle_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticle, 76, txtArticle.Text);
            }
        }

        //품명
        private void btnPfArticle_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticle, 76, txtArticle.Text);
        }

        #endregion

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {

            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                TabItem NowTI = tabconGrid.SelectedItem as TabItem;

                re_Search(0, NowTI);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);


        }

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

        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {

        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            TabItem NowTI = tabconGrid.SelectedItem as TabItem;
            DataTable dt = null;
            string Name = string.Empty;
            ExportExcelxaml ExpExc = null;

            if (NowTI.Header.ToString().Equals("전체"))
            {
                string[] lstAll = new string[6];
                lstAll[0] = "불량시정건_전체_불량유형";
                lstAll[1] = "불량시정건_전체_불량원인";
                lstAll[2] = "불량시정건_전체_업체";
                lstAll[3] = dgdAllDefectSymptom.Name;
                lstAll[4] = dgdAllDefectReason.Name;
                lstAll[5] = dgdAllDefectCustom.Name;

                ExpExc = new ExportExcelxaml(lstAll);
                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdAllDefectSymptom.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdAllDefectSymptom);
                        else
                            dt = lib.DataGirdToDataTable(dgdAllDefectSymptom);

                        Name = dgdAllDefectSymptom.Name;

                        if (lib.GenerateExcel(dt, Name))
                        {
                            lib.excel.Visible = true;
                            lib.ReleaseExcelObject(lib.excel);
                        }
                    }
                    else if (ExpExc.choice.Equals(dgdAllDefectReason.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdAllDefectReason);
                        else
                            dt = lib.DataGirdToDataTable(dgdAllDefectReason);

                        Name = dgdAllDefectReason.Name;

                        if (lib.GenerateExcel(dt, Name))
                        {
                            lib.excel.Visible = true;
                            lib.ReleaseExcelObject(lib.excel);
                        }
                    }
                    else if (ExpExc.choice.Equals(dgdAllDefectCustom.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdAllDefectCustom);
                        else
                            dt = lib.DataGirdToDataTable(dgdAllDefectCustom);

                        Name = dgdAllDefectCustom.Name;

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
            else if (NowTI.Header.ToString().Equals("일별"))
            {
                string[] lstDay = new string[8];
                lstDay[0] = "불량시정건_일별_조회";
                lstDay[1] = "불량시정건_일별_불량유형";
                lstDay[2] = "불량시정건_일별_불량원인";
                lstDay[3] = "불량시정건_일별_업체";
                lstDay[4] = dgdDefectRepairDaily.Name;
                lstDay[5] = dgdDailyDefectSymptom.Name;
                lstDay[6] = dgdDailyDefectReason.Name;
                lstDay[7] = dgdDailyDefectCustom.Name;

                ExpExc = new ExportExcelxaml(lstDay);
                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdDefectRepairDaily.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdDefectRepairDaily);
                        else
                            dt = lib.DataGirdToDataTable(dgdDefectRepairDaily);

                        Name = dgdDefectRepairDaily.Name;

                        if (lib.GenerateExcel(dt, Name))
                        {
                            lib.excel.Visible = true;
                            lib.ReleaseExcelObject(lib.excel);
                        }
                    }
                    else if (ExpExc.choice.Equals(dgdDailyDefectSymptom.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdDailyDefectSymptom);
                        else
                            dt = lib.DataGirdToDataTable(dgdDailyDefectSymptom);

                        Name = dgdDailyDefectSymptom.Name;

                        if (lib.GenerateExcel(dt, Name))
                        {
                            lib.excel.Visible = true;
                            lib.ReleaseExcelObject(lib.excel);
                        }
                    }
                    else if (ExpExc.choice.Equals(dgdDailyDefectReason.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdDailyDefectReason);
                        else
                            dt = lib.DataGirdToDataTable(dgdDailyDefectReason);

                        Name = dgdDailyDefectReason.Name;

                        if (lib.GenerateExcel(dt, Name))
                        {
                            lib.excel.Visible = true;
                            lib.ReleaseExcelObject(lib.excel);
                        }
                    }
                    else if (ExpExc.choice.Equals(dgdDailyDefectCustom.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdDailyDefectCustom);
                        else
                            dt = lib.DataGirdToDataTable(dgdDailyDefectCustom);

                        Name = dgdDailyDefectCustom.Name;

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
            else    //월별 분석
            {
                string[] lstDay = new string[8];
                lstDay[0] = "불량시정건_월별_조회";
                lstDay[1] = "불량시정건_월별_불량유형";
                lstDay[2] = "불량시정건_월별_불량원인";
                lstDay[3] = "불량시정건_월별_업체";
                lstDay[4] = dgdDefectRepairMonth.Name;
                lstDay[5] = dgdMonthDefectSymptom.Name;
                lstDay[6] = dgdMonthDefectReason.Name;
                lstDay[7] = dgdMonthDefectCustom.Name;

                ExpExc = new ExportExcelxaml(lstDay);
                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdDefectRepairMonth.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdDefectRepairMonth);
                        else
                            dt = lib.DataGirdToDataTable(dgdDefectRepairMonth);

                        Name = dgdDefectRepairMonth.Name;

                        if (lib.GenerateExcel(dt, Name))
                        {
                            lib.excel.Visible = true;
                            lib.ReleaseExcelObject(lib.excel);
                        }
                    }
                    else if (ExpExc.choice.Equals(dgdMonthDefectSymptom.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdMonthDefectSymptom);
                        else
                            dt = lib.DataGirdToDataTable(dgdMonthDefectSymptom);

                        Name = dgdMonthDefectSymptom.Name;

                        if (lib.GenerateExcel(dt, Name))
                        {
                            lib.excel.Visible = true;
                            lib.ReleaseExcelObject(lib.excel);
                        }
                    }
                    else if (ExpExc.choice.Equals(dgdMonthDefectReason.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdMonthDefectReason);
                        else
                            dt = lib.DataGirdToDataTable(dgdMonthDefectReason);

                        Name = dgdMonthDefectReason.Name;

                        if (lib.GenerateExcel(dt, Name))
                        {
                            lib.excel.Visible = true;
                            lib.ReleaseExcelObject(lib.excel);
                        }
                    }
                    else if (ExpExc.choice.Equals(dgdMonthDefectCustom.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = lib.DataGridToDTinHidden(dgdMonthDefectCustom);
                        else
                            dt = lib.DataGirdToDataTable(dgdMonthDefectCustom);

                        Name = dgdMonthDefectCustom.Name;

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
        }

        private void re_Search(int selectIndex, TabItem tabItem)
        {
            if (tabItem.Header.ToString().Equals("전체"))
            {
                FillGridAllTab();

                if (dgdAllDefectCustom.Items.Count > 0)
                {
                    dgdAllDefectCustom.SelectedIndex = selectIndex;
                }

                if (dgdAllDefectReason.Items.Count > 0)
                {
                    dgdAllDefectReason.SelectedIndex = selectIndex;
                }

                if (dgdAllDefectSymptom.Items.Count > 0)
                {
                    dgdAllDefectSymptom.SelectedIndex = selectIndex;
                }
            }
            else if (tabItem.Header.ToString().Equals("일별"))
            {
                FillGridDailyTab_Daily();
                FillGridDailyTab_Bottom();

                if (dgdDailyDefectCustom.Items.Count > 0)
                {
                    dgdDailyDefectCustom.SelectedIndex = selectIndex;
                }

                if (dgdDailyDefectReason.Items.Count > 0)
                {
                    dgdDailyDefectReason.SelectedIndex = selectIndex;
                }

                if (dgdDailyDefectSymptom.Items.Count > 0)
                {
                    dgdDailyDefectSymptom.SelectedIndex = selectIndex;
                }
            }
            else if (tabItem.Header.ToString().Equals("월별"))
            {
                FillGridMonthTab_Monthly();
                FillGridMonthTab_Bottom();

                if (dgdMonthDefectCustom.Items.Count > 0)
                {
                    dgdMonthDefectCustom.SelectedIndex = selectIndex;
                }

                if (dgdMonthDefectReason.Items.Count > 0)
                {
                    dgdMonthDefectReason.SelectedIndex = selectIndex;
                }

                if (dgdMonthDefectSymptom.Items.Count > 0)
                {
                    dgdMonthDefectSymptom.SelectedIndex = selectIndex;
                }
            }
        }

        private void FillGridAllTab()
        {
            if (dgdAllDefectCustom.Items.Count > 0)
            {
                dgdAllDefectCustom.Items.Clear();
            }
            if (dgdAllDefectReason.Items.Count > 0)
            {
                dgdAllDefectReason.Items.Clear();
            }
            if (dgdAllDefectSymptom.Items.Count > 0)
            {
                dgdAllDefectSymptom.Items.Clear();
            }

            try
            {
                DataSet ds1 = null;
                DataSet ds2 = null;
                DataSet ds3 = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                int totalDefectQty = 0;
                int totalRepairQty = 0;
                double totalRepairRate = 0.00;
                sqlParameter.Clear();
                sqlParameter.Add("nchkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("StartDate", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("EndDate", chkDate.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("nchkDefectStep", chkDefectOccurStep.IsChecked == true ? 1 : 0);
                sqlParameter.Add("DefectStep", chkDefectOccurStep.IsChecked == true ? cboOccurStepSrh.SelectedValue.ToString() : "");

                sqlParameter.Add("nchkCustom", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true ? txtCustom.Tag.ToString() : "");
                sqlParameter.Add("nchkArticleID", chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticle.IsChecked == true ? txtArticle.Tag.ToString() : "");
                //sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");
                sqlParameter.Add("nchkBuyerArticleNo", chkArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkArticleNo.IsChecked == true ? (txtArticleNo.Tag != null ? txtArticleNo.Tag.ToString() : "") : "");
                sqlParameter.Add("sGrouping", 1);   //불량유형
                ds1 = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectRepair_Sum", sqlParameter, false);

                sqlParameter.Remove("sGrouping");
                sqlParameter.Add("sGrouping", 4);   //불량원인
                ds2 = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectRepair_Sum", sqlParameter, false);

                sqlParameter.Remove("sGrouping");
                sqlParameter.Add("sGrouping", 2);   //업체


                ds3 = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectRepair_Sum", sqlParameter, false);

                if (ds1 != null && ds1.Tables.Count > 0)
                {
                    DataTable dt = ds1.Tables[0];
                    int i = 0;
                    totalDefectQty = 0;
                    totalRepairQty = 0;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DT_SYMPTOM = null;
                        DT_SYMPTOM = dt;

                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var DefectSymptom = new Win_Qul_DefectRepair_Q_CodeView()
                            {
                                Num = i + 1,
                                cls = dr["cls"].ToString(),
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                GroupingName = dr["GroupingName"].ToString(),
                                RepairQty = stringFormatN0(dr["RepairQty"]),
                                RepairRate = dr["RepairRate"].ToString(),
                                RepairRate1 = dr["RepairRate1"].ToString()
                            };

                            totalDefectQty += Convert.ToInt32(dr["DefectQty"].ToString());
                            totalRepairQty += Convert.ToInt32(dr["RepairQty"].ToString());

                            dgdAllDefectSymptom.Items.Add(DefectSymptom);
                            i++;
                        }

                        if (totalDefectQty == 0 || totalRepairQty == 0)
                        {
                            totalRepairRate = 0;
                        }
                        else
                        {
                            totalRepairRate = ((double)totalRepairQty / (double)totalDefectQty) * 100;
                        }

                        //총계그리드 비우고
                        dgdTotal.Items.Clear();

                        var DefectSymptom2 = new Win_Qul_DefectRepair_Q_CodeView()
                        {
                            GroupingName = "총계",
                            DefectQty = stringFormatN0(totalDefectQty),
                            RepairQty = stringFormatN0(totalRepairQty),
                            RepairRate = totalRepairRate.ToString(),
                            ColorBlue = "true"
                        };

                        if (Lib.Instance.IsNumOrAnother(DefectSymptom2.RepairRate))
                        {
                            DefectSymptom2.RepairRate = string.Format("{0:N2}", double.Parse(DefectSymptom2.RepairRate));
                        }
                        //채우기
                        dgdTotal.Items.Add(DefectSymptom2);
                    }
                }

                if (ds2 != null && ds2.Tables.Count > 0)
                {
                    DataTable dt = ds2.Tables[0];
                    int i = 0;
                    totalDefectQty = 0;
                    totalRepairQty = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DT_REASON = null;
                        DT_REASON = dt;

                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var DefectSymptom = new Win_Qul_DefectRepair_Q_CodeView()
                            {
                                Num = i + 1,
                                cls = dr["cls"].ToString(),
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                GroupingName = dr["GroupingName"].ToString(),
                                RepairQty = stringFormatN0(dr["RepairQty"]),
                                RepairRate = dr["RepairRate"].ToString(),
                                RepairRate1 = dr["RepairRate1"].ToString()
                            };

                            totalDefectQty += Convert.ToInt32(dr["DefectQty"].ToString());
                            totalRepairQty += Convert.ToInt32(dr["RepairQty"].ToString());

                            dgdAllDefectReason.Items.Add(DefectSymptom);
                            i++;
                        }

                        if (totalDefectQty == 0 || totalRepairQty == 0)
                        {
                            totalRepairRate = 0;
                        }
                        else
                        {
                            totalRepairRate = ((double)totalRepairQty / (double)totalDefectQty) * 100;
                        }


                        var DefectSymptom2 = new Win_Qul_DefectRepair_Q_CodeView()
                        {
                            //cls = dr["cls"].ToString(),
                            GroupingName = "총계",
                            DefectQty = stringFormatN0(totalDefectQty),
                            RepairQty = stringFormatN0(totalRepairQty),
                            RepairRate = totalRepairRate.ToString(),
                            ColorBlue = "true"
                        };

                        if (Lib.Instance.IsNumOrAnother(DefectSymptom2.RepairRate))
                        {
                            DefectSymptom2.RepairRate = string.Format("{0:N2}", double.Parse(DefectSymptom2.RepairRate));
                        }
                        //dgdAllDefectReason.Items.Add(DefectSymptom2);
                    }
                }

                if (ds3 != null && ds3.Tables.Count > 0)
                {
                    DataTable dt = ds3.Tables[0];
                    int i = 0;
                    totalDefectQty = 0;
                    totalRepairQty = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DT_CUSTOM = null;
                        DT_CUSTOM = dt;

                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var DefectSymptom = new Win_Qul_DefectRepair_Q_CodeView()
                            {
                                Num = i + 1,
                                cls = dr["cls"].ToString(),
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                GroupingName = dr["GroupingName"].ToString(),
                                RepairQty = stringFormatN0(dr["RepairQty"]),
                                RepairRate = dr["RepairRate"].ToString(),
                                RepairRate1 = dr["RepairRate1"].ToString()
                            };

                            totalDefectQty += Convert.ToInt32(dr["DefectQty"].ToString());
                            totalRepairQty += Convert.ToInt32(dr["RepairQty"].ToString());

                            dgdAllDefectCustom.Items.Add(DefectSymptom);
                            i++;
                        }

                        if (totalDefectQty == 0 || totalRepairQty == 0)
                        {
                            totalRepairRate = 0;
                        }
                        else
                        {
                            totalRepairRate = ((double)totalRepairQty / (double)totalDefectQty) * 100;
                        }

                        var DefectSymptom2 = new Win_Qul_DefectRepair_Q_CodeView()
                        {
                            //cls = dr["cls"].ToString(),
                            GroupingName = "총계",
                            DefectQty = stringFormatN0(totalDefectQty),
                            RepairQty = stringFormatN0(totalRepairQty),
                            RepairRate = totalRepairRate.ToString(),
                            ColorBlue = "true"
                        };

                        if (Lib.Instance.IsNumOrAnother(DefectSymptom2.RepairRate))
                        {
                            DefectSymptom2.RepairRate = string.Format("{0:N2}", double.Parse(DefectSymptom2.RepairRate));
                        }
                        //dgdAllDefectCustom.Items.Add(DefectSymptom2);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        private void FillGridDailyTab_Daily()
        {
            if (dgdDefectRepairDaily.Items.Count > 0)
            {
                dgdDefectRepairDaily.Items.Clear();
            }

            if (lvcDayChart.Series != null && lvcDayChart.Series.Count > 0)
            {
                lvcDayChart.Series.Clear();
            }

            dgdtpeDate29.Visibility = Visibility.Visible;
            dgdtpeDate30.Visibility = Visibility.Visible;
            dgdtpeDate31.Visibility = Visibility.Visible;

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("nchkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("FromMonth", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMM") : "");
                sqlParameter.Add("nchkDefectStep", 1);
                sqlParameter.Add("DefectStep", cboOccurStepSrh.SelectedValue.ToString());

                sqlParameter.Add("nchkCustom", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true ? txtCustom.Tag.ToString() : "");
                sqlParameter.Add("nchkArticleID", 0); // chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", ""); // chkArticle.IsChecked == true ? txtArticle.Tag.ToString() : "");
                                                   //sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");
                sqlParameter.Add("nchkBuyerArticleNo", chkArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkArticleNo.IsChecked == true ? (txtArticleNo.Tag != null ? txtArticleNo.Tag.ToString() : "") : "");
                ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectRepair_Daily", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinDaily = new Win_Qul_DefectRepair_Q_Daily_CodeView()
                            {
                                ProdQty1 = Lib.Instance.returnNumString(dr["ProdQty1"].ToString()),
                                ProdQty2 = Lib.Instance.returnNumString(dr["ProdQty2"].ToString()),
                                ProdQty3 = Lib.Instance.returnNumString(dr["ProdQty3"].ToString()),
                                ProdQty4 = Lib.Instance.returnNumString(dr["ProdQty4"].ToString()),
                                ProdQty5 = Lib.Instance.returnNumString(dr["ProdQty5"].ToString()),
                                ProdQty6 = Lib.Instance.returnNumString(dr["ProdQty6"].ToString()),
                                ProdQty7 = Lib.Instance.returnNumString(dr["ProdQty7"].ToString()),
                                ProdQty8 = Lib.Instance.returnNumString(dr["ProdQty8"].ToString()),
                                ProdQty9 = Lib.Instance.returnNumString(dr["ProdQty9"].ToString()),
                                ProdQty10 = Lib.Instance.returnNumString(dr["ProdQty10"].ToString()),
                                ProdQty11 = Lib.Instance.returnNumString(dr["ProdQty11"].ToString()),
                                ProdQty12 = Lib.Instance.returnNumString(dr["ProdQty12"].ToString()),
                                ProdQty13 = Lib.Instance.returnNumString(dr["ProdQty13"].ToString()),
                                ProdQty14 = Lib.Instance.returnNumString(dr["ProdQty14"].ToString()),
                                ProdQty15 = Lib.Instance.returnNumString(dr["ProdQty15"].ToString()),
                                ProdQty16 = Lib.Instance.returnNumString(dr["ProdQty16"].ToString()),
                                ProdQty17 = Lib.Instance.returnNumString(dr["ProdQty17"].ToString()),
                                ProdQty18 = Lib.Instance.returnNumString(dr["ProdQty18"].ToString()),
                                ProdQty19 = Lib.Instance.returnNumString(dr["ProdQty19"].ToString()),
                                ProdQty20 = Lib.Instance.returnNumString(dr["ProdQty20"].ToString()),
                                ProdQty21 = Lib.Instance.returnNumString(dr["ProdQty21"].ToString()),
                                ProdQty22 = Lib.Instance.returnNumString(dr["ProdQty22"].ToString()),
                                ProdQty23 = Lib.Instance.returnNumString(dr["ProdQty23"].ToString()),
                                ProdQty24 = Lib.Instance.returnNumString(dr["ProdQty24"].ToString()),
                                ProdQty25 = Lib.Instance.returnNumString(dr["ProdQty25"].ToString()),
                                ProdQty26 = Lib.Instance.returnNumString(dr["ProdQty26"].ToString()),
                                ProdQty27 = Lib.Instance.returnNumString(dr["ProdQty27"].ToString()),
                                ProdQty28 = Lib.Instance.returnNumString(dr["ProdQty28"].ToString()),
                                ProdQty29 = Lib.Instance.returnNumString(dr["ProdQty29"].ToString()),
                                ProdQty30 = Lib.Instance.returnNumString(dr["ProdQty30"].ToString()),
                                ProdQty31 = Lib.Instance.returnNumString(dr["ProdQty31"].ToString()),

                                DefectQty1 = Lib.Instance.returnNumString(dr["DefectQty1"].ToString()),
                                DefectQty2 = Lib.Instance.returnNumString(dr["DefectQty2"].ToString()),
                                DefectQty3 = Lib.Instance.returnNumString(dr["DefectQty3"].ToString()),
                                DefectQty4 = Lib.Instance.returnNumString(dr["DefectQty4"].ToString()),
                                DefectQty5 = Lib.Instance.returnNumString(dr["DefectQty5"].ToString()),
                                DefectQty6 = Lib.Instance.returnNumString(dr["DefectQty6"].ToString()),
                                DefectQty7 = Lib.Instance.returnNumString(dr["DefectQty7"].ToString()),
                                DefectQty8 = Lib.Instance.returnNumString(dr["DefectQty8"].ToString()),
                                DefectQty9 = Lib.Instance.returnNumString(dr["DefectQty9"].ToString()),
                                DefectQty10 = Lib.Instance.returnNumString(dr["DefectQty10"].ToString()),
                                DefectQty11 = Lib.Instance.returnNumString(dr["DefectQty11"].ToString()),
                                DefectQty12 = Lib.Instance.returnNumString(dr["DefectQty12"].ToString()),
                                DefectQty13 = Lib.Instance.returnNumString(dr["DefectQty13"].ToString()),
                                DefectQty14 = Lib.Instance.returnNumString(dr["DefectQty14"].ToString()),
                                DefectQty15 = Lib.Instance.returnNumString(dr["DefectQty15"].ToString()),
                                DefectQty16 = Lib.Instance.returnNumString(dr["DefectQty16"].ToString()),
                                DefectQty17 = Lib.Instance.returnNumString(dr["DefectQty17"].ToString()),
                                DefectQty18 = Lib.Instance.returnNumString(dr["DefectQty18"].ToString()),
                                DefectQty19 = Lib.Instance.returnNumString(dr["DefectQty19"].ToString()),
                                DefectQty20 = Lib.Instance.returnNumString(dr["DefectQty20"].ToString()),
                                DefectQty21 = Lib.Instance.returnNumString(dr["DefectQty21"].ToString()),
                                DefectQty22 = Lib.Instance.returnNumString(dr["DefectQty22"].ToString()),
                                DefectQty23 = Lib.Instance.returnNumString(dr["DefectQty23"].ToString()),
                                DefectQty24 = Lib.Instance.returnNumString(dr["DefectQty24"].ToString()),
                                DefectQty25 = Lib.Instance.returnNumString(dr["DefectQty25"].ToString()),
                                DefectQty26 = Lib.Instance.returnNumString(dr["DefectQty26"].ToString()),
                                DefectQty27 = Lib.Instance.returnNumString(dr["DefectQty27"].ToString()),
                                DefectQty28 = Lib.Instance.returnNumString(dr["DefectQty28"].ToString()),
                                DefectQty29 = Lib.Instance.returnNumString(dr["DefectQty29"].ToString()),
                                DefectQty30 = Lib.Instance.returnNumString(dr["DefectQty30"].ToString()),
                                DefectQty31 = Lib.Instance.returnNumString(dr["DefectQty31"].ToString()),

                                RepairRate1 = Lib.Instance.returnNumStringTwo(dr["RepairRate1"].ToString()),
                                RepairRate2 = Lib.Instance.returnNumStringTwo(dr["RepairRate2"].ToString()),
                                RepairRate3 = Lib.Instance.returnNumStringTwo(dr["RepairRate3"].ToString()),
                                RepairRate4 = Lib.Instance.returnNumStringTwo(dr["RepairRate4"].ToString()),
                                RepairRate5 = Lib.Instance.returnNumStringTwo(dr["RepairRate5"].ToString()),
                                RepairRate6 = Lib.Instance.returnNumStringTwo(dr["RepairRate6"].ToString()),
                                RepairRate7 = Lib.Instance.returnNumStringTwo(dr["RepairRate7"].ToString()),
                                RepairRate8 = Lib.Instance.returnNumStringTwo(dr["RepairRate8"].ToString()),
                                RepairRate9 = Lib.Instance.returnNumStringTwo(dr["RepairRate9"].ToString()),
                                RepairRate10 = Lib.Instance.returnNumStringTwo(dr["RepairRate10"].ToString()),
                                RepairRate11 = Lib.Instance.returnNumStringTwo(dr["RepairRate11"].ToString()),
                                RepairRate12 = Lib.Instance.returnNumStringTwo(dr["RepairRate12"].ToString()),
                                RepairRate13 = Lib.Instance.returnNumStringTwo(dr["RepairRate13"].ToString()),
                                RepairRate14 = Lib.Instance.returnNumStringTwo(dr["RepairRate14"].ToString()),
                                RepairRate15 = Lib.Instance.returnNumStringTwo(dr["RepairRate15"].ToString()),
                                RepairRate16 = Lib.Instance.returnNumStringTwo(dr["RepairRate16"].ToString()),
                                RepairRate17 = Lib.Instance.returnNumStringTwo(dr["RepairRate17"].ToString()),
                                RepairRate18 = Lib.Instance.returnNumStringTwo(dr["RepairRate18"].ToString()),
                                RepairRate19 = Lib.Instance.returnNumStringTwo(dr["RepairRate19"].ToString()),
                                RepairRate20 = Lib.Instance.returnNumStringTwo(dr["RepairRate20"].ToString()),
                                RepairRate21 = Lib.Instance.returnNumStringTwo(dr["RepairRate21"].ToString()),
                                RepairRate22 = Lib.Instance.returnNumStringTwo(dr["RepairRate22"].ToString()),
                                RepairRate23 = Lib.Instance.returnNumStringTwo(dr["RepairRate23"].ToString()),
                                RepairRate24 = Lib.Instance.returnNumStringTwo(dr["RepairRate24"].ToString()),
                                RepairRate25 = Lib.Instance.returnNumStringTwo(dr["RepairRate25"].ToString()),
                                RepairRate26 = Lib.Instance.returnNumStringTwo(dr["RepairRate26"].ToString()),
                                RepairRate27 = Lib.Instance.returnNumStringTwo(dr["RepairRate27"].ToString()),
                                RepairRate28 = Lib.Instance.returnNumStringTwo(dr["RepairRate28"].ToString()),
                                RepairRate29 = Lib.Instance.returnNumStringTwo(dr["RepairRate29"].ToString()),
                                RepairRate30 = Lib.Instance.returnNumStringTwo(dr["RepairRate30"].ToString()),
                                RepairRate31 = Lib.Instance.returnNumStringTwo(dr["RepairRate31"].ToString()),

                                RepairQty1 = Lib.Instance.returnNumString(dr["RepairQty1"].ToString()),
                                RepairQty2 = Lib.Instance.returnNumString(dr["RepairQty2"].ToString()),
                                RepairQty3 = Lib.Instance.returnNumString(dr["RepairQty3"].ToString()),
                                RepairQty4 = Lib.Instance.returnNumString(dr["RepairQty4"].ToString()),
                                RepairQty5 = Lib.Instance.returnNumString(dr["RepairQty5"].ToString()),
                                RepairQty6 = Lib.Instance.returnNumString(dr["RepairQty6"].ToString()),
                                RepairQty7 = Lib.Instance.returnNumString(dr["RepairQty7"].ToString()),
                                RepairQty8 = Lib.Instance.returnNumString(dr["RepairQty8"].ToString()),
                                RepairQty9 = Lib.Instance.returnNumString(dr["RepairQty9"].ToString()),
                                RepairQty10 = Lib.Instance.returnNumString(dr["RepairQty10"].ToString()),
                                RepairQty11 = Lib.Instance.returnNumString(dr["RepairQty11"].ToString()),
                                RepairQty12 = Lib.Instance.returnNumString(dr["RepairQty12"].ToString()),
                                RepairQty13 = Lib.Instance.returnNumString(dr["RepairQty13"].ToString()),
                                RepairQty14 = Lib.Instance.returnNumString(dr["RepairQty14"].ToString()),
                                RepairQty15 = Lib.Instance.returnNumString(dr["RepairQty15"].ToString()),
                                RepairQty16 = Lib.Instance.returnNumString(dr["RepairQty16"].ToString()),
                                RepairQty17 = Lib.Instance.returnNumString(dr["RepairQty17"].ToString()),
                                RepairQty18 = Lib.Instance.returnNumString(dr["RepairQty18"].ToString()),
                                RepairQty19 = Lib.Instance.returnNumString(dr["RepairQty19"].ToString()),
                                RepairQty20 = Lib.Instance.returnNumString(dr["RepairQty20"].ToString()),
                                RepairQty21 = Lib.Instance.returnNumString(dr["RepairQty21"].ToString()),
                                RepairQty22 = Lib.Instance.returnNumString(dr["RepairQty22"].ToString()),
                                RepairQty23 = Lib.Instance.returnNumString(dr["RepairQty23"].ToString()),
                                RepairQty24 = Lib.Instance.returnNumString(dr["RepairQty24"].ToString()),
                                RepairQty25 = Lib.Instance.returnNumString(dr["RepairQty25"].ToString()),
                                RepairQty26 = Lib.Instance.returnNumString(dr["RepairQty26"].ToString()),
                                RepairQty27 = Lib.Instance.returnNumString(dr["RepairQty27"].ToString()),
                                RepairQty28 = Lib.Instance.returnNumString(dr["RepairQty28"].ToString()),
                                RepairQty29 = Lib.Instance.returnNumString(dr["RepairQty29"].ToString()),
                                RepairQty30 = Lib.Instance.returnNumString(dr["RepairQty30"].ToString()),
                                RepairQty31 = Lib.Instance.returnNumString(dr["RepairQty31"].ToString()),

                                AvgProdQty = Lib.Instance.returnNumStringTwo(dr["AvgProdQty"].ToString()),
                                AvgDefectQty = Lib.Instance.returnNumStringTwo(dr["AvgDefectQty"].ToString()),
                                AvgRepairQty = Lib.Instance.returnNumStringTwo(dr["AvgRepairQty"].ToString()),
                                AvgRepairRate = Lib.Instance.returnNumStringTwo(dr["AvgRepairRate"].ToString()),
                                DayCount = dr["DayCount"].ToString(),

                                TDefectQty = Lib.Instance.returnNumString(dr["TDefectQty"].ToString()),
                                TProdQty = Lib.Instance.returnNumString(dr["TProdQty"].ToString()),
                                TRepairQty = Lib.Instance.returnNumString(dr["TRepairQty"].ToString()),
                                TRepairRate = Lib.Instance.returnNumStringTwo(dr["TRepairRate"].ToString()),
                                OccurDate = dr["OccurDate"].ToString(),
                                step = dr["step"].ToString(),

                                Num1 = 1,
                                Num2 = 2,
                                Num3 = 3,
                                Num4 = 4,

                                ////콤보박스 선택에 따른 수량 명칭 변경
                                strGubun1 = cboOccurStepSrh.SelectedValue.Equals("1") ? "입고수량" :
                                            cboOccurStepSrh.SelectedValue.Equals("3") ? "검사수량" :
                                            cboOccurStepSrh.SelectedValue.Equals("2") ? "생산수량" :
                                            cboOccurStepSrh.SelectedValue.Equals("4") ? "검사수량" : "출고수량",

                                strGubun2 = "불량수",
                                strGubun3 = "시정수량",
                                strGubun4 = "시정률(%)"
                            };

                            strLastDay = WinDaily.DayCount;

                            if (Lib.Instance.IsNumOrAnother(WinDaily.TRepairRate))
                            {
                                WinDaily.TRepairRate = string.Format("{0:N2}", double.Parse(WinDaily.TRepairRate));
                            }

                            if (Lib.Instance.IsNumOrAnother(WinDaily.AvgRepairRate))
                            {
                                WinDaily.AvgRepairRate = string.Format("{0:N2}", double.Parse(WinDaily.AvgRepairRate));
                            }

                            double day1 = 0;
                            double day2 = 0;
                            double day3 = 0;
                            double day4 = 0;
                            double day5 = 0;
                            double day6 = 0;
                            double day7 = 0;
                            double day8 = 0;
                            double day9 = 0;
                            double day10 = 0;
                            double day11 = 0;
                            double day12 = 0;
                            double day13 = 0;
                            double day14 = 0;
                            double day15 = 0;
                            double day16 = 0;
                            double day17 = 0;
                            double day18 = 0;
                            double day19 = 0;
                            double day20 = 0;
                            double day21 = 0;
                            double day22 = 0;
                            double day23 = 0;
                            double day24 = 0;
                            double day25 = 0;
                            double day26 = 0;
                            double day27 = 0;
                            double day28 = 0;
                            double day29 = 0;
                            double day30 = 0;
                            double day31 = 0;

                            day1 = Convert.ToDouble(dr["RepairRate1"].ToString());
                            day2 = Convert.ToDouble(dr["RepairRate2"].ToString());
                            day3 = Convert.ToDouble(dr["RepairRate3"].ToString());
                            day4 = Convert.ToDouble(dr["RepairRate4"].ToString());
                            day5 = Convert.ToDouble(dr["RepairRate5"].ToString());
                            day6 = Convert.ToDouble(dr["RepairRate6"].ToString());
                            day7 = Convert.ToDouble(dr["RepairRate7"].ToString());
                            day8 = Convert.ToDouble(dr["RepairRate8"].ToString());
                            day9 = Convert.ToDouble(dr["RepairRate9"].ToString());
                            day10 = Convert.ToDouble(dr["RepairRate10"].ToString());
                            day11 = Convert.ToDouble(dr["RepairRate11"].ToString());
                            day12 = Convert.ToDouble(dr["RepairRate12"].ToString());
                            day13 = Convert.ToDouble(dr["RepairRate13"].ToString());
                            day14 = Convert.ToDouble(dr["RepairRate14"].ToString());
                            day15 = Convert.ToDouble(dr["RepairRate15"].ToString());
                            day16 = Convert.ToDouble(dr["RepairRate16"].ToString());
                            day17 = Convert.ToDouble(dr["RepairRate17"].ToString());
                            day18 = Convert.ToDouble(dr["RepairRate18"].ToString());
                            day19 = Convert.ToDouble(dr["RepairRate19"].ToString());
                            day20 = Convert.ToDouble(dr["RepairRate20"].ToString());
                            day21 = Convert.ToDouble(dr["RepairRate21"].ToString());
                            day22 = Convert.ToDouble(dr["RepairRate22"].ToString());
                            day23 = Convert.ToDouble(dr["RepairRate23"].ToString());
                            day24 = Convert.ToDouble(dr["RepairRate24"].ToString());
                            day25 = Convert.ToDouble(dr["RepairRate25"].ToString());
                            day26 = Convert.ToDouble(dr["RepairRate26"].ToString());
                            day27 = Convert.ToDouble(dr["RepairRate27"].ToString());
                            day28 = Convert.ToDouble(dr["RepairRate28"].ToString());
                            day29 = Convert.ToDouble(dr["RepairRate29"].ToString());
                            day30 = Convert.ToDouble(dr["RepairRate30"].ToString());
                            day31 = Convert.ToDouble(dr["RepairRate31"].ToString());


                            //day25 = 50000;

                            if (WinDaily.DayCount.Equals("28"))
                            {
                                SeriesCollection SeriesCollection = new SeriesCollection
                                {
                                    new ColumnSeries
                                    {
                                        Title = "일별 불량율",
                                        Values = new ChartValues<double>
                                        {
                                            0, day1, day2, day3, day4, day5, day6, day7, day8, day9, day10,
                                            day11, day12, day13, day14, day15, day16, day17, day18, day19, day20,
                                            day21, day22, day23, day24, day25, day26, day27, day28
                                        },
                                        DataLabels = false
                                    }
                                };
                                lvcDayChart.Series = SeriesCollection;

                            }
                            else if (WinDaily.DayCount.Equals("29"))
                            {
                                SeriesCollection SeriesCollection = new SeriesCollection
                                {
                                    new ColumnSeries
                                    {
                                        Title = "일별 불량율",
                                        Values = new ChartValues<double>
                                        {
                                            0, day1, day2, day3, day4, day5, day6, day7, day8, day9, day10,
                                            day11, day12, day13, day14, day15, day16, day17, day18, day19, day20,
                                            day21, day22, day23, day24, day25, day26, day27, day28, day29
                                        },
                                        DataLabels = false
                                    }
                                };
                                lvcDayChart.Series = SeriesCollection;
                            }
                            else if (WinDaily.DayCount.Equals("30"))
                            {
                                SeriesCollection SeriesCollection = new SeriesCollection
                                {
                                    new ColumnSeries
                                    {
                                        Title = "일별 불량율",
                                        Values = new ChartValues<double>
                                        {
                                            0, day1, day2, day3, day4, day5, day6, day7, day8, day9, day10,
                                            day11, day12, day13, day14, day15, day16, day17, day18, day19, day20,
                                            day21, day22, day23, day24, day25, day26, day27, day28, day29, day30
                                        },
                                        DataLabels = false
                                    }
                                };
                                lvcDayChart.Series = SeriesCollection;
                            }
                            else
                            {
                                SeriesCollection SeriesCollection = new SeriesCollection
                                {
                                    new ColumnSeries
                                    {
                                        Title = "일별 불량율",
                                        Values = new ChartValues<double>
                                        {
                                            0, day1, day2, day3, day4, day5, day6, day7, day8, day9, day10,
                                            day11, day12, day13, day14, day15, day16, day17, day18, day19, day20,
                                            day21, day22, day23, day24, day25, day26, day27, day28, day29, day30
                                        },
                                        DataLabels = false
                                    }
                                };
                                lvcDayChart.Series = SeriesCollection;
                            }

                            //SeriesCollection SeriesCollection = new SeriesCollection(columnSeries);
                            //lvcDayChart.Series = SeriesCollection;

                            dgdDefectRepairDaily.Items.Add(WinDaily);
                            i++;
                        }

                        if (strLastDay.Equals("28"))
                        {
                            dgdtpeDate29.Visibility = Visibility.Hidden;
                            dgdtpeDate30.Visibility = Visibility.Hidden;
                            dgdtpeDate31.Visibility = Visibility.Hidden;
                        }
                        else if (strLastDay.Equals("29"))
                        {
                            dgdtpeDate30.Visibility = Visibility.Hidden;
                            dgdtpeDate31.Visibility = Visibility.Hidden;
                        }
                        else if (strLastDay.Equals("30"))
                        {
                            dgdtpeDate31.Visibility = Visibility.Hidden;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        private void FillGridDailyTab_Bottom()
        {
            if (dgdDailyDefectCustom.Items.Count > 0)
            {
                dgdDailyDefectCustom.Items.Clear();
            }
            if (dgdDailyDefectReason.Items.Count > 0)
            {
                dgdDailyDefectReason.Items.Clear();
            }
            if (dgdDailyDefectSymptom.Items.Count > 0)
            {
                dgdDailyDefectSymptom.Items.Clear();
            }

            try
            {
                DataSet ds1 = null;
                DataSet ds2 = null;
                DataSet ds3 = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                int totalDefectQty;
                int totalRepairQty;
                double totalRepairRate = 0.00;
                sqlParameter.Clear();
                sqlParameter.Add("nchkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("StartMonth", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyyMM") : "");
                sqlParameter.Add("nchkDefectStep", 1);
                sqlParameter.Add("DefectStep", cboOccurStepSrh.SelectedValue.ToString());

                sqlParameter.Add("nchkCustom", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true ? txtCustom.Tag.ToString() : "");
                sqlParameter.Add("nchkArticleID", 0); //chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", ""); // chkArticle.IsChecked == true ? txtArticle.Tag.ToString() : "");
                sqlParameter.Add("sGrouping", 1);
                //sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");
                sqlParameter.Add("nchkBuyerArticleNo", chkArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkArticleNo.IsChecked == true ? (txtArticleNo.Tag != null ? txtArticleNo.Tag.ToString() : "") : "");
                ds1 = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectRepair_Sum_Daily", sqlParameter, false);

                sqlParameter.Remove("sGrouping");
                sqlParameter.Add("sGrouping", 4);   //불량원인
                ds2 = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectRepair_Sum_Daily", sqlParameter, false);

                sqlParameter.Remove("sGrouping");
                sqlParameter.Add("sGrouping", 2);   //업체
                ds3 = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectRepair_Sum_Daily", sqlParameter, false);

                if (ds1 != null && ds1.Tables.Count > 0)
                {
                    DataTable dt = ds1.Tables[0];
                    int i = 0;
                    totalDefectQty = 0;
                    totalRepairQty = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var DailyDefectSymptom = new Win_Qul_DefectRepair_Q_Sum_Daily_CodeView()
                            {
                                Num = i + 1,
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                GroupingName = dr["GroupingName"].ToString(),
                                RepairQty = stringFormatN0(dr["RepairQty"]),
                                RepairRate = dr["RepairRate"].ToString(),
                                InspectDate = DatePickerFormat(dr["InspectDate"].ToString())

                            };

                            if (Lib.Instance.IsNumOrAnother(DailyDefectSymptom.RepairRate))
                            {
                                DailyDefectSymptom.RepairRate = string.Format("{0:N2}", double.Parse(DailyDefectSymptom.RepairRate));
                            }

                            totalDefectQty += Convert.ToInt32(dr["DefectQty"].ToString());
                            totalRepairQty += Convert.ToInt32(dr["RepairQty"].ToString());

                            dgdDailyDefectSymptom.Items.Add(DailyDefectSymptom);
                            i++;
                        }

                        if (totalDefectQty == 0 || totalRepairQty == 0)
                        {
                            totalRepairRate = 0;
                        }
                        else
                        {
                            totalRepairRate = ((double)totalRepairQty / (double)totalDefectQty) * 100;
                        }

                        dgdTotal2.Items.Clear();
                        var DefectSymptom2 = new Win_Qul_DefectRepair_Q_CodeView()
                        {
                            //cls = dr["cls"].ToString(),
                            GroupingName = "총계",
                            DefectQty = stringFormatN0(totalDefectQty),
                            RepairQty = stringFormatN0(totalRepairQty),
                            RepairRate = totalRepairRate.ToString(),
                            ColorBlue = "true"
                        };

                        if (Lib.Instance.IsNumOrAnother(DefectSymptom2.RepairRate))
                        {
                            DefectSymptom2.RepairRate = string.Format("{0:N2}", double.Parse(DefectSymptom2.RepairRate));
                        }

                        dgdTotal2.Items.Add(DefectSymptom2);
                    }
                }

                if (ds2 != null && ds2.Tables.Count > 0)
                {
                    DataTable dt = ds2.Tables[0];
                    int i = 0;
                    totalDefectQty = 0;
                    totalRepairQty = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var DailyDefectReason = new Win_Qul_DefectRepair_Q_Sum_Daily_CodeView()
                            {
                                Num = i + 1,
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                GroupingName = dr["GroupingName"].ToString(),
                                RepairQty = stringFormatN0(dr["RepairQty"]),
                                RepairRate = dr["RepairRate"].ToString(),
                                InspectDate = DatePickerFormat(dr["InspectDate"].ToString())
                            };

                            if (Lib.Instance.IsNumOrAnother(DailyDefectReason.RepairRate))
                            {
                                DailyDefectReason.RepairRate = string.Format("{0:N2}", double.Parse(DailyDefectReason.RepairRate));
                            }

                            totalDefectQty += Convert.ToInt32(dr["DefectQty"].ToString());
                            totalRepairQty += Convert.ToInt32(dr["RepairQty"].ToString());

                            dgdDailyDefectReason.Items.Add(DailyDefectReason);
                            i++;
                        }

                        if (totalDefectQty == 0 || totalRepairQty == 0)
                        {
                            totalRepairRate = 0;
                        }
                        else
                        {
                            totalRepairRate = ((double)totalRepairQty / (double)totalDefectQty) * 100;
                        }

                        var DefectSymptom2 = new Win_Qul_DefectRepair_Q_CodeView()
                        {
                            //cls = dr["cls"].ToString(),
                            GroupingName = "총계",
                            DefectQty = stringFormatN0(totalDefectQty),
                            RepairQty = stringFormatN0(totalRepairQty),
                            RepairRate = totalRepairRate.ToString(),
                            ColorBlue = "true"
                        };

                        if (Lib.Instance.IsNumOrAnother(DefectSymptom2.RepairRate))
                        {
                            DefectSymptom2.RepairRate = string.Format("{0:N2}", double.Parse(DefectSymptom2.RepairRate));
                        }

                        //dgdDailyDefectReason.Items.Add(DefectSymptom2);
                    }
                }

                if (ds3 != null && ds3.Tables.Count > 0)
                {
                    DataTable dt = ds3.Tables[0];
                    int i = 0;
                    totalDefectQty = 0;
                    totalRepairQty = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var DailyDefectCustom = new Win_Qul_DefectRepair_Q_Sum_Daily_CodeView()
                            {
                                Num = i + 1,
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                GroupingName = dr["GroupingName"].ToString(),
                                RepairQty = stringFormatN0(dr["RepairQty"]),
                                RepairRate = dr["RepairRate"].ToString(),
                                InspectDate = DatePickerFormat(dr["InspectDate"].ToString())
                            };

                            if (Lib.Instance.IsNumOrAnother(DailyDefectCustom.RepairRate))
                            {
                                DailyDefectCustom.RepairRate = string.Format("{0:N2}", double.Parse(DailyDefectCustom.RepairRate));
                            }

                            totalDefectQty += Convert.ToInt32(dr["DefectQty"].ToString());
                            totalRepairQty += Convert.ToInt32(dr["RepairQty"].ToString());

                            dgdDailyDefectCustom.Items.Add(DailyDefectCustom);
                            i++;
                        }

                        if (totalDefectQty == 0 || totalRepairQty == 0)
                        {
                            totalRepairRate = 0;
                        }
                        else
                        {
                            totalRepairRate = ((double)totalRepairQty / (double)totalDefectQty) * 100;
                        }

                        var DefectSymptom2 = new Win_Qul_DefectRepair_Q_CodeView()
                        {
                            //cls = dr["cls"].ToString(),
                            GroupingName = "총계",
                            DefectQty = stringFormatN0(totalDefectQty),
                            RepairQty = stringFormatN0(totalRepairQty),
                            RepairRate = totalRepairRate.ToString(),
                            ColorBlue = "true"
                        };

                        if (Lib.Instance.IsNumOrAnother(DefectSymptom2.RepairRate))
                        {
                            DefectSymptom2.RepairRate = string.Format("{0:N2}", double.Parse(DefectSymptom2.RepairRate));
                        }

                        //dgdDailyDefectCustom.Items.Add(DefectSymptom2);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        private void FillGridMonthTab_Monthly()
        {
            if (dgdDefectRepairMonth.Items.Count > 0)
            {
                dgdDefectRepairMonth.Items.Clear();
            }

            if (lvcMonthChart.Series != null && lvcMonthChart.Series.Count > 0)
            {
                lvcMonthChart.Series.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("nchkDate", chkDate.IsChecked == true ? 1 : 0);
                //sqlParameter.Add("StartYYYYMM", dtpSDate.SelectedDate.Value.ToString("yyyyMM"));
                //sqlParameter.Add("EndYYYYMM", dtpEDate.SelectedDate.Value.ToString("yyyyMM"));
                sqlParameter.Add("YYYY", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyy") : "");

                sqlParameter.Add("nchkDefectStep", 1);
                sqlParameter.Add("DefectStep", cboOccurStepSrh.SelectedValue.ToString());
                sqlParameter.Add("nchkCustom", chkCustom.IsChecked == true ? 1 : 0);

                sqlParameter.Add("CustomID", chkCustom.IsChecked == true ? txtCustom.Tag.ToString() : "");
                sqlParameter.Add("nchkArticleID", 0);// chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", "");// chkArticle.IsChecked == true ? txtArticle.Tag.ToString() : "");
                //sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");
                sqlParameter.Add("nchkBuyerArticleNo", chkArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkArticleNo.IsChecked == true ? (txtArticleNo.Tag != null ? txtArticleNo.Tag.ToString() : "") : "");
                ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectRepair_Monthly", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinMonthly = new Win_Qul_DefectRepair_Q_Month_CodeView()
                            {
                                ProdQty1 = Lib.Instance.returnNumString(dr["ProdQty1"].ToString()),
                                ProdQty2 = Lib.Instance.returnNumString(dr["ProdQty2"].ToString()),
                                ProdQty3 = Lib.Instance.returnNumString(dr["ProdQty3"].ToString()),
                                ProdQty4 = Lib.Instance.returnNumString(dr["ProdQty4"].ToString()),
                                ProdQty5 = Lib.Instance.returnNumString(dr["ProdQty5"].ToString()),
                                ProdQty6 = Lib.Instance.returnNumString(dr["ProdQty6"].ToString()),
                                ProdQty7 = Lib.Instance.returnNumString(dr["ProdQty7"].ToString()),
                                ProdQty8 = Lib.Instance.returnNumString(dr["ProdQty8"].ToString()),
                                ProdQty9 = Lib.Instance.returnNumString(dr["ProdQty9"].ToString()),
                                ProdQty10 = Lib.Instance.returnNumString(dr["ProdQty10"].ToString()),
                                ProdQty11 = Lib.Instance.returnNumString(dr["ProdQty11"].ToString()),
                                ProdQty12 = Lib.Instance.returnNumString(dr["ProdQty12"].ToString()),

                                DefectQty1 = Lib.Instance.returnNumString(dr["DefectQty1"].ToString()),
                                DefectQty2 = Lib.Instance.returnNumString(dr["DefectQty2"].ToString()),
                                DefectQty3 = Lib.Instance.returnNumString(dr["DefectQty3"].ToString()),
                                DefectQty4 = Lib.Instance.returnNumString(dr["DefectQty4"].ToString()),
                                DefectQty5 = Lib.Instance.returnNumString(dr["DefectQty5"].ToString()),
                                DefectQty6 = Lib.Instance.returnNumString(dr["DefectQty6"].ToString()),
                                DefectQty7 = Lib.Instance.returnNumString(dr["DefectQty7"].ToString()),
                                DefectQty8 = Lib.Instance.returnNumString(dr["DefectQty8"].ToString()),
                                DefectQty9 = Lib.Instance.returnNumString(dr["DefectQty9"].ToString()),
                                DefectQty10 = Lib.Instance.returnNumString(dr["DefectQty10"].ToString()),
                                DefectQty11 = Lib.Instance.returnNumString(dr["DefectQty11"].ToString()),
                                DefectQty12 = Lib.Instance.returnNumString(dr["DefectQty12"].ToString()),

                                RepairRate1 = Lib.Instance.returnNumStringTwo(dr["RepairRate1"].ToString()),
                                RepairRate2 = Lib.Instance.returnNumStringTwo(dr["RepairRate2"].ToString()),
                                RepairRate3 = Lib.Instance.returnNumStringTwo(dr["RepairRate3"].ToString()),
                                RepairRate4 = Lib.Instance.returnNumStringTwo(dr["RepairRate4"].ToString()),
                                RepairRate5 = Lib.Instance.returnNumStringTwo(dr["RepairRate5"].ToString()),
                                RepairRate6 = Lib.Instance.returnNumStringTwo(dr["RepairRate6"].ToString()),
                                RepairRate7 = Lib.Instance.returnNumStringTwo(dr["RepairRate7"].ToString()),
                                RepairRate8 = Lib.Instance.returnNumStringTwo(dr["RepairRate8"].ToString()),
                                RepairRate9 = Lib.Instance.returnNumStringTwo(dr["RepairRate9"].ToString()),
                                RepairRate10 = Lib.Instance.returnNumStringTwo(dr["RepairRate10"].ToString()),
                                RepairRate11 = Lib.Instance.returnNumStringTwo(dr["RepairRate11"].ToString()),
                                RepairRate12 = Lib.Instance.returnNumStringTwo(dr["RepairRate12"].ToString()),

                                RepairQty1 = Lib.Instance.returnNumString(dr["RepairQty1"].ToString()),
                                RepairQty2 = Lib.Instance.returnNumString(dr["RepairQty2"].ToString()),
                                RepairQty3 = Lib.Instance.returnNumString(dr["RepairQty3"].ToString()),
                                RepairQty4 = Lib.Instance.returnNumString(dr["RepairQty4"].ToString()),
                                RepairQty5 = Lib.Instance.returnNumString(dr["RepairQty5"].ToString()),
                                RepairQty6 = Lib.Instance.returnNumString(dr["RepairQty6"].ToString()),
                                RepairQty7 = Lib.Instance.returnNumString(dr["RepairQty7"].ToString()),
                                RepairQty8 = Lib.Instance.returnNumString(dr["RepairQty8"].ToString()),
                                RepairQty9 = Lib.Instance.returnNumString(dr["RepairQty9"].ToString()),
                                RepairQty10 = Lib.Instance.returnNumString(dr["RepairQty10"].ToString()),
                                RepairQty11 = Lib.Instance.returnNumString(dr["RepairQty11"].ToString()),
                                RepairQty12 = Lib.Instance.returnNumString(dr["RepairQty12"].ToString()),

                                AvgProdQty = Lib.Instance.returnNumStringTwo(dr["AvgProdQty"].ToString()),
                                AvgDefectQty = Lib.Instance.returnNumStringTwo(dr["AvgDefectQty"].ToString()),
                                AvgRepairQty = Lib.Instance.returnNumStringTwo(dr["AvgRepairQty"].ToString()),
                                AvgRepairRate = Lib.Instance.returnNumStringTwo(dr["AvgRepairRate"].ToString()),

                                TDefectQty = Lib.Instance.returnNumString(dr["TDefectQty"].ToString()),
                                TProdQty = Lib.Instance.returnNumString(dr["TProdQty"].ToString()),
                                TRepairQty = Lib.Instance.returnNumString(dr["TRepairQty"].ToString()),
                                TRepairRate = Lib.Instance.returnNumStringTwo(dr["TRepairRate"].ToString()),
                                Blank = dr["Blank"].ToString(),
                                step = dr["step"].ToString(),

                                Num1 = 1,
                                Num2 = 2,
                                Num3 = 3,
                                Num4 = 4,

                                //콤보박스 선택에 따른 수량 명칭 변경
                                strGubun1 = cboOccurStepSrh.SelectedValue.Equals("1") ? "입고수량" :
                                            cboOccurStepSrh.SelectedValue.Equals("3") ? "생산수량" :
                                            cboOccurStepSrh.SelectedValue.Equals("4") ? "검사수량" : "출고수량",

                                strGubun2 = "불량수",
                                strGubun3 = "시정수량",
                                strGubun4 = "시정률(%)",
                            };

                            if (Lib.Instance.IsNumOrAnother(WinMonthly.AvgRepairRate))
                            {
                                WinMonthly.AvgRepairRate = string.Format("{0:N2}", double.Parse(WinMonthly.AvgRepairRate));
                            }

                            if (Lib.Instance.IsNumOrAnother(WinMonthly.TRepairRate))
                            {
                                WinMonthly.TRepairRate = string.Format("{0:N2}", double.Parse(WinMonthly.TRepairRate));
                            }

                            dgdDefectRepairMonth.Items.Add(WinMonthly);


                            double MON01 = 0;
                            double MON02 = 0;
                            double MON03 = 0;
                            double MON04 = 0;
                            double MON05 = 0;
                            double MON06 = 0;
                            double MON07 = 0;
                            double MON08 = 0;
                            double MON09 = 0;
                            double MON10 = 0;
                            double MON11 = 0;
                            double MON12 = 0;

                            MON01 = Convert.ToDouble(dr["RepairRate1"].ToString());
                            MON02 = Convert.ToDouble(dr["RepairRate2"].ToString());
                            MON03 = Convert.ToDouble(dr["RepairRate3"].ToString());
                            MON04 = Convert.ToDouble(dr["RepairRate4"].ToString());
                            MON05 = Convert.ToDouble(dr["RepairRate5"].ToString());
                            MON06 = Convert.ToDouble(dr["RepairRate6"].ToString());
                            MON07 = Convert.ToDouble(dr["RepairRate7"].ToString());
                            MON08 = Convert.ToDouble(dr["RepairRate8"].ToString());
                            MON09 = Convert.ToDouble(dr["RepairRate9"].ToString());
                            MON10 = Convert.ToDouble(dr["RepairRate10"].ToString());
                            MON11 = Convert.ToDouble(dr["RepairRate11"].ToString());
                            MON12 = Convert.ToDouble(dr["RepairRate12"].ToString());

                            SeriesCollection SeriesCollection = new SeriesCollection
                            {
                                new ColumnSeries
                                {
                                    Title = "월별 불량율",
                                    Values = new ChartValues<double>
                                    {
                                        0, MON01, MON02, MON03, MON04, MON05, MON06,
                                        MON07, MON08, MON09, MON10, MON11, MON12
                                    },
                                    DataLabels = false
                                }

                            };
                            lvcMonthChart.Series = SeriesCollection;
                            i++;
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        private void FillGridMonthTab_Bottom()
        {
            try
            {
                if (dgdMonthDefectSymptom.Items.Count > 0)
                {
                    dgdMonthDefectSymptom.Items.Clear();
                }
                if (dgdMonthDefectReason.Items.Count > 0)
                {
                    dgdMonthDefectReason.Items.Clear();
                }
                if (dgdMonthDefectCustom.Items.Count > 0)
                {
                    dgdMonthDefectCustom.Items.Clear();
                }

                DataSet ds1 = null;
                DataSet ds2 = null;
                DataSet ds3 = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                int totalDefectQty;
                int totalRepairQty;
                double totalRepairRate = 0.00;
                sqlParameter.Clear();
                sqlParameter.Add("nchkDate", chkDate.IsChecked == true ? 1 : 0);
                sqlParameter.Add("StartYear", chkDate.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyy") : "");
                sqlParameter.Add("nchkDefectStep", 1);
                sqlParameter.Add("DefectStep", cboOccurStepSrh.SelectedValue.ToString());

                sqlParameter.Add("nchkCustom", chkCustom.IsChecked == true ? 1 : 0);
                sqlParameter.Add("CustomID", chkCustom.IsChecked == true ? txtCustom.Tag.ToString() : "");
                sqlParameter.Add("nchkArticleID", 0); // chkArticle.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", ""); // chkArticle.IsChecked == true ? txtArticle.Tag.ToString() : "");
                                                   //sqlParameter.Add("BuyerArticleNo", chkArticle.IsChecked == true ? txtArticle.Text : "");
                sqlParameter.Add("nchkBuyerArticleNo", chkArticleNo.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyerArticleNo", chkArticleNo.IsChecked == true ? (txtArticleNo.Tag != null ? txtArticleNo.Tag.ToString() : "") : "");
                sqlParameter.Add("sGrouping", 1);
                ds1 = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectRepair_Sum_Monthly", sqlParameter, false);

                sqlParameter.Remove("sGrouping");
                sqlParameter.Add("sGrouping", 4);   //불량원인
                ds2 = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectRepair_Sum_Monthly", sqlParameter, false);

                sqlParameter.Remove("sGrouping");
                sqlParameter.Add("sGrouping", 2);   //업체
                ds3 = DataStore.Instance.ProcedureToDataSet("xp_Qul_sStsDefectRepair_Sum_Monthly", sqlParameter, false);

                if (ds1 != null && ds1.Tables.Count > 0)
                {
                    DataTable dt = ds1.Tables[0];
                    int i = 0;
                    totalDefectQty = 0;
                    totalRepairQty = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var MonthDefectSymptom = new Win_Qul_DefectRepair_Q_Sum_Month_CodeView()
                            {
                                Num = i + 1,
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                GroupingName = dr["GroupingName"].ToString(),
                                RepairQty = stringFormatN0(dr["RepairQty"]),
                                RepairRate = dr["RepairRate"].ToString(),
                                InspectMonth = dr["InspectMonth"].ToString()
                            };

                            if (Lib.Instance.IsNumOrAnother(MonthDefectSymptom.RepairRate))
                            {
                                MonthDefectSymptom.RepairRate = string.Format("{0:N2}", double.Parse(MonthDefectSymptom.RepairRate));
                            }

                            totalDefectQty += Convert.ToInt32(dr["DefectQty"].ToString());
                            totalRepairQty += Convert.ToInt32(dr["RepairQty"].ToString());

                            dgdMonthDefectSymptom.Items.Add(MonthDefectSymptom);
                            i++;
                        }

                        if (totalDefectQty == 0 || totalRepairQty == 0)
                        {
                            totalRepairRate = 0;
                        }
                        else
                        {
                            totalRepairRate = ((double)totalRepairQty / (double)totalDefectQty) * 100;
                        }
                        dgdTotal3.Items.Clear();
                        var DefectSymptom2 = new Win_Qul_DefectRepair_Q_CodeView()
                        {
                            DefectQty = stringFormatN0(totalDefectQty),
                            GroupingName = "총계",
                            RepairQty = stringFormatN0(totalRepairQty),
                            RepairRate = totalRepairRate.ToString(),
                            ColorBlue = "true"
                        };

                        if (Lib.Instance.IsNumOrAnother(DefectSymptom2.RepairRate))
                        {
                            DefectSymptom2.RepairRate = string.Format("{0:N2}", double.Parse(DefectSymptom2.RepairRate));
                        }

                        dgdTotal3.Items.Add(DefectSymptom2);
                        //dgdMonthDefectSymptom.Items.Add(DefectSymptom2);
                    }
                }

                if (ds2 != null && ds2.Tables.Count > 0)
                {
                    DataTable dt = ds2.Tables[0];
                    int i = 0;
                    totalDefectQty = 0;
                    totalRepairQty = 0;
                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var MonthDefectReason = new Win_Qul_DefectRepair_Q_Sum_Month_CodeView()
                            {
                                Num = i + 1,
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                GroupingName = dr["GroupingName"].ToString(),
                                RepairQty = stringFormatN0(dr["RepairQty"]),
                                RepairRate = dr["RepairRate"].ToString(),
                                InspectMonth = dr["InspectMonth"].ToString()
                            };

                            if (Lib.Instance.IsNumOrAnother(MonthDefectReason.RepairRate))
                            {
                                MonthDefectReason.RepairRate = string.Format("{0:N2}", double.Parse(MonthDefectReason.RepairRate));
                            }

                            totalDefectQty += Convert.ToInt32(dr["DefectQty"].ToString());
                            totalRepairQty += Convert.ToInt32(dr["RepairQty"].ToString());

                            dgdMonthDefectReason.Items.Add(MonthDefectReason);
                            i++;
                        }

                        if (totalDefectQty == 0 || totalRepairQty == 0)
                        {
                            totalRepairRate = 0;
                        }
                        else
                        {
                            totalRepairRate = ((double)totalRepairQty / (double)totalDefectQty) * 100;
                        }

                        var DefectSymptom2 = new Win_Qul_DefectRepair_Q_CodeView()
                        {
                            GroupingName = "총계",
                            DefectQty = stringFormatN0(totalDefectQty),
                            RepairQty = stringFormatN0(totalRepairQty),
                            RepairRate = totalRepairRate.ToString(),
                            ColorBlue = "true"
                        };

                        if (Lib.Instance.IsNumOrAnother(DefectSymptom2.RepairRate))
                        {
                            DefectSymptom2.RepairRate = string.Format("{0:N2}", double.Parse(DefectSymptom2.RepairRate));
                        }

                        //dgdMonthDefectReason.Items.Add(DefectSymptom2);
                    }
                }

                if (ds3 != null && ds3.Tables.Count > 0)
                {
                    DataTable dt = ds3.Tables[0];
                    int i = 0;
                    totalDefectQty = 0;
                    totalRepairQty = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var MonthDefectCustom = new Win_Qul_DefectRepair_Q_Sum_Month_CodeView()
                            {
                                Num = i + 1,
                                DefectQty = stringFormatN0(dr["DefectQty"]),
                                GroupingName = dr["GroupingName"].ToString(),
                                RepairQty = stringFormatN0(dr["RepairQty"]),
                                RepairRate = dr["RepairRate"].ToString(),
                                InspectMonth = dr["InspectMonth"].ToString()
                            };

                            if (Lib.Instance.IsNumOrAnother(MonthDefectCustom.RepairRate))
                            {
                                MonthDefectCustom.RepairRate = string.Format("{0:N2}", double.Parse(MonthDefectCustom.RepairRate));
                            }

                            totalDefectQty += Convert.ToInt32(dr["DefectQty"].ToString());
                            totalRepairQty += Convert.ToInt32(dr["RepairQty"].ToString());

                            dgdMonthDefectCustom.Items.Add(MonthDefectCustom);
                            i++;
                        }

                        if (totalDefectQty == 0 || totalRepairQty == 0)
                        {
                            totalRepairRate = 0;
                        }
                        else
                        {
                            totalRepairRate = ((double)totalRepairQty / (double)totalDefectQty) * 100;
                        }

                        var DefectSymptom2 = new Win_Qul_DefectRepair_Q_CodeView()
                        {
                            GroupingName = "총계",
                            DefectQty = stringFormatN0(totalDefectQty),
                            RepairQty = stringFormatN0(totalRepairQty),
                            RepairRate = totalRepairRate.ToString(),
                            ColorBlue = "true"
                        };

                        if (Lib.Instance.IsNumOrAnother(DefectSymptom2.RepairRate))
                        {
                            DefectSymptom2.RepairRate = string.Format("{0:N2}", double.Parse(DefectSymptom2.RepairRate));
                        }

                        //dgdMonthDefectCustom.Items.Add(DefectSymptom2);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        private void btnSymptomGraph_Click(object sender, RoutedEventArgs e)
        {
            if (dgdAllDefectSymptom.Items.Count <= 0)
            {
                MessageBox.Show("유형별 불량건에 대한 조회결과가 없습니다.");
                return;
            }
            if (cboOccurStepSrh.SelectedValue.ToString().Equals("0"))
            {
                MessageBox.Show("불량발생단계를 선택하시고 조회하세요.");
                return;
            }
            PopUp.ShowCircleGraph SCG = new PopUp.ShowCircleGraph(DT_SYMPTOM, 3);
            SCG.ShowDialog();
        }

        private void btnReasonGraph_Click(object sender, RoutedEventArgs e)
        {
            if (dgdAllDefectReason.Items.Count <= 0)
            {
                MessageBox.Show("원인별 불량건에 대한 조회결과가 없습니다.");
                return;
            }
            if (cboOccurStepSrh.SelectedValue.ToString().Equals("0"))
            {
                MessageBox.Show("불량발생단계를 선택하시고 조회하세요.");
                return;
            }
            PopUp.ShowCircleGraph SCG = new PopUp.ShowCircleGraph(DT_REASON, 3);
            SCG.ShowDialog();
        }

        private void btnCustomGraph_Click(object sender, RoutedEventArgs e)
        {
            if (dgdAllDefectCustom.Items.Count <= 0)
            {
                MessageBox.Show("업체별 불량건에 대한 조회결과가 없습니다.");
                return;
            }
            if (cboOccurStepSrh.SelectedValue.ToString().Equals("0"))
            {
                MessageBox.Show("불량발생단계를 선택하시고 조회하세요.");
                return;
            }
            PopUp.ShowCircleGraph SCG = new PopUp.ShowCircleGraph(DT_CUSTOM, 3);
            SCG.ShowDialog();
        }

        private void chkDefectOccurStep_Checked(object sender, RoutedEventArgs e)
        {
            //cboOccurStepSrh.IsEnabled = true;
        }

        private void chkDefectOccurStep_Unchecked(object sender, RoutedEventArgs e)
        {
            //cboOccurStepSrh.IsEnabled = false;
        }

        //일별 탭
        private void TabItem_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            TabItem tabItem = sender as TabItem;
            if (tabItem != null && tabItem.IsSelected)
            {
                dtpEDate.Visibility = Visibility.Hidden;
                btnLastSixMonth.Visibility = Visibility.Hidden;
            }

            e.Handled = true;
        }

        //월별 탭
        private void Month_TabItem_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            TabItem tabItem = sender as TabItem;
            if (tabItem != null && tabItem.IsSelected)
            {

                dtpEDate.Visibility = Visibility.Hidden;
                btnLastSixMonth.Visibility = Visibility.Hidden;
            }

            e.Handled = true;
        }

        //전체 탭
        private void All_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            TabItem tabItem = sender as TabItem;
            if (tabItem != null && tabItem.IsSelected)
            {
                dtpEDate.Visibility = Visibility.Visible;
                btnLastSixMonth.Visibility = Visibility.Visible;
            }
            e.Handled = true;
        }

        private void lblArticleNo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleNo.IsChecked == true)
            {
                txtArticleNo.IsEnabled = false;
                btnPfArticle.IsEnabled = false;
                chkArticleNo.IsChecked = false;
            }
            else
            {
                txtArticleNo.IsEnabled = true;
                btnPfArticle.IsEnabled = true;
                chkArticleNo.IsChecked = true;
            }
        }


        private void txtArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtArticleNo, 77, txtArticleNo.Text);
            }

        }

        private void btnPfArticleNo_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleNo, 77, txtArticleNo.Text);
        }

        private void chkArticleNo_Checked(object sender, RoutedEventArgs e)
        {
            if (chkArticleNo.IsChecked == true)
            {
                txtArticleNo.IsEnabled = true;
                btnPfArticleNo.IsEnabled = true;
                chkArticleNo.IsChecked = true;
            }
        }

        private void chkArticleNo_Unchecked(object sender, RoutedEventArgs e)
        {
            if (chkArticleNo.IsChecked == false)
            {
                txtArticleNo.IsEnabled = false;
                btnPfArticleNo.IsEnabled = false;
                chkArticleNo.IsChecked = false;
            }
        }

        // 천단위 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
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

        private void lblDate_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (chkDate.IsChecked == true)
            {
                chkDate.IsChecked = false;
                dtpSDate.IsEnabled = false;
                dtpEDate.IsEnabled = false;
            }
            else
            {
                chkDate.IsChecked = true;
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
        }

        private void chkDate_Click(object sender, RoutedEventArgs e)
        {
            if (chkDate.IsChecked == true)
            {
                chkDate.IsChecked = true;
                dtpSDate.IsEnabled = true;
                dtpEDate.IsEnabled = true;
            }
            else
            {
                chkDate.IsChecked = false;
                dtpSDate.IsEnabled = false;
                dtpEDate.IsEnabled = false;
            }
        }

        //private void CommonControl_Click(object sender, MouseButtonEventArgs e)
        //{
        //    lib.CommonControl_Click(sender, e);
        //}

        //private void CommonControl_Click(object sender, RoutedEventArgs e)
        //{
        //    lib.CommonControl_Click(sender, e);
        //}




        public class Win_Qul_DefectRepair_Q_CodeView : BaseView
        {
            public int Num { get; set; }
            public string cls { get; set; }
            public string DefectQty { get; set; }
            public string RepairQty { get; set; }
            public string GroupingName { get; set; }
            public string RepairRate { get; set; }
            public string RepairRate1 { get; set; }
            public string ColorBlue { get; set; }
        }

        public class Win_Qul_DefectRepair_Q_Daily_CodeView : BaseView
        {
            public string step { get; set; }
            public string OccurDate { get; set; }

            public int Num1 { get; set; }
            public int Num2 { get; set; }
            public int Num3 { get; set; }
            public int Num4 { get; set; }

            public string strGubun1 { get; set; }
            public string strGubun2 { get; set; }
            public string strGubun3 { get; set; }
            public string strGubun4 { get; set; }

            public string YYYYMMDD1 { get; set; }
            public string ProdQty1 { get; set; }
            public string DefectQty1 { get; set; }
            public string RepairQty1 { get; set; }
            public string RepairRate1 { get; set; }

            public string YYYYMMDD2 { get; set; }
            public string ProdQty2 { get; set; }
            public string DefectQty2 { get; set; }
            public string RepairQty2 { get; set; }
            public string RepairRate2 { get; set; }

            public string YYYYMMDD3 { get; set; }
            public string ProdQty3 { get; set; }
            public string DefectQty3 { get; set; }
            public string RepairQty3 { get; set; }
            public string RepairRate3 { get; set; }

            public string YYYYMMDD4 { get; set; }
            public string ProdQty4 { get; set; }
            public string DefectQty4 { get; set; }
            public string RepairQty4 { get; set; }
            public string RepairRate4 { get; set; }

            public string YYYYMMDD5 { get; set; }
            public string ProdQty5 { get; set; }
            public string DefectQty5 { get; set; }
            public string RepairQty5 { get; set; }
            public string RepairRate5 { get; set; }

            public string YYYYMMDD6 { get; set; }
            public string ProdQty6 { get; set; }
            public string DefectQty6 { get; set; }
            public string RepairQty6 { get; set; }
            public string RepairRate6 { get; set; }

            public string YYYYMMDD7 { get; set; }
            public string ProdQty7 { get; set; }
            public string DefectQty7 { get; set; }
            public string RepairQty7 { get; set; }
            public string RepairRate7 { get; set; }

            public string YYYYMMDD8 { get; set; }
            public string ProdQty8 { get; set; }
            public string DefectQty8 { get; set; }
            public string RepairQty8 { get; set; }
            public string RepairRate8 { get; set; }

            public string YYYYMMDD9 { get; set; }
            public string ProdQty9 { get; set; }
            public string DefectQty9 { get; set; }
            public string RepairQty9 { get; set; }
            public string RepairRate9 { get; set; }

            public string YYYYMMDD10 { get; set; }
            public string ProdQty10 { get; set; }
            public string DefectQty10 { get; set; }
            public string RepairQty10 { get; set; }
            public string RepairRate10 { get; set; }

            public string YYYYMMDD11 { get; set; }
            public string ProdQty11 { get; set; }
            public string DefectQty11 { get; set; }
            public string RepairQty11 { get; set; }
            public string RepairRate11 { get; set; }

            public string YYYYMMDD12 { get; set; }
            public string ProdQty12 { get; set; }
            public string DefectQty12 { get; set; }
            public string RepairQty12 { get; set; }
            public string RepairRate12 { get; set; }

            public string YYYYMMDD13 { get; set; }
            public string ProdQty13 { get; set; }
            public string DefectQty13 { get; set; }
            public string RepairQty13 { get; set; }
            public string RepairRate13 { get; set; }

            public string YYYYMMDD14 { get; set; }
            public string ProdQty14 { get; set; }
            public string DefectQty14 { get; set; }
            public string RepairQty14 { get; set; }
            public string RepairRate14 { get; set; }

            public string YYYYMMDD15 { get; set; }
            public string ProdQty15 { get; set; }
            public string DefectQty15 { get; set; }
            public string RepairQty15 { get; set; }
            public string RepairRate15 { get; set; }

            public string YYYYMMDD16 { get; set; }
            public string ProdQty16 { get; set; }
            public string DefectQty16 { get; set; }
            public string RepairQty16 { get; set; }
            public string RepairRate16 { get; set; }

            public string YYYYMMDD17 { get; set; }
            public string ProdQty17 { get; set; }
            public string DefectQty17 { get; set; }
            public string RepairQty17 { get; set; }
            public string RepairRate17 { get; set; }

            public string YYYYMMDD18 { get; set; }
            public string ProdQty18 { get; set; }
            public string DefectQty18 { get; set; }
            public string RepairQty18 { get; set; }
            public string RepairRate18 { get; set; }

            public string YYYYMMDD19 { get; set; }
            public string ProdQty19 { get; set; }
            public string DefectQty19 { get; set; }
            public string RepairQty19 { get; set; }
            public string RepairRate19 { get; set; }

            public string YYYYMMDD20 { get; set; }
            public string ProdQty20 { get; set; }
            public string DefectQty20 { get; set; }
            public string RepairQty20 { get; set; }
            public string RepairRate20 { get; set; }

            public string YYYYMMDD21 { get; set; }
            public string ProdQty21 { get; set; }
            public string DefectQty21 { get; set; }
            public string RepairQty21 { get; set; }
            public string RepairRate21 { get; set; }

            public string YYYYMMDD22 { get; set; }
            public string ProdQty22 { get; set; }
            public string DefectQty22 { get; set; }
            public string RepairQty22 { get; set; }
            public string RepairRate22 { get; set; }

            public string YYYYMMDD23 { get; set; }
            public string ProdQty23 { get; set; }
            public string DefectQty23 { get; set; }
            public string RepairQty23 { get; set; }
            public string RepairRate23 { get; set; }

            public string YYYYMMDD24 { get; set; }
            public string ProdQty24 { get; set; }
            public string DefectQty24 { get; set; }
            public string RepairQty24 { get; set; }
            public string RepairRate24 { get; set; }

            public string YYYYMMDD25 { get; set; }
            public string ProdQty25 { get; set; }
            public string DefectQty25 { get; set; }
            public string RepairQty25 { get; set; }
            public string RepairRate25 { get; set; }

            public string YYYYMMDD26 { get; set; }
            public string ProdQty26 { get; set; }
            public string DefectQty26 { get; set; }
            public string RepairQty26 { get; set; }
            public string RepairRate26 { get; set; }

            public string YYYYMMDD27 { get; set; }
            public string ProdQty27 { get; set; }
            public string DefectQty27 { get; set; }
            public string RepairQty27 { get; set; }
            public string RepairRate27 { get; set; }

            public string YYYYMMDD28 { get; set; }
            public string ProdQty28 { get; set; }
            public string DefectQty28 { get; set; }
            public string RepairQty28 { get; set; }
            public string RepairRate28 { get; set; }

            public string YYYYMMDD29 { get; set; }
            public string ProdQty29 { get; set; }
            public string DefectQty29 { get; set; }
            public string RepairQty29 { get; set; }
            public string RepairRate29 { get; set; }

            public string YYYYMMDD30 { get; set; }
            public string ProdQty30 { get; set; }
            public string DefectQty30 { get; set; }
            public string RepairQty30 { get; set; }
            public string RepairRate30 { get; set; }

            public string YYYYMMDD31 { get; set; }
            public string ProdQty31 { get; set; }
            public string DefectQty31 { get; set; }
            public string RepairQty31 { get; set; }
            public string RepairRate31 { get; set; }

            public string TProdQty { get; set; }
            public string TDefectQty { get; set; }
            public string TRepairQty { get; set; }
            public string TRepairRate { get; set; }

            public string AvgProdQty { get; set; }
            public string AvgDefectQty { get; set; }
            public string AvgRepairQty { get; set; }
            public string AvgRepairRate { get; set; }

            public string DayCount { get; set; }
        }

        public class Win_Qul_DefectRepair_Q_Sum_Daily_CodeView : BaseView
        {
            public int Num { get; set; }
            public string InspectDate { get; set; }
            public string GroupingName { get; set; }
            public string DefectQty { get; set; }
            public string RepairQty { get; set; }
            public string RepairRate { get; set; }
        }

        public class Win_Qul_DefectRepair_Q_Month_CodeView : BaseView
        {
            public string step { get; set; }
            public string Blank { get; set; }

            public int Num1 { get; set; }
            public int Num2 { get; set; }
            public int Num3 { get; set; }
            public int Num4 { get; set; }

            public string strGubun1 { get; set; }
            public string strGubun2 { get; set; }
            public string strGubun3 { get; set; }
            public string strGubun4 { get; set; }

            public string Month01 { get; set; }
            public string ProdQty1 { get; set; }
            public string DefectQty1 { get; set; }
            public string RepairQty1 { get; set; }
            public string RepairRate1 { get; set; }

            public string Month02 { get; set; }
            public string ProdQty2 { get; set; }
            public string DefectQty2 { get; set; }
            public string RepairQty2 { get; set; }
            public string RepairRate2 { get; set; }

            public string Month03 { get; set; }
            public string ProdQty3 { get; set; }
            public string DefectQty3 { get; set; }
            public string RepairQty3 { get; set; }
            public string RepairRate3 { get; set; }

            public string Month04 { get; set; }
            public string ProdQty4 { get; set; }
            public string DefectQty4 { get; set; }
            public string RepairQty4 { get; set; }
            public string RepairRate4 { get; set; }

            public string Month05 { get; set; }
            public string ProdQty5 { get; set; }
            public string DefectQty5 { get; set; }
            public string RepairQty5 { get; set; }
            public string RepairRate5 { get; set; }

            public string Month06 { get; set; }
            public string ProdQty6 { get; set; }
            public string DefectQty6 { get; set; }
            public string RepairQty6 { get; set; }
            public string RepairRate6 { get; set; }

            public string Month07 { get; set; }
            public string ProdQty7 { get; set; }
            public string DefectQty7 { get; set; }
            public string RepairQty7 { get; set; }
            public string RepairRate7 { get; set; }

            public string Month08 { get; set; }
            public string ProdQty8 { get; set; }
            public string DefectQty8 { get; set; }
            public string RepairQty8 { get; set; }
            public string RepairRate8 { get; set; }

            public string Month09 { get; set; }
            public string ProdQty9 { get; set; }
            public string DefectQty9 { get; set; }
            public string RepairQty9 { get; set; }
            public string RepairRate9 { get; set; }

            public string Month10 { get; set; }
            public string ProdQty10 { get; set; }
            public string DefectQty10 { get; set; }
            public string RepairQty10 { get; set; }
            public string RepairRate10 { get; set; }

            public string Month11 { get; set; }
            public string ProdQty11 { get; set; }
            public string DefectQty11 { get; set; }
            public string RepairQty11 { get; set; }
            public string RepairRate11 { get; set; }

            public string Month12 { get; set; }
            public string ProdQty12 { get; set; }
            public string DefectQty12 { get; set; }
            public string RepairQty12 { get; set; }
            public string RepairRate12 { get; set; }

            public string TProdQty { get; set; }
            public string TDefectQty { get; set; }
            public string TRepairQty { get; set; }
            public string TRepairRate { get; set; }

            public string AvgProdQty { get; set; }
            public string AvgDefectQty { get; set; }
            public string AvgRepairQty { get; set; }
            public string AvgRepairRate { get; set; }

            public string DayCount { get; set; }
        }

        public class Win_Qul_DefectRepair_Q_Sum_Month_CodeView : BaseView
        {
            public int Num { get; set; }
            public string InspectMonth { get; set; }
            public string GroupingName { get; set; }
            public string DefectQty { get; set; }
            public string RepairQty { get; set; }
            public string RepairRate { get; set; }
        }

    }
}
