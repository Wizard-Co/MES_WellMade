using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_com_QulDefectGoal_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_Qul_QulDefectGoal_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        string sInspectPoint = string.Empty;
        string strFlag = string.Empty;
        string strRemain = string.Empty;

        int Wh_Ar_SelectedLastIndex = 0;        // 그리드 마지막 선택 줄 임시저장 그릇

        ObservableCollection<Win_Qul_QulDefectGoal_U_CodeView> ovcDefectGoalYear
            = new ObservableCollection<Win_Qul_QulDefectGoal_U_CodeView>();
        ObservableCollection<Win_Qul_QulDefectGoal_U_Sub_CodeView> ovcDefectGoal1
            = new ObservableCollection<Win_Qul_QulDefectGoal_U_Sub_CodeView>();
        ObservableCollection<Win_Qul_QulDefectGoal_U_Sub_CodeView> ovcDefectGoal2
            = new ObservableCollection<Win_Qul_QulDefectGoal_U_Sub_CodeView>();

        Lib lib = new Lib();

        public Win_Qul_QulDefectGoal_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            TbnJaju_Click(tbnJaju, null);
            chkYear.IsChecked = true;
            BtnLast5Years_Click(null, null);
            txtCallYear.Text = DateTime.Today.ToString("yyyy");
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
                tbnCustom.IsChecked = false;
                if (!sInspectPoint.Equals("1"))
                {
                    dgdMain.ItemsSource = null;
                    dgdMain.Refresh();
                    dgdSub1.ItemsSource = null;
                    dgdSub1.Refresh();
                    dgdSub2.ItemsSource = null;
                    dgdSub2.Refresh();

                }
                sInspectPoint = "1";
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
                tbnCustom.IsChecked = false;
                if (!sInspectPoint.Equals("5"))
                {
                    dgdMain.ItemsSource = null;
                    dgdMain.Refresh();
                    dgdSub1.ItemsSource = null;
                    dgdSub1.Refresh();
                    dgdSub2.ItemsSource = null;
                    dgdSub2.Refresh();
                }
                sInspectPoint = "5";
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
                tbnCustom.IsChecked = false;
                if (!sInspectPoint.Equals("3"))
                {
                    dgdMain.ItemsSource = null;
                    dgdMain.Refresh();
                    dgdSub1.ItemsSource = null;
                    dgdSub1.Refresh();
                    dgdSub2.ItemsSource = null;
                    dgdSub2.Refresh();
                }
                sInspectPoint = "3";
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
                tbnCustom.IsChecked = false;
                if (!sInspectPoint.Equals("9"))
                {
                    dgdMain.ItemsSource = null;
                    dgdMain.Refresh();
                    dgdSub1.ItemsSource = null;
                    dgdSub1.Refresh();
                    dgdSub2.ItemsSource = null;
                    dgdSub2.Refresh();
                }
                sInspectPoint = "9";
            }
            else
            {
                toggleButton.IsChecked = true;
            }
        }

        /// <summary>
        /// 고객
        /// </summary>
        private void TbnCustom_Click(object sender, RoutedEventArgs e)
        {
            ToggleButton toggleButton = sender as ToggleButton;
            if (toggleButton.IsChecked == true)
            {
                tbnInCome.IsChecked = false;
                tbnOutCome.IsChecked = false;
                tbnProcessCycle.IsChecked = false;
                tbnJaju.IsChecked = false;
                if (!sInspectPoint.Equals("7"))
                {
                    dgdMain.ItemsSource = null;
                    dgdMain.Refresh();
                    dgdSub1.ItemsSource = null;
                    dgdSub1.Refresh();
                    dgdSub2.ItemsSource = null;
                    dgdSub2.Refresh();
                }
                sInspectPoint = "7";
            }
            else
            {
                toggleButton.IsChecked = true;
            }
        }

        /// <summary>
        /// 기간
        /// </summary>
        private void LblYear_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkYear.IsChecked == true) { chkYear.IsChecked = false; }
            else { chkYear.IsChecked = true; }
        }

        /// <summary>
        /// 기간
        /// </summary>
        private void ChkYear_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
        }

        /// <summary>
        /// 기간
        /// </summary>
        private void ChkYear_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
        }

        /// <summary>
        /// 최근 5년
        /// </summary>
        private void BtnLast5Years_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today.AddYears(-5);
            dtpEDate.SelectedDate = DateTime.Today;
        }

        /// <summary>
        /// 최근 10년
        /// </summary>
        private void BtnLast10Years_Click(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today.AddYears(-10);
            dtpEDate.SelectedDate = DateTime.Today;
        }

        /// <summary>
        /// 추가,수정 시 동작 모음
        /// </summary>
        private void ControlVisibleAndEnable_AU()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            dgdMain.IsHitTestVisible = false;
            grdInput.IsHitTestVisible = true;
        }

        /// <summary>
        /// 저장,취소 시 동작 모음
        /// </summary>
        private void ControlVisibleAndEnable_SC()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            dgdMain.IsHitTestVisible = true;
            grdInput.IsHitTestVisible = false;
        }

        /// <summary>
        /// 추가
        /// </summary>
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count > 0)
            {
                Wh_Ar_SelectedLastIndex = dgdMain.SelectedIndex;
            }
            else
            {
                Wh_Ar_SelectedLastIndex = 0;
            }

            this.DataContext = null;
            strFlag = "I";
            ControlVisibleAndEnable_AU();
            tbkMsg.Text = "자료 입력 중";
            txtYear.Text = DateTime.Today.ToString("yyyy");
            FillGridEmptySub(txtYear.Text, ovcDefectGoal1);
            dgdSub1.ItemsSource = ovcDefectGoal1;

            // 여기까지가 한계인듯.... CellContent가 여전히 null이네..ㅜ..
            //var dataGridCellInfo = new DataGridCellInfo(dgdSub1.Items[0], dgdSub1.Columns[2]);
            //var CellContent = dataGridCellInfo.Column.GetCellContent(dataGridCellInfo.Item);
            //if (CellContent != null)
            //{
            //    DataGridCell dgc = (DataGridCell)CellContent.Parent;
            //    dgc.Focus();
            //}
        }

        /// <summary>
        /// 수정
        /// </summary>
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var GoalMain = dgdMain.SelectedItem as Win_Qul_QulDefectGoal_U_CodeView;
            if (GoalMain == null)
            {
                MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
                return;
            }
            else
            {
                Wh_Ar_SelectedLastIndex = dgdMain.SelectedIndex;
                strFlag = "U";
                ControlVisibleAndEnable_AU();
                tbkMsg.Text = "자료 수정 중";
            }
        }

        /// <summary>
        /// 삭제
        /// </summary>
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var GoalMain = dgdMain.SelectedItem as Win_Qul_QulDefectGoal_U_CodeView;
                if (GoalMain == null)
                {
                    MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                    return;
                }
                else
                {
                    if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "D");

                        if (dgdMain.Items.Count > 0 && dgdMain.SelectedItem != null)
                        {
                            Wh_Ar_SelectedLastIndex = dgdMain.SelectedIndex;
                        }

                        if (Procedure.Instance.DeleteData(sInspectPoint, GoalMain.YYYY,
                            "sInspectGubun", "sYYYY", "xp_Qul_dDefectGoal"))
                        {
                            Wh_Ar_SelectedLastIndex -= 1;
                            FillGrid();
                            if (dgdMain.Items.Count > 0)
                            {
                                dgdMain.SelectedIndex = Wh_Ar_SelectedLastIndex;
                                dgdMain.Focus();
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

        /// <summary>
        /// 닫기
        /// </summary>
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        /// <summary>
        /// 조회
        /// </summary>
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

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

        /// <summary>
        /// 저장
        /// </summary>
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (CheckData())
            {
                if (SaveData())
                {
                    ControlVisibleAndEnable_SC();
                    FillGrid();

                    dgdMain.ItemsSource = null;
                    dgdMain.ItemsSource = ovcDefectGoalYear;
                    dgdMain.Items.Refresh();

                    if (strFlag == "I")     //1. 추가 > 저장했다면,
                    {
                        if (dgdMain.Items.Count > 0)
                        {
                            dgdMain.SelectedIndex = dgdMain.Items.Count - 1;
                            dgdMain.Focus();
                        }
                    }
                    else        //2. 수정 > 저장했다면,
                    {
                        dgdMain.SelectedIndex = Wh_Ar_SelectedLastIndex;
                        dgdMain.Focus();
                    }

                    strFlag = string.Empty; // 추가했는지, 수정했는지 알려면 맨 마지막에 flag 값을 비어야 한다.
                }
            }

        }

        private bool CheckData()
        {
            bool result = false;
            if (strFlag.Equals("I"))
            {
                foreach (Win_Qul_QulDefectGoal_U_CodeView MainGridClass in dgdMain.Items)
                {
                    string CheckYYYY = MainGridClass.YYYY;
                    if (CheckYYYY == txtYear.Text)
                    {
                        MessageBox.Show("이미 지정되어 있는 년도입니다. 추가할 년도를 변경해주세요.");
                        return result;
                    }
                }
            }

            result = true;
            return result;
        }




        /// <summary>
        /// 취소
        /// </summary>
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            ControlVisibleAndEnable_SC();
            FillGrid();

            if (strFlag == "I") // 1. 추가하다가 취소했다면,
            {
                if (dgdMain.Items.Count > 0)
                {
                    dgdMain.SelectedIndex = Wh_Ar_SelectedLastIndex;
                    dgdMain.Focus();
                }
            }
            else        //2. 수정하다가 취소했다면
            {
                dgdMain.SelectedIndex = Wh_Ar_SelectedLastIndex;
                dgdMain.Focus();
            }
            strFlag = string.Empty; // 추가했는지, 수정했는지 알려면 맨 마지막에 flag 값을 비어야 한다.
        }

        /// <summary>
        /// 엑셀
        /// </summary>
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;
            Lib lib = new Lib();

            string[] dgdStr = new string[6];
            dgdStr[0] = "연도별 불량률 평균";
            dgdStr[1] = "선택년도 월별 불량률";
            dgdStr[2] = "선택년도의 전년 월별 불량률";
            dgdStr[3] = dgdMain.Name;
            dgdStr[4] = dgdSub1.Name;
            dgdStr[5] = dgdSub2.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdMain);
                    else
                        dt = lib.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;
                    if (lib.GenerateExcel(dt, Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdSub1.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdSub1);
                    else
                        dt = lib.DataGirdToDataTable(dgdSub1);

                    Name = dgdSub1.Name;
                    if (lib.GenerateExcel(dt, Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
                    else
                        return;
                }
                else if (ExpExc.choice.Equals(dgdSub2.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = lib.DataGridToDTinHidden(dgdSub2);
                    else
                        dt = lib.DataGirdToDataTable(dgdSub2);

                    Name = dgdSub2.Name;
                    if (lib.GenerateExcel(dt, Name))
                    {
                        lib.excel.Visible = true;
                        lib.ReleaseExcelObject(lib.excel);
                    }
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
            lib = null;
        }

        /// <summary>
        /// 실조회
        /// </summary>
        private void FillGrid()
        {
            try
            {
                ovcDefectGoalYear.Clear();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sInspectGubun", sInspectPoint);
                sqlParameter.Add("nChkDate", chkYear.IsChecked == true ? 1 : 0);
                sqlParameter.Add("sFromYYYY", chkYear.IsChecked == true ? dtpSDate.SelectedDate.Value.ToString("yyyy") : "");
                sqlParameter.Add("sToYYYY", chkYear.IsChecked == true ? dtpEDate.SelectedDate.Value.ToString("yyyy") : "");
                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Qul_sDefectGoalYear", sqlParameter, true, "R");

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
                            var DefectGoalYear = new Win_Qul_QulDefectGoal_U_CodeView
                            {
                                YYYY = dr["YYYY"].ToString(),
                                DefectGoalAvg = dr["DefectGoalAvg"].ToString(),
                                Num = i
                            };
                            DefectGoalYear.DefectGoalAvg = Lib.Instance.returnNumStringZero(DefectGoalYear.DefectGoalAvg);
                            ovcDefectGoalYear.Add(DefectGoalYear);
                        }

                        dgdMain.ItemsSource = ovcDefectGoalYear;
                    }
                    else
                    {
                        dgdSub1.ItemsSource = null;
                        dgdSub2.ItemsSource = null;
                    }
                }
                else
                {
                    dgdSub1.ItemsSource = null;
                    dgdSub2.ItemsSource = null;
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

            dgdMain.SelectedIndex = 0;
        }

        /// <summary>
        /// 실조회(선택한 년도)
        /// </summary>
        private void FillGridSub1(string strYear)
        {
            try
            {
                ovcDefectGoal1.Clear();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sInspectGubun", sInspectPoint);
                sqlParameter.Add("sYYYY", strYear);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sDefectGoal", sqlParameter, false);

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
                            var DefectGoalMonth = new Win_Qul_QulDefectGoal_U_Sub_CodeView
                            {
                                YYYY = dr["YYYY"].ToString(),
                                Num = i,
                                MM = dr["MM"].ToString(),
                                DefectGoal = dr["DefectGoal"].ToString(),
                                sortMM = dr["sortMM"].ToString(),
                                InspectGubun = dr["InspectGubun"].ToString(),
                                AvgDefectGoal = dr["AvgDefectGoal"].ToString()
                            };
                            DefectGoalMonth.MM = DefectGoalMonth.MM + "월";
                            DefectGoalMonth.DefectGoal = Lib.Instance.returnNumStringZero(DefectGoalMonth.DefectGoal);
                            ovcDefectGoal1.Add(DefectGoalMonth);
                        }
                    }
                    else
                    {
                        FillGridEmptySub(strYear, ovcDefectGoal1);
                    }

                    dgdSub1.ItemsSource = ovcDefectGoal1;
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

        /// <summary>
        /// 실조회(선택한 년도의 전년)
        /// </summary>
        private void FillGridSub2(string strYear)
        {
            try
            {
                ovcDefectGoal2.Clear();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sInspectGubun", sInspectPoint);
                sqlParameter.Add("sYYYY", strYear);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Qul_sDefectGoal", sqlParameter, false);

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
                            var DefectGoalMonth = new Win_Qul_QulDefectGoal_U_Sub_CodeView
                            {
                                YYYY = dr["YYYY"].ToString(),
                                Num = i,
                                MM = dr["MM"].ToString(),
                                DefectGoal = dr["DefectGoal"].ToString(),
                                sortMM = dr["sortMM"].ToString(),
                                InspectGubun = dr["InspectGubun"].ToString(),
                                AvgDefectGoal = dr["AvgDefectGoal"].ToString()
                            };
                            DefectGoalMonth.MM = DefectGoalMonth.MM + "월";
                            DefectGoalMonth.DefectGoal = Lib.Instance.returnNumStringZero(DefectGoalMonth.DefectGoal);
                            ovcDefectGoal2.Add(DefectGoalMonth);
                        }
                    }
                    else
                    {
                        FillGridEmptySub(strYear, ovcDefectGoal2);
                    }

                    dgdSub2.ItemsSource = ovcDefectGoal2;
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

        private void FillGridEmptySub(string strYear, ObservableCollection<Win_Qul_QulDefectGoal_U_Sub_CodeView> ovc)
        {
            ovc.Clear();
            for (int i = 0; i < 12; i++)
            {
                var DefectGoalMonth = new Win_Qul_QulDefectGoal_U_Sub_CodeView
                {
                    YYYY = strYear,
                    Num = i + 1,
                    MM = (i + 1).ToString() + "월",
                    DefectGoal = string.Empty,
                    sortMM = string.Format("{0:00}", (i + 1)),
                    InspectGubun = sInspectPoint,
                    AvgDefectGoal = string.Empty
                };
                ovc.Add(DefectGoalMonth);
            }
        }

        //
        private void DgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var GoalMain = dgdMain.SelectedItem as Win_Qul_QulDefectGoal_U_CodeView;
            int OneYearAgo = 0;

            if (GoalMain != null)
            {
                this.DataContext = GoalMain;
                if (GoalMain.YYYY != null)
                {
                    OneYearAgo = int.Parse(GoalMain.YYYY) - 1;
                }

                FillGridSub1(GoalMain.YYYY);
                FillGridSub2(OneYearAgo.ToString());
            }
        }

        //
        private bool SaveData()
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                if (strFlag.Equals("I"))
                {
                    for (int i = 0; i < dgdSub1.Items.Count; i++)
                    {
                        var SubGoal = dgdSub1.Items[i] as Win_Qul_QulDefectGoal_U_Sub_CodeView;

                        if (SubGoal != null)
                        {
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("sInspectGubun", sInspectPoint);
                            sqlParameter.Add("YYYY", txtYear.Text);
                            sqlParameter.Add("MM", SubGoal.MM.Replace("월", ""));
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            if (SubGoal.DefectGoal.Replace(",", "").Equals(""))
                            {
                                sqlParameter.Add("DefectGoal", 0);
                            }
                            else
                            {
                                sqlParameter.Add("DefectGoal", SubGoal.DefectGoal.Replace(",", ""));
                            }
                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Qul_iDefectGoal";
                            pro1.OutputUseYN = "N";
                            pro1.OutputName = "YYYY";
                            pro1.OutputLength = "10";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);
                        }
                    }

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"C");
                    if (Confirm[0] != "success")
                    {
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    }
                    else
                    {
                        flag = true;
                    }
                }
                else
                {
                    sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sInspectGubun", sInspectPoint);
                    sqlParameter.Add("sYYYY", txtYear.Text);

                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_Qul_dDefectGoal";
                    pro1.OutputUseYN = "N";
                    pro1.OutputName = "YYYY";
                    pro1.OutputLength = "10";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);

                    for (int i = 0; i < dgdSub1.Items.Count; i++)
                    {
                        var SubGoal = dgdSub1.Items[i] as Win_Qul_QulDefectGoal_U_Sub_CodeView;

                        if (SubGoal != null)
                        {
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("sInspectGubun", sInspectPoint);
                            sqlParameter.Add("YYYY", txtYear.Text);
                            sqlParameter.Add("MM", SubGoal.MM.Replace("월", ""));
                            sqlParameter.Add("DefectGoal", SubGoal.DefectGoal.Replace(",", ""));
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_Qul_iDefectGoal";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "YYYY";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }
                    }

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"U");
                    if (Confirm[0] != "success")
                    {
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    }
                    else
                    {
                        flag = true;
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

            return flag;
        }

        private void YearText(TextBox textBox, bool Flag)
        {
            int tNumber = 0;
            if (!textBox.Text.Equals(string.Empty))
            {
                if (Lib.Instance.IsIntOrAnother(textBox.Text))
                {
                    if (Flag)
                    {
                        tNumber = int.Parse(textBox.Text) + 1;
                    }
                    else
                    {
                        tNumber = int.Parse(textBox.Text) - 1;
                    }

                    textBox.Text = tNumber.ToString();
                }

            }
        }

        private void BtnYearPlus_Click(object sender, RoutedEventArgs e)
        {
            string tFlag = (sender as Button).Tag.ToString();
            if (tFlag.Equals("1"))
            {
                YearText(txtYear, true);
            }
            else
            {
                YearText(txtCallYear, true);
            }
        }

        private void BtnYearMinus_Click(object sender, RoutedEventArgs e)
        {
            string tFlag = (sender as Button).Tag.ToString();
            if (tFlag.Equals("1"))
            {
                YearText(txtYear, false);
            }
            else
            {
                YearText(txtCallYear, false);
            }
        }

        private void BtnCallOldYearGoal_Click(object sender, RoutedEventArgs e)
        {
            FillGridSub1(txtCallYear.Text);
        }

        private void BtnLastYearDefect_Click(object sender, RoutedEventArgs e)
        {
            int OneYearAgo = 0;
            OneYearAgo = int.Parse(txtCallYear.Text) - 1;
            FillGridSub2(OneYearAgo.ToString());
        }

        private void DataGridCell_KeyDown(object sender, KeyEventArgs e)
        {
            DataGridCell dataGridCell = sender as DataGridCell;
            var GoalMon = (dataGridCell.Content as ContentPresenter).Content as Win_Qul_QulDefectGoal_U_Sub_CodeView;
            int currentIndex = GoalMon.Num - 1;
            DataGrid dataGrid = Lib.Instance.GetParent<DataGrid>(dataGridCell);

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (currentIndex == dataGrid.Items.Count - 1)
                {
                    btnSave.Focus();
                }
                else
                {
                    dataGrid.SelectedIndex = currentIndex + 1;
                    dataGrid.CurrentCell =
                        new DataGridCellInfo(dataGrid.Items[currentIndex + 1], dataGrid.Columns[2]);
                    var AutoMon = dataGrid.Items[currentIndex + 1] as Win_Qul_QulDefectGoal_U_Sub_CodeView;
                    if (strFlag.Equals("I") && !strRemain.Equals(string.Empty))
                    {
                        AutoMon.DefectGoal = strRemain;
                        strRemain = string.Empty;
                    }
                }
            }
            else if (e.Key == Key.Down)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (currentIndex == dataGrid.Items.Count - 1)
                {
                    btnSave.Focus();
                }
                else
                {
                    dataGrid.SelectedIndex = currentIndex + 1;
                    dataGrid.CurrentCell =
                        new DataGridCellInfo(dataGrid.Items[currentIndex + 1], dataGrid.Columns[2]);
                }
            }
            else if (e.Key == Key.Up)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (currentIndex == 0)
                {
                    btnSave.Focus();
                }
                else
                {
                    dataGrid.SelectedIndex = currentIndex - 1;
                    dataGrid.CurrentCell =
                        new DataGridCellInfo(dataGrid.Items[currentIndex - 1], dataGrid.Columns[2]);
                }
            }
            else if (e.Key.Equals(Key.ImeProcessed))
            {
                e.Handled = true;
            }
        }

        private void TextBoxFocusInDataGrid(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINControlFocus(sender, e);
        }

        private void TextBoxFocusInDataGrid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINBothByMouseUP(sender, e);
        }

        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;
                cell.IsEditing = true;
            }
        }

        private void DataGridSubCell_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down || e.Key == Key.Up)
            {
                DataGridCell_KeyDown(sender, e);
            }
        }

        private void dgdtpetxtDefectGoal_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                var DefectGoal = dgdSub1.CurrentItem as Win_Qul_QulDefectGoal_U_Sub_CodeView;

                if (DefectGoal != null)
                {
                    TextBox tb1 = sender as TextBox;

                    if (tb1 != null)
                    {
                        DefectGoal.DefectGoal = tb1.Text;
                        strRemain = DefectGoal.DefectGoal;
                    }

                    sender = tb1;
                }
            }
        }

        private void dgdtpetxtDefectGoal_KeyDown(object sender, KeyEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                if (e.Key == Key.Enter)
                {
                    var DefectGoal = dgdSub1.CurrentItem as Win_Qul_QulDefectGoal_U_Sub_CodeView;
                    if (DefectGoal != null)
                    {
                        TextBox tb1 = sender as TextBox;

                        if (tb1 != null)
                        {
                            DefectGoal.DefectGoal = tb1.Text;
                            strRemain = DefectGoal.DefectGoal;
                        }

                        sender = tb1;
                    }
                }
                else if (e.Key.Equals(Key.ImeProcessed))
                {
                    e.Handled = true;
                }
            }
        }

        private void DgdMain_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
                btnUpdate_Click(btnUpdate, null);
        }

        private void dgdtpetxtDefectGoal_Preview(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric((TextBox)sender, e);
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
}
