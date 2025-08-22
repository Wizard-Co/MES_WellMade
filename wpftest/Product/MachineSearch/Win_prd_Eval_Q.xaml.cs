using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    /// Win_com_MCEvalCal_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_Eval_Q : UserControl
    {
        string strFlag = string.Empty;
        int rowNum = 0;
        Lib lib = new Lib();
        Win_com_MCEvalCal_U_CodeView McEvalCal = new Win_com_MCEvalCal_U_CodeView();

        ObservableCollection<Win_com_MCEvalCal_U_CodeView> ovcMC = new ObservableCollection<Win_com_MCEvalCal_U_CodeView>();

        /// <summary>
        /// 고장시 타공저에 미치는 영향
        /// </summary>
        ObservableCollection<CodeView> ovcOne = new ObservableCollection<CodeView>();
        /// <summary>
        /// 대체 설비 유.무
        /// </summary>
        ObservableCollection<CodeView> ovcTwo = new ObservableCollection<CodeView>();
        /// <summary>
        /// 설비에 의해 공정이 품질에 미치는 영향
        /// </summary>
        ObservableCollection<CodeView> ovcThree = new ObservableCollection<CodeView>();
        /// <summary>
        /// 손실발생 빈도
        /// </summary>
        ObservableCollection<CodeView> ovcFour = new ObservableCollection<CodeView>();
        /// <summary>
        /// 고장으로 인한 인체 및 환경에 미치는 영향
        /// </summary>
        ObservableCollection<CodeView> ovcFive = new ObservableCollection<CodeView>();

        DataTable thisDT = new DataTable();
        DataTable thisDTClone = new DataTable();

        int MaxScore1 = 0;
        int MaxScore2 = 0;
        int MaxScore3 = 0;
        int MaxScore4 = 0;
        int MaxScore5 = 0;
        int MaxScore6 = 0;
        int MaxScore7 = 0;
        int MaxScore8 = 0;

        int MinScore1 = 0;
        int MinScore2 = 0;
        int MinScore3 = 0;
        int MinScore4 = 0;
        int MinScore5 = 0;
        int MinScore6 = 0;
        int MinScore7 = 0;
        int MinScore8 = 0;

        public Win_prd_Eval_Q()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            lib.UiLoading(sender);
            chkMcEvalDaySrh.IsChecked = true;
            dtpSDate.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            dtpEDate.SelectedDate = lib.BringThisMonthDatetimeList()[1];
            txtPerson.Text = MainWindow.CurrentPerson;
            txtPerson.Tag = MainWindow.CurrentPersonID;
            FillDataTable();

            ovcOne = SetComboBox("생산성", "고장시 타공정에 미치는 영향");
            ovcTwo = SetComboBox("생산성", "대체 설비 유.무");
            ovcThree = SetComboBox("품질", "설비에 의해 공정이 품질에 미치는 영향");
            ovcFour = SetComboBox("품질", "손실발생 빈도");
            ovcFive = SetComboBox("안전성", "고장으로 인한 인체 및 환경에 미치는 영향");
        }

        private void FillDataTable()
        {
            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                ds = DataStore.Instance.ProcedureToDataSet("xp_mc_sMcEvalBasis", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        thisDT = dt;
                        thisDTClone = thisDT.Clone();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //일단 만들어보자
        private ObservableCollection<CodeView> SetComboBox(string strGroupName, string strMCEvalName)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();

            try
            {
                string Sql = "  select  EvalSpecMin , MCEvalSpec  from mt_MCEvalBasis ";
                Sql += " where 1=1 ";
                Sql += " and MCGroupName=  '" + strGroupName + "' ";
                Sql += " and MCEvalName=  '" + strMCEvalName + "' ";

                DataSet ds = DataStore.Instance.QueryToDataSet(Sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0) { }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {

                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item["EvalSpecMin"].ToString().Trim(),
                                code_name = item["MCEvalSpec"].ToString().Trim(),
                            };

                            mCodeView.code_id = lib.returnNumStringZero(mCodeView.code_id);
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("에러 발생, 에러 내용 : " + ex.ToString());
            }

            return retunCollection;
        }

        //설비평가일 라벨 클릭시
        private void lblMcEvalDaySrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMcEvalDaySrh.IsChecked == true) { chkMcEvalDaySrh.IsChecked = false; }
            else { chkMcEvalDaySrh.IsChecked = true; }
        }

        //설비평가일 라벨 in 체크박스 체크시
        private void chkMcEvalDaySrh_Checked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = true;
            dtpEDate.IsEnabled = true;
            btnLastMonth.IsEnabled = true;
            btnThisMonth.IsEnabled = true;
            btnYesterDay.IsEnabled = true;
            btnToday.IsEnabled = true;
        }

        //설비평가일 라벨 in 체크박스 언체크시
        private void chkMcEvalDaySrh_Unchecked(object sender, RoutedEventArgs e)
        {
            dtpSDate.IsEnabled = false;
            dtpEDate.IsEnabled = false;
            btnLastMonth.IsEnabled = false;
            btnThisMonth.IsEnabled = false;
            btnYesterDay.IsEnabled = false;
            btnToday.IsEnabled = false;
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


        //설비명 라벨 클릭시
        private void lblMcEvalNameSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkMcEvalNameSrh.IsChecked == true) { chkMcEvalNameSrh.IsChecked = false; }
            else { chkMcEvalNameSrh.IsChecked = true; }
        }

        //설비명 라벨 in 체크박스 체크시
        private void chkMcEvalNameSrh_Checked(object sender, RoutedEventArgs e)
        {
            txtMcEvalNameSrh.IsEnabled = true;
        }

        //설비명 라벨 in 체크박스 언체크시
        private void chkMcEvalNameSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            txtMcEvalNameSrh.IsEnabled = false;
        }

        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count == 0)
            {
                MessageBox.Show("수정할 자료가 없습니다.");
            }
            else
            {
                strFlag = "U";
                btnNewMCEvalCal.IsEnabled = false;
                btnUpdate.IsEnabled = false;
                btnMCEvalCal.IsEnabled = true;
                btnRowAdd.IsEnabled = true;
                btnRowDel.IsEnabled = true;
                chkAllCheck.IsEnabled = false;
                chkAllCheck.IsChecked = false;
            }
        }

        //뉴평가
        private void btnNewMCEvalCal_Click(object sender, RoutedEventArgs e)
        {
            strFlag = "I";
            btnNewMCEvalCal.IsEnabled = false;
            btnUpdate.IsEnabled = false;
            btnMCEvalCal.IsEnabled = true;
            btnRowAdd.IsEnabled = true;
            btnRowDel.IsEnabled = true;
            chkAllCheck.IsEnabled = false;
            chkAllCheck.IsChecked = false;
            FillData();
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            bool delFlag = true;
            List<Win_com_MCEvalCal_U_CodeView> lstMCEvalCal = new List<Win_com_MCEvalCal_U_CodeView>();

            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                McEvalCal = dgdMain.Items[i] as Win_com_MCEvalCal_U_CodeView;

                if (McEvalCal.chkData)
                {
                    lstMCEvalCal.Add(McEvalCal);
                }
            }

            if (lstMCEvalCal.Count <= 0)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdMain.Items.Count > 0 && dgdMain.SelectedItem != null)
                    {
                        rowNum = dgdMain.SelectedIndex;
                    }

                    if (!DeleteData())
                    {
                        MessageBox.Show("삭제를 실패하여 중단됩니다.");
                        delFlag = false;
                    }

                    if (delFlag)
                    {
                        lstMCEvalCal.Clear();
                        rowNum = 0;
                        re_Search(rowNum);
                    }
                }
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {

            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                strFlag = string.Empty;
                btnNewMCEvalCal.IsEnabled = true;
                btnUpdate.IsEnabled = true;
                btnMCEvalCal.IsEnabled = false;
                btnRowAdd.IsEnabled = false;
                btnRowDel.IsEnabled = false;
                chkAllCheck.IsEnabled = true;
                chkAllCheck.IsChecked = false;


                rowNum = 0;
                re_Search(rowNum);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
            

        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "설비등급평가처리";
            lst[1] = dgdMain.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdMain.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdMain);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdMain);

                    Name = dgdMain.Name;
                    Lib.Instance.GenerateExcel(dt, Name);
                    Lib.Instance.excel.Visible = true;
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

        //등급 평가처리
        private void btnMCEvalCal_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData())
            {
                FillGrid();
                btnMCEvalCal.IsEnabled = false;
                btnUpdate.IsEnabled = true;
                btnNewMCEvalCal.IsEnabled = true;
                strFlag = string.Empty;
                btnRowAdd.IsEnabled = false;
                btnRowDel.IsEnabled = false;
                chkAllCheck.IsEnabled = true;
                chkAllCheck.IsChecked = false;
            }
        }

        //설비 등급조회
        private void btnMCEvalBasis_Click(object sender, RoutedEventArgs e)
        {
            int k = 0;
            foreach (MenuViewModel mvm in MainWindow.mMenulist)
            {
                if (mvm.Menu.Equals("★설비등급 평가기준 등록"))
                {
                    break;
                }
                k++;
            }

            if (MainWindow.MainMdiContainer.Children.Contains(MainWindow.mMenulist[k].subProgramID as MdiChild))
            {
                (MainWindow.mMenulist[k].subProgramID as MdiChild).Focus();
            }
            else
            {
                Type type = Type.GetType("WizMes_WellMade." + MainWindow.mMenulist[k].ProgramID.Trim(), true);
                object uie = Activator.CreateInstance(type);

                MainWindow.mMenulist[k].subProgramID = new MdiChild()
                {
                    Title = "(주)WellMade [" + MainWindow.mMenulist[k].MenuID.Trim() + "] " + MainWindow.mMenulist[k].Menu.Trim() + 
                            " (→" + MainWindow.mMenulist[k].ProgramID + ")",
                    Height = SystemParameters.PrimaryScreenHeight * 0.8,
                    MaxHeight = SystemParameters.PrimaryScreenHeight * 0.85,
                    Width = SystemParameters.WorkArea.Width * 0.85,
                    MaxWidth = SystemParameters.WorkArea.Width,
                    Content = uie as UIElement,
                    Tag = MainWindow.mMenulist[k]
                };
                lib.AllMenuLogInsert(MainWindow.mMenulist[k].MenuID, MainWindow.mMenulist[k].Menu, MainWindow.mMenulist[k].subProgramID);
                MainWindow.MainMdiContainer.Children.Add(MainWindow.mMenulist[k].subProgramID as MdiChild);
            }
        }

        /// <summary>
        /// 재검색(수정,삭제,추가 저장후에 자동 재검색)
        /// </summary>
        /// <param name="selectedIndex"></param>
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdMain.Items.Count == 0)
            {
                this.DataContext = null;
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        /// <summary>
        /// 실조회
        /// </summary>
        private void FillGrid()
        {
            try
            {
                ovcMC.Clear();

                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("chkMCID", chkMcEvalNameSrh.IsChecked==true ? 1:0);
                sqlParameter.Add("sMCID", chkMcEvalNameSrh.IsChecked == true ? txtMcEvalNameSrh.Text:"");
                sqlParameter.Add("chkDate", chkMcEvalDaySrh.IsChecked==true? 1:0);
                sqlParameter.Add("sStartDate", chkMcEvalDaySrh.IsChecked == true ? 
                    (dtpSDate.SelectedDate !=null ? dtpSDate.SelectedDate.Value.ToString("yyyyMMdd"): "" ): "" );
                sqlParameter.Add("sEndDate", chkMcEvalDaySrh.IsChecked == true ?
                    (dtpEDate.SelectedDate != null ? dtpEDate.SelectedDate.Value.ToString("yyyyMMdd") : "") : "");

                ds = DataStore.Instance.ProcedureToDataSet("xp_mc_sMCEval", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinMCEval = new Win_com_MCEvalCal_U_CodeView()
                            {
                                Num = i + 1,
                                MCEvalID = dr["MCEvalID"].ToString(),
                                MCID = dr["MCID"].ToString(),
                                EvalDate = dr["EvalDate"].ToString(),
                                MCNAME = dr["MCNAME"].ToString(),
                                MCoperationRate = dr["MCoperationRate"].ToString(),
                                OtherProceeEffectBroken = dr["OtherProceeEffectBroken"].ToString(),
                                ReplacementMC = dr["ReplacementMC"].ToString(),
                                MCProcessQualtityEffect = dr["MCProcessQualtityEffect"].ToString(),
                                FrequencyOfLoss = dr["FrequencyOfLoss"].ToString(),
                                FrequencyOfFailure = dr["FrequencyOfFailure"].ToString(),
                                FaultStopTime = dr["FaultStopTime"].ToString(),
                                HumanEnvironImpact = dr["HumanEnvironImpact"].ToString(),

                                MCoperationRateScore = dr["MCoperationRateScore"].ToString(),
                                OtherProceeEffectBrokenScore = dr["OtherProceeEffectBrokenScore"].ToString(),
                                ReplacementMCScore = dr["ReplacementMCScore"].ToString(),
                                MCProcessQualtityEffectScore = dr["MCProcessQualtityEffectScore"].ToString(),
                                FrequencyOfLossScore = dr["FrequencyOfLossScore"].ToString(),
                                FrequencyOfFailureScore = dr["FrequencyOfFailureScore"].ToString(),
                                FaultStopTimeScore = dr["FaultStopTimeScore"].ToString(),
                                HumanEnvironImpactScore = dr["HumanEnvironImpactScore"].ToString(),

                                Score = dr["Score"].ToString(),
                                EvalGrade = dr["EvalGrade"].ToString(),
                                EvalPersonName = dr["EvalPersonName"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                CreateDate = dr["CreateDate"].ToString(),

                                OvcOtherProceeEffectBroken = ovcOne,
                                OvcReplacementMC = ovcTwo,
                                OvcMCProcessQualtityEffect = ovcThree,
                                OvcFrequencyOfLoss = ovcFour,
                                OvcHumanEnvironImpact = ovcFive,

                                chkData = false,
                                OvcMachine = ComboBoxUtil.Instance.Get_MCID()
                            };

                            WinMCEval.MCoperationRate = Lib.Instance.returnNumStringZero(WinMCEval.MCoperationRate);
                            WinMCEval.OtherProceeEffectBroken = Lib.Instance.returnNumStringZero(WinMCEval.OtherProceeEffectBroken);
                            WinMCEval.ReplacementMC = Lib.Instance.returnNumStringZero(WinMCEval.ReplacementMC);
                            WinMCEval.MCProcessQualtityEffect = Lib.Instance.returnNumStringZero(WinMCEval.MCProcessQualtityEffect);
                            WinMCEval.FrequencyOfLoss = Lib.Instance.returnNumStringZero(WinMCEval.FrequencyOfLoss);
                            WinMCEval.FrequencyOfFailure = Lib.Instance.returnNumStringZero(WinMCEval.FrequencyOfFailure);
                            WinMCEval.FaultStopTime = Lib.Instance.returnNumStringZero(WinMCEval.FaultStopTime);
                            WinMCEval.HumanEnvironImpact = Lib.Instance.returnNumStringZero(WinMCEval.HumanEnvironImpact);

                            WinMCEval.MCoperationRateScore = Lib.Instance.returnNumStringZero(WinMCEval.MCoperationRateScore);
                            WinMCEval.OtherProceeEffectBrokenScore = Lib.Instance.returnNumStringZero(WinMCEval.OtherProceeEffectBrokenScore);
                            WinMCEval.ReplacementMCScore = Lib.Instance.returnNumStringZero(WinMCEval.ReplacementMCScore);
                            WinMCEval.MCProcessQualtityEffectScore = Lib.Instance.returnNumStringZero(WinMCEval.MCProcessQualtityEffectScore);
                            WinMCEval.FrequencyOfLossScore = Lib.Instance.returnNumStringZero(WinMCEval.FrequencyOfLossScore);
                            WinMCEval.FrequencyOfFailureScore = Lib.Instance.returnNumStringZero(WinMCEval.FrequencyOfFailureScore);
                            WinMCEval.FaultStopTimeScore = Lib.Instance.returnNumStringZero(WinMCEval.FaultStopTimeScore);
                            WinMCEval.HumanEnvironImpactScore = Lib.Instance.returnNumStringZero(WinMCEval.HumanEnvironImpactScore);

                            foreach (CodeView cv in WinMCEval.OvcOtherProceeEffectBroken)
                            {
                                if (cv.code_id.Equals(WinMCEval.OtherProceeEffectBroken))
                                    WinMCEval.OtherProceeEffectBroken_CB = cv.code_name;
                            }

                            foreach (CodeView cv in WinMCEval.OvcReplacementMC)
                            {
                                if (cv.code_id.Equals(WinMCEval.ReplacementMC))
                                    WinMCEval.ReplacementMC_CB = cv.code_name;
                            }

                            foreach (CodeView cv in WinMCEval.OvcMCProcessQualtityEffect)
                            {
                                if (cv.code_id.Equals(WinMCEval.MCProcessQualtityEffect))
                                    WinMCEval.MCProcessQualtityEffect_CB = cv.code_name;
                            }

                            foreach (CodeView cv in WinMCEval.OvcFrequencyOfLoss)
                            {
                                if (cv.code_id.Equals(WinMCEval.FrequencyOfLoss))
                                    WinMCEval.FrequencyOfLoss_CB = cv.code_name;
                            }

                            foreach (CodeView cv in WinMCEval.OvcHumanEnvironImpact)
                            {
                                if (cv.code_id.Equals(WinMCEval.HumanEnvironImpact))
                                    WinMCEval.HumanEnvironImpact_CB = cv.code_name;
                            }

                            WinMCEval.Score = Lib.Instance.returnNumStringZero(WinMCEval.Score);
                            //dgdMain.Items.Add(WinMCEval);
                            ovcMC.Add(WinMCEval);
                            i++;
                        }

                        tbkCount.Text = "▶ 검색결과 : " + i.ToString() + " 건";
                        
                        dgdMain.ItemsSource = ovcMC;
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

        /// <summary>
        /// 실삭제
        /// </summary>
        /// <returns></returns>
        private bool DeleteData()
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    McEvalCal = dgdMain.Items[i] as Win_com_MCEvalCal_U_CodeView;

                    if (McEvalCal.chkData)
                    {
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("MCEvalID", McEvalCal.MCEvalID);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_mc_dMCEval";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "MoldEvalID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);
                    }
                }

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    flag = false;
                    //return false;
                }
                else
                {
                    flag = true;
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

        private bool SaveData()
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
            bool DataCheck = true;

            try
            {
                if (CheckData())
                {
                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        for (int i = 0; i < dgdMain.Items.Count; i++)
                        {
                            var WinMCEval = dgdMain.Items[i] as Win_com_MCEvalCal_U_CodeView;
                            if (CheckDataCodeView(WinMCEval, i + 1))
                            {
                                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();
                                sqlParameter.Add("MCEvalID", "");
                                sqlParameter.Add("MCID", WinMCEval.MCID);
                                sqlParameter.Add("EvalDate", dtpEvalCal.SelectedDate.Value.ToString("yyyyMMdd"));
                                sqlParameter.Add("MCoperationRate", WinMCEval.MCoperationRate);
                                sqlParameter.Add("OtherProceeEffectBroken", WinMCEval.OtherProceeEffectBroken);
                                sqlParameter.Add("ReplacementMC", WinMCEval.ReplacementMC);
                                sqlParameter.Add("MCProcessQualtityEffect", WinMCEval.MCProcessQualtityEffect);
                                sqlParameter.Add("FrequencyOfLoss", WinMCEval.FrequencyOfLoss);
                                sqlParameter.Add("FrequencyOfFailure", WinMCEval.FrequencyOfFailure.Replace(",", ""));
                                sqlParameter.Add("FaultStopTime", WinMCEval.FaultStopTime.Replace(",", ""));
                                sqlParameter.Add("HumanEnvironImpact", WinMCEval.HumanEnvironImpact);

                                sqlParameter.Add("MCoperationRateScore", WinMCEval.MCoperationRateScore);
                                sqlParameter.Add("OtherProceeEffectBrokenScore", WinMCEval.OtherProceeEffectBrokenScore);
                                sqlParameter.Add("ReplacementMCScore", WinMCEval.ReplacementMCScore);
                                sqlParameter.Add("MCProcessQualtityEffectScore", WinMCEval.MCProcessQualtityEffectScore);
                                sqlParameter.Add("FrequencyOfLossScore", WinMCEval.FrequencyOfLossScore);
                                sqlParameter.Add("FrequencyOfFailureScore", WinMCEval.FrequencyOfFailureScore.Replace(",", ""));
                                sqlParameter.Add("FaultStopTimeScore", WinMCEval.FaultStopTimeScore.Replace(",", ""));
                                sqlParameter.Add("HumanEnvironImpactScore", WinMCEval.HumanEnvironImpactScore);

                                sqlParameter.Add("EvalPersonName", WinMCEval.EvalPersonName);
                                sqlParameter.Add("sComments", WinMCEval.Comments != null ? WinMCEval.Comments : "");
                                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                                Procedure pro1 = new Procedure();
                                pro1.Name = "xp_MC_iMCEval";
                                pro1.OutputUseYN = "N";
                                pro1.OutputName = "MCID";
                                pro1.OutputLength = "10";

                                Prolist.Add(pro1);
                                ListParameter.Add(sqlParameter);
                            }
                            else
                            {
                                DataCheck = false;
                                break;
                            }
                        }

                        if (DataCheck)
                        {
                            string[] Confirm = new string[2];
                            Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                            if (Confirm[0] != "success")
                            {
                                MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                                flag = false;
                                //return false;
                            }
                            else
                            {
                                flag = true;
                            }
                        }
                        else
                        {
                            return false;
                        }
                    }

                    #endregion

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        for (int i = 0; i < dgdMain.Items.Count; i++)
                        {
                            var WinMCEval = dgdMain.Items[i] as Win_com_MCEvalCal_U_CodeView;
                            if (CheckDataCodeView(WinMCEval, i + 1))
                            {
                                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();
                                sqlParameter.Add("MCEvalID", WinMCEval.MCEvalID);
                                sqlParameter.Add("MCID", WinMCEval.MCID);
                                sqlParameter.Add("EvalDate", dtpEvalCal.SelectedDate.Value.ToString("yyyyMMdd"));
                                sqlParameter.Add("MCoperationRate", WinMCEval.MCoperationRate);
                                sqlParameter.Add("OtherProceeEffectBroken", WinMCEval.OtherProceeEffectBroken);
                                sqlParameter.Add("ReplacementMC", WinMCEval.ReplacementMC);
                                sqlParameter.Add("MCProcessQualtityEffect", WinMCEval.MCProcessQualtityEffect);
                                sqlParameter.Add("FrequencyOfLoss", WinMCEval.FrequencyOfLoss);
                                sqlParameter.Add("FrequencyOfFailure", WinMCEval.FrequencyOfFailure.Replace(",", ""));
                                sqlParameter.Add("FaultStopTime", WinMCEval.FaultStopTime.Replace(",", ""));
                                sqlParameter.Add("HumanEnvironImpact", WinMCEval.HumanEnvironImpact);

                                sqlParameter.Add("MCoperationRateScore", WinMCEval.MCoperationRateScore);
                                sqlParameter.Add("OtherProceeEffectBrokenScore", WinMCEval.OtherProceeEffectBrokenScore);
                                sqlParameter.Add("ReplacementMCScore", WinMCEval.ReplacementMCScore);
                                sqlParameter.Add("MCProcessQualtityEffectScore", WinMCEval.MCProcessQualtityEffectScore);
                                sqlParameter.Add("FrequencyOfLossScore", WinMCEval.FrequencyOfLossScore);
                                sqlParameter.Add("FrequencyOfFailureScore", WinMCEval.FrequencyOfFailureScore.Replace(",", ""));
                                sqlParameter.Add("FaultStopTimeScore", WinMCEval.FaultStopTimeScore.Replace(",", ""));
                                sqlParameter.Add("HumanEnvironImpactScore", WinMCEval.HumanEnvironImpactScore);

                                sqlParameter.Add("EvalPersonName", WinMCEval.EvalPersonName);
                                sqlParameter.Add("sComments", WinMCEval.Comments != null ? WinMCEval.Comments : "");
                                sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                                Procedure pro1 = new Procedure();
                                pro1.Name = "xp_MC_uMCEval";
                                pro1.OutputUseYN = "N";
                                pro1.OutputName = "MCID";
                                pro1.OutputLength = "10";

                                Prolist.Add(pro1);
                                ListParameter.Add(sqlParameter);
                            }
                            else
                            {
                                DataCheck = false;
                                break;
                            }
                        }

                        if (DataCheck)
                        {
                            string[] Confirm = new string[2];
                            Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                            if (Confirm[0] != "success")
                            {
                                MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                                flag = false;
                                //return false;
                            }
                            else
                            {
                                flag = true;
                            }
                        }
                        else
                        {
                            return false;
                        }
                    }

                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생,오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }

        private bool CheckData()
        {
            bool flag = true;

            if (dtpEvalCal.SelectedDate == null)
            {
                MessageBox.Show("평가일자가 선택되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (txtPerson.Text.Length <= 0 || txtPerson.Text.Equals(""))
            {
                MessageBox.Show("평가자가 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            return flag;
        }

        private bool CheckDataCodeView(Win_com_MCEvalCal_U_CodeView MCEvalCodeView,int Index)
        {
            bool flag = true;

            if (MCEvalCodeView.MCoperationRateScore.Equals(string.Empty))
            {
                MessageBox.Show("설비가동현황평가가 입력되지 않았습니다. " + Index  + "번째 줄");
                flag = false;
                return flag;
            }

            if (MCEvalCodeView.OtherProceeEffectBrokenScore.Equals(string.Empty))
            {
                MessageBox.Show("고장시 타공정에 미치는 영향 평가가 입력되지 않았습니다. " + Index  + "번째 줄");
                flag = false;
                return flag;
            }

            if (MCEvalCodeView.ReplacementMCScore.Equals(string.Empty))
            {
                MessageBox.Show("대체 설비 유.무 평가가 입력되지 않았습니다. " + Index  + "번째 줄");
                flag = false;
                return flag;
            }

            if (MCEvalCodeView.MCProcessQualtityEffectScore.Equals(string.Empty))
            {
                MessageBox.Show("설비에 의해 공정이 품질에 미치는 영향 평가가 입력되지 않았습니다. " + Index  + "번째 줄");
                flag = false;
                return flag;
            }

            if (MCEvalCodeView.FrequencyOfLossScore.Equals(string.Empty))
            {
                MessageBox.Show("손실발생 빈도 평가가 입력되지 않았습니다. " + Index  + "번째 줄");
                flag = false;
                return flag;
            }

            if (MCEvalCodeView.FrequencyOfFailureScore.Equals(string.Empty))
            {
                MessageBox.Show("고장빈도 여부 평가가 입력되지 않았습니다. " + Index  + "번째 줄");
                flag = false;
                return flag;
            }

            if (MCEvalCodeView.FaultStopTimeScore.Equals(string.Empty))
            {
                MessageBox.Show("고장 정지 시간 평가가 입력되지 않았습니다. " + Index  + "번째 줄");
                flag = false;
                return flag;
            }

            if (MCEvalCodeView.HumanEnvironImpactScore.Equals(string.Empty))
            {
                MessageBox.Show("고장으로 인한 인체 및 환경에 미치는 영향 평가가 입력되지 않았습니다. " + Index  + "번째 줄");
                flag = false;
                return flag;
            }

            return flag;
        }
        
        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            McEvalCal = dgdMain.SelectedItem as Win_com_MCEvalCal_U_CodeView;

            if (McEvalCal != null)
            {
                if(McEvalCal.chkData)
                    McEvalCal.chkData = false;
                else
                    McEvalCal.chkData = true;
            }
        }

        //행추가
        private void BtnRowAdd_Click(object sender, RoutedEventArgs e)
        {
            var WinMCEval = new Win_com_MCEvalCal_U_CodeView()
            {
                Num = dgdMain.Items.Count + 1,
                MCID = "0001",
                MCNAME = "포밍01 DS-06",
                EvalDate = DateTime.Today.ToString("yyyyMMdd"),
                EvalPersonName = txtPerson.Text,

                MCoperationRate = "0",
                OtherProceeEffectBroken = "0",
                ReplacementMC = "0",
                MCProcessQualtityEffect = "0",
                FrequencyOfLoss = "0",
                FrequencyOfFailure = "0",
                FaultStopTime = "0",
                HumanEnvironImpact = "0",

                OvcOtherProceeEffectBroken = ovcOne,
                OvcReplacementMC = ovcTwo,
                OvcMCProcessQualtityEffect = ovcThree,
                OvcFrequencyOfLoss = ovcFour,
                OvcHumanEnvironImpact = ovcFive,
                OvcMachine = ComboBoxUtil.Instance.Get_MCID()
            };

            WinMCEval.MCoperationRateScore = Lib.Instance.CheckNullZero(GetScore("설비가동현황", WinMCEval.MCoperationRate, 1));
            WinMCEval.OtherProceeEffectBrokenScore =
                Lib.Instance.CheckNullZero(GetScore("고장시 타공정에 미치는 영향", WinMCEval.OtherProceeEffectBroken, 2));
            WinMCEval.ReplacementMCScore = Lib.Instance.CheckNullZero(GetScore("대체 설비 유.무", WinMCEval.ReplacementMC, 3));
            WinMCEval.MCProcessQualtityEffectScore =
                Lib.Instance.CheckNullZero(GetScore("설비에 의해 공정이 품질에 미치는 영향", WinMCEval.MCProcessQualtityEffect, 4));
            WinMCEval.FrequencyOfLossScore = Lib.Instance.CheckNullZero(GetScore("손실발생 빈도", WinMCEval.FrequencyOfLoss, 5));
            WinMCEval.FrequencyOfFailureScore = Lib.Instance.CheckNullZero(GetScore("고장빈도 여부", WinMCEval.FrequencyOfFailure, 6));
            WinMCEval.FaultStopTimeScore = Lib.Instance.CheckNullZero(GetScore("고장 정지 시간", WinMCEval.FaultStopTime, 7));
            WinMCEval.HumanEnvironImpactScore =
                Lib.Instance.CheckNullZero(GetScore("고장으로 인한 인체 및 환경에 미치는 영향", WinMCEval.HumanEnvironImpact, 8));

            WinMCEval.Score = TotalScore(WinMCEval.MCoperationRateScore, WinMCEval.OtherProceeEffectBrokenScore,
                WinMCEval.ReplacementMCScore, WinMCEval.MCProcessQualtityEffectScore, WinMCEval.FrequencyOfLossScore
                , WinMCEval.FrequencyOfFailureScore, WinMCEval.FaultStopTimeScore, WinMCEval.HumanEnvironImpactScore);
            WinMCEval.EvalGrade = Lib.Instance.ReturnGrade(WinMCEval.Score);

            foreach (CodeView cv in WinMCEval.OvcOtherProceeEffectBroken)
            {
                if (cv.code_id.Equals(WinMCEval.OtherProceeEffectBroken))
                    WinMCEval.OtherProceeEffectBroken_CB = cv.code_name;
            }

            foreach (CodeView cv in WinMCEval.OvcReplacementMC)
            {
                if (cv.code_id.Equals(WinMCEval.ReplacementMC))
                    WinMCEval.ReplacementMC_CB = cv.code_name;
            }

            foreach (CodeView cv in WinMCEval.OvcMCProcessQualtityEffect)
            {
                if (cv.code_id.Equals(WinMCEval.MCProcessQualtityEffect))
                    WinMCEval.MCProcessQualtityEffect_CB = cv.code_name;
            }

            foreach (CodeView cv in WinMCEval.OvcFrequencyOfLoss)
            {
                if (cv.code_id.Equals(WinMCEval.FrequencyOfLoss))
                    WinMCEval.FrequencyOfLoss_CB = cv.code_name;
            }

            foreach (CodeView cv in WinMCEval.OvcHumanEnvironImpact)
            {
                if (cv.code_id.Equals(WinMCEval.HumanEnvironImpact))
                    WinMCEval.HumanEnvironImpact_CB = cv.code_name;
            }

            ovcMC.Add(WinMCEval);
        }

        //행삭제
        private void BtnRowDel_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count > 0)
            {
                if (dgdMain.SelectedItem != null)
                {
                    ovcMC.Remove(dgdMain.SelectedItem as Win_com_MCEvalCal_U_CodeView);
                }
                else
                {
                    ovcMC.Remove(dgdMain.Items[dgdMain.Items.Count - 1] as Win_com_MCEvalCal_U_CodeView);
                }
            }
        }

        private void FillData()
        {
            try
            {
                ovcMC.Clear();

                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                ds = DataStore.Instance.ProcedureToDataSet("xp_mc_sNewMCEval", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        for (int k = 0; k < drc.Count; k++)
                        {
                            DataRow dr = drc[k];
                            var WinNewMC = new NewMCEval()
                            {
                                EvalDate = dr["EvalDate"].ToString(),
                                MCID = dr["MCID"].ToString(),
                                MCName = dr["MCName"].ToString(),

                                MCoperationRate = dr["MCoperationRate"].ToString(),
                                OtherProceeEffectBroken = dr["OtherProceeEffectBroken"].ToString(),
                                ReplacementMC = dr["ReplacementMC"].ToString(),
                                MCProcessQualtityEffect = dr["MCProcessQualtityEffect"].ToString(),
                                FrequencyOfLoss = dr["FrequencyOfLoss"].ToString(),
                                FrequencyOfFailure = dr["FrequencyOfFailure"].ToString(),
                                FaultStopTime = dr["FaultStopTime"].ToString(),
                                HumanEnvironImpact = dr["HumanEnvironImpact"].ToString()
                            };

                            var WinMCEval = new Win_com_MCEvalCal_U_CodeView()
                            {
                                Num = i + 1,
                                MCID = WinNewMC.MCID,
                                MCNAME = WinNewMC.MCName,
                                EvalDate = WinNewMC.EvalDate,
                                EvalPersonName = txtPerson.Text,

                                MCoperationRate = WinNewMC.MCoperationRate,
                                OtherProceeEffectBroken = WinNewMC.OtherProceeEffectBroken,
                                ReplacementMC = WinNewMC.ReplacementMC,
                                MCProcessQualtityEffect = WinNewMC.MCProcessQualtityEffect,
                                FrequencyOfLoss = WinNewMC.FrequencyOfLoss,
                                FrequencyOfFailure = WinNewMC.FrequencyOfFailure,
                                FaultStopTime = WinNewMC.FaultStopTime,
                                HumanEnvironImpact = WinNewMC.HumanEnvironImpact,

                                OvcOtherProceeEffectBroken = ovcOne,
                                OvcReplacementMC = ovcTwo,
                                OvcMCProcessQualtityEffect = ovcThree,
                                OvcFrequencyOfLoss = ovcFour,
                                OvcHumanEnvironImpact = ovcFive,
                                OvcMachine = ComboBoxUtil.Instance.Get_MCID()
                            };

                            WinMCEval.MCoperationRateScore = Lib.Instance.CheckNullZero(GetScore("설비가동현황", WinMCEval.MCoperationRate, 1));
                            WinMCEval.OtherProceeEffectBrokenScore =
                                Lib.Instance.CheckNullZero(GetScore("고장시 타공정에 미치는 영향", WinMCEval.OtherProceeEffectBroken, 2));
                            WinMCEval.ReplacementMCScore = Lib.Instance.CheckNullZero(GetScore("대체 설비 유.무", WinMCEval.ReplacementMC, 3));
                            WinMCEval.MCProcessQualtityEffectScore =
                                Lib.Instance.CheckNullZero(GetScore("설비에 의해 공정이 품질에 미치는 영향", WinMCEval.MCProcessQualtityEffect, 4));
                            WinMCEval.FrequencyOfLossScore = Lib.Instance.CheckNullZero(GetScore("손실발생 빈도", WinMCEval.FrequencyOfLoss, 5));
                            WinMCEval.FrequencyOfFailureScore = Lib.Instance.CheckNullZero(GetScore("고장빈도 여부", WinMCEval.FrequencyOfFailure, 6));
                            WinMCEval.FaultStopTimeScore = Lib.Instance.CheckNullZero(GetScore("고장 정지 시간", WinMCEval.FaultStopTime, 7));
                            WinMCEval.HumanEnvironImpactScore =
                                Lib.Instance.CheckNullZero(GetScore("고장으로 인한 인체 및 환경에 미치는 영향", WinMCEval.HumanEnvironImpact, 8));

                            WinMCEval.Score = TotalScore(WinMCEval.MCoperationRateScore, WinMCEval.OtherProceeEffectBrokenScore,
                                WinMCEval.ReplacementMCScore, WinMCEval.MCProcessQualtityEffectScore, WinMCEval.FrequencyOfLossScore
                                , WinMCEval.FrequencyOfFailureScore, WinMCEval.FaultStopTimeScore, WinMCEval.HumanEnvironImpactScore);
                            WinMCEval.EvalGrade = Lib.Instance.ReturnGrade(WinMCEval.Score);

                            foreach (CodeView cv in WinMCEval.OvcOtherProceeEffectBroken)
                            {
                                if (cv.code_id.Equals(WinMCEval.OtherProceeEffectBroken))
                                    WinMCEval.OtherProceeEffectBroken_CB = cv.code_name;
                            }

                            foreach (CodeView cv in WinMCEval.OvcReplacementMC)
                            {
                                if (cv.code_id.Equals(WinMCEval.ReplacementMC))
                                    WinMCEval.ReplacementMC_CB = cv.code_name;
                            }

                            foreach (CodeView cv in WinMCEval.OvcMCProcessQualtityEffect)
                            {
                                if (cv.code_id.Equals(WinMCEval.MCProcessQualtityEffect))
                                    WinMCEval.MCProcessQualtityEffect_CB = cv.code_name;
                            }

                            foreach (CodeView cv in WinMCEval.OvcFrequencyOfLoss)
                            {
                                if (cv.code_id.Equals(WinMCEval.FrequencyOfLoss))
                                    WinMCEval.FrequencyOfLoss_CB = cv.code_name;
                            }

                            foreach (CodeView cv in WinMCEval.OvcHumanEnvironImpact)
                            {
                                if (cv.code_id.Equals(WinMCEval.HumanEnvironImpact))
                                    WinMCEval.HumanEnvironImpact_CB = cv.code_name;
                            }

                            ovcMC.Add(WinMCEval);
                            //dgdMain.Items.Add(WinMCEval);

                            i++;
                        }

                        dgdMain.ItemsSource = ovcMC;
                    }

                    int colCount = dgdMain.Columns.IndexOf(dgdtpeMCoperationRate);
                    dgdMain.Focus();
                    dgdMain.CurrentCell = new DataGridCellInfo
                        (dgdMain.Items[0], dgdMain.Columns[colCount]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생,오류 내용 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        private string TotalScore(string MCoperationRateScore, string OtherProceeEffectBrokenScore, string ReplacementMCScore
            , string MCProcessQualtityEffectScore, string FrequencyOfLossScore, string FrequencyOfFailureScore
            , string FaultStopTimeScore, string HumanEnvironImpactScore)
        {
            string strTotalScore = string.Empty;

            if (MCoperationRateScore.Equals(string.Empty))
                MCoperationRateScore = "0";

            if (OtherProceeEffectBrokenScore.Equals(string.Empty))
                OtherProceeEffectBrokenScore = "0";

            if (ReplacementMCScore.Equals(string.Empty))
                ReplacementMCScore = "0";

            if (MCProcessQualtityEffectScore.Equals(string.Empty))
                MCProcessQualtityEffectScore = "0";

            if (FrequencyOfLossScore.Equals(string.Empty))
                FrequencyOfLossScore = "0";

            if (FrequencyOfFailureScore.Equals(string.Empty))
                FrequencyOfFailureScore = "0";

            if (FaultStopTimeScore.Equals(string.Empty))
                FaultStopTimeScore = "0";

            if (HumanEnvironImpactScore.Equals(string.Empty))
                HumanEnvironImpactScore = "0";


            strTotalScore = (int.Parse(MCoperationRateScore) + int.Parse(OtherProceeEffectBrokenScore)
                                + int.Parse(ReplacementMCScore) + int.Parse(MCProcessQualtityEffectScore)
                                + int.Parse(FrequencyOfLossScore) + int.Parse(FrequencyOfFailureScore) 
                                + int.Parse(FaultStopTimeScore) + int.Parse(HumanEnvironImpactScore)).ToString();

            return strTotalScore;
        }

        //
        private void DataGridMainCell_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Left || e.Key == Key.Right)
                {
                    DataGridMainCell_KeyDown(sender, e);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //
        private void DataGridMainCell_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                McEvalCal = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;
                int rowCount = dgdMain.Items.IndexOf(dgdMain.CurrentItem);
                int colCount = dgdMain.Columns.IndexOf(dgdMain.CurrentCell.Column);
                int lastColcount = dgdMain.Columns.IndexOf(dgdtpeComments);
                int startColcount = dgdMain.Columns.IndexOf(dgdtpeMCoperationRateScore);
                
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (lastColcount == colCount && dgdMain.Items.Count - 1 > rowCount)
                    {
                        dgdMain.SelectedIndex = rowCount + 1;
                        dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount + 1], dgdMain.Columns[startColcount]);
                    }
                    else if (lastColcount > colCount && dgdMain.Items.Count - 1 > rowCount)
                    {
                        dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount], dgdMain.Columns[colCount + 1]);
                    }
                    else if (lastColcount == colCount && dgdMain.Items.Count - 1 == rowCount)
                    {
                        btnMCEvalCal.Focus();
                    }
                    else if (lastColcount > colCount && dgdMain.Items.Count - 1 == rowCount)
                    {
                        dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount], dgdMain.Columns[colCount + 1]);
                    }
                    else
                    {
                        MessageBox.Show("있으면 찾아보자...");
                    }
                }
                else if (e.Key == Key.Down)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (dgdMain.Items.Count - 1 > rowCount)
                    {
                        dgdMain.SelectedIndex = rowCount + 1;
                        dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount + 1], dgdMain.Columns[colCount]);
                    }
                    else if (dgdMain.Items.Count - 1 == rowCount)
                    {
                        if (lastColcount > colCount)
                        {
                            dgdMain.SelectedIndex = 0;
                            dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[0], dgdMain.Columns[colCount + 1]);
                        }
                        else
                        {
                            btnMCEvalCal.Focus();
                        }
                    }
                }
                else if (e.Key == Key.Up)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (rowCount > 0)
                    {
                        dgdMain.SelectedIndex = rowCount - 1;
                        dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount - 1], dgdMain.Columns[colCount]);
                    }
                }
                else if (e.Key == Key.Left)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (colCount > 0)
                    {
                        dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount], dgdMain.Columns[colCount - 1]);
                    }
                }
                else if (e.Key == Key.Right)
                {
                    e.Handled = true;
                    (sender as DataGridCell).IsEditing = false;

                    if (lastColcount > colCount)
                    {
                        dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount], dgdMain.Columns[colCount + 1]);
                    }
                    else if (lastColcount == colCount)
                    {
                        if (dgdMain.Items.Count - 1 > rowCount)
                        {
                            dgdMain.SelectedIndex = rowCount + 1;
                            dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount + 1], dgdMain.Columns[startColcount]);
                        }
                        else
                        {
                            btnMCEvalCal.Focus();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //
        private void TextBoxFocusInDataGrid(object sender, KeyEventArgs e)
        {
            Lib.Instance.DataGridINControlFocus(sender, e);
        }

        //
        private void TextBoxFocusInDataGrid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            Lib.Instance.DataGridINBothByMouseUP(sender, e);
        }

        //
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            if (strFlag.Equals("I")||strFlag.Equals("U"))
            {
                DataGridCell cell = sender as DataGridCell;
                cell.IsEditing = true;
            }
        }

        //
        private string GetScore(string strName,string strValue, int number)
        {
            string retunString = string.Empty;
            string sql = string.Empty;
            if (thisDTClone.Rows.Count > 0)
            {
                thisDTClone.Rows.Clear();
            }

            if ((strFlag.Equals("I") || strFlag.Equals("U")) && Lib.Instance.IsNumOrAnother(strValue.Replace(",", "")))
            {
                string ColName = string.Empty;
                ColName = thisDT.Columns[2].Caption;

                if (ColName != null && !ColName.Equals(string.Empty))
                {
                    sql = ColName + " = '" + strName+"' ";

                    foreach (DataRow dr in thisDT.Select(sql))
                    {
                        thisDTClone.Rows.Add(dr.ItemArray);
                        if (int.Parse(strValue.Replace(",", "")) >= (int)(double.Parse(dr["EvalSpecMin"].ToString())) &&
                            int.Parse(strValue.Replace(",", "")) <= (int)(double.Parse(dr["EvalSpecMax"].ToString())))
                        {
                            retunString = dr["MCEvalScore"].ToString();
                            break;
                        }
                    }

                    foreach (DataRow dr in thisDTClone.Rows)
                    {
                        int Level = dr.Field<int>("MCEvalScore");
                        if (number == 1)
                        {
                            MaxScore1 = Math.Max(MaxScore1, Level);
                            MinScore1 = Math.Min(MinScore1, Level);
                        }
                        else if (number == 2)
                        {
                            MaxScore2 = Math.Max(MaxScore2, Level);
                            MinScore2 = Math.Min(MinScore2, Level);
                        }
                        else if (number == 3)
                        {
                            MaxScore3 = Math.Max(MaxScore3, Level);
                            MinScore3 = Math.Min(MinScore3, Level);
                        }
                        else if (number == 4)
                        {
                            MaxScore4 = Math.Max(MaxScore4, Level);
                            MinScore4 = Math.Min(MinScore4, Level);
                        }
                        else if (number == 5)
                        {
                            MaxScore5 = Math.Max(MaxScore5, Level);
                            MinScore5 = Math.Min(MinScore5, Level);
                        }
                        else if (number == 6)
                        {
                            MaxScore6 = Math.Max(MaxScore6, Level);
                            MinScore6 = Math.Min(MinScore6, Level);
                        }
                        else if (number == 7)
                        {
                            MaxScore7 = Math.Max(MaxScore7, Level);
                            MinScore7 = Math.Min(MinScore7, Level);
                        }
                        else if (number == 8)
                        {
                            MaxScore8 = Math.Max(MaxScore8, Level);
                            MinScore8 = Math.Min(MinScore8, Level);
                        }
                    }
                }
            }

            return retunString;
        }

        //
        private void MaxAndMinLimit(string strValue, int number)
        {
            if (strValue.Equals(string.Empty))
            {
                return;
            }

            if (number == 1)
            {
                if (int.Parse(strValue) > MaxScore1)
                {
                    MessageBox.Show("해당 평가의 최고점수는 " + MaxScore1 + " 입니다.");
                    return;
                }
                if (int.Parse(strValue) < MinScore1)
                {
                    MessageBox.Show("해당 평가의 최저점수는 " + MinScore1 + " 입니다.");
                    return;
                }
            }
            else if (number == 2)
            {
                if (int.Parse(strValue) > MaxScore2)
                {
                    MessageBox.Show("해당 평가의 최고점수는 " + MaxScore2 + " 입니다.");
                    return;
                }
                if (int.Parse(strValue) < MinScore2)
                {
                    MessageBox.Show("해당 평가의 최저점수는 " + MinScore2 + " 입니다.");
                    return;
                }
            }
            else if (number == 3)
            {
                if (int.Parse(strValue) > MaxScore3)
                {
                    MessageBox.Show("해당 평가의 최고점수는 " + MaxScore3 + " 입니다.");
                    return;
                }
                if (int.Parse(strValue) < MinScore3)
                {
                    MessageBox.Show("해당 평가의 최저점수는 " + MinScore3 + " 입니다.");
                    return;
                }
            }
            else if (number == 4)
            {
                if (int.Parse(strValue) > MaxScore4)
                {
                    MessageBox.Show("해당 평가의 최고점수는 " + MaxScore4 + " 입니다.");
                    return;
                }
                if (int.Parse(strValue) < MinScore4)
                {
                    MessageBox.Show("해당 평가의 최저점수는 " + MinScore4 + " 입니다.");
                    return;
                }
            }
            else if (number == 5)
            {
                if (int.Parse(strValue) > MaxScore5)
                {
                    MessageBox.Show("해당 평가의 최고점수는 " + MaxScore5 + " 입니다.");
                    return;
                }
                if (int.Parse(strValue) < MinScore5)
                {
                    MessageBox.Show("해당 평가의 최저점수는 " + MinScore5 + " 입니다.");
                    return;
                }
            }
            else if (number == 6)
            {
                if (int.Parse(strValue) > MaxScore6)
                {
                    MessageBox.Show("해당 평가의 최고점수는 " + MaxScore6 + " 입니다.");
                    return;
                }
                if (int.Parse(strValue) < MinScore6)
                {
                    MessageBox.Show("해당 평가의 최저점수는 " + MinScore6 + " 입니다.");
                    return;
                }
            }
            else if (number == 7)
            {
                if (int.Parse(strValue) > MaxScore7)
                {
                    MessageBox.Show("해당 평가의 최고점수는 " + MaxScore7 + " 입니다.");
                    return;
                }
                if (int.Parse(strValue) < MinScore7)
                {
                    MessageBox.Show("해당 평가의 최저점수는 " + MinScore7 + " 입니다.");
                    return;
                }
            }
            else if (number == 8)
            {
                if (int.Parse(strValue) > MaxScore8)
                {
                    MessageBox.Show("해당 평가의 최고점수는 " + MaxScore8 + " 입니다.");
                    return;
                }
                if (int.Parse(strValue) < MinScore8)
                {
                    MessageBox.Show("해당 평가의 최저점수는 " + MinScore8 + " 입니다.");
                    return;
                }
            }
        }

        //설비 가동 현황
        private void dgdtpetxtMCoperationRate_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;

                    if (MCEval != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        //MCEval.MCoperationRate = tb1.Text;
                        MCEval.MCoperationRate = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        MCEval.MCoperationRateScore = GetScore("설비가동현황", MCEval.MCoperationRate, 1);
                        tb1.Tag = MCEval.MCoperationRateScore;
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = tb1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : "+ex.ToString());
            }
        }

        //설비 가동 현황 평가
        private void dgdtpetxtMCoperationRateScore_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;

                    if (MCEval != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        //MCEval.MCoperationRate = tb1.Text;
                        MCEval.MCoperationRateScore = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        //MaxAndMinLimit(tb1.Text.Replace(",", ""), 1);
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        tb1.Tag = MCEval.Score;
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = tb1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //고장시 타공정에 미치는 영향
        private void dgdtpecboOtherProceeEffectBroken_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;
                ComboBox cboBase = (ComboBox)sender;

                if (MCEval == null)
                {
                    MCEval = dgdMain.Items[rowNum] as Win_com_MCEvalCal_U_CodeView;
                }

                if (cboBase.SelectedValue != null && !cboBase.SelectedValue.ToString().Equals(string.Empty))
                {
                    var theView = cboBase.SelectedItem as CodeView;
                    if (theView != null)
                    {
                        MCEval.OtherProceeEffectBroken = theView.code_id;
                        MCEval.OtherProceeEffectBroken_CB = theView.code_name;
                        MCEval.OtherProceeEffectBrokenScore = GetScore("고장시 타공정에 미치는 영향", MCEval.OtherProceeEffectBroken, 2);
                        cboBase.Tag = MCEval.OtherProceeEffectBrokenScore;
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = cboBase;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //고장시 타공정에 미치는 영향
        private void dgdtpecboOtherProceeEffectBroken_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DataGridCell cell = Lib.Instance.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

        //고장시 타공정에 미치는 영향 평가
        private void dgdtpetxtOtherProceeEffectBrokenScore_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;

                    if (MCEval != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        //MCEval.MCoperationRate = tb1.Text;
                        MCEval.OtherProceeEffectBrokenScore = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        //MaxAndMinLimit(tb1.Text.Replace(",", ""), 2);
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        tb1.Tag = MCEval.OtherProceeEffectBrokenScore;
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = tb1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //대체 설비 유.무
        private void dgdtpecboReplacementMC_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;
                ComboBox cboBase = (ComboBox)sender;

                if (MCEval == null)
                {
                    MCEval = dgdMain.Items[rowNum] as Win_com_MCEvalCal_U_CodeView;
                }

                if (cboBase.SelectedValue != null && !cboBase.SelectedValue.ToString().Equals(string.Empty))
                {
                    var theView = cboBase.SelectedItem as CodeView;
                    if (theView != null)
                    {
                        MCEval.ReplacementMC = theView.code_id;
                        MCEval.ReplacementMC_CB = theView.code_name;
                        MCEval.ReplacementMCScore = GetScore("대체 설비 유.무", MCEval.ReplacementMC, 3);
                        cboBase.Tag = MCEval.ReplacementMCScore;
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = cboBase;
                    }

                    //sender = cboBase;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //대체 설비 유.무
        private void dgdtpecboReplacementMC_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DataGridCell cell = Lib.Instance.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

        //대체 설비 유.무 평가
        private void dgdtpetxtReplacementMCScore_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;

                    if (MCEval != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        //MCEval.MCoperationRate = tb1.Text;
                        MCEval.ReplacementMCScore = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        //MaxAndMinLimit(tb1.Text.Replace(",", ""), 3);
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        tb1.Tag = MCEval.Score;
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = tb1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }            
        }

        //설비에 의해 공정이 품질에 미치는 영향
        private void dgdtpecboMCProcessQualtityEffect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;
                ComboBox cboBase = (ComboBox)sender;

                if (MCEval == null)
                {
                    MCEval = dgdMain.Items[rowNum] as Win_com_MCEvalCal_U_CodeView;
                }

                if (cboBase.SelectedValue != null && !cboBase.SelectedValue.ToString().Equals(string.Empty))
                {
                    var theView = cboBase.SelectedItem as CodeView;
                    if (theView != null)
                    {
                        MCEval.MCProcessQualtityEffect = theView.code_id;
                        MCEval.MCProcessQualtityEffect_CB = theView.code_name;
                        MCEval.MCProcessQualtityEffectScore =
                            GetScore("설비에 의해 공정이 품질에 미치는 영향", MCEval.MCProcessQualtityEffect, 4);
                        cboBase.Tag = MCEval.MCProcessQualtityEffectScore;
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);

                        sender = cboBase;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //설비에 의해 공정이 품질에 미치는 영향
        private void dgdtpecboMCProcessQualtityEffect_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DataGridCell cell = Lib.Instance.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

        //설비에 의해 공정이 품질에 미치는 영향 평가
        private void dgdtpetxtMCProcessQualtityEffectScore_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;

                    if (MCEval != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        //MCEval.MCoperationRate = tb1.Text;
                        MCEval.MCProcessQualtityEffectScore = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        //MaxAndMinLimit(tb1.Text.Replace(",", ""), 4);
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        tb1.Tag = MCEval.Score;
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = tb1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //손실발생 빈도
        private void dgdtpecboFrequencyOfLoss_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;
                ComboBox cboBase = (ComboBox)sender;

                if (MCEval == null)
                {
                    MCEval = dgdMain.Items[rowNum] as Win_com_MCEvalCal_U_CodeView;
                }

                if (cboBase.SelectedValue != null && !cboBase.SelectedValue.ToString().Equals(string.Empty))
                {
                    var theView = cboBase.SelectedItem as CodeView;
                    if (theView != null)
                    {
                        MCEval.FrequencyOfLoss = theView.code_id;
                        MCEval.FrequencyOfLoss_CB = theView.code_name;
                        MCEval.FrequencyOfLossScore = GetScore("손실발생 빈도", MCEval.FrequencyOfLoss, 5);
                        cboBase.Tag = MCEval.FrequencyOfLossScore;
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);

                        sender = cboBase;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //손실발생 빈도
        private void dgdtpecboFrequencyOfLoss_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DataGridCell cell = Lib.Instance.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

        //손실발생 빈도 평가
        private void dgdtpetxtFrequencyOfLossScore_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;

                    if (MCEval != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        //MCEval.MCoperationRate = tb1.Text;
                        MCEval.FrequencyOfLossScore = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        //MaxAndMinLimit(tb1.Text.Replace(",", ""), 5);
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        tb1.Tag = MCEval.Score;
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = tb1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //고장빈도여부
        private void dgdtpetxtFrequencyOfFailure_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;

                    if (MCEval != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        //MCEval.MCoperationRate = tb1.Text;
                        MCEval.FrequencyOfFailure = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        MCEval.FrequencyOfFailureScore = GetScore("고장빈도 여부", MCEval.FrequencyOfFailure, 6);
                        tb1.Tag = MCEval.FrequencyOfFailureScore;
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = tb1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //고장빈도여부 평가
        private void dgdtpeFrequencyOfFailureScore_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;

                    if (MCEval != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        //MCEval.MCoperationRate = tb1.Text;
                        MCEval.FrequencyOfFailureScore = Lib.Instance.returnNumStringZero(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        //MaxAndMinLimit(tb1.Text.Replace(",", ""), 6);
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        tb1.Tag = MCEval.Score;
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = tb1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //고장정지시간
        private void dgdtpetxtFaultStopTime_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;

                    if (MCEval != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        //MCEval.MCoperationRate = tb1.Text;
                        MCEval.FaultStopTime = Lib.Instance.returnNumStringTwoExceptDot(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        MCEval.FaultStopTimeScore = GetScore("고장 정지 시간", MCEval.FaultStopTime, 7);
                        tb1.Tag = MCEval.FaultStopTimeScore;
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = tb1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //고장정지시간 평가
        private void dgdtpetxtFaultStopTimeScore_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;

                    if (MCEval != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        //MCEval.MCoperationRate = tb1.Text;
                        MCEval.FaultStopTimeScore = Lib.Instance.returnNumStringTwoExceptDot(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        //MaxAndMinLimit(tb1.Text.Replace(",", ""), 7);
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        tb1.Tag = MCEval.Score;
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = tb1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //고장으로 인한 인체 및 환경에 미치는 영향
        private void dgdtpecboHumanEnvironImpact_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;
                ComboBox cboBase = (ComboBox)sender;

                if (MCEval == null)
                {
                    MCEval = dgdMain.Items[rowNum] as Win_com_MCEvalCal_U_CodeView;
                }

                if (cboBase.SelectedValue != null && !cboBase.SelectedValue.ToString().Equals(string.Empty))
                {
                    var theView = cboBase.SelectedItem as CodeView;
                    if (theView != null)
                    {
                        MCEval.HumanEnvironImpact = theView.code_id;
                        MCEval.HumanEnvironImpact_CB = theView.code_name;
                        MCEval.HumanEnvironImpactScore =
                            GetScore("고장으로 인한 인체 및 환경에 미치는 영향", MCEval.HumanEnvironImpact, 8);
                        cboBase.Tag = MCEval.HumanEnvironImpactScore;
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = cboBase;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //고장으로 인한 인체 및 환경에 미치는 영향
        private void dgdtpecboHumanEnvironImpact_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                DataGridCell cell = Lib.Instance.GetParent<DataGridCell>(sender as ComboBox);
                cell.Focus();
            }
        }

        //고장으로 인한 인체 및 환경에 미치는 영향 평가
        private void dgdtpetxtHumanEnvironImpactScore_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;

                    if (MCEval != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        //MCEval.MCoperationRate = tb1.Text;
                        MCEval.HumanEnvironImpactScore = Lib.Instance.returnNumStringTwoExceptDot(tb1.Text);
                        tb1.SelectionStart = tb1.Text.Length;
                        //MaxAndMinLimit(tb1.Text.Replace(",", ""), 8);
                        MCEval.Score = TotalScore(MCEval.MCoperationRateScore, MCEval.OtherProceeEffectBrokenScore,
                                    MCEval.ReplacementMCScore, MCEval.MCProcessQualtityEffectScore, MCEval.FrequencyOfLossScore
                                    , MCEval.FrequencyOfFailureScore, MCEval.FaultStopTimeScore, MCEval.HumanEnvironImpactScore);
                        tb1.Tag = MCEval.Score;
                        MCEval.EvalGrade = Lib.Instance.ReturnGrade(MCEval.Score);
                        sender = tb1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //비고
        private void dgdtpetxtComments_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (strFlag.Equals("I") || strFlag.Equals("U"))
                {
                    var MCEval = dgdMain.CurrentItem as Win_com_MCEvalCal_U_CodeView;

                    if (MCEval != null)
                    {
                        TextBox tb1 = sender as TextBox;
                        //MCEval.MCoperationRate = tb1.Text;
                        MCEval.Comments = tb1.Text;
                        sender = tb1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
            }
        }

        //숫자만 입력
        private void dgdtpetxt_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric(sender as TextBox, e);
        }

        private void ChkAllCheck_Checked(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count > 0)
            {
                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    var Check = dgdMain.Items[i] as Win_com_MCEvalCal_U_CodeView;
                    if (Check != null)
                    {
                        Check.chkData = true;
                    }
                }
            }
        }

        private void ChkAllCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count > 0)
            {
                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    var Check = dgdMain.Items[i] as Win_com_MCEvalCal_U_CodeView;
                    if (Check != null)
                    {
                        Check.chkData = false;
                    }
                }
            }
        }

        
    }

    class Win_com_MCEvalCal_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public int Num { get; set; }
        public string MCEvalID { get; set; }
        public string MCID { get; set; }       
        public string EvalDate { get; set; }
        public string MCNAME { get; set; }
        /// <summary>
        /// 설비 가동 현황
        /// </summary>
        public string MCoperationRate { get; set; }
        /// <summary>
        /// 고장시 타공정에 미치는 영향
        /// </summary>
        public string OtherProceeEffectBroken { get; set; }
        /// <summary>
        /// 대체 설비 유.무
        /// </summary>
        public string ReplacementMC { get; set; }
        /// <summary>
        /// 설비에 의해 공정이 품질에 미치는 영향
        /// </summary>
        public string MCProcessQualtityEffect { get; set; }
        /// <summary>
        /// 손실발생 빈도
        /// </summary>
        public string FrequencyOfLoss { get; set; }
        /// <summary>
        /// 고장빈도여부
        /// </summary>
        public string FrequencyOfFailure { get; set; }
        /// <summary>
        /// 고장정지시간
        /// </summary>
        public string FaultStopTime { get; set; }
        /// <summary>
        /// 고장으로 인한 인체 및 환경에 미치는 영향
        /// </summary>
        public string HumanEnvironImpact { get; set; }

        /// <summary>
        /// 설비 가동 현황
        /// </summary>
        public string MCoperationRateScore { get; set; }
        /// <summary>
        /// 고장시 타공정에 미치는 영향
        /// </summary>
        public string OtherProceeEffectBrokenScore { get; set; }
        /// <summary>
        /// 대체 설비 유.무
        /// </summary>
        public string ReplacementMCScore { get; set; }
        /// <summary>
        /// 설비에 의해 공정이 품질에 미치는 영향
        /// </summary>
        public string MCProcessQualtityEffectScore { get; set; }
        /// <summary>
        /// 손실발생 빈도
        /// </summary>
        public string FrequencyOfLossScore { get; set; }
        /// <summary>
        /// 고장빈도여부
        /// </summary>
        public string FrequencyOfFailureScore { get; set; }
        /// <summary>
        /// 고장정지시간
        /// </summary>
        public string FaultStopTimeScore { get; set; }
        /// <summary>
        /// 고장으로 인한 인체 및 환경에 미치는 영향
        /// </summary>
        public string HumanEnvironImpactScore { get; set; }

        public string OtherProceeEffectBroken_CB { get; set; }
        public string ReplacementMC_CB { get; set; }
        public string MCProcessQualtityEffect_CB { get; set; }
        public string FrequencyOfLoss_CB { get; set; }
        public string HumanEnvironImpact_CB { get; set; }

        public ObservableCollection<CodeView> OvcOtherProceeEffectBroken { get; set; }
        public ObservableCollection<CodeView> OvcReplacementMC { get; set; }
        public ObservableCollection<CodeView> OvcMCProcessQualtityEffect { get; set; }
        public ObservableCollection<CodeView> OvcFrequencyOfLoss { get; set; }
        public ObservableCollection<CodeView> OvcHumanEnvironImpact { get; set; }

        public string Score { get; set; }
        public string EvalGrade { get; set; }
        public string EvalPersonName { get; set; }
        public string Comments { get; set; }
        public string CreateDate { get; set; }

        public bool chkData { get; set; }
        public ObservableCollection<CodeView> OvcMachine { get; set; }
    }

    class NewMCEval : BaseView
    {
        public string EvalDate { get; set; }
        public string MCID { get; set; }
        public string MCName { get; set; }

        public string MCoperationRate { get; set; }
        public string OtherProceeEffectBroken { get; set; }
        public string ReplacementMC { get; set; }
        public string MCProcessQualtityEffect { get; set; }
        public string FrequencyOfLoss { get; set; }
        public string FrequencyOfFailure { get; set; }
        public string FaultStopTime { get; set; }
        public string HumanEnvironImpact { get; set; }
    }
}
