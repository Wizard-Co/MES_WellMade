/**
 * 
 * @details     공정패턴 코드 등록
 * @author      정승학
 * @date        2019-07-29
 * @version     1.0
 * 
 * @section MODIFYINFO 수정정보
 * - 수정일        - 수정자       : 수정내역
 * - 2021-10-14    - 정승학       : 소스 재작성
 * 
 * 
 * */

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
using WizMes_WellMade.PopUp;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_prd_PatternCode_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_PatternCode_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = MainWindow.pf;

        string strFlag = string.Empty;
        int rowNum = 0;


        public Win_prd_PatternCode_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            lib.UiLoading(sender);
            SetComboBox();
        }

        #region 콤보박스
        private void SetComboBox()
        {
            ObservableCollection<CodeView> ovcArticleGrp = ComboBoxUtil.Instance.GetArticleCode_SetComboBox("", 0);
            this.ComboBoxArticleGroup.ItemsSource = ovcArticleGrp;
            this.ComboBoxArticleGroup.DisplayMemberPath = "code_name";
            this.ComboBoxArticleGroup.SelectedValuePath = "code_id";
        }
        #endregion

        #region 조회
        private void FillGrid()
        {
            if (DataGridPattern.Items.Count > 0)
            {
                DataGridPattern.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("nChkWorkID", 0);
                sqlParameter.Add("sWorkID", "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_prd_sPattern", sqlParameter, true, "R");

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
                            var winPattern = new Win_prd_PatternCode_U_CodeView()
                            {
                                Num = i + 1,
                                Pattern = dr["Pattern"].ToString(),
                                PatternID = dr["PatternID"].ToString(),
                                ArticleGrpID = dr["ArticleGrpID"].ToString(),
                                WorkID = dr["WorkID"].ToString(),
                                WorkName = dr["WorkName"].ToString(),
                                ArticleGrp = dr["ArticleGrp"].ToString(),
                            };

                            DataGridPattern.Items.Add(winPattern);

                            i++;
                        }
                    }

                    TextBlockCountPattern.Text = " ▶ 검색 결과 : " + i + " 건";
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
        #endregion

        #region 전체 공정 조회
        private void FillGridlAllProcess()
        {
            if (DataGridAllProcess.Items.Count > 0)
            {
                DataGridAllProcess.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ProcessID", "");
                ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sProcess", sqlParameter, false);

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
                            var winPatternAllProcess = new Win_prd_PatternCode_U_Process_CodeView()
                            {
                                Num = i + 1,
                                Process = dr["Process"].ToString(),
                                ProcessID = dr["ProcessID"].ToString()
                            };

                            DataGridAllProcess.Items.Add(winPatternAllProcess);
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
        #endregion

        #region PatternID로 패턴의 공정과 공정순서 가져오기
        private void FillGrid_OrderAndProcess(string strPatternID)
        {
            if (DataGridPatternProcess.Items.Count > 0)
            {
                DataGridPatternProcess.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sPatternID", strPatternID);

                ds = DataStore.Instance.ProcedureToDataSet("xp_Pattern_sPatternSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;
                    string strProcessPattern = string.Empty;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("조회된 데이터가 없습니다.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinPatternProcess = new Win_prd_PatternCode_U_Process_CodeView()
                            {
                                Num = i + 1,
                                Process = dr["Process"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                PatternSeq = dr["PatternSeq"].ToString()
                            };

                            DataGridPatternProcess.Items.Add(WinPatternProcess);
                            i++;

                            if (i == drc.Count)
                            {
                                strProcessPattern += WinPatternProcess.Process;
                            }
                            else
                            {
                                strProcessPattern += WinPatternProcess.Process + "→";
                            }
                        }

                        if (DataGridProcessOrder.Items.Count > 0)
                        {
                            DataGridProcessOrder.Items.Clear();
                        }

                        var WinProcessOrder = new Win_prd_PatternCode_U_Order_CodeView() { ProcessOrder = strProcessPattern };
                        DataGridProcessOrder.Items.Add(WinProcessOrder);
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
        #endregion

        #region 삭제
        private bool DeleteData(string strPatternId)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sPatternID", strPatternId);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Pattern_dPattern", sqlParameter, "D");
                DataStore.Instance.CloseConnection();

                if (result[0].Equals("success"))
                {
                    flag = true;
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }

        #endregion

        #region 데이터 체크
        private bool CheckData()
        {
            bool flag = true;

            if (ComboBoxArticleGroup.SelectedIndex == -1 || ComboBoxArticleGroup.SelectedValue == null)
            {
                MessageBox.Show("제품그룹이 선택되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (TextBoxPattern.Text.Length <= 0 || TextBoxPattern.Text.Equals(""))
            {
                MessageBox.Show("패턴 설명이 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            //if (DataGridPatternProcess.Items.Count <= 0)
            //{
            //    MessageBox.Show("선택된 공정이 없습니다.");
            //    flag = false;
            //    return flag;
            //}

            if (TextBoxPattern.Text.Length > 20)
            {
                MessageBox.Show("패턴설명은 최대 20글자 까지 가능합니다.");
                flag = false;
                return flag;
            }

            return flag;
        }
        #endregion

        #region 저장
        private bool SaveData(string strFlag, string strPatternID)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    string sPatternID = string.Empty;
                    string sPattern = string.Empty;
                    string sWorkID = "0001"; //고정값으로 들어가게
                    string sArticleGrpID = string.Empty;

                    sPatternID = strPatternID;
                    sPattern = TextBoxPattern.Text;
                    sArticleGrpID = ComboBoxArticleGroup.SelectedValue == null ? "" : ComboBoxArticleGroup.SelectedValue.ToString();

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sPatternID", strPatternID);
                    sqlParameter.Add("sPattern", sPattern);
                    sqlParameter.Add("sWorkID", sWorkID);
                    sqlParameter.Add("sArticleGrpID", sArticleGrpID);

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Pattern_iPattern";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "sPatternID";
                        pro1.OutputLength = "2";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < DataGridPatternProcess.Items.Count; i++)
                        {
                            DataGridRow dgr = lib.GetRow(i, DataGridPatternProcess);
                            
                            var winPatternProcess = dgr.Item as Win_prd_PatternCode_U_Process_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("sPatternID", sPatternID);
                            sqlParameter.Add("nPatternSeq", i + 1);
                            sqlParameter.Add("sProcessID", winPatternProcess.ProcessID);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_Pattern_iPatternSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "sPatternID";
                            pro2.OutputLength = "2";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter,"C");
                        string sGetPatternID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "sPatternID")
                                {
                                    sGetPatternID = kv.value;
                                    flag = true;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                            //return false;
                        }
                    }
                    #endregion

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Pattern_uPattern";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sPatternID";
                        pro1.OutputLength = "2";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < DataGridPatternProcess.Items.Count; i++)
                        {
                            DataGridRow dgr = lib.GetRow(i, DataGridPatternProcess);
                            var winPatternProcess = dgr.Item as Win_prd_PatternCode_U_Process_CodeView;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("sPatternID", strPatternID);
                            sqlParameter.Add("nPatternSeq", i + 1);
                            sqlParameter.Add("sProcessID", winPatternProcess.ProcessID);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_Pattern_iPatternSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "sPatternID";
                            pro2.OutputLength = "2";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"U");
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
                    #endregion
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return flag;
        }
        #endregion

        #region 활성화 & 비활성화
        /// <summary>
        /// 추가, 수정모드(추가 및 수정 버튼 클릭했을 때)
        /// 저장, 취소 버튼만 활성화 나머지 비활성화
        /// </summary>
        private void AddUpdateMode()
        {
            lib.UiButtonEnableChange_SCControl(this);

            GridInputArea1.IsHitTestVisible = true;
            GridInputArea2.IsHitTestVisible = true;

            DataGridPattern.IsHitTestVisible = false;

        }

        /// <summary>
        /// 조회모드(저장 및 취소 버튼 클릭했을 때)
        /// 저장, 취소 버튼만 비활성화 나머지 활성화
        /// </summary>
        private void SaveCancelMode()
        {
            lib.UiButtonEnableChange_IUControl(this);

            GridInputArea1.IsHitTestVisible = false;
            GridInputArea2.IsHitTestVisible = false;

            DataGridPattern.IsHitTestVisible = true;

        }
        #endregion

        #region 입력창 초기화
        private void ClearData()
        {
            this.DataContext = null;

            TextBoxCode.Text = string.Empty;
            ComboBoxArticleGroup.SelectedIndex = -1;
            TextBoxPattern.Text = string.Empty;

        }

        #endregion

        #region ReSearch
        //재조회
        private void Re_Search()
        {
            FillGrid();

            if (DataGridPattern.Items.Count <= 0)
            {
                MessageBox.Show("조회된 내용이 없습니다.");
            }
            else
            {
                DataGridPattern.SelectedIndex = rowNum;
            }
        }


        private void BeSave()
        {
            if (SaveData(strFlag, TextBoxCode.Text))
            {
                SaveCancelMode();
                Re_Search();
                strFlag = "";
            }
        }

        private void BeDelete()
        {
            var patternInfo = new Win_prd_PatternCode_U_CodeView();

            if (patternInfo != null)
            {
                if (DeleteData(patternInfo.PatternID))
                {
                    Re_Search();
                }
            }
        }

        private void BeCancel()
        {
            SaveCancelMode();
            strFlag = "";

            Re_Search();
        }

        #endregion

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                AddUpdateMode();
                strFlag = "I";

                lblMsg.Visibility = Visibility.Visible;
                tbkMsg.Text = "자료 입력 중";
                rowNum = DataGridPattern.SelectedIndex;
                this.DataContext = null;
                ClearData();

                if(DataGridPatternProcess.Items.Count > 0)
                {
                    DataGridPatternProcess.Items.Clear();
                }

                FillGridlAllProcess();
                ComboBoxArticleGroup.IsDropDownOpen = true;
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var winPattern = DataGridPattern.SelectedItem as Win_prd_PatternCode_U_CodeView;

                if (winPattern != null)
                {
                    rowNum = DataGridPattern.SelectedIndex;
                    DataGridPattern.IsHitTestVisible = false;
                    tbkMsg.Text = "자료 수정 중";
                    lblMsg.Visibility = Visibility.Visible;
                    AddUpdateMode();
                    strFlag = "U";

                    FillGridlAllProcess();
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var winPattern = DataGridPattern.SelectedItem as Win_prd_PatternCode_U_CodeView;

                if (winPattern == null)
                {
                    MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
                }
                else
                {
                    if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        if (DataGridPattern.Items.Count > 0 && DataGridPattern.SelectedItem != null)
                        {
                            rowNum = DataGridPattern.SelectedIndex;
                        }

                        if (DeleteData(winPattern.PatternID))
                        {
                            rowNum -= 1;
                            Re_Search();
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
                this.DataContext = null;

                lib.ChildMenuClose(this.ToString());
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            // 검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                try
                {
                    rowNum = 0;
                    using (Loading lw = new Loading(Re_Search))
                    {
                        lw.ShowDialog();
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show("예외처리 - " + ee.ToString());
                }

            }), System.Windows.Threading.DispatcherPriority.Background);

            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
            
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (Loading lw = new Loading(BeSave))
                {
                    lw.ShowDialog();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (Loading lw = new Loading(BeCancel))
                {
                    lw.ShowDialog();
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataTable dt = null;
                string Name = string.Empty;

                string[] lst = new string[6];
                lst[0] = "패턴";
                lst[1] = "전체 공정";
                lst[2] = "선택된 공정(패턴 공정)";
                lst[3] = DataGridPattern.Name;
                lst[4] = DataGridAllProcess.Name;
                lst[5] = DataGridPatternProcess.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(lst);
                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(DataGridPattern.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(DataGridPattern);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(DataGridPattern);

                        Name = DataGridPattern.Name;
                        Lib.Instance.GenerateExcel(dt, Name);
                        Lib.Instance.excel.Visible = true;
                    }
                    else if (ExpExc.choice.Equals(DataGridAllProcess.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(DataGridAllProcess);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(DataGridAllProcess);

                        Name = DataGridAllProcess.Name;
                        Lib.Instance.GenerateExcel(dt, Name);
                        Lib.Instance.excel.Visible = true;
                    }
                    else if (ExpExc.choice.Equals(DataGridPatternProcess.Name))
                    {
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(DataGridPatternProcess);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(DataGridPatternProcess);

                        Name = DataGridPatternProcess.Name;
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
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
            finally
            {
                lib.ReleaseExcelObject(lib.workSheet);
                lib.ReleaseExcelObject(lib.workBook);
                lib.ReleaseExcelObject(lib.excel);
            }
        }

        //공정이동-오른쪽으로
        private void ButtonMoveRight_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var winPatternAllProcess = DataGridAllProcess.SelectedItem as Win_prd_PatternCode_U_Process_CodeView;
                bool flag = true;

                if (winPatternAllProcess != null) 
                {
                    for (int i = 0; i < DataGridPatternProcess.Items.Count; i++) //패턴공장에 아무것도 없으면 0번이라 실행X
                    {
                        var WinPP = DataGridPatternProcess.Items[i] as Win_prd_PatternCode_U_Process_CodeView;

                        if (WinPP.Process == winPatternAllProcess.Process) //패턴공장 공정영이랑 winpp공정명 같으면(이미있으면) flag flase반환 
                        {
                            flag = false;
                            break;
                        }
                    }

                    if (flag) 
                    {
                        if (DataGridPatternProcess.Items.Count > 0)
                        {
                            winPatternAllProcess.Num = DataGridPatternProcess.Items.Count + 1;
                        }

                        DataGridPatternProcess.Items.Add(winPatternAllProcess);
                    }
                    else //flag=false 패턴공장 공정명이랑 winpp공정명 같으면(이미있으면) 
                    {
                        MessageBox.Show("같은 이름의 공정이 추가되어있습니다.");
                    }
                }
                else //winPatternAllProcess != null
                {
                    MessageBox.Show("패턴에 추가할 공정이 선택되지 않았습니다.");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //공정이동-왼쪽으로
        private void ButtonMoveLeft_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var winPatternProcess = DataGridPatternProcess.SelectedItem as Win_prd_PatternCode_U_Process_CodeView;

                if (winPatternProcess != null)
                {
                    DataGridPatternProcess.Items.Remove(winPatternProcess);
                }
                else
                {
                    MessageBox.Show("패턴에서 제외할 공정이 선택되지 않았습니다.");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //패턴공정 순서 변경
        private void ButtonStepUpDown_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Button senderBtn = sender as Button;

                Win_prd_PatternCode_U_Process_CodeView AppStepE = new Win_prd_PatternCode_U_Process_CodeView();

                // 아래 버튼 클릭시
                if (senderBtn.Tag.ToString().Equals("Down"))
                {
                    var StepFrom = DataGridPatternProcess.SelectedItem as Win_prd_PatternCode_U_Process_CodeView;

                    if (StepFrom != null)
                    {
                        int currRow = DataGridPatternProcess.SelectedIndex;

                        int goalRow = currRow + 1;
                        int maxRow = DataGridPatternProcess.Items.Count - 1;

                        if (goalRow <= maxRow)
                        {
                            var StepTo = DataGridPatternProcess.Items[goalRow] as Win_prd_PatternCode_U_Process_CodeView;

                            if (StepTo != null)
                            {
                                DataGridPatternProcess.Items.RemoveAt(currRow); // 선택한 행 지우고
                                DataGridPatternProcess.Items.RemoveAt(currRow); // 바로 밑의 행 지우고

                                StepTo.Num = currRow + 1;
                                DataGridPatternProcess.Items.Insert(currRow, StepTo);

                                StepFrom.Num = goalRow + 1;
                                DataGridPatternProcess.Items.Insert(goalRow, StepFrom);

                                DataGridPatternProcess.SelectedIndex = goalRow;
                            }
                        }
                    }
                }
                else // 위 버튼 클릭시
                {
                    var StepFrom = DataGridPatternProcess.SelectedItem as Win_prd_PatternCode_U_Process_CodeView;

                    if (StepFrom != null)
                    {
                        int currRow = DataGridPatternProcess.SelectedIndex;

                        int goalRow = currRow - 1;

                        if (goalRow >= 0)
                        {
                            var StepTo = DataGridPatternProcess.Items[goalRow] as Win_prd_PatternCode_U_Process_CodeView;

                            if (StepTo != null)
                            {
                                DataGridPatternProcess.Items.RemoveAt(goalRow); // 선택한 행 지우고
                                DataGridPatternProcess.Items.RemoveAt(goalRow); // 바로 밑의 행 지우고

                                StepTo.Num = currRow + 1;
                                DataGridPatternProcess.Items.Insert(goalRow, StepTo);

                                StepFrom.Num = goalRow + 1;
                                DataGridPatternProcess.Items.Insert(goalRow, StepFrom);

                                DataGridPatternProcess.SelectedIndex = goalRow;
                            }
                        }
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //더블 클릭으로 공정이동
        private void MouseLeftButtonDownDataGridAllProcess(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (e.ClickCount == 2)
                {
                    TextBlock tblSender = sender as TextBlock;

                    var Process = tblSender.DataContext as Win_prd_PatternCode_U_Process_CodeView;

                    if (Process != null)
                    {
                        for (int i = 0; i < DataGridPatternProcess.Items.Count; i++)
                        {
                            var Compare = DataGridPatternProcess.Items[i] as Win_prd_PatternCode_U_Process_CodeView;

                            if (Compare != null)
                            {
                                if (Compare.ProcessID.Trim().Equals(Process.ProcessID))
                                {
                                    return;
                                }
                            }
                        }

                        var newP = new Win_prd_PatternCode_U_Process_CodeView()
                        {
                            Num = DataGridPatternProcess.Items.Count + 1,
                            ProcessID = Process.ProcessID,
                            Process = Process.Process
                        };

                        //Process.Num = dgdPatternProcess.Items.Count + 1;
                        DataGridPatternProcess.Items.Add(newP);
                        //(dgdPatternProcess.Items[dgdPatternProcess.Items.Count - 1] as Win_prd_PatternCode_U_Process_CodeView).Num = dgdPatternProcess.Items.Count;
                    }
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //더블클릭으로 패턴공정에서 공정 빼기
        private void MouseLeftButtonDownDataGridPatternProcess(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (e.ClickCount == 2)
                {
                    TextBlock tblSender = sender as TextBlock;

                    var Process = tblSender.DataContext as Win_prd_PatternCode_U_Process_CodeView;

                    if (Process != null)
                    {
                        DataGridPatternProcess.Items.Remove(Process);
                        SettingNum();
                    }
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //순번 재배치
        private void SettingNum()
        {
            for (int i = 0; i < DataGridPatternProcess.Items.Count; i++)
            {
                var Process = DataGridPatternProcess.Items[i] as Win_prd_PatternCode_U_Process_CodeView;
                if (Process != null)
                {
                    Process.Num = i + 1;
                }
            }
        }

        private void DataGridPattern_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var winPattern = DataGridPattern.SelectedItem as Win_prd_PatternCode_U_CodeView;

                if (winPattern != null)
                {
                    FillGrid_OrderAndProcess(winPattern.PatternID);
                    this.DataContext = winPattern;
                }
            }
            catch(Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        #region 입력창 이동 이벤트

        private void ComboBoxArticleGroup_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                TextBoxPattern.Focus();
            }
        }

        private void ComboBoxArticleGroup_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                TextBoxPattern.Focus();
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        #endregion
    }

    #region CodeView
    class Win_prd_PatternCode_U_CodeView
    {
        public int Num { get; set; }
        public string PatternID { get; set; }
        public string Pattern { get; set; }
        public string WorkID { get; set; }
        public string WorkName { get; set; }
        public string ArticleGrpID { get; set; }
        public string ArticleGrp { get; set; }
    }

    class Win_prd_PatternCode_U_Process_CodeView : BaseView
    {
        public int Num { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string DetailProcessYN { get; set; }
        public string PatternSeq { get; set; }
    }

    class Win_prd_PatternCode_U_Order_CodeView
    {
        public string ProcessOrder { get; set; }
    }

    #endregion
}
