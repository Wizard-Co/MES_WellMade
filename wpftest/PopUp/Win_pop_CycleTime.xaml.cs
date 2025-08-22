using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// RheoChoice.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_pop_CycleTime : Window
    {

        private int rowNum = 0;
        private List<string> lstMachineS = new List<string>();
        private string MachineS = ""; // 해당 호기들

        private int MCnt = 1; // 선택한 갯수

        public Win_pop_CycleTime_CodeView ReceiveData = new Win_pop_CycleTime_CodeView();

        public List<Win_prd_MCRunningGoal_U_Sub_CodeView> lstSub = new List<Win_prd_MCRunningGoal_U_Sub_CodeView>();

        private bool saveFlag = false;

        public Win_pop_CycleTime()
        {
            InitializeComponent();
        }

        private void Win_pop_CycleTime_Loaded(object sender, RoutedEventArgs e)
        {
            SetComboBox();

            re_Search(rowNum);
        }

        private void SetComboBox()
        {
            //자동 여부 : Y - 자동 / N - 수동
            List<string[]> listYN = new List<string[]>();
            string[] YN00 = new string[] { "", "전체" };
            string[] YN01 = new string[] { "Y", "자동" };
            string[] YN02 = new string[] { "N", "수동" };
            listYN.Add(YN00);
            listYN.Add(YN01);
            listYN.Add(YN02);

            ObservableCollection<CodeView> ovcYN = ComboBoxUtil.Instance.Direct_SetComboBox(listYN);
            this.cboAutoMcYN.ItemsSource = ovcYN;
            this.cboAutoMcYN.DisplayMemberPath = "code_name";
            this.cboAutoMcYN.SelectedValuePath = "code_id";

            if (ReceiveData.AutoPassive.Trim().Equals("1"))
            {
                cboAutoMcYN.SelectedIndex = 2;
            }
            else
            {
                cboAutoMcYN.SelectedIndex = 1;
            }
        }

        #region 주요 버튼 이벤트 - 일괄저장, 닫기

        // 일괄저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            lstMachineS.Clear();
            MachineS = "";

            // 호기 세팅
            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                var Main = dgdMain.Items[i] as Win_pop_CycleTime_CodeView;
                if (Main != null)
                {
                    if (Main.Chk == true)
                    {
                        lstMachineS.Add(Main.MachineID);

                        if (lstMachineS.Count == 1)
                        {
                            MachineS += Main.MachineID;
                        }
                        else
                        {
                            MachineS += (", " + Main.MachineID);
                        }
                    }
                }
            }

            if (lstMachineS.Count == 0)
            {
                MessageBox.Show("선택된 호기가 없습니다.");
                return;
            }

            if (MessageBox.Show("일괄 저장하는 경우, 해당 호기들의 이전 데이터는 사라집니다.\r이대로 진행 하시겠습니까??", "저장 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                using (Loading lw = new Loading(SaveData))
                {
                    lw.ShowDialog();
                }

                if (saveFlag == true)
                {
                    this.DialogResult = true;
                }
            }
        }



        //닫기
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        #endregion // 주요 버튼 이벤트

        #region 주요 메서드 모음

        private void re_Search(int rowNum)
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = rowNum;
            }
            else
            {
                this.DataContext = null;
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        #region 조회

        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("ProcessID", ReceiveData.ProcessID);
                sqlParameter.Add("AutoYN", cboAutoMcYN.SelectedValue != null ? cboAutoMcYN.SelectedValue.ToString().Trim() : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sMachineForCycleTime", sqlParameter, false);

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

                            var Machine = new Win_pop_CycleTime_CodeView()
                            {
                                Num = i.ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                                MachineID = dr["MachineID"].ToString(),
                                AutoMCName = dr["AutoMCName"].ToString(),
                            };

                            if (Machine.MachineID.Trim().Equals(ReceiveData.MachineID))
                            {
                                Machine.Chk = true;
                                Machine.FontColor = true;
                            }

                            dgdMain.Items.Add(Machine);
                        }

                        tblCount.Text = "▶선택한 갯수 : " + 1 + "건";
                    }
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("조회 오류 : " + ee.Message);
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        #endregion

        #region 저장
        /// <summary>
        /// 실저장
        /// </summary>
        /// <param name="strFlag"></param>
        /// <param name="strArticleID"></param>
        /// <returns></returns>
        private void SaveData()
        {
            //bool flag = false;
            saveFlag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

            try
            {
                if (CheckData())
                {

                    sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("YYYY", ReceiveData.YYYY);
                    sqlParameter.Add("ProcessID", ReceiveData.ProcessID);
                    sqlParameter.Add("MachineS", @MachineS);

                    Procedure pro0 = new Procedure();
                    pro0.Name = "xp_MachineGoal_dMachineGoalAll_SelectSave";
                    pro0.OutputUseYN = "N";
                    pro0.OutputName = "sArticleID";
                    pro0.OutputLength = "5";

                    Prolist.Add(pro0);
                    ListParameter.Add(sqlParameter);

                    for (int j = 0; j < lstMachineS.Count; j++)
                    {
                        string MachineID = lstMachineS[j].Trim();

                        for (int i = 0; i < lstSub.Count; i++)
                        {
                            var WinSub = lstSub[i] as Win_prd_MCRunningGoal_U_Sub_CodeView;

                            if (WinSub != null
                                && WinSub.ArticleID != null)
                            {
                                sqlParameter = new Dictionary<string, object>();
                                sqlParameter.Clear();

                                sqlParameter.Add("YYYY", ReceiveData.YYYY);
                                sqlParameter.Add("ProcessID", ReceiveData.ProcessID);
                                sqlParameter.Add("MachineID", MachineID);
                                sqlParameter.Add("ArticleID", WinSub.ArticleID);
                                sqlParameter.Add("CT", ConvertDouble(WinSub.CT));
                                sqlParameter.Add("AutoPassive", ReceiveData.AutoPassive);

                                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                                Procedure pro1 = new Procedure();
                                pro1.Name = "xp_MachineGoal_iMachineGoal_SelectSave";
                                pro1.OutputUseYN = "N";
                                pro1.OutputName = "sArticleID";
                                pro1.OutputLength = "5";

                                Prolist.Add(pro1);
                                ListParameter.Add(sqlParameter);
                            }
                        }
                    }
                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                    if (Confirm[0] != "success")
                    {
                        MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                        //flag = false;
                        saveFlag = false;
                        //return false;
                    }
                    else
                    {
                        //flag = true;
                        saveFlag = true;
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

            //return flag;
        }
        #endregion

        #endregion

        #region 유효성 검사

        private bool CheckData()
        {
            bool flag = true;

            return flag;
        }

        #endregion

        #region 데이터 그리드 체크박스 이벤트

        // 팝업창 체크박스 이벤트
        private void CHK_Click_Sub(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            var Machine = chkSender.DataContext as Win_pop_CycleTime_CodeView;

            if (Machine != null)
            {
                if (chkSender.IsChecked == true)
                {
                    Machine.Chk = true;
                    Machine.FontColor = true;

                    MCnt++;

                    tblCount.Text = "▶선택한 갯수 : " + MCnt + "건";
                }
                else
                {
                    Machine.Chk = false;
                    Machine.FontColor = false;

                    MCnt--;

                    tblCount.Text = "▶선택한 갯수 : " + MCnt + "건";
                }
            }
        }

        #endregion // 데이터 그리드 체크박스 이벤트

        #region 전체 선택 체크박스 이벤트

        // 전체 선택 체크박스 체크 이벤트
        private void AllCheck_Checked(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Visibility == Visibility.Visible)
            {
                MCnt = 0;
                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    var Machine = dgdMain.Items[i] as Win_pop_CycleTime_CodeView;
                    if (Machine != null)
                    {
                        Machine.Chk = true;
                        Machine.FontColor = true;

                        MCnt++;
                    }
                }

                tblCount.Text = "▶선택한 갯수 : " + MCnt + "건";
            }
        }

        // 전체 선택 체크박스 언체크 이벤트
        private void AllCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Visibility == Visibility.Visible)
            {
                MCnt = 0;
                for (int i = 0; i < dgdMain.Items.Count; i++)
                {
                    var Machine = dgdMain.Items[i] as Win_pop_CycleTime_CodeView;
                    if (Machine != null)
                    {
                        Machine.Chk = false;
                        Machine.FontColor = false;
                    }

                }

                tblCount.Text = "▶선택한 갯수 : " + 0 + "건";
            }
        }

        #endregion // 전체 선택 체크박스 이벤트

        #region 기타 메서드

        // 천마리 콤마, 소수점 버리기
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







        #endregion // 기타 메서드

        private void cboAutoMcYN_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            re_Search(0);
        }
    }

    public class Win_pop_CycleTime_CodeView : BaseView
    {
        public string Num { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string MachineNo { get; set; }
        public string MachineID { get; set; }

        public string AutoMCName { get; set; }
        public string AutoPassive { get; set; } // 0 자동 → Y / 1 수동 → N
        public string YYYY { get; set; }

        public bool Chk { get; set; }
        public bool FontColor { get; set; }
    }
}
