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
    /// Win_com_MCEvalBasis_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_prd_EvalBasis_U : UserControl
    {
        string strFlag = string.Empty;
        int rowNum = 0;
        Win_com_MCEvalBasis_U_CodeView MCEvalBasis = new Win_com_MCEvalBasis_U_CodeView();
        Lib lib = new Lib();

        /// <summary>
        /// 추가 누른 후 작업하면 모두 삭제후 추가한다,
        /// Items.Clear() 전 임시 저장하고 추가->저장 시 삭제시 사용
        /// </summary>
        List<Win_com_MCEvalBasis_U_CodeView> lstMCEvaluBasis = new List<Win_com_MCEvalBasis_U_CodeView>();

        public Win_prd_EvalBasis_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
        }

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {            
            //gbxMcPart.IsEnabled = false;
            Lib.Instance.UiButtonEnableChange_IUControl(this);

            btnMainAdd.IsEnabled = false;
            btnMainDel.IsEnabled = false;
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            //gbxMcPart.IsEnabled = true;
            Lib.Instance.UiButtonEnableChange_SCControl(this);

            btnMainAdd.IsEnabled = true;
            btnMainDel.IsEnabled = true;
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            CantBtnControl();
            strFlag = "I";

            lblMsg.Visibility = Visibility.Visible;
            tbkMsg.Text = "자료 입력 중";
            rowNum = dgdMain.SelectedIndex;
            this.DataContext = null;

            //추가 전 삭제할거 모아모아 저장
            for (int i = 0; i < dgdMain.Items.Count; i++)
            {
                MCEvalBasis = dgdMain.Items[i] as Win_com_MCEvalBasis_U_CodeView;
                lstMCEvaluBasis.Add(MCEvalBasis);
            }

            //화면상에서 행 삭제
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            //첫행 자동 추가
            if (dgdMain.Items.Count == 0)
            {
                MainAddRow();
            }
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            MCEvalBasis = dgdMain.SelectedItem as Win_com_MCEvalBasis_U_CodeView;

            if (MCEvalBasis != null)
            {
                rowNum = dgdMain.SelectedIndex;
                //dgdMain.IsEnabled = false;
                tbkMsg.Text = "자료 수정 중";
                lblMsg.Visibility = Visibility.Visible;
                CantBtnControl();
                strFlag = "U";
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            //MCEvalBasis = dgdMain.SelectedItem as Win_com_MCEvalBasis_U_CodeView;

            if (dgdMain.Items.Count <= 0)
            {
                MessageBox.Show("삭제할 데이터가 존재하지 않습니다.");
            }
            else
            {
                if (MessageBox.Show("모든 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (DeleteData())
                    {
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
                rowNum = 0;
                re_Search(rowNum);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (DeleteData())
            {
                if (SaveData(strFlag))
                {
                    CanBtnControl();
                    lblMsg.Visibility = Visibility.Hidden;
                    rowNum = 0;
                    lstMCEvaluBasis.Clear();
                    re_Search(rowNum);
                }
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();

            if (!strFlag.Equals(string.Empty))
            {
                re_Search(rowNum);
            }
            
            lstMCEvaluBasis.Clear();
            strFlag = string.Empty;
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "설비등급평가기준";
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

        /// <summary>
        /// 재검색(수정,삭제,추가 저장후에 자동 재검색)
        /// </summary>
        /// <param name="selectedIndex"></param>
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = selectedIndex;
            }
        }

        /// <summary>
        /// 실조회
        /// </summary>
        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }

            try
            {
                DataSet ds = null;
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                ds = DataStore.Instance.ProcedureToDataSet("xp_mc_sMcEvalBasis", sqlParameter, false);

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
                            var WinMCEvalBasis = new Win_com_MCEvalBasis_U_CodeView()
                            {
                                MCGroupName = dr["MCGroupName"].ToString(),
                                MCSeq = dr["MCSeq"].ToString(),
                                AutoMCEvaluYN = dr["AutoMCEvaluYN"].ToString(),
                                MCEvalName = dr["MCEvalName"].ToString(),
                                EvalSpecMin = dr["EvalSpecMin"].ToString(),
                                EvalSpecMax = dr["EvalSpecMax"].ToString(),
                                MCEvalSpec = dr["MCEvalSpec"].ToString(),
                                MCEvalScore = dr["MCEvalScore"].ToString()
                            };

                            if (WinMCEvalBasis.EvalSpecMin.Contains("."))
                            {
                                if (int.Parse(WinMCEvalBasis.EvalSpecMin.Substring(WinMCEvalBasis.EvalSpecMin.IndexOf(".") + 1)) > 0)
                                {
                                    WinMCEvalBasis.EvalSpecMin = Lib.Instance.returnNumStringTwo(WinMCEvalBasis.EvalSpecMin);
                                }
                                else
                                {
                                    WinMCEvalBasis.EvalSpecMin = Lib.Instance.returnNumStringZero(WinMCEvalBasis.EvalSpecMin);
                                }
                            }
                            else
                            {
                                WinMCEvalBasis.EvalSpecMin = Lib.Instance.returnNumStringZero(WinMCEvalBasis.EvalSpecMin);
                            }

                            if (WinMCEvalBasis.EvalSpecMax.Contains("."))
                            {
                                if (int.Parse(WinMCEvalBasis.EvalSpecMax.Substring(WinMCEvalBasis.EvalSpecMax.IndexOf(".") + 1)) > 0)
                                {
                                    WinMCEvalBasis.EvalSpecMax = Lib.Instance.returnNumStringTwo(WinMCEvalBasis.EvalSpecMax);
                                }
                                else
                                {
                                    WinMCEvalBasis.EvalSpecMax = Lib.Instance.returnNumStringZero(WinMCEvalBasis.EvalSpecMax);
                                }
                            }
                            else
                            {
                                WinMCEvalBasis.EvalSpecMax = Lib.Instance.returnNumStringZero(WinMCEvalBasis.EvalSpecMax);
                            }

                            if (WinMCEvalBasis.MCEvalScore.Contains("."))
                            {
                                if (int.Parse(WinMCEvalBasis.MCEvalScore.Substring(WinMCEvalBasis.MCEvalScore.IndexOf(".") + 1)) > 0)
                                {
                                    WinMCEvalBasis.MCEvalScore = Lib.Instance.returnNumStringTwo(WinMCEvalBasis.MCEvalScore);
                                }
                                else
                                {
                                    WinMCEvalBasis.MCEvalScore = Lib.Instance.returnNumStringZero(WinMCEvalBasis.MCEvalScore);
                                }
                            }
                            else
                            {
                                WinMCEvalBasis.MCEvalScore = Lib.Instance.returnNumStringZero(WinMCEvalBasis.MCEvalScore);
                            }

                            dgdMain.Items.Add(WinMCEvalBasis);
                            i++;
                        }
                        tbkCount.Text = "▶ 검색결과 : " + i.ToString() + " 건";
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

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                //sqlParameter.Add("MCEvalCode", strCode);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_mc_dMcEvalBasis", sqlParameter, false);

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
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

        /// <summary>
        /// 저장
        /// </summary>
        /// <param name="strFlag"></param>
        /// <returns></returns>
        private bool SaveData(string strFlag)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
            
            try
            { 
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                foreach (Win_com_MCEvalBasis_U_CodeView McEB in dgdMain.Items)
                {
                    if (strFlag.Equals("I") || strFlag.Equals("U"))
                    {
                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("MCSeq", McEB.MCSeq);
                        sqlParameter.Add("MCGroupName", McEB.MCGroupName);
                        sqlParameter.Add("MCEvalName", McEB.MCEvalName);
                        sqlParameter.Add("EvalSpecMin", McEB.EvalSpecMin.Replace(",", ""));
                        sqlParameter.Add("EvalSpecMax", McEB.EvalSpecMax.Replace(",", ""));
                        sqlParameter.Add("MCEvalSpec", McEB.MCEvalSpec);
                        sqlParameter.Add("MCEvalScore", McEB.MCEvalScore.Replace(",", ""));
                        sqlParameter.Add("AutoMCEvaluYN", McEB.AutoMCEvaluYN);
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_mc_iMcEvalBasis";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "MCSeq";
                        pro1.OutputLength = "4";

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

        /// <summary>
        /// 입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData(Win_com_MCEvalBasis_U_CodeView McEB)
        {
            bool flag = true;
            return flag;
        }

        //행추가 클릭
        private void btnMainAdd_Click(object sender, RoutedEventArgs e)
        {
            MainAddRow();
        }

        //행추가
        private void MainAddRow()
        {
            int i = 1;

            if (dgdMain.Items.Count > 0)
            {
                i = dgdMain.Items.Count + 1;
            }

            var WinMCEval = new Win_com_MCEvalBasis_U_CodeView()
            {
                MCSeq = i.ToString(),
                MCGroupName = "",
                MCEvalSpec ="",
                MCEvalName = "",
                EvalSpecMin = "",
                EvalSpecMax = "",
                MCEvalScore = "",
                AutoMCEvaluYN=""
            };

            dgdMain.Items.Add(WinMCEval);
        }

        //행삭제 클릭
        private void btnMainDel_Click(object sender, RoutedEventArgs e)
        {
            MCEvalBasis = dgdMain.SelectedItem as Win_com_MCEvalBasis_U_CodeView;
            MainRemoveRow();
            //if (strFlag.Equals("U"))
            //{
            //    lstMCEvaluBasis.Add(MCEvalBasis);
            //    MainRemoveRow();
            //}
            //else
            //{
            //    MainRemoveRow();
            //}
        }

        //행 삭제
        private void MainRemoveRow()
        {
            if (dgdMain.Items.Count > 0)
            {
                if (dgdMain.CurrentItem != null)
                {
                    dgdMain.Items.Remove((dgdMain.CurrentItem as Win_com_MCEvalBasis_U_CodeView));
                }
                else
                {
                    dgdMain.Items.Remove((dgdMain.Items[dgdMain.Items.Count - 1]) as Win_com_MCEvalBasis_U_CodeView);
                }
                dgdMain.Refresh();
            }
        }

        //
        private void DataGridCell_KeyDown(object sender, KeyEventArgs e)
        {
            MCEvalBasis = dgdMain.CurrentItem as Win_com_MCEvalBasis_U_CodeView;
            int rowCount = dgdMain.Items.IndexOf(dgdMain.CurrentItem);
            int colCount = dgdMain.Columns.IndexOf(dgdMain.CurrentCell.Column);

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (dgdMain.Columns.Count - 1 == colCount && dgdMain.Items.Count - 1 > rowCount)
                {
                    dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount + 1], dgdMain.Columns[1]);
                }
                else if (dgdMain.Columns.Count - 1 > colCount && dgdMain.Items.Count - 1 > rowCount)
                {
                    dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount], dgdMain.Columns[colCount + 1]);
                }
                else if (dgdMain.Columns.Count - 1 == colCount && dgdMain.Items.Count - 1 == rowCount)
                {
                    btnSave.Focus();
                }
                else if (dgdMain.Columns.Count - 1 > colCount && dgdMain.Items.Count - 1 == rowCount)
                {
                    dgdMain.CurrentCell = new DataGridCellInfo(dgdMain.Items[rowCount], dgdMain.Columns[colCount + 1]);
                }
                else
                {
                    MessageBox.Show("있으면 찾아보자...");
                }
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
            if (lblMsg.Visibility == Visibility.Visible)
            {
                DataGridCell cell = sender as DataGridCell;
                cell.IsEditing = true;
            }
        }

        //
        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Lib.Instance.CheckIsNumeric(sender as TextBox, e);
        }

        //구분
        private void dgdtpetxtMCGroupName_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                MCEvalBasis = dgdMain.CurrentItem as Win_com_MCEvalBasis_U_CodeView;

                if (MCEvalBasis != null)
                {
                    TextBox tb1 = sender as TextBox;
                    MCEvalBasis.MCGroupName = tb1.Text;
                    sender = tb1;
                }
            }
        }

        //항목명
        private void dgdtpeMCEvalName_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                MCEvalBasis = dgdMain.CurrentItem as Win_com_MCEvalBasis_U_CodeView;

                if (MCEvalBasis != null)
                {
                    TextBox tb1 = sender as TextBox;
                    MCEvalBasis.MCEvalName = tb1.Text;
                    sender = tb1;
                }
            }
        }

        //Spec
        private void dgdtpetxtMCEvalSpec_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                MCEvalBasis = dgdMain.CurrentItem as Win_com_MCEvalBasis_U_CodeView;

                if (MCEvalBasis != null)
                {
                    TextBox tb1 = sender as TextBox;
                    MCEvalBasis.MCEvalSpec = tb1.Text;
                    sender = tb1;
                }
            }
        }

        //EvalSpecMin
        private void dgdtpetxtEvalSpecMin_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                MCEvalBasis = dgdMain.CurrentItem as Win_com_MCEvalBasis_U_CodeView;

                if (MCEvalBasis != null)
                {
                    TextBox tb1 = sender as TextBox;
                    MCEvalBasis.EvalSpecMin = Lib.Instance.returnNumStringTwoExceptDot(tb1.Text);
                    tb1.SelectionStart = tb1.Text.Length;
                    sender = tb1;
                }
            }
        }

        //EvalSpecMax
        private void dgdtpetxtEvalSpecMax_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                MCEvalBasis = dgdMain.CurrentItem as Win_com_MCEvalBasis_U_CodeView;

                if (MCEvalBasis != null)
                {
                    TextBox tb1 = sender as TextBox;
                    MCEvalBasis.EvalSpecMax = Lib.Instance.returnNumStringTwoExceptDot(tb1.Text);
                    tb1.SelectionStart = tb1.Text.Length;
                    sender = tb1;
                }
            }
        }

        //AutoMCEvaluYN
        private void dgdtpetxtAutoMCEvaluYN_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                MCEvalBasis = dgdMain.CurrentItem as Win_com_MCEvalBasis_U_CodeView;

                if (MCEvalBasis != null)
                {
                    TextBox tb1 = sender as TextBox;
                    MCEvalBasis.AutoMCEvaluYN = tb1.Text;
                    sender = tb1;
                }
            }
        }

        //MCEvalScore
        private void dgdtpetxtMCEvalScore_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lblMsg.Visibility == Visibility.Visible)
            {
                MCEvalBasis = dgdMain.CurrentItem as Win_com_MCEvalBasis_U_CodeView;

                if (MCEvalBasis != null)
                {
                    TextBox tb1 = sender as TextBox;
                    MCEvalBasis.MCEvalScore = Lib.Instance.returnNumStringTwoExceptDot(tb1.Text);
                    tb1.SelectionStart = tb1.Text.Length;
                    sender = tb1;
                }
            }
        }
    }

    class Win_com_MCEvalBasis_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        //public int Num { get; set; }
        /// <summary>
        /// pkey
        /// </summary>
        public string MCSeq { get; set; }
        public string MCGroupName { get; set; }
        public string MCEvalName { get; set; }
        public string EvalSpecMin { get; set; }
        public string EvalSpecMax { get; set; }
        public string MCEvalScore { get; set; }
        public string MCEvalSpec { get; set; }
        public string AutoMCEvaluYN { get; set; }
    }
}
