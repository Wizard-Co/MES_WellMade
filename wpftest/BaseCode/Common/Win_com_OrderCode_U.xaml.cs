using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_com_OrderCode_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_OrderCode_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        string strFlag = string.Empty;
        int rowNum = 0;
        Win_com_OrderCode_U_CodeView winWork = new Win_com_OrderCode_U_CodeView();
        string CD_WORK = "Work";

        public Win_com_OrderCode_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);
        }

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            // 수정 후, 수정 취소시 다시 활성화
            //txtWorkID.IsReadOnly = false;
            //txtWorkID.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#fff2d2");

            // 코드 설명 초기화
            infoWorkID.Text = "";

            gbxInput.IsHitTestVisible = false;
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            dgdWork.IsHitTestVisible = true;
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            gbxInput.IsHitTestVisible = true;
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            dgdWork.IsHitTestVisible = false;
        }

        // 검색조건 - 가공구분명
        private void lblWorkSrh_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkWorkSrh.IsChecked == true)
            {
                chkWorkSrh.IsChecked = false;
            }
            else
            {
                chkWorkSrh.IsChecked = true;
            }
        }
        private void chkWorkSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkWorkSrh.IsChecked = true;
            txtWorkSrh.IsEnabled = true;
        }
        private void chkWorkSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkWorkSrh.IsChecked = false;
            txtWorkSrh.IsEnabled = false;
        }

        // 검색시 - 사용불가 포함 체크박스
        private void chkNoUseClss_Checked(object sender, RoutedEventArgs e)
        {
            chkNoUseClss.IsChecked = true;
        }
        // 검색시 - 사용불가 포함 체크박스
        private void chkNoUseClss_UnChecked(object sender, RoutedEventArgs e)
        {
            chkNoUseClss.IsChecked = false;
        }

        //// 사용불가 여부 체크박스
        //private void chkUseClss_Checked(object sender, RoutedEventArgs e)
        //{
        //    chkUseClss.IsChecked = true;
        //}
        //// 사용불가 여부 체크박스
        //private void chkUseClss_UnChecked(object sender, RoutedEventArgs e)
        //{
        //    chkUseClss.IsChecked = false;
        //}

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            CantBtnControl();
            //txtWorkID.IsReadOnly = false;   //추가시에는 입력가능하게
            strFlag = "I";

            lblMsg.Visibility = Visibility.Visible;
            tbkMsg.Text = "자료 입력 중";
            rowNum = dgdWork.SelectedIndex;
            this.DataContext = null;

            chkUseClss.IsChecked = false;

            // 코드 자릿수 (4자리 이하의 숫자) 설명
            //infoWorkID.Text = "* 코드는 4자리 이하의 숫자를 입력해주세요.";

            txtWorkName.Focus();
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            winWork = dgdWork.SelectedItem as Win_com_OrderCode_U_CodeView;

            if (winWork != null)
            {
                rowNum = dgdWork.SelectedIndex;
                dgdWork.IsEnabled = false;
                tbkMsg.Text = "자료 수정 중";
                lblMsg.Visibility = Visibility.Visible;
                CantBtnControl();
                strFlag = "U";

                // 수정시에는 사용못하게
                //txtWorkID.IsReadOnly = true;   
                //txtWorkID.Background = (SolidColorBrush)new BrushConverter().ConvertFrom("#c2fdc3");

                // 코드 자릿수 (4자리 이하의 숫자) 설명
                //infoWorkID.Text = "* 코드는 수정이 불가능합니다.";
            }

            txtWorkName.Focus();
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            winWork = dgdWork.SelectedItem as Win_com_OrderCode_U_CodeView;

            if (winWork == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    //삭제 전 체크
                    if (!DeleteDataCheck(winWork.WorkID))
                        return;

                    if (dgdWork.Items.Count > 0 && dgdWork.SelectedItem != null)
                    {
                        rowNum = dgdWork.SelectedIndex;
                    }

                    if (DeleteData(winWork.WorkID))
                    {
                        rowNum -= 1;
                        re_Search(rowNum);
                    }
                }
            }
        }

        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            rowNum = 0;
            re_Search(rowNum);
        }

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            //winWork = dgdWork.SelectedItem as Win_com_OrderCode_U_CodeView;

            if (SaveData(strFlag))
            {
                CanBtnControl();
                lblMsg.Visibility = Visibility.Hidden;
                rowNum = 0;
                dgdWork.IsEnabled = true;
                re_Search(rowNum);
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();
            strFlag = string.Empty;
            dgdWork.IsEnabled = true;
            re_Search(rowNum);
        }

        /// <summary>
        /// 엑셀로 보내기
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] dgdStr = new string[2];
            dgdStr[0] = "가공구분";
            dgdStr[1] = dgdWork.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdWork.Name))
                {
                    DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdWork);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdWork);

                    Name = dgdWork.Name;
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

        /// <summary>
        /// 재검색(수정,삭제,추가 저장후에 자동 재검색)
        /// </summary>
        /// <param name="selectedIndex"></param>
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdWork.Items.Count > 0)
            {
                dgdWork.SelectedIndex = selectedIndex;
            }
        }

        /// <summary>
        /// 실조회
        /// </summary>
        private void FillGrid()
        {
            if (dgdWork.Items.Count > 0)
            {
                dgdWork.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                // 부서
                sqlParameter.Clear();
                sqlParameter.Add("WorkName", chkWorkSrh.IsChecked == true && !txtWorkSrh.Text.Trim().Equals("") ? txtWorkSrh.Text : "");
                sqlParameter.Add("sUseClss", chkNoUseClss.IsChecked == true ? 1 : 0);

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Code_sWork", sqlParameter, true, "R");
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WorkInfo = new Win_com_OrderCode_U_CodeView()
                            {
                                Num = i,
                                WorkID = dr["WorkID"].ToString(),
                                WorkName = dr["WorkName"].ToString(),
                                UseClss = dr["UseClss"].ToString()
                            };

                            dgdWork.Items.Add(WorkInfo);
                        }

                        // 2019.08.28 검색결과에 갯수 추가
                        //tbkBuseoCount.Text = "▶검색 결과 : " + i + "건";
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
        /// 메인 그리드 로우 선택 시
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgdWork_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var winWorkOrder = dgdWork.SelectedItem as Win_com_OrderCode_U_CodeView;

            if (winWorkOrder != null)
            {
                winWork.WorkID = winWorkOrder.WorkID;
                winWork.WorkName = winWorkOrder.WorkName;
                winWork.UseClss = winWorkOrder.UseClss.Trim();

                this.DataContext = winWork;

                if (winWorkOrder.UseClss.Trim().Equals(""))
                {
                    chkUseClss.IsChecked = false;
                }
                else
                {
                    chkUseClss.IsChecked = true;
                }
            }

            this.DataContext = winWork;
        }

        //삭제체크
        private bool DeleteDataCheck(string strWorkID)
        {
            bool Flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sWorkID", strWorkID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Code_dCodeWork_Check", sqlParameter, false);

                if (result[0].Equals("success") && result[1].Equals(""))
                {
                    //MessageBox.Show("성공 *^^*");
                    Flag = true;
                }
                else
                {
                    MessageBox.Show(result[1]);
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

            return Flag;
        }
        /// <summary>
        /// 실삭제
        /// </summary>
        /// <param name="strWorkID"></param>
        /// <returns></returns>
        private bool DeleteData(string strWorkID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sTableName", CD_WORK);
                sqlParameter.Add("sCodeID", CD_WORK + "ID");
                sqlParameter.Add("sID", strWorkID);
                sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Code_uDeleteCode", sqlParameter, "D");
                DataStore.Instance.CloseConnection();

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
        /// 실저장
        /// </summary>
        /// <param name="strFlag"></param>
        /// <param name="strWorID"></param>
        /// <param name="strWorkName"></param>
        /// <returns></returns>
        private bool SaveData(string strFlag)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sID", txtWorkID.Text);
                    sqlParameter.Add("sData", txtWorkName.Text);
                    sqlParameter.Add("UseClss", chkUseClss.IsChecked == true ? "*" : "");

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Code_iCodeWork";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        string[] Confirm = new string[2];
                        Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"C");
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

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {
                        sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Code_uCodeWork";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

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
        /// 데이터 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            // 코드 숫자만 입력되도록 유효성 검사 → 코드 자동 생성으로 변경 : 이건 사용안함
            //int chkNum = 0;
            //if (Int32.TryParse(txtWorkID.Text, out chkNum) == false)
            //{
            //    MessageBox.Show("코드는 숫자만 입력 가능합니다.");
            //    flag = false;
            //    return flag;
            //}

            //// 코드 체크 (4자리의 자릿수로 입력되도록)
            //if (txtWorkID.Text.Length > 4)
            //{
            //    MessageBox.Show("코드 자릿수를 초과하셨습니다.");
            //    flag = false;
            //    return flag;
            //}
            //else
            //{
            //    // 입력되지 않았을 시
            //    if (txtWorkID.Text.Length <= 0 || txtWorkID.Text.Equals(""))
            //    {
            //        MessageBox.Show("코드가 입력되지 않았습니다.");
            //        flag = false;
            //        return flag;
            //    }

            //    // 만약 입력한 숫자가 4자리 미만이라면 앞에 0을 추가
            //    if (txtWorkID.Text.Length < 4)
            //    {
            //        for (int i = txtWorkID.Text.Length; i < 4; i++)
            //        {
            //            txtWorkID.Text = txtWorkID.Text.Insert(0, "0");
            //        }
            //    }
            //}

            // 코드 중복체크 → 코드 자동 생성으로 변경 : 이건 사용안함
            //if (strFlag.Equals("I") && CheckCode(CD_WORK, txtWorkID.Text) == false)
            //{
            //    MessageBox.Show("입력하신 코드가 이미 존재합니다.");
            //    flag = false;
            //    return flag;
            //}

            if (txtWorkName.Text.Length <= 0 || txtWorkName.Text.Equals(""))
            {
                MessageBox.Show("가공구분이 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            return flag;
        }

        // 추가시 코드 중복 체크
        public bool CheckCode(string tableName, string code)
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("TableName", "mt_" + tableName);
            sqlParameter.Add("sCodeID", tableName + "ID");
            sqlParameter.Add("sID", code);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sCheckDepartID", sqlParameter, false);
            DataTable dt = ds.Tables[0];
            DataRow dr = dt.Rows[0];

            int count = Convert.ToInt32(dr["num"].ToString());

            // 코드 갯수가 0보다 크다면 false 반환
            if (count > 0) { return false; }

            return true;
        }

        private void txtWorkName_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                btnSave.Focus();
            }
        }

        private void lblNoUseClss_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkNoUseClss.IsChecked == true)
            {
                chkNoUseClss.IsChecked = false;
            }
            else
            {
                chkNoUseClss.IsChecked = true;
            }
        }

        //가공구분 텍스트박스 엔터키 이벤트
        private void txtWorkSrh_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                pf.ReturnCode(txtWorkSrh, 89, txtWorkSrh.Text);
            }
        }
    }

    class Win_com_OrderCode_U_CodeView : BaseView
    {
        public int Num { get; set; }
        public string WorkID { get; set; }
        public string WorkName { get; set; }
        public string UseClss { get; set; }
    }
}
