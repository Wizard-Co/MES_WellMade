using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_com_Bank_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_Bank_U : UserControl
    {
        #region 변수 선언 및 로드

        string strFlag = string.Empty;
        int rowNum = 0;
        string strFindID = string.Empty;
        string strFindString = string.Empty;
        Win_com_Bank_U_CodeView winBank = new Win_com_Bank_U_CodeView();
        Lib lib = new Lib();

        public Win_com_Bank_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
        }

        #endregion

        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            gbxInput.IsHitTestVisible = false;
            dgdBank.IsHitTestVisible = true;
        }

        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            gbxInput.IsHitTestVisible = true;
            dgdBank.IsHitTestVisible = false;
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (dgdBank.SelectedItem != null)
            {
                //취소대비
                rowNum = dgdBank.SelectedIndex;
            }

            CantBtnControl();
            tbkMsg.Text = "자료 입력 중";
            strFlag = "I";
            this.DataContext = null;
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            winBank = dgdBank.SelectedItem as Win_com_Bank_U_CodeView;

            if (winBank != null)
            {
                rowNum = dgdBank.SelectedIndex;
                CantBtnControl();
                tbkMsg.Text = "자료 수정 중";
                strFlag = "U";
            }
            else
            {
                MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            winBank = dgdBank.SelectedItem as Win_com_Bank_U_CodeView;

            if (winBank == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (dgdBank.Items.Count > 0 && dgdBank.SelectedItem != null)
                    {
                        rowNum = dgdBank.SelectedIndex;
                    }

                    if (Procedure.Instance.DeleteData(winBank.BankID, "BankID", "xp_CodeBank_dBank"))
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
            Lib.Instance.ChildMenuClose(this.ToString());
        }

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

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
            winBank = dgdBank.SelectedItem as Win_com_Bank_U_CodeView;

            if (SaveData(txtBankID.Text, strFlag))
            {
                CanBtnControl();
                re_Search(rowNum);
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();
            strFlag = string.Empty;
            re_Search(rowNum);
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] dgdStr = new string[2];
            dgdStr[0] = "은행 정보";
            dgdStr[1] = dgdBank.Name;

            ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
            ExpExc.ShowDialog();

            if (ExpExc.DialogResult.HasValue)
            {
                if (ExpExc.choice.Equals(dgdBank.Name))
                {
                    if (ExpExc.Check.Equals("Y"))
                        dt = Lib.Instance.DataGridToDTinHidden(dgdBank);
                    else
                        dt = Lib.Instance.DataGirdToDataTable(dgdBank);

                    Name = dgdBank.Name;
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

        //재검색
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdBank.Items.Count > 0)
            {
                if (strFindString.Equals(string.Empty))
                {
                    dgdBank.SelectedIndex = selectedIndex;
                }
                else
                {
                    dgdBank.SelectedIndex = Lib.Instance.ReTrunIndex(dgdBank, strFindString);
                }
            }
            else
            {
                this.DataContext = null;
            }

            strFindID = string.Empty;
            strFindString = string.Empty;
        }

        //조회
        private void FillGrid()
        {
            if (dgdBank.Items.Count > 0)
            {
                dgdBank.Items.Clear();
            }

            try
            {
                DataGrid test = new DataGrid();

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("UseYN", chkIncDelete.IsChecked == true ? "" : "Y");
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_CodeBank_sBank", sqlParameter, false);

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
                            var dgdBankInfo = new Win_com_Bank_U_CodeView()
                            {
                                BankID = dr["BankID"].ToString(),
                                BankName = dr["BankName"].ToString(),
                                BankNameEng = dr["BankNameEng"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                Use_YN = dr["Use_YN"].ToString(),
                                Create_Date = dr["Create_Date"].ToString(),
                                Create_User_ID = dr["Create_User_ID"].ToString(),
                                Update_Date = dr["Update_Date"].ToString(),
                                Update_User_ID = dr["Update_User_ID"].ToString()
                            };

                            if (!strFindID.Equals(string.Empty))
                            {
                                if (strFindID.Equals(dgdBankInfo.BankID))
                                {
                                    strFindString = dgdBankInfo.ToString();
                                }
                            }

                            dgdBank.Items.Add(dgdBankInfo);
                        }

                        tbkIndexCount.Text = "▶검색결과 : " + i.ToString() + " 건";
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

        //저장
        private bool SaveData(string strBankID, string strFlag)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Add("BankID", strBankID);
                    sqlParameter.Add("BankName", txtBankName.Text);
                    sqlParameter.Add("BankNameEng", txtBankNameEng.Text);
                    sqlParameter.Add("Use_YN", chkNotUse.IsChecked == true ? "N" : "Y");
                    sqlParameter.Add("Comments", txtComments.Text);

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_CodeBank_iBank";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "BankID";
                        pro1.OutputLength = "5";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        string sGetBankID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "BankID")
                                {
                                    sGetBankID = kv.value;
                                    strFindID = sGetBankID;
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
                    else
                    {
                        sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_CodeBank_uBank";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "BankID";
                        pro1.OutputLength = "5";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

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
                }
                else { flag = false; }
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

        //데이터 체크
        private bool CheckData()
        {
            bool flag = true;

            if (txtBankName.Text.Length <= 0 || txtBankName.Text.Equals(""))
            {
                MessageBox.Show("은행명 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            return flag;
        }

        private void dgdBank_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            winBank = dgdBank.SelectedItem as Win_com_Bank_U_CodeView;
            this.DataContext = winBank;
        }
    }

    class Win_com_Bank_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string BankID { get; set; }
        public string BankName { get; set; }
        public string BankNameEng { get; set; }
        public string Comments { get; set; }
        public string Use_YN { get; set; }
        public string Create_Date { get; set; }
        public string Create_User_ID { get; set; }
        public string Update_Date { get; set; }
        public string Update_User_ID { get; set; }
    }
}
