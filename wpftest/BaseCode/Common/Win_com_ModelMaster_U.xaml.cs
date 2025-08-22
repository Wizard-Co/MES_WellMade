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
    /// Win_com_CarModelMaster_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_ModelMaster_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        PlusFinder pf = new PlusFinder();

        string strBasisID = string.Empty;

        string InspectName = string.Empty;

        string AASS = string.Empty;

        string strFlag = string.Empty;
        int rowNum = 0;
        string strFindID = string.Empty;
        string strFindString = string.Empty;
        Win_com_CarModelMaster_U_CodeView winModel = new Win_com_CarModelMaster_U_CodeView();

        public Win_com_ModelMaster_U()
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

        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            grdInput.IsHitTestVisible = false;
            dgdModel.IsHitTestVisible = true;
        }

        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            grdInput.IsHitTestVisible = true;
            dgdModel.IsHitTestVisible = false;
        }

        // 검색 조건 - 모델
        private void lblModelSrh_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkModelSrh.IsChecked == true)
            {
                chkModelSrh.IsChecked = false;
            }
            else
            {
                chkModelSrh.IsChecked = true;
            }
        }
        private void chkModelSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkModelSrh.IsChecked = true;
            txtModelSrh.IsEnabled = true;
        }
        private void chkModelSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkModelSrh.IsChecked = false;
            txtModelSrh.IsEnabled = false;
        }

        //추가
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (dgdModel.SelectedItem != null)
            {
                rowNum = dgdModel.SelectedIndex;
            }
            ClearData();
            CantBtnControl();
            tbkMsg.Text = "자료 입력 중";
            strFlag = "I";
            this.DataContext = null;

            txtName.Focus();
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            winModel = dgdModel.SelectedItem as Win_com_CarModelMaster_U_CodeView;

            if (winModel != null)
            {
                //rowNum = dgdModel.SelectedIndex;
                CantBtnControl();
                tbkMsg.Text = "자료 수정 중";
                strFlag = "U";
            }
            else
            {
                MessageBox.Show("수정할 자료를 선택하고 눌러주십시오.");
            }

            txtName.Focus();
        }

        // 모델 텍스트박스 엔터 → 비고 활성화
        private void txtName_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                txtComments.Focus();
            }
        }

        // 비고 텍스트박스 엔터 → 저장버튼 활성화
        private void txtComments_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                btnSave.Focus();
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            winModel = dgdModel.SelectedItem as Win_com_CarModelMaster_U_CodeView;

            if (winModel == null)
            {
                MessageBox.Show("삭제할 데이터가 지정되지 않았습니다. 삭제데이터를 지정하고 눌러주세요");
            }
            else
            {
                if (MessageBox.Show("선택하신 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    //삭제 전 체크
                    if (!DeleteDataCheck(winModel.ModelID))
                        return;

                    if (dgdModel.Items.Count > 0 && dgdModel.SelectedItem != null)
                    {
                        rowNum = dgdModel.SelectedIndex;
                    }

                    if (DeleteData(winModel.ModelID))
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
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                ClearData();
                rowNum = 0;
                re_Search(rowNum);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        ////저장
        //private void btnSave_Click(object sender, RoutedEventArgs e)
        //{
        //    winModel = dgdModel.SelectedItem as Win_com_CarModelMaster_U_CodeView;

        //    if (SaveData(txtCode.Text, strFlag))
        //    {
        //        CanBtnControl();
        //        strFlag = string.Empty;
        //        re_Search(rowNum);
        //    }
        //}


        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData(txtCode.Text, strFlag))
            {
                CanBtnControl();
                strBasisID = string.Empty;
                lblMsg.Visibility = Visibility.Hidden;

                if (strFlag.Equals("I"))
                {
                    InspectName = txtCode.ToString();
                    //InspectName = txtKCustom.ToString();
                    //InspectDate = dtpInspectDate.SelectedDate.ToString().Substring(0, 10);

                    rowNum = 0;
                    re_Search(rowNum);
                }
                else
                {
                    rowNum = dgdModel.SelectedIndex;
                }
            }

            int i = 0;

            foreach (Win_com_CarModelMaster_U_CodeView WMRIC in dgdModel.Items)
            {

                string a = WMRIC.ModelID.ToString();
                string b = AASS;


                if (a == b)
                {
                    System.Diagnostics.Debug.WriteLine("데이터 같음");

                    break;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("다름");
                }

                i++;
            }

            rowNum = i;
            re_Search(rowNum);
        }


        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            ClearData();
            CanBtnControl();
            strFlag = string.Empty;
            re_Search(rowNum);
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataTable dt = null;
                string Name = string.Empty;

                string[] dgdStr = new string[2];
                dgdStr[0] = "차종 정보";
                dgdStr[1] = dgdModel.Name;

                ExportExcelxaml ExpExc = new ExportExcelxaml(dgdStr);
                ExpExc.ShowDialog();

                if (ExpExc.DialogResult.HasValue)
                {
                    if (ExpExc.choice.Equals(dgdModel.Name))
                    {
                        DataStore.Instance.InsertLogByForm(this.GetType().Name, "E");
                        if (ExpExc.Check.Equals("Y"))
                            dt = Lib.Instance.DataGridToDTinHidden(dgdModel);
                        else
                            dt = Lib.Instance.DataGirdToDataTable(dgdModel);

                        Name = dgdModel.Name;
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
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - 엑셀버튼 클릭 : " + ee.ToString());
            }

        }

        //재검색
        private void re_Search(int selectedIndex)
        {
            FillGrid();

            if (dgdModel.Items.Count > 0)
            {
                if (strFindString.Equals(string.Empty))
                {
                    dgdModel.SelectedIndex = selectedIndex;
                }
                else
                {
                    dgdModel.SelectedIndex = Lib.Instance.ReTrunIndex(dgdModel, strFindString);
                }
            }
            else
            {
                this.DataContext = null;
            }

            strFindID = string.Empty;
            strFindString = string.Empty;
        }

        //검색
        private void FillGrid()
        {
            if (dgdModel.Items.Count > 0)
            {
                dgdModel.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sModelID", "");
                sqlParameter.Add("sModel", chkModelSrh.IsChecked == true ? txtModelSrh.Text : "");
                sqlParameter.Add("sIncNotUse", chkNotUseSrh.IsChecked == true ? "1" : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Model_sModel", sqlParameter, true, "R");

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
                            var dgdModelInfo = new Win_com_CarModelMaster_U_CodeView()
                            {
                                Num = i.ToString(),
                                ModelID = dr["ModelID"].ToString(),
                                Model = dr["Model"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                useclss = dr["UseClss"].ToString()
                            };

                            if (strFindID.Equals(dgdModelInfo.ModelID))
                            {
                                strFindString = dgdModelInfo.ToString();
                            }

                            dgdModel.Items.Add(dgdModelInfo);
                        }

                        tbkIndexCount.Text = "▶ 검색결과 : " + i.ToString() + " 건";
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

        //삭제(라기 보단 사용안함으로 체인지)
        private bool DeleteData(string strModelID)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sModelID", strModelID);
                sqlParameter.Add("DeleteUserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_Model_dModel", sqlParameter, "D");

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

        //저장
        private bool SaveData(string strModelID, string strFlag)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {
                    if (strFlag.Equals("I"))
                    {
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("sNewModelID", strModelID.Replace(" ", ""));
                        sqlParameter.Add("sModel", txtName.Text);
                        sqlParameter.Add("sComments", txtComments.Text);
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Model_iModel";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "sNewModelID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter,"C");
                        string sGetBankID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "sNewModelID")
                                {
                                    sGetBankID = kv.value;
                                    strFindID = sGetBankID;


                                    AASS = kv.value;

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
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("sModelID", strModelID);
                        sqlParameter.Add("sModel", txtName.Text);
                        sqlParameter.Add("sComments", txtComments.Text);
                        sqlParameter.Add("sUseClss", chkNotUse.IsChecked == true ? "*" : "");
                        sqlParameter.Add("UpdateUserID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Model_uModel";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "sModelID";
                        pro1.OutputLength = "5";

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

        //데이터 체크
        private bool CheckData()
        {
            bool flag = true;

            if (txtName.Text.Equals("") || txtName.Text.Length <= 0)
            {
                MessageBox.Show("모델명 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            return flag;
        }
        //삭제체크
        private bool DeleteDataCheck(string strModelID)
        {
            bool Flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sModelID", strModelID);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Model_dModel_Check", sqlParameter, false);

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
        private void dgdModel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgdModel.SelectedValue != null)
            {
                winModel = dgdModel.SelectedItem as Win_com_CarModelMaster_U_CodeView;

                if (winModel.useclss.Equals("*"))
                {
                    chkNotUse.IsChecked = true;
                }
                else
                {
                    chkNotUse.IsChecked = false;
                }

                this.DataContext = winModel;
            }
        }

        private void lblNotUseSrh_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (chkNotUseSrh.IsChecked == true)
            {
                chkNotUseSrh.IsChecked = false;
            }
            else
            {
                chkNotUseSrh.IsChecked = true;
            }
        }

        //모델 검색 텍스트박스 키다운 이벤트
        private void txtModelSrh_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                pf.ReturnCode(txtModelSrh, 88, txtModelSrh.Text);
            }
        }

        private void ClearData()
        {
            txtCode.Clear();
            txtName.Clear();
            txtComments.Clear();
            chkNotUse.IsChecked = false;
        }



    }

    class Win_com_CarModelMaster_U_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Num { get; set; }
        public string ModelID { get; set; }
        public string Model { get; set; }
        public string Comments { get; set; }
        public string useclss { get; set; }
    }
}
