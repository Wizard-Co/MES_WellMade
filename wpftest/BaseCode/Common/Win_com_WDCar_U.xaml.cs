using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WizMes_WellMade.PopUP;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_com_WDCar_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_WDCar_U : UserControl
    {
        string strFlag = string.Empty;
        int rowNum = 0;
        Win_com_WDCar_U_CodeView WinWDCar = new Win_com_WDCar_U_CodeView();
        Lib lib = new Lib();
        public Win_com_WDCar_U()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            Lib.Instance.UiLoading(sender);
        }

        //사용안함 포함
        private void lblUseClssSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkNotUseClssSrh.IsChecked == true) { chkNotUseClssSrh.IsChecked = false; }
            else { chkNotUseClssSrh.IsChecked = true; }
        }

        /// <summary>
        /// 수정,추가 저장 후
        /// </summary>
        private void CanBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_IUControl(this);
            grdInput.IsHitTestVisible = false;
            dgdMain.IsEnabled = true;
        }

        /// <summary>
        /// 수정,추가 진행 중
        /// </summary>
        private void CantBtnControl()
        {
            Lib.Instance.UiButtonEnableChange_SCControl(this);
            grdInput.IsHitTestVisible = true;
            dgdMain.IsEnabled = false;
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
        }

        //수정
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            WinWDCar = dgdMain.SelectedItem as Win_com_WDCar_U_CodeView;

            if (WinWDCar != null)
            {
                rowNum = dgdMain.SelectedIndex;
                tbkMsg.Text = "자료 수정 중";
                lblMsg.Visibility = Visibility.Visible;
                CantBtnControl();
                strFlag = "U";
            }
        }

        //삭제
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            WinWDCar = dgdMain.SelectedItem as Win_com_WDCar_U_CodeView;

            if (WinWDCar == null)
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

                    if (DeleteData(WinWDCar.WDID))
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
            if (SaveData(strFlag, txtWDID.Text))
            {
                CanBtnControl();
                lblMsg.Visibility = Visibility.Hidden;
                rowNum = 0;
                dgdMain.IsEnabled = true;

                re_Search(rowNum);
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            CanBtnControl();
            strFlag = string.Empty;
            dgdMain.IsEnabled = true;

            re_Search(rowNum);
        }

        //엑셀
        private void btnExcel_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = null;
            string Name = string.Empty;

            string[] lst = new string[2];
            lst[0] = "대차코드";
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

        //
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
                sqlParameter.Add("WDNo", txtWDIDSrh.Text);
                sqlParameter.Add("UseClss", chkNotUseClssSrh.IsChecked == true ? 1 : 0);
                ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sWDCar", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        //MessageBox.Show("조회된 데이터가 없습니다.");
                        this.DataContext = null;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var WinWd = new Win_com_WDCar_U_CodeView()
                            {
                                Num = i + 1,
                                WDID = dr["WDID"].ToString(),
                                WDNo = dr["WDNo"].ToString(),
                                WDQty = dr["WDQty"].ToString(),
                                Comments = dr["Comments"].ToString(),
                                UseClss = dr["UseClss"].ToString()
                            };

                            dgdMain.Items.Add(WinWd);
                            i++;
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

        //
        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            WinWDCar = dgdMain.SelectedItem as Win_com_WDCar_U_CodeView;

            if (WinWDCar != null)
            {
                if (WinWDCar.UseClss.Replace(" ", "").Equals(string.Empty))
                {
                    chkNotUseClss.IsChecked = false;
                }
                else
                {
                    chkNotUseClss.IsChecked = true;
                }

                this.DataContext = WinWDCar;
            }
        }

        /// <summary>
        /// 실삭제
        /// </summary>
        /// <param name="strID"></param>
        /// <returns></returns>
        private bool DeleteData(string strID)
        {
            bool flag = false;

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("WDID", strID);

            string[] result = DataStore.Instance.ExecuteProcedure("xp_Code_dWDCar", sqlParameter, false);
            DataStore.Instance.CloseConnection();

            if (result[0].Equals("success"))
            {
                //MessageBox.Show("성공 *^^*");
                flag = true;
            }

            return flag;
        }

        /// <summary>
        /// 실저장
        /// </summary>
        /// <param name="strFlag"></param>
        /// <param name="strArticleID"></param>
        /// <returns></returns>
        private bool SaveData(string strFlag, string strID)
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
                    sqlParameter.Add("WDID", strID);
                    sqlParameter.Add("WDNo", txtWDNo.Text);
                    sqlParameter.Add("WDQty", txtWDQty.Text);
                    sqlParameter.Add("Commtents", txtComments.Text);
                    sqlParameter.Add("UseClss", chkNotUseClss.IsChecked == true ? "*" : "");
                    sqlParameter.Add("sUserID", MainWindow.CurrentUser);

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Code_iWDCar";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "WDID";
                        pro1.OutputLength = "5";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "WDID")
                                {
                                    sGetID = kv.value;
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
                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_Code_uWDCar";
                        pro1.OutputUseYN = "N";
                        pro1.OutputName = "WDID";
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

                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return flag;
        }

        /// <summary>
        /// 입력 데이터 체크
        /// </summary>
        /// <returns></returns>
        private bool CheckData()
        {
            bool flag = true;

            if (txtWDNo.Text.Length <= 0 || txtWDNo.Text.Equals(""))
            {
                MessageBox.Show("대차번호가 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            if (txtWDQty.Text.Length == -1 || txtWDQty.Text.Equals(""))
            {
                MessageBox.Show("장입량이 입력되지 않았습니다.");
                flag = false;
                return flag;
            }

            return flag;
        }


    }

    class Win_com_WDCar_U_CodeView : BaseView
    {
        public int Num { get; set; }
        public string WDID { get; set; }
        public string WDNo { get; set; }
        public string WDQty { get; set; }
        public string UseClss { get; set; }
        public string Comments { get; set; }
    }
}
