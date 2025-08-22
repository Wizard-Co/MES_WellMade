using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;

namespace WizMes_WellMade.PopUp
{
    public partial class Stuffin_Inspect : Window
    {
        List<string> lstStuffinID = new List<string>();

        public Stuffin_Inspect(List<string> lstStuffinID)
        {
            InitializeComponent();

            this.lstStuffinID = lstStuffinID;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            txtInspector.Text = MainWindow.CurrentPerson;
            dtpInspectDate.SelectedDate = DateTime.Today;

            SetComboBox();
        }

        // 콤보박스 세팅
        private void SetComboBox()
        {
            List<string[]> strValueYN = new List<string[]>();
            strValueYN.Add(new string[] { "Y", "Y" });
            //strValueYN.Add(new string[] { "N", "N" });

            ObservableCollection<CodeView> ovcYN = ComboBoxUtil.Instance.Direct_SetComboBox(strValueYN);
            this.cboInspectApprovalYN.ItemsSource = ovcYN;
            this.cboInspectApprovalYN.DisplayMemberPath = "code_name";
            this.cboInspectApprovalYN.SelectedValuePath = "code_id";

            this.cboInspectApprovalYN.SelectedIndex = 0;
        }

        public void setStuffinID(List<string> lstStuffinID)
        {
            this.lstStuffinID = lstStuffinID;
        }

        // 저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (lstStuffinID.Count > 0)
            {
                if (CheckData())
                {
                    if (MessageBox.Show("선택한 항목들을 입고검수처리 하시겠습니까?", "입고검수승인 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        if (UpdateStuffinInspect())
                        {
                            DialogResult = true;
                        }
                        else
                        {
                            DialogResult = false;
                        }
                    }
                }
            }
        }

        // 입고검수처리 메서드
        private bool UpdateStuffinInspect()
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                for (int i = 0; i < lstStuffinID.Count; i++)
                {
                    sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("StuffInID", lstStuffinID[i]);
                    sqlParameter.Add("StuffInSubSeq", 1);
                    sqlParameter.Add("sInspector", txtInspector.Text);
                    sqlParameter.Add("sInspectDate", dtpInspectDate.SelectedDate.Value.ToString("yyyyMMdd"));
                    sqlParameter.Add("sInspectApprovalYN", cboInspectApprovalYN.SelectedValue != null ? cboInspectApprovalYN.SelectedValue.ToString() : "N");
                    sqlParameter.Add("sUserID", MainWindow.CurrentUser);
                    sqlParameter.Add("sInspector1", txtInspector1.Text);

                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_StuffIN_uStuffINSub_Inspect";
                    pro1.OutputUseYN = "N";
                    pro1.OutputName = "StuffInID";
                    pro1.OutputLength = "12";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);
                }

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                    flag = false;
                }
                else
                {
                    flag = true;
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

            return flag;
        }


        // 취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private bool CheckData()
        {
            bool flag = true;

            // 검수자, 검수일자, 검사자, 승인여부
            // 검수자
            if (txtInspector.Text.Trim().Equals(""))
            {
                MessageBox.Show("검수자가 입력되지 않았습니다.");
                flag = false;
                return flag;
            }
            // 검수일자
            if (dtpInspectDate.SelectedDate == null)
            {
                MessageBox.Show("검수일자가 선택되지 않았습니다.");
                flag = false;
                return flag;
            }
            // 검사자
            if (txtInspector1.Text.Trim().Equals(""))
            {
                MessageBox.Show("검사자가 입력되지 않았습니다.");
                flag = false;
                return flag;
            }
            // 승인여부
            if (cboInspectApprovalYN.SelectedValue == null
                || cboInspectApprovalYN.SelectedValue.ToString().Trim().Equals(""))
            {
                MessageBox.Show("승인여부를 선택해주세요.");
                flag = false;
                return flag;
            }

            return flag;
        }
    }
}
