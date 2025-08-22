using System;
using System.Collections.Generic;
using System.Windows;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// Win_comMCRunningTarget_PopUP.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_comMCRunningTarget_PopUP : Window
    {
        public Win_comMCRunningTarget_PopUP()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dtpSDate.SelectedDate = DateTime.Today.AddYears(-1);
            dtpEDate.SelectedDate = DateTime.Today;
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            if (CopyDataCheck())
            {
                if (!CopyData())
                {
                    MessageBox.Show("복사등록이 실패하였습니다.");
                    DialogResult = false;
                }
                else
                {
                    DialogResult = true;
                }
            }

        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
        private bool CopyDataCheck()
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("cYYYY", dtpSDate.SelectedDate.Value.ToString("yyyy"));
                sqlParameter.Add("vYYYY", dtpEDate.SelectedDate.Value.ToString("yyyy"));

                string[] result = DataStore.Instance.ExecuteProcedure("xp_MachineGoal_iCopyMachineGoal_Check", sqlParameter, false);
                string[] resultSplit;

                if (result[0].Equals("success") && result[1].Equals(""))
                {
                    flag = true;
                }
                else
                {
                    resultSplit = result[1].Split('/');

                    if (resultSplit.Length == 2)
                    {
                        if (Convert.ToInt32(resultSplit[0]) == 20)
                        {
                            if (MessageBox.Show(resultSplit[1] + " \n그래도 계속 복사 등록작업을 진행하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                            {
                                flag = true;
                            }
                        }
                        else
                        {
                            MessageBox.Show(resultSplit[1]);
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
        private bool CopyData()
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                #region 추가

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("cYYYY", dtpSDate.SelectedDate.Value.ToString("yyyy"));
                sqlParameter.Add("vYYYY", dtpEDate.SelectedDate.Value.ToString("yyyy"));
                sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_MachineGoal_iCopyMachineGoal";
                pro1.OutputUseYN = "N";
                pro1.OutputName = "sArticleID";
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

                #endregion
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
    }
}
