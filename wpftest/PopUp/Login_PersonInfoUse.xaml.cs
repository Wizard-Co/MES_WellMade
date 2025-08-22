using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// LoginPage.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Login_PersonInfoUse : Window
    {

        //private string userID = "";
        private string name = "";

        public Login_PersonInfoUse()
        {
            InitializeComponent();
        }

        public void setData(string name)
        {
            //this.userID = userID;
            this.name = name;

            txtUserID.Text = MainWindow.CurrentUser;
            txtUserName.Text = this.name;
        }

        // 개인정보활용 동의 메서드
        private bool PersonInfoUse_Y()
        {
            bool flag = false;


            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            if (CheckData())
            {
                // [xp_Code_uPassWord]
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("PersonID", MainWindow.CurrentPersonID);
                sqlParameter.Add("NessaryAcptYN", chkAccessControl.IsChecked == true ? "Y" : "N");
                sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentPersonID);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_Person_uPersonInfoUse";
                pro1.OutputUseYN = "N";
                pro1.OutputName = "REQ_ID";
                pro1.OutputLength = "10";

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
                    //MessageBox.Show("성공");
                    flag = true;
                }
            }


            return flag;
        }

        // 유효성 검사
        private bool CheckData()
        {
            bool flag = true;

            // 사용자 ID 가 일치하지 않을 때
            if (!MainWindow.CurrentUser.Trim().Equals(txtUserID.Text.Trim()))
            {
                MessageBox.Show("사용자 ID가 일치하지 않습니다.");
                flag = false;
                return flag;
            }

            // 이름이 일치하지 않을 때
            if (!name.Trim().Equals(txtUserName.Text.Trim()))
            {
                MessageBox.Show("이름이 일치하지 않습니다.");
                flag = false;
                return flag;
            }

            // 개인정보활용 동의 체크하지 않았을 때 chkAccessControl
            if (chkAccessControl.IsChecked == false)
            {
                MessageBox.Show("개인정보활용 동의가 체크되지 않았습니다.");
                flag = false;
                return flag;
            }

            return flag;
        }

        // 확인 버튼 이벤트
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (PersonInfoUse_Y())
            {
                //MessageBox.Show("");
                DialogResult = true;
            }
            else
            {

            }
        }

        // 취소 버튼 이벤트
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {

            DialogResult = false;
        }

        // 개인정보활용동의 체크 이벤트
        private void tblAccessControl_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkAccessControl.IsChecked == true)
            {
                chkAccessControl.IsChecked = false;
            }
            else
            {
                chkAccessControl.IsChecked = true;
            }
        }
        private void chkAccessControl_Checked(object sender, RoutedEventArgs e)
        {
            chkAccessControl.IsChecked = true;
            btnSave.IsEnabled = true;
        }
        private void chkAccessControl_Unchecked(object sender, RoutedEventArgs e)
        {
            chkAccessControl.IsChecked = false;
            btnSave.IsEnabled = false;
        }

    }
}
