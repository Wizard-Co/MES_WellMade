using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// LoginPage.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Login_PwChange : Window
    {
        private string exPassword = "";

        private Regex regex = new Regex("^(?=.+[A-Za-z])(?=.+\\d)(?=.+[$@$!%*#?&])[A-Za-z\\d$@$!%*#?&]{8,}$");

        public Login_PwChange()
        {
            InitializeComponent();
        }

        public void setPw(string exPassword)
        {
            this.exPassword = exPassword;
        }

        public void setInitChangeMode()
        {
            btnNext.IsEnabled = true;
            txtInfo1.Text = "관리자에 의해 비밀번호가 세팅되었습니다.";
            txtInfo2.Text = "개인정보보호를 위해 비밀번호를 변경해 주세요.";
            txtInfo3_one.Text = "비밀번호 변경 후 시스템을 정상적으로 이용 가능합니다.";
            txtInfo3_two.Text = "";
        }

        public void setChangePwMode()
        {
            btnNext.IsEnabled = true;
            txtInfo1.Text = "비밀번호를 변경하신지 3개월이 지난 경우 아래와 같이 변경 안내를 드리고 있습니다.";
            txtInfo2.Text = "'다음에 변경하기' 버튼을 눌러 변경을 연기하시면 3개월 후에 다시 안내해 드립니다.";
            txtInfo3_one.Text = "조금 불편하시더라도 ";
            txtInfo3_two.Text = "지금 비밀번호를 변경해주세요.";
        }

        // 비밀번호 변경
        private bool goPwChange(bool changeFlag)
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            if (CheckData(changeFlag))
            {
                // [xp_Code_uPassWord]
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("PersonID", MainWindow.CurrentPersonID);
                sqlParameter.Add("PassWord", changeFlag == true ? txtNewPw.Password.Trim() : exPassword);
                sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentPersonID);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_Code_uPassWord";
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

        private bool goPwChangeNext(bool changeFlag)
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            if (CheckDataNext(changeFlag))
            {
                // [xp_Code_uPassWord]
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("PersonID", MainWindow.CurrentPersonID);

                sqlParameter.Add("PassWord", changeFlag == true ? txtPw.Password.Trim() : txtPw.Password.ToString());
                sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentPersonID);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_Code_uPassWord";
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


        private bool CheckData(bool changeFlag)
        {
            #region 2021-08-23 암호화 전의 소스
            //bool flag = true;

            //// 현재 비밀번호를 입력하지 않았을 때
            //if (txtPw.Password.Trim().Equals(""))
            //{
            //    MessageBox.Show("현재 비밀번호를 입력해주세요.");
            //    flag = false;
            //    return flag;
            //}

            //// 현재 비밀번호가 일치하지 않을 때
            //if (!txtPw.Password.Trim().Equals(exPassword))
            //{
            //    MessageBox.Show("현재 비밀번호가 일치하지 않습니다.");
            //    flag = false;
            //    return flag;
            //}

            //if (changeFlag == true) // 다음에 변경은 새 비밀번호 입력안해도 됨여
            //{
            //    // 새 비밀번호를 입력하지 않았을 때
            //    if (txtNewPw.Password.Trim().Equals(""))
            //    {
            //        MessageBox.Show("새 비밀번호를 입력해주세요.");
            //        flag = false;
            //        return flag;
            //    }

            //    // 현재 비밀번호와 새 비밀번호가 일치할 때
            //    if (txtNewPw.Password.Trim().Equals(txtPw.Password.Trim()))
            //    {
            //        MessageBox.Show("새 비밀번호가 현재 비밀번호와 같습니다.\r현재 비밀번호와 다르게 입력해주세요.");
            //        flag = false;
            //        return flag;
            //    }

            //    // 새 비밀번호 확인을 입력하지 않았을 때
            //    if (txtNewPwConfirm.Password.Trim().Equals(""))
            //    {
            //        MessageBox.Show("새 비밀번호 확인 입력해주세요.");
            //        flag = false;
            //        return flag;
            //    }

            //    // 새 비밀번호와 새 비밀번호 확인이 일치하지 않을 때
            //    if (!txtNewPw.Password.Trim().Equals(txtNewPwConfirm.Password.Trim()))
            //    {
            //        MessageBox.Show("새 비밀번호와 새 비밀번호 확인이 일치하지 않습니다.");
            //        flag = false;
            //        return flag;
            //    }


            //}   

            //return flag;
            #endregion

            try
            {
                bool flag = true;

                // 현재 비밀번호를 입력하지 않았을 때
                if (txtPw.Password.Trim().Equals(""))
                {
                    MessageBox.Show("현재 비밀번호를 입력해주세요.");

                    flag = false;
                    return flag;
                }

                //// 현재 비밀번호가 일치하지 않을 때
                //if (!txtPw.Password.Trim().Equals(exPassword))
                //{
                //    MessageBox.Show("현재 비밀번호가 일치하지 않습니다.");
                //    flag = false;
                //    return flag;
                //}
                // 현재 비밀번호가 일치하지 않을때 메서드

                if (changeFlag == true) // 다음에 변경은 새 비밀번호 입력안해도 됨여
                {
                    if (CheckPW(changeFlag) == false)
                    {
                        //MessageBox.Show("현재 비밀번호가 일치하지 않습니다.");
                        flag = false;
                        return flag;
                    }

                    // 새 비밀번호를 입력하지 않았을 때
                    if (txtNewPw.Password.Trim().Equals(""))
                    {
                        MessageBox.Show("새 비밀번호를 입력해주세요.");
                        flag = false;
                        return flag;
                    }

                    // 현재 비밀번호와 새 비밀번호가 일치할 때
                    //if (txtNewPw.Password.Trim().Equals(txtPw.Password.Trim()))
                    //{
                    //    MessageBox.Show("새 비밀번호가 현재 비밀번호와 같습니다.\r현재 비밀번호와 다르게 입력해주세요.");
                    //    flag = false;
                    //    return flag;
                    //}

                    // 새 비밀번호 확인을 입력하지 않았을 때
                    if (txtNewPwConfirm.Password.Trim().Equals(""))
                    {
                        MessageBox.Show("새 비밀번호 확인 입력해주세요.");
                        flag = false;
                        return flag;
                    }

                    // 새 비밀번호와 새 비밀번호 확인이 일치하지 않을 때
                    if (!txtNewPw.Password.Trim().Equals(txtNewPwConfirm.Password.Trim()))
                    {
                        MessageBox.Show("새 비밀번호와 새 비밀번호 확인이 일치하지 않습니다.");
                        flag = false;
                        return flag;
                    }


                }
                else
                {
                    //if (CheckPW(changeFlag) == false)
                    //{
                    //    //MessageBox.Show("현재 비밀번호가 일치하지 않습니다.");
                    //    flag = false;
                    //    return flag;
                    //}
                }

                return flag;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return true;

        }

        private bool CheckDataNext(bool changeFlag)
        {
            bool flag = true;

            // 현재 비밀번호를 입력하지 않았을 때
            if (txtPw.Password.Trim().Equals(""))
            {
                MessageBox.Show("현재 비밀번호를 입력해주세요.");
                flag = false;
                return flag;
            }


            if (changeFlag == true) // 다음에 변경은 새 비밀번호 입력안해도 됨여
            {
                if (CheckPW(changeFlag) == false)
                {
                    //MessageBox.Show("현재 비밀번호가 일치하지 않습니다.");
                    flag = false;
                    return flag;
                }


            }
            else
            {

            }

            return flag;
        }

        // 비밀번호 변경
        private void btnChangePW_Click(object sender, RoutedEventArgs e)
        {
            if (goPwChange(true))
            {
                MessageBox.Show("비밀번호가 성공적으로 변경되었습니다.");
                DialogResult = true;
            }
            else
            {

            }
        }

        // 다음에 변경
        private void btnNext_Click(object sender, RoutedEventArgs e)
        {
            if (goPwChangeNext(true))
            {
                DialogResult = true;
            }
            else
            {

            }
        }

        // 
        private void txtNewPw_KeyUp(object sender, KeyEventArgs e)
        {
            if (txtNewPw.Password.Length > 7)
            {
                // 특수문자 포함해서 적었는지 체크

                if (!regex.IsMatch(txtNewPw.Password.ToString()))
                {
                    tblMsg.Text = "특수문자 2자 이상, 숫자를\r포함해서 8자 이상 입력해주세요.";
                    tblMsg.Foreground = Brushes.Red;
                    tblMsg.Visibility = Visibility.Visible;
                    btnChangePW.IsEnabled = false;
                }
                else
                {
                    if (txtNewPw.Password.Trim().Equals(txtNewPwConfirm.Password.Trim()))
                    {
                        tblMsg.Text = "비밀번호가 일치합니다.";
                        tblMsg.Foreground = Brushes.Blue;
                        tblMsg.Visibility = Visibility.Visible;
                        btnChangePW.IsEnabled = true;
                    }
                    else // 비밀번호가 일치하지 않을때
                    {
                        tblMsg.Text = "비밀번호가 일치하지 않습니다.";
                        tblMsg.Foreground = Brushes.Red;
                        tblMsg.Visibility = Visibility.Visible;
                        btnChangePW.IsEnabled = false;
                    }
                }
            }
            else
            {
                tblMsg.Text = "최소 8자 이상 입력해주세요.";
                tblMsg.Foreground = Brushes.Red;
                tblMsg.Visibility = Visibility.Visible;
            }

        }


        private void txtPw_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtNewPw.Focus();
            }
        }

        private void txtNewPw_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                txtNewPwConfirm.Focus();
            }
        }

        private void txtNewPwConfirm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (goPwChange(true))
                {
                    MessageBox.Show("비밀번호가 성공적으로 변경되었습니다.");
                    DialogResult = true;
                }
            }
        }
        //2021-08-23 현재 비밀번호가 맞는지 체크하는 메서드
        public bool CheckPW(bool flag)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sPersonID", MainWindow.CurrentPersonID);
                sqlParameter.Add("sPassword", txtPw.Password.Trim());
                sqlParameter.Add("nNewPassword", flag == true ? 1 : 0);
                sqlParameter.Add("sNewPassword", txtNewPw.Password.Trim());

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Person_sCheckPassword", sqlParameter, false);
                DataTable dt = ds.Tables[0];
                //DataRow dr = dt.Rows[0];

                //int count = Convert.ToInt32(dr["num"].ToString());

                //// 코드 갯수가 0보다 크다면 false 반환
                //if (count > 0)
                //{
                //    return true;
                //}
                //else
                //{
                //    return false;
                //}

                if (dt.Columns.Count == 1)
                {
                    if (dt.Rows[0]["Result"].ToString().Trim().ToUpper().Equals("TRUE"))
                    {
                        return true;
                    }
                    else
                    {
                        MessageBox.Show(dt.Rows[0].ItemArray[0].ToString());
                        return false;
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return true;
        }
    }
}
