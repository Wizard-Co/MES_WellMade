using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Input;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// LoginPage.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class LoginPage : Window
    {
        public string strLogRegID = string.Empty;
        public string strUserName = string.Empty;

        public string PersonID = "";
        public string exPassword = "";
        public int dayDiff = 0;
        public string AccessControl = "";
        public string UserName = "";

        public string initChange = "Y";

        int fail = 0;

        string stDate = string.Empty;
        string stTime = string.Empty;

        public LoginPage()
        {
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");
            GetInfo();
        }

        //로그인
        private void btnLogin_Click(object sender, RoutedEventArgs e)
        {
            if (Log(txtUserID.Text))
            {
                strLogRegID = txtUserID.Text;
                Lib.Instance.SetLogResitry(strLogRegID);
                DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
                DialogResult = true;
            }
            else
            {
                txtPassWd.Password = "";
                return;
            }
        }

        //취소
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            DialogResult = false;
        }

        private bool Log(string strID)
        {
            #region 20210823 암호화 이전의 소스
            //bool flag = true;

            //DataSet ds = null;
            //Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            //sqlParameter.Clear();
            //sqlParameter.Add("UserID", strID);
            //ds = DataStore.Instance.ProcedureToDataSet("xp_Common_Login", sqlParameter, false);


            //if (ds != null && ds.Tables.Count > 0)
            //{
            //    DataTable dt = ds.Tables[0];

            //    if (dt.Rows.Count <= 0)
            //    {
            //        MessageBox.Show("존재하지 않는 ID 입니다.");
            //        flag = false;
            //        return flag;
            //    }
            //    else
            //    {
            //        if (!dt.Rows[0]["Password"].ToString().Equals(txtPassWd.Password))
            //        {
            //            MessageBox.Show("비밀번호가 잘못되었습니다.");
            //            flag = false;
            //            return flag;
            //        }

            //        //if (!dt.Rows[0]["Name"].Equals("20150401") && !dt.Rows[0]["Name"].Equals("admin"))
            //        //{
            //        //    MessageBox.Show("권한이 없는 사용자입니다.");
            //        //    return flag;
            //        //}

            //        // 비밀번호 변경 추가
            //        if (CheckConvertDateTime(dt.Rows[0]["PasswordChangeDate"].ToString()) == true)
            //        {
            //            DateTime setDate = DateTime.Parse(dt.Rows[0]["PasswordChangeDate"].ToString());

            //            TimeSpan timeDiff = DateTime.Today - setDate;
            //            dayDiff = timeDiff.Days;

            //            if (dayDiff > 90)
            //            {
            //                exPassword = dt.Rows[0]["Password"].ToString();
            //            }
            //        }
            //        else if (dt.Rows[0]["PasswordChangeDate"].ToString().Trim().Equals("")) // 초기 비밀번호가 세팅되지 않았다면
            //        {
            //            initChange = "N";
            //            exPassword = dt.Rows[0]["Password"].ToString();
            //        }

            //        // 개인정보활용 동의 여부 추가, 이름도 추가
            //        AccessControl = dt.Rows[0]["NessaryAcptYN"].ToString();
            //        UserName = dt.Rows[0]["Name"].ToString();

            //        //PersonID
            //        PersonID = dt.Rows[0]["PersonID"].ToString();
            //    }
            //}

            //return flag;
            #endregion

            bool flag = true;

            DataSet ds = null;
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("UserID", strID);
            sqlParameter.Add("Password", txtPassWd.Password);
            ds = DataStore.Instance.ProcedureToDataSet("xp_Common_Login", sqlParameter, false);


            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];

                if (dt.Rows.Count > 0)
                {
                    if (dt.Columns.Count == 1)
                    {
                        fail++;

                        if (dt.Rows[0].ItemArray[0].Equals("계정이 차단되었습니다. 관리자에게 문의해 주시기 바랍니다."))
                        {
                            MessageBox.Show(dt.Rows[0].ItemArray[0].ToString());
                            DialogResult = false;
                            return false;
                        }

                        if (fail >= 5)
                        {
                            strLogRegID = strID;
                            Lib.Instance.SetLogResitry(strLogRegID);
                            //DialogResult = true;
                            goPwChange(strID);  //비밀번호변경


                            //Login_PwChange pwChange = new Login_PwChange();

                        }
                        //dt.Rows[0].ItemArray[0].ToString();

                        if (dt.Rows[0].ItemArray[0].Equals("계정이 차단되었습니다. 관리자에게 문의해 주시기 바랍니다."))
                        {
                            MessageBox.Show(dt.Rows[0].ItemArray[0].ToString());
                            return false;
                        }
                        else
                        {
                            if (fail < 5)
                            {
                                MessageBox.Show(dt.Rows[0].ItemArray[0].ToString() + "(" + fail + " / 5 회" + ")");

                            }
                            else
                            {
                                MessageBox.Show(dt.Rows[0].ItemArray[0].ToString() + "(" + fail + " / 5 회" + ") \r\n" + "계정이 차단되었으니 관리자에게 문의해주시기 바랍니다.");
                                DialogResult = false;
                            }
                            return false;
                        }

                    }
                    else
                    {
                        //if (!dt.Rows[0]["Name"].Equals("20150401") && !dt.Rows[0]["Name"].Equals("admin"))
                        //{
                        //    MessageBox.Show("권한이 없는 사용자입니다.");
                        //    return flag;
                        //}

                        // 비밀번호 변경 추가
                        if (CheckConvertDateTime(dt.Rows[0]["PasswordChangeDate"].ToString()) == true)
                        {
                            DateTime setDate = DateTime.Parse(dt.Rows[0]["PasswordChangeDate"].ToString());

                            TimeSpan timeDiff = DateTime.Today - setDate;
                            dayDiff = timeDiff.Days;

                            if (dayDiff > 90)
                            {
                                //exPassword = dt.Rows[0]["Password"].ToString();
                            }
                        }
                        else if (dt.Rows[0]["PasswordChangeDate"].ToString().Trim().Equals("")) // 초기 비밀번호가 세팅되지 않았다면
                        {
                            initChange = "N";
                            //exPassword = dt.Rows[0]["Password"].ToString();
                        }

                        // 개인정보활용 동의 여부 추가, 이름도 추가
                        AccessControl = dt.Rows[0]["NessaryAcptYN"].ToString();
                        UserName = dt.Rows[0]["Name"].ToString();

                        //PersonID
                        PersonID = dt.Rows[0]["PersonID"].ToString();
                    }

                    strUserName = UserName;
                }
            }
            DataStore.Instance.CloseConnection(); //2021-09-13 현달씨 DBClose

            return flag;

        }


        // 비밀번호 변경
        private bool goPwChange(string strID)
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            // [xp_Code_uPassWord]
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("UserID", strID);
            sqlParameter.Add("PassWord", "0000");
            sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentPersonID);

            Procedure pro1 = new Procedure();
            pro1.Name = "xp_Code_uPassWord_Fail";
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
                //return false;fail
            }
            else
            {
                //MessageBox.Show("성공");
                //fail = 0;
                flag = true;
            }



            return flag;


        }


        private void GetInfo()
        {
            txtUserID.Text = Lib.Instance.GetLogResitry();

            if (txtUserID.Text.Equals(""))
            {
                txtUserID.Focus();
            }
            else
            {
                txtPassWd.Focus();
            }
        }

        private void txtPassWd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (Log(txtUserID.Text))
                {
                    strLogRegID = txtUserID.Text;
                    Lib.Instance.SetLogResitry(strLogRegID);
                    DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
                    DialogResult = true;
                }
                else
                {
                    txtPassWd.Password = "";
                    return;
                }
            }
        }

        #region 기타 메서드

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // DateTime 으로 변환 가능한지 체크
        private bool CheckConvertDateTime(string str)
        {
            bool flag = false;

            DateTime chkDt;

            if (!str.Trim().Equals(""))
            {
                if (str.Length == 8)
                {
                    str = DatePickerFormat(str);

                    if (DateTime.TryParse(str, out chkDt) == true)
                    {
                        flag = true;
                        return flag;
                    }
                }
                else
                {
                    if (DateTime.TryParse(str, out chkDt) == true)
                    {
                        flag = true;
                        return flag;
                    }
                }
            }

            return flag;
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
    }
}
