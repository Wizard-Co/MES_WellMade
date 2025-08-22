using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// ReFresh.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ReFresh : Window
    {

        public string strDate = string.Empty;
        public int ButtonClick = 0; //2021-11-02 어떤 버튼을 눌렀는지 판단하는 변수
        // 테스트용
        public string flag = "";
        public string WorkDate = "";
        public string WorkStartDateTime = "";
        public string WorkEndDateTime = "";
        public string ProcessID = "";
        public string MachineNo = "";
        public string Name = "";

        public ReFresh()
        {
            InitializeComponent();
        }

        public ReFresh(string WorkDate, string ProcessID, string MachineNo, string Name)
        {
            InitializeComponent();
            flag = "Test";

            this.WorkDate = WorkDate;
            this.ProcessID = ProcessID;
            this.MachineNo = MachineNo;
            this.Name = Name;
        }

        public ReFresh(string WorkStartDateTime, string WorkEndDateTime, string ProcessID, string MachineNo, string Name)
        {
            InitializeComponent();
            flag = "Test";

            this.WorkStartDateTime = WorkStartDateTime;
            this.WorkEndDateTime = WorkEndDateTime;
            this.ProcessID = ProcessID;
            this.MachineNo = MachineNo;
            this.Name = Name;
        }


        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DateTime now = DateTime.Now;

            FromDay.SelectedDate = now;
            ToDay.SelectedDate = now;

        }


        #region 기타 메서드 모음

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN1(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천마리 콤마, 소수점 두자리
        private string stringFormatN2(object obj)
        {
            return string.Format("{0:N2}", obj);
        }

        // 천마리 콤마, 소수점 두자리
        private string stringFormatNDigit(object obj, int digit)
        {
            return string.Format("{0:N" + digit + "}", obj);
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

        // 시간 형식 6글자라면! 11:11:11
        private string DateTimeFormat(string str)
        {
            str = str.Replace(":", "").Trim();

            if (str.Length == 6)
            {
                string Hour = str.Substring(0, 2);
                string Min = str.Substring(2, 2);
                string Sec = str.Substring(4, 2);

                str = Hour + ":" + Min + ":" + Sec;
            }

            return str;
        }

        // 시간 분 → 11:12 형식으로 변환
        private string DateTimeMinToTime(string str)
        {
            str = str.Replace(":", "").Trim();

            int num = 0;
            if (int.TryParse(str, out num) == true)
            {
                string hour = (num / 60).ToString();
                string min = (num % 60).ToString();

                if (min.Length == 1)
                {
                    min = "0" + min;
                }

                str = hour + ":" + min;
            }

            return str;
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


        #endregion

        private void btnDayUpdate_Click(object sender, RoutedEventArgs e)
        {
            ButtonClick = 1;
            btnDayUpdate.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                using (Loading lw = new Loading(WorkTimeUpdate))
                {
                    lw.ShowDialog();
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnDayUpdate.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }
                    
        private void WorkTimeUpdate()
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();

            if (ButtonClick == 1)
            {
                sqlParameter.Add("sFromDate", FromDay.SelectedDate != null ? FromDay.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("sToDate", ToDay.SelectedDate != null ? ToDay.SelectedDate.Value.ToString("yyyyMMdd") : "");
                string[] result = DataStore.Instance.ExecuteProcedure("xp_Batch_iWorkTime_NowUpdate", sqlParameter, false);
                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("이상발생, 관리자에게 문의하세요.");
                    return;
                }
                else
                {
                    MessageBox.Show(FromDay.SelectedDate.Value.ToString("yyyy-MM-dd") + " ~ " + ToDay.SelectedDate.Value.ToString("yyyy-MM-dd") + "까지 업데이트 되었습니다.");
                    DataStore.Instance.CloseConnection();
                }
            }
            else
            {
                MessageBox.Show("버튼클릭, 버튼을 눌러주세요");
                return;
            }
        }

    }

}
