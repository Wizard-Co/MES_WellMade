using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using WizMes_WellMade;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// NoRunningMachineInfo.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class NoWorkInfo : Window
    {
        ObservableCollection<NoRunningMachineInfo_CodeView> ovcProcessResult_Q =
            new ObservableCollection<NoRunningMachineInfo_CodeView>();

        public string strSDate = string.Empty;
        public string strEDate = string.Empty;
        public string strMCID = string.Empty;

        // 테스트용
        public string flag = "";
        public string WorkDate = "";
        public string WorkStartDateTime = "";
        public string WorkEndDateTime = "";
        public string ProcessID = "";
        public string MachineNo = "";
        public string Name = "";

        public NoWorkInfo()
        {
            InitializeComponent();
        }

        public NoWorkInfo(string SDate, string EDate, string MCID)
        {
            InitializeComponent();
            strSDate = SDate;
            strEDate = EDate;
            strMCID = MCID;
        }

        public NoWorkInfo(string WorkDate, string ProcessID, string MachineNo, string Name)
        {
            InitializeComponent();
            flag = "Test";

            this.WorkDate = WorkDate;
            this.ProcessID = ProcessID;
            this.MachineNo = MachineNo;
            this.Name = Name;
        }

        public NoWorkInfo(string WorkStartDateTime, string WorkEndDateTime, string ProcessID, string MachineNo, string Name)
        {
            InitializeComponent();
            flag = "Test";

            this.WorkStartDateTime = WorkStartDateTime;
            this.WorkEndDateTime = WorkEndDateTime;
            this.ProcessID = ProcessID;
            this.MachineNo = MachineNo;
            this.Name = Name;
        }

        #region 설비 무작업 조회

        private void FillGridSub()
        {
            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("WorkDate", WorkDate);
                sqlParameter.Add("ProcessID", ProcessID);
                sqlParameter.Add("MachineNo", MachineNo);
                sqlParameter.Add("Name", Name);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sNoWorkSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WinNoR = new NoWorkInfo_CodeView()
                            {
                                Num = i.ToString(),
                                WorkDate = dr["WorkDate"].ToString(),
                                WorkDate_CV = DatePickerFormat(dr["WorkDate"].ToString()),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                WorkStartTime = dr["WorkStartTime"].ToString(),
                                WorkStartTime_CV = DateTimeFormat(dr["WorkStartTime"].ToString()),
                                WorkEndTime = dr["WorkEndTime"].ToString(),
                                WorkEndTime_CV = DateTimeFormat(dr["WorkEndTime"].ToString()),
                                WorkRealTime = DateTimeMinToTime(dr["WorkRealTime"].ToString()),
                                WorkerName = dr["WorkerName"].ToString(),
                                NoReworkName = dr["NoReworkName"].ToString(),
                                NoReworkReason = dr["NoReworkReason"].ToString(),
                                Process = dr["Process"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                            };

                            dgdResult.Items.Add(WinNoR);
                        }
                    }
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
        }

        #endregion

        #region 설비 무작업 조회2

        private void FillGridSub2()
        {
            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("WorkStartDateTime", WorkStartDateTime.Replace(" ", "").Replace("-", "").Replace(":", "").Replace("/", ""));
                sqlParameter.Add("WorkEndDateTime", WorkEndDateTime.Replace(" ", "").Replace("-", "").Replace(":", "").Replace("/", ""));
                sqlParameter.Add("ProcessID", ProcessID);
                sqlParameter.Add("MachineNo", MachineNo);
                sqlParameter.Add("Name", Name);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sNoWorkDiff", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WinNoR = new NoWorkInfo_CodeView()
                            {
                                Num = i.ToString(),
                                WorkDate = dr["WorkDate"].ToString(),
                                WorkDate_CV = DatePickerFormat(dr["WorkDate"].ToString()),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                WorkStartTime = dr["WorkStartTime"].ToString(),
                                WorkStartTime_CV = DateTimeFormat(dr["WorkStartTime"].ToString()),
                                WorkEndTime = dr["WorkEndTime"].ToString(),
                                WorkEndTime_CV = DateTimeFormat(dr["WorkEndTime"].ToString()),
                                WorkRealTime = DateTimeMinToTime(dr["WorkRealTime"].ToString()),
                                WorkerName = dr["WorkerName"].ToString(),
                                NoReworkName = dr["NoReworkName"].ToString(),
                                NoReworkReason = dr["NoReworkReason"].ToString(),
                                Process = dr["Process"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                            };

                            dgdResult.Items.Add(WinNoR);
                        }
                    }
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
        }

        #endregion

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (flag != null
                && flag.Equals(""))
            {
                FillGrid();
            }
            else if (flag.Equals("Test"))
            {
                FillGridSub2();
            }
        }

        private void FillGrid()
        {
            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sFromDate", strSDate);
                sqlParameter.Add("sToDate", strEDate);
                sqlParameter.Add("MCID", strMCID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sNoWork", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var WinNoR = new NoWorkInfo_CodeView()
                            {
                                Num = i.ToString(),
                                WorkDate = dr["WorkDate"].ToString(),
                                WorkDate_CV = DatePickerFormat(dr["WorkDate"].ToString()),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                WorkStartTime = dr["WorkStartTime"].ToString(),
                                WorkStartTime_CV = DateTimeFormat(dr["WorkStartTime"].ToString()),
                                WorkEndTime = dr["WorkEndTime"].ToString(),
                                WorkEndTime_CV = DateTimeFormat(dr["WorkEndTime"].ToString()),
                                WorkRealTime = DateTimeMinToTime(dr["WorkRealTime"].ToString()),
                                WorkerName = dr["WorkerName"].ToString(),
                                NoReworkName = dr["NoReworkName"].ToString(),
                                NoReworkReason = dr["NoReworkReason"].ToString(),
                                Process = dr["Process"].ToString(),
                                MachineNo = dr["MachineNo"].ToString(),
                            };

                            dgdResult.Items.Add(WinNoR);
                        }
                    }
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

    }

    public class NoWorkInfo_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string Num { get; set; }
        public string WorkDate_CV { get; set; }
        public string WorkDate { get; set; }
        public string BuyerArticleNo { get; set; }
        public string Article { get; set; }

        public string WorkStartTime_CV { get; set; }
        public string WorkStartTime { get; set; }
        public string WorkEndTime_CV { get; set; }
        public string WorkEndTime { get; set; }
        public string WorkRealTime { get; set; }

        public string WorkerName { get; set; }
        public string NoReworkName { get; set; }
        public string NoReworkReason { get; set; }
        public string Process { get; set; }
        public string MachineNo { get; set; }
    }
}
