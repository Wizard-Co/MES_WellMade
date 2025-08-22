using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows;

namespace WizMes_WellMade.PopUp
{
    /// <summary>
    /// NoRunningMachineInfo.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class NoRunningMachineInfo : Window
    {
        ObservableCollection<NoRunningMachineInfo_CodeView> ovcProcessResult_Q =
            new ObservableCollection<NoRunningMachineInfo_CodeView>();

        public string strSDate = string.Empty;
        public string strEDate = string.Empty;
        public string strProcessID = string.Empty;
        public string strMachinID = string.Empty;

        public NoRunningMachineInfo()
        {
            InitializeComponent();
        }

        public NoRunningMachineInfo(string SDate, string EDate, string strProcess, string strMachine)
        {
            InitializeComponent();
            strSDate = SDate;
            strEDate = EDate;
            strProcessID = strProcess;
            strMachinID = strMachine;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            FillGrid();
        }

        private void FillGrid()
        {
            try
            {
                //if (cboProcess.SelectedValue != null)
                //{
                //    strProcess = cboProcess.SelectedValue.ToString();
                //}

                //if (cboMachine.SelectedIndex > 0)
                //{
                //    if (cboMachine.SelectedValue != null)
                //    {
                //        strProcess = cboMachine.SelectedValue.ToString().Split('/')[0];
                //        strMachine = cboMachine.SelectedValue.ToString().Split('/')[1];
                //    }
                //}

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sFromDate", strSDate);
                sqlParameter.Add("sToDate", strEDate);
                sqlParameter.Add("sProcessID", strProcessID);
                sqlParameter.Add("sMachineID", strMachinID);
                sqlParameter.Add("ArticleID", "");

                sqlParameter.Add("CustomID", "");
                sqlParameter.Add("nOrderID", 0);
                sqlParameter.Add("sOrderID", "");
                sqlParameter.Add("nJobGbn", 1);
                sqlParameter.Add("sJobGubun", "2");

                sqlParameter.Add("nBuyerModel", 0);
                sqlParameter.Add("sBuyerModel", "");
                sqlParameter.Add("nBuyerArticleNo", 0);
                sqlParameter.Add("sBuyerArticleNo", "");
                sqlParameter.Add("ndefect", 0);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_MIS_sMCRunningRate_NoWork", sqlParameter, false);

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
                            var WinProcessResult = new NoRunningMachineInfo_CodeView()
                            {
                                Num = i,
                                cls = dr["cls"].ToString(),
                                ScanDate = dr["ScanDate"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                OrderID = dr["OrderID"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                AcptDate = dr["AcptDate"].ToString(),
                                OrderQty = dr["OrderQty"].ToString(),
                                MachineID = dr["MachineID"].ToString(),
                                InstDate = dr["InstDate"].ToString(),
                                InstQty = dr["InstQty"].ToString(),
                                WorkQty = dr["WorkQty"].ToString(),
                                WorkPersonID = dr["WorkPersonID"].ToString(),
                                ScanTime = dr["ScanTime"].ToString(),
                                BuyerModelID = dr["BuyerModelID"].ToString(),
                                BuyerModel = dr["BuyerModel"].ToString(),
                                CustomID = dr["CustomID"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                Worker = dr["Worker"].ToString(),
                                Article = dr["Article"].ToString(),
                                LabelID = dr["LabelID"].ToString(),
                                JobGbn = dr["JobGbn"].ToString(),
                                JobGbnname = dr["JobGbnname"].ToString(),
                                WorkStartDate = dr["WorkStartDate"].ToString(),
                                WorkStartTime = dr["WorkStartTime"].ToString(),
                                WorkEndDate = dr["WorkEndDate"].ToString(),
                                WorkEndTime = dr["WorkEndTime"].ToString(),
                                WorkHour = dr["WorkHour"].ToString(),
                                WorkMinute = dr["WorkMinute"].ToString(),
                                JobID = dr["JobID"].ToString(),
                                Articleid = dr["Articleid"].ToString(),
                                WorkCnt = dr["WorkCnt"].ToString(),
                                NoReworkCode = dr["NoReworkCode"].ToString(),
                                NoReworkName = dr["NoReworkName"].ToString(),
                                FourMID = dr["4MID"].ToString(),
                                FourMSubject = dr["4MSubject"].ToString()
                            };

                            WinProcessResult.InstQty = Lib.Instance.returnNumStringZero(WinProcessResult.InstQty);
                            WinProcessResult.OrderQty = Lib.Instance.returnNumStringZero(WinProcessResult.OrderQty);
                            WinProcessResult.WorkQty = Lib.Instance.returnNumStringZero(WinProcessResult.WorkQty);

                            if (WinProcessResult.cls.Equals("1"))
                            {
                                WinProcessResult.Time = StartTimeAndEndTime(WinProcessResult.WorkStartTime, WinProcessResult.WorkEndTime);
                            }

                            if (WinProcessResult.cls.Equals("1"))
                            {
                                ovcProcessResult_Q.Add(WinProcessResult);
                            }
                            //dgdResult.Items.Add(WinProcessResult);
                        }
                        dgdResult.ItemsSource = ovcProcessResult_Q;
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

        private string StartTimeAndEndTime(string STime, string ETime)
        {
            string STandET = string.Empty;

            if (STime.Equals(string.Empty))
            {
                STandET += "None ~ ";
            }
            else
            {
                STandET += STime.Substring(0, 2) + ":" + STime.Substring(2, 2) + " ~ ";
            }

            if (ETime.Equals(string.Empty))
            {
                STandET += "None";
            }
            else
            {
                STandET += ETime.Substring(0, 2) + ":" + ETime.Substring(2, 2);
            }

            return STandET;
        }
    }

    public class NoRunningMachineInfo_CodeView : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public string cls { get; set; }
        public string ScanDate { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string OrderID { get; set; }

        public string OrderNo { get; set; }
        public string AcptDate { get; set; }
        public string OrderQty { get; set; }
        public string MachineID { get; set; }
        public string InstDate { get; set; }

        public string InstQty { get; set; }
        public string WorkQty { get; set; }
        public string WorkPersonID { get; set; }
        public string ScanTime { get; set; }
        public string BuyerModelID { get; set; }

        public string BuyerModel { get; set; }
        public string CustomID { get; set; }
        public string KCustom { get; set; }
        public string Worker { get; set; }
        public string Article { get; set; }

        public string LabelID { get; set; }
        public string JobGbn { get; set; }
        public string JobGbnname { get; set; }
        public string WorkStartDate { get; set; }
        public string WorkStartTime { get; set; }

        public string WorkEndDate { get; set; }
        public string WorkEndTime { get; set; }
        public string WorkHour { get; set; }
        public string WorkMinute { get; set; }
        public string JobID { get; set; }

        public string Articleid { get; set; }
        public string WorkCnt { get; set; }
        public string NoReworkCode { get; set; }
        public string NoReworkName { get; set; }
        public string FourMID { get; set; }

        public string FourMSubject { get; set; }
        public int Num { get; set; }
        public string Time { get; set; }
    }
}
