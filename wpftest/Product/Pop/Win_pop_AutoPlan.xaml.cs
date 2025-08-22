using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WizMes_BooKyong.PopUp
{
    /// <summary>
    /// RheoChoice.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_pop_AutoPlan : Window
    {
        int rowNum = 0;
        string stDate = string.Empty;
        string stTime = string.Empty;
        public static bool AutoSave = false;

        string AutoID = ""; // 작업지시 PK
        Lib lib = new Lib();
        

        public Win_pop_AutoPlan()
        {
            InitializeComponent();
        }

        private void AutoPlan_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");
            Lib.Instance.UiLoading(sender);
        }

        #region 버튼 이벤트 -  닫기, 검색

        public List<Win_mtr_LotStockControl_U_CodeView> lstLotStock = new List<Win_mtr_LotStockControl_U_CodeView>();


        //닫기
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            this.Close();
        }
        
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (Loading lw = new Loading(SaveData))
                {
                    lw.ShowDialog();
                    re_Search();
                    Win_prd_PlanInputAuto.planAuto.dgdMain.SelectedIndex = 0;
                    Win_prd_PlanInputAuto.planAuto.FillGrid();
                    
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("예외처리 - " + ee.ToString());
            }
        }

        //검색
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>

            {
                //로직
                using (Loading lw = new Loading(re_Search))
                {
                    lw.ShowDialog();
                }

                if (dgdMain.Items.Count == 0)
                {
                    dgdPattern.Items.Clear();

                    MessageBox.Show("조회된 데이터가 없습니다.");
                }
                else
                {
                    dgdMain.SelectedIndex = 0;
                }

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        #endregion // 주요 버튼 이벤트


        #region Header 부분 - 검색조건


       
      

        #endregion // Header 부분 - 검색조건

        #region 주요 메서드 모음

        private void re_Search()
        {
            FillGrid();

            if (dgdMain.Items.Count > 0)
            {
                dgdMain.SelectedIndex = rowNum;
            }
            else
            {
                this.DataContext = null;
                MessageBox.Show("조회된 데이터가 없습니다.");
                return;
            }
        }

        #region 조회 - 생성대상조회 
        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
            }
            if (dgdPattern.Items.Count > 0)
            {
                dgdPattern.Items.Clear();

            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("ChkDate",0);
                sqlParameter.Add("SDate", "");
                sqlParameter.Add("EDate", "");
                sqlParameter.Add("ChkCustomID",  0);
                sqlParameter.Add("CustomID", "");
                sqlParameter.Add("ChkArticleID", 0);
                sqlParameter.Add("ArticleID", "");
                sqlParameter.Add("ChkOrder",  0);
                sqlParameter.Add("Order",  "");
                sqlParameter.Add("ChkIncPlComplete",  0);
                sqlParameter.Add("ChkCloseClss", 0);
                sqlParameter.Add("ChkBuyerArticleNo",  0);
                sqlParameter.Add("BuyerArticleNoID", "");


                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_AutoPlan_sOrder", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var AutoPlan = new Win_prd_AutoPlan_CodeView()
                            {
                                AcptDate = dr["AcptDate"].ToString(),
                                KCustom = dr["KCustom"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                Article = dr["Article"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                OrderNo = dr["OrderNo"].ToString(),
                                OrderID = dr["OrderID"].ToString(),
                                DvlyDate =dr["DvlyDate"].ToString(),
                                OrderQty = stringFormatN0(dr["OrderQty"]),
                                OrderInstQy = stringFormatN0(dr["OrderInstQy"]),
                                notOrderInstQty = stringFormatN0(dr["notOrderInstQty"]),
                                OrderSeq = dr["OrderSeq"].ToString(),

                            };

                            dgdMain.Items.Add(AutoPlan);
                        }
                        //tbkCount.Text = "▶ 검색결과 : " + i.ToString() + " 건";
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
        #endregion



        #endregion


        #region 생산계획편성 - 편성처리 
        //
        private void SaveData()
        {
            AutoSave = false;

            if (dgdPattern.Items.Count > 0)
            {
                dgdPattern.Items.Clear();

            }
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();
            var AutoPlan = dgdMain.SelectedItem as Win_prd_AutoPlan_CodeView;

            if (AutoPlan != null)
            {
                try
                {
                    if (CheckData(AutoPlan.ArticleID))
                    {
                        Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                        sqlParameter.Clear();

                        sqlParameter.Add("AutoID", "");
                        //sqlParameter.Add("InstDate", AutoPlan.AcptDate);
                        sqlParameter.Add("AutoDate", DateTime.Now.ToString("yyyyMMdd"));
                        sqlParameter.Add("OrderID", AutoPlan.OrderID);
                        sqlParameter.Add("OrderSeq", AutoPlan.OrderSeq);
                        sqlParameter.Add("InstRoll", "0");
                        sqlParameter.Add("InstQty", AutoPlan.notOrderInstQty.Replace(",", ""));
                        sqlParameter.Add("ExpectDate", AutoPlan.DvlyDate);
                        sqlParameter.Add("PersonID", MainWindow.CurrentUser);
                        sqlParameter.Add("Remark", "");
                        sqlParameter.Add("MtrExceptYN", "N");
                        sqlParameter.Add("OutwareExceptYN", "N");
                        sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);
                        sqlParameter.Add("AutoPlanYN", chkAutoInput.IsChecked == true ? "Y" : "N");

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_PlanInput_iAutoPlan";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "AutoID";
                        pro1.OutputLength = "12";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdPattern.Items.Count; i++)
                        {
                            var AutoPattern = dgdPattern.Items[i] as Win_prd_AutoPattern_CodeView;
                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Add("AutoID", "");
                            sqlParameter.Add("AutoDate", DateTime.Now.ToString("yyyyMMdd"));
                            sqlParameter.Add("ProcSeq", AutoPattern.PatternSeq);
                            sqlParameter.Add("ArticleID", AutoPattern.ArticleID);
                            sqlParameter.Add("ProcessID", AutoPattern.ProcessID);

                            sqlParameter.Add("InstRemark", "");
                            sqlParameter.Add("InstQty", AutoPlan.notOrderInstQty.Replace(",", "")); //내일오면 이거 메인그리드 미계획량으로 바꾸자 
                            sqlParameter.Add("StartDate", DateTime.Now.ToString("yyyyMMdd"));
                            sqlParameter.Add("EndDate", AutoPattern.EndDate == null ? "" : AutoPattern.EndDate);
                            sqlParameter.Add("Remark", AutoPattern.Remark == null ? "" : AutoPattern.Remark);

                            sqlParameter.Add("MachineID", AutoPattern.MachineID == null ? "" : AutoPattern.MachineID);
                            sqlParameter.Add("MtrExceptYN", AutoPattern.MtrExceptYN == null ? "" : AutoPattern.MtrExceptYN);
                            sqlParameter.Add("FirstInFirstOutYN", AutoPattern.FirstInFirstOutYN == null ? "" : AutoPattern.FirstInFirstOutYN);
                            sqlParameter.Add("CreateUserID", MainWindow.CurrentUser);
                            sqlParameter.Add("AutoPlanYN", chkAutoInput.IsChecked == true ? "Y" : "N");

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_PlanInput_iAutoPlanSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "AutoID";
                            pro2.OutputLength = "12";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);
                        }

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter, "C");
                        string sGetID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "AutoID")
                                {
                                    AutoID = kv.value;
                                    sGetID = kv.value;
                                    flag = true;
                                    AutoSave = true;
                                }
                            }

                            if (flag && chkAutoInput.IsChecked == true)
                            {
                                UpdatePattern(AutoPlan.OrderID, AutoPlan.ArticleID);
                            }

                            MessageBox.Show("편성이 완료 되었습니다");

                          
                        }
                        else
                        {
                            MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                            flag = false;
                            
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

            } else
            {
                MessageBox.Show("대상을 선택 해주세요");
            }

        }
        #endregion

        #region UpdatePattern
        //
        private bool UpdatePattern(string strOrderID, string strArticleID)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("OrderID", strOrderID);
                sqlParameter.Add("ArticleID", strArticleID);
                sqlParameter.Add("PatternID", "");
                sqlParameter.Add("StuffCloseClss", "");
                sqlParameter.Add("LastUpdateUserID", MainWindow.CurrentUser);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_PlanInput_uOrderPatternID";
                pro1.OutputUseYN = "N";
                pro1.OutputName = "OrderID";
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
        #endregion


        #region 유효성 검사

        private bool CheckData(string sArticleID)
        {
            bool flag = true;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ArticleID", sArticleID);
                sqlParameter.Add("sOutMessage", "");
            
                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_prd_sAutoPattern", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count > 0 && dt.Columns.Count > 1)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;

                            var Pattern = new Win_prd_AutoPattern_CodeView()
                            {
                                PatternSeq = dr["PatternSeq"].ToString(),
                                ProcessID = dr["ProcessID"].ToString(),
                                Process = dr["Process"].ToString(),
                                Qty = dr["Qty"].ToString(),
                                ArticleID = dr["ArticleID"].ToString(),
                                Article = dr["Article"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                                LVL = dr["LVL"].ToString(),
                                ChildBuyerArticleNo = dr["ChildBuyerArticleNo"].ToString(),
                             
                            };

                            dgdPattern.Items.Add(Pattern);
                        }
                    } else if (dt.Columns.Count == 1)
                    {
                        DataRow dr = dt.Rows[0];
                        string error = dr["error"].ToString();

                        flag = false;

                        MessageBox.Show(error);
                    }
             
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("오류 발생, 오류 내용 : " + ex.ToString());
                MessageBox.Show(ex.ToString());
                flag = false;
                return flag;
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
            return flag;
        }

        #endregion

        #region 전체 선택 체크박스 이벤트

        // 전체 선택 체크박스 체크 이벤트
        private void AllCheck_Checked(object sender, RoutedEventArgs e)
        {
            //ovcMoveSub.Clear();

            //if (dgdMain.Visibility == Visibility.Visible)
            //{
            //    for (int i = 0; i < dgdMain.Items.Count; i++)
            //    {
            //        var MoveSub = dgdMain.Items[i] as Win_mtr_Move_U_CodeViewSub;
            //        MoveSub.Chk = true;
            //        MoveSub.FontColor = true;

            //        ovcMoveSub.Add(MoveSub);
            //    }
            //}
        }

        // 전체 선택 체크박스 언체크 이벤트
        private void AllCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            //ovcMoveSub.Clear();

            //if (dgdMain.Visibility == Visibility.Visible)
            //{
            //    for (int i = 0; i < dgdMain.Items.Count; i++)
            //    {
            //        var MoveSub = dgdMain.Items[i] as Win_mtr_Move_U_CodeViewSub;
            //        MoveSub.Chk = false;
            //        MoveSub.FontColor = false;
            //    }
            //}
        }

        #endregion // 전체 선택 체크박스 이벤트

        #region 기타 메서드

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
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




        //2021-05-29(2021-07-12 해제도 추가)
        private void BtnAllChoice_Click(object sender, RoutedEventArgs e)
        {
            if (dgdMain.Items.Count > 0)
            {
                foreach (Win_mtr_LotStockControl_U_CodeView Silsadata in dgdMain.Items)
                {

                    if (Silsadata != null && Silsadata.Chk == false)
                    {
                        Silsadata.Chk = true;
                    }
                    else
                    {
                        Silsadata.Chk = false;
                    }

                }

                dgdMain.Items.Refresh();
            }
        }


     

    }

    public class Win_prd_AutoPlan_CodeView
    {
        public int Num { get; set; }
        public string AcptDate { get; set; }            
        public string KCustom { get; set; }           
        public string BuyerArticleNo { get; set; }          
        public string ArticleID { get; set; }          
        public string Article { get; set; }            
        public string OrderNo { get; set; }     
        public string OrderID { get; set; }             
        public string OrderSeq { get; set; }             
        public string DvlyDate { get; set; }          
        public string OrderQty { get; set; }            //수주량
        public string OrderInstQy { get; set; }         //계획량
        public string notOrderInstQty { get; set; }     //미계획량 
    }

    public class Win_prd_AutoPattern_CodeView
    {
        public string PatternSeq { get; set; }
        public string ProcessID { get; set; }
        public string Process { get; set; }
        public string Qty { get; set; }
        public string ArticleID { get; set; }
        public string Article { get; set; }
        public string BuyerArticleNo { get; set; }
        public string LVL { get; set; }
        public string ChildBuyerArticleNo { get; set; }            
        public string InstQty { get; set; }            
        public string StartDate { get; set; }            
        public string EndDate { get; set; }            
        public string Remark { get; set; }            
        public string MachineID { get; set; }            
        public string MtrExceptYN { get; set; }            
        public string FirstInFirstOutYN { get; set; }            
    }
}
