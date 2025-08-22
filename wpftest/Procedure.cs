using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows;

namespace WizMes_WellMade
{
    public class Procedure
    {
        private static Procedure mProcedure = null;

        public static Procedure Instance
        {
            get
            {
                if (mProcedure == null)
                {
                    mProcedure = new Procedure();
                }

                return mProcedure;
            }
        }

        public string Name { get; set; }
        public string OutputUseYN { get; set; }
        public string OutputName { get; set; }
        public string OutputLength { get; set; }
        public int OutputCount { get; set; }
        public List<string> list_OutputName { get; set; }
        public List<string> list_OutputLength { get; set; }
        //2021-07-09 추가
        public Dictionary<SqlParameter, object> dicOutputList { get; set; }

        /// <summary>
        /// xp_PlanInput_sPlanInput 사용
        /// </summary>
        /// <param name="numOrderDay"></param>
        /// <param name="SDate"></param>
        /// <param name="EDate"></param>
        /// <param name="numCustomID"></param>
        /// <param name="strCustomID"></param>
        /// <param name="numArticleID"></param>
        /// <param name="strArticleID"></param>
        /// <param name="numOrderID"></param>
        /// <param name="strOrderID"></param>
        /// <param name="numComplete"></param>
        /// <returns></returns>
        public DataTable GetPlanInput(int numOrderDay, string SDate, string EDate, int numCustomID, string strCustomID
            , int numArticleID, string strArticleID, int numOrderID, string strOrderID, int numComplete)
        {
            DataTable dataTable = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkDate", numOrderDay);
                sqlParameter.Add("SDate", SDate);
                sqlParameter.Add("EDate", EDate);
                sqlParameter.Add("ChkCustomID", numCustomID);
                sqlParameter.Add("CustomID", strCustomID);

                sqlParameter.Add("ChkArticleID", numArticleID);
                sqlParameter.Add("ArticleID", strArticleID);
                sqlParameter.Add("ChkOrder", numOrderID);
                sqlParameter.Add("Order", strOrderID);
                sqlParameter.Add("ChkPlanComplete", numComplete);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_PlanInput_sPlanInput", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
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

            return dataTable;
        }


        /// <summary>
        /// xp_PlanInput_sPlanInputDet 사용
        /// </summary>
        /// <param name="numDate"></param>
        /// <param name="SDate"></param>
        /// <param name="EDate"></param>
        /// <param name="numCustomID"></param>
        /// <param name="strCustomID"></param>
        /// <param name="numArticleID"></param>
        /// <param name="strArticleID"></param>
        /// <param name="numOrderID"></param>
        /// <param name="strOrderID"></param>
        /// <param name="numInstID"></param>
        /// <param name="strInstID"></param>
        /// <returns></returns>
        public DataTable GetPlanInputDet(int numDate, string SDate, string EDate, int numCustomID, string strCustomID
            , int numArticleID, string strArticleID, int numOrderID, string strOrderID, int numInstID, string strInstID)
        {
            DataTable dataTable = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("ChkDate", numDate);
                sqlParameter.Add("SDate", SDate);
                sqlParameter.Add("EDate", EDate);
                sqlParameter.Add("ChkCustomID", numCustomID);
                sqlParameter.Add("CustomID", strCustomID);

                sqlParameter.Add("ChkArticleID", numArticleID);
                sqlParameter.Add("ArticleID", strArticleID);
                sqlParameter.Add("nChkOrder", numOrderID);
                sqlParameter.Add("Order", strOrderID);
                sqlParameter.Add("nChkInstID", numInstID);
                sqlParameter.Add("sInstID", strInstID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_prd_sPlanInputDet", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
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

            return dataTable;
        }

        /// <summary>
        /// xp_PlanInput_sPlanInputDetArticle 사용
        /// </summary>
        /// <param name="strInstID"></param>
        /// <param name="numInstSeq"></param>
        /// <returns></returns>
        public DataTable GetPlanInputDetArticleChild(string strInstID, int numInstSeq)
        {
            DataTable dataTable = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sInstID", strInstID);
                sqlParameter.Add("sInstSeq", numInstSeq);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_PlanInput_sPlanInputDetArticle", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
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

            return dataTable;
        }


        /// <summary>
        /// xp_Article_sArticle 사용
        /// </summary>
        /// <param name="strArticleGrpID"></param>
        /// <param name="strArticleID"></param>
        /// <param name="strDirection"></param>
        /// <param name="numNotUse"></param>
        /// <returns></returns>
        public DataTable GetArticle(string strArticleID, int numNotUse, string strArticleGrpID, string strSupplyType)
        {
            DataTable dataTable = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sArticle", strArticleID);
                sqlParameter.Add("iIncNotUse", numNotUse);
                sqlParameter.Add("sArticleGrpID", strArticleGrpID);
                sqlParameter.Add("sSupplyType", strSupplyType);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticle", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
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

            return dataTable;
        }

        /// <summary>
        /// xp_Custom_sCustomArticle 사용
        /// </summary>
        /// <param name="strCustom"></param>
        /// <param name="strCustomID"></param>
        /// <param name="strArticleID"></param>
        /// <param name="strYN"></param>
        /// <returns></returns>
        public DataTable GetCustomArticle(int numCustom, string strCustomID, string strArticleID, string strYN)
        {
            DataTable dataTable = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nCustom", numCustom);
                sqlParameter.Add("sCustom", strCustomID);
                sqlParameter.Add("sArticleID", strArticleID);
                sqlParameter.Add("sShowArticleYN", strYN);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Custom_sCustomArticle", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
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

            return dataTable;
        }

        // 2019.10.11 곽동운 : 
        // 거래처별 등록 품목( CustomArticle_U ) 화면에서 / 해당 거래처의 / 선택된 데이터만 가져오기 위해서 추가 
        public DataTable GetCustomArticleSelection(string sCustomID) //(int numCustom, string strCustom, string strArticleID, string strYN)
        {
            DataTable dataTable = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                //sqlParameter.Add("chkCustom", numCustom);
                sqlParameter.Add("sCustomID", sCustomID);
                //sqlParameter.Add("sArticleID", strArticleID);
                //sqlParameter.Add("sShowArticleYN", strYN);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Custom_sCustomArticleSelection", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
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

            return dataTable;
        }

        /// <summary>
        /// xp_Article_sArticleBOM 사용
        /// </summary>
        /// <param name="strArticleGrpID"></param>
        /// <param name="strArticleID"></param>
        /// <param name="strDirection"></param>
        /// <param name="strNotUse"></param>
        /// <returns></returns>
        public DataTable GetArticleBOM(string strArticleGrpID, string strArticleID, string strDirection, string strNotUse)
        {
            DataTable dataTable = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sArticleGrpID", strArticleGrpID);
                sqlParameter.Add("sArticleID", strArticleID);
                sqlParameter.Add("sDirection", strDirection);
                sqlParameter.Add("sIncNotuse", strNotUse);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Article_sArticleBOM", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
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

            return dataTable;
        }





        /// <summary>
        /// 해당 Table의 최대값에+1 한 ID 리턴
        /// </summary>
        public string GetMaxValue(string Field, string Table)
        {
            string ReturnString = string.Empty;
            try
            {
                string sql = "SELECT MAX(" + Field + ") FROM  " + Table + " ";
                sql += " where 1=1 ";
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count >= 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        ReturnString = drc[dt.Rows.Count - 1].ItemArray[0].ToString();
                        if (Lib.Instance.IsIntOrAnother(ReturnString))
                        {
                            ReturnString = string.Format("{0,2:00}", (int.Parse(ReturnString) + 1));
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

            return ReturnString;
        }

        /// <summary>
        /// 변수가 하나인 삭제
        /// </summary>
        public bool DeleteData(string strID, string strParameter, string strDelProcedure)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add(strParameter, strID);

                string[] result = DataStore.Instance.ExecuteProcedure(strDelProcedure, sqlParameter, false);

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
                }
                else
                {
                    MessageBox.Show("삭제 실패 , 내용 : " + result[1]);
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

        /// <summary>
        /// 변수가 2개인 삭제
        /// </summary>
        public bool DeleteData(string strID, string strOneVariable, string strParameter, string strOneParameter, string strDelProcedure)
        {
            bool flag = false;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add(strParameter, strID);
                sqlParameter.Add(strOneParameter, strOneVariable);

                string[] result = DataStore.Instance.ExecuteProcedure(strDelProcedure, sqlParameter, false);

                if (result[0].Equals("success"))
                {
                    //MessageBox.Show("성공 *^^*");
                    flag = true;
                }
                else
                {
                    MessageBox.Show("삭제 실패 , 내용 : " + result[1]);
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

        /// <summary>
        /// xp_Code_sDefect 사용
        /// </summary>
        public DataTable GetDefect()
        {
            DataTable dataTable = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sBasisID", "%");
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sDefect", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
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

            return dataTable;
        }

        /// <summary>
        /// xp_Code_sBasis 사용
        /// </summary>
        /// <returns></returns>
        public DataTable GetBasis()
        {
            DataTable dataTable = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sBasis", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
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

            return dataTable;
        }

        /// <summary>
        /// xp_Code_sGrade 사용
        /// </summary>
        public DataTable GetGrade()
        {
            DataTable dataTable = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sGrade", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
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

            return dataTable;
        }

        /// <summary>
        /// xp_Order_sArticleData 사용
        /// </summary>
        public DataTable GetArticleData(string strArticleID)
        {
            DataTable dataTable = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", strArticleID);
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleData", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
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

            return dataTable;
        }
        /// <summary>
        /// xp_Custom_sCustomArticle 사용
        /// </summary>
        /// <param name="strCustom"></param>
        /// <param name="strCustomID"></param>
        /// <param name="strArticleID"></param>
        /// <param name="strYN"></param>
        /// <returns></returns>
        public DataTable GetCustomArticle(int numCustom, string strCustomID, string strCustomGubun, int numArticleID, string strArticleID, string strYN, int numBuyCustom, string strBuyCustom)
        {
            DataTable dataTable = null;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("nCustom", numCustom);
                sqlParameter.Add("sCustom", strCustomID);
                sqlParameter.Add("sTradeID", strCustomGubun);
                sqlParameter.Add("nArticleID ", numArticleID);
                sqlParameter.Add("sArticleID", strArticleID);

                sqlParameter.Add("sShowArticleYN", strYN);
                sqlParameter.Add("nCustomBuyArticle", numBuyCustom);
                sqlParameter.Add("sCustomBuyArticle", strBuyCustom);



                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Custom_sCustomArticle", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
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

            return dataTable;
        }
    }

}
