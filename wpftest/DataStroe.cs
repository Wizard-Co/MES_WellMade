using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Windows;
using WizMes_WellMade.Properties;

//*******************************************************************************
//프로그램명    DataStore.cs
//메뉴ID        
//설명          데이터 처리 클래스
//작성일        2017.12.28
//개발자        박명회
//*******************************************************************************
// 변경일자     변경자      요청자      요구사항ID          요청 및 작업내용
//*******************************************************************************
//
//
//*******************************************************************************

namespace WizMes_WellMade
{
    /// <summary> 
    /// 데이터 처리 클래스   s
    /// </summary>
    public class DataStore
    {
        Lib lib = new Lib();
        //private string CONN_STR = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=nanokem;Password=nanokem;Initial Catalog=Nanokem_POP;Data Source=192.168.46.21"; //MSSQL ConnectionString <- 진짜
        private SqlConnection p_Connection;
        private SqlCommand p_Command;
        private SqlConnection L_Connection;
        private SqlCommand L_Command;
        private SqlConnection Z_Connection;
        private SqlCommand Z_Command;
        private const string ConnectionStringDataSource = "Data Source=";
        //private const string ConnectionStringCatalogAndID = ";Initial Catalog=KRBPOP_C;UID=";
        private const string ConnectionStringPWD = ";PWD=";
        private const string ConnectionStringTimeout = "; Connection Timeout= 0";

        //private string ConnectionString = "Data Source=wizis.iptime.org,20150;Initial Catalog=MES_DaeWon;UID=DBUser;PWD=Wizardis; Connection Timeout=0";
        //private string LogConnectionString = "Data Source=wizis.iptime.org,20150;Initial Catalog=WizLog;UID=DBUser;PWD=Wizardis; Connection Timeout=0";

#if DEBUG
        //private string ConnectionString = "Data Source=121.254.224.196, 20220;Initial Catalog=MES_WellMade;UID=DBUser;PWD=Wizardis; Connection Timeout=180";
        //private string LogConnectionString = "Data Source=121.254.224.196, 20220;Initial Catalog=WizLog;UID=DBUser;PWD=Wizardis; Connection Timeout=180";
        private string ConnectionString = "Data Source=wizis.iptime.org,20220;Initial Catalog=MES_WellMade;UID=DBUser;PWD=Wizardis; Connection Timeout=180";
        private string LogConnectionString = "Data Source=wizis.iptime.org,20220;Initial Catalog=WizLog;UID=DBUser;PWD=Wizardis; Connection Timeout=180";
        //private string ConnectionString = "Data Source=" + LoadINI.server + ";Initial Catalog=" + LoadINI.Database + ";UID=DBUser;PWD=Wizardis; Connection Timeout= 0";
        //private string LogConnectionString = "Data Source=" + LoadINI.server + ";Initial Catalog=WizLog;UID=DBUser;PWD=Wizardis; Connection Timeout= 0";
#else
        private string ConnectionString = "Data Source=" + LoadINI.server + ";Initial Catalog=" + LoadINI.Database + ";UID=DBUser;PWD=Wizardis; Connection Timeout= 0";
        private string LogConnectionString = "Data Source=" + LoadINI.server + ";Initial Catalog=WizLog;UID=DBUser;PWD=Wizardis; Connection Timeout= 0";
#endif

        private string ZipConnectionString = "Data Source=wizis.iptime.org,1433;Initial Catalog=ZipDB;UID= wizard;PWD=wizard2013; Connection Timeout= 0";

        private static DataStore p_dataStore = new DataStore();
        private static DataStore L_dataStore = new DataStore();
        private static DataStore Z_dataStore = new DataStore();

        public static DataStore Log_Instance
        {
            get
            {
                if (L_dataStore == null)
                {
                    L_dataStore = new DataStore();
                }

                if (L_dataStore.L_Connection == null)
                {
                    // 주소가 유효하지 않을경우 오류 발생
                    L_dataStore.L_Connection = new SqlConnection(L_dataStore.LogConnectionString);
                    L_dataStore.L_Command = L_dataStore.L_Connection.CreateCommand();
                }

                if (L_dataStore.L_Command == null)
                {
                    L_dataStore.L_Command = L_dataStore.L_Connection.CreateCommand();
                }

                return L_dataStore;
            }
        }
        public static DataStore Instance
        {
            get
            {

                if (p_dataStore == null)
                {
                    p_dataStore = new DataStore();
                }

                if (p_dataStore.p_Connection == null)
                {
                    // 주소가 유효하지 않을경우 오류 발생
                    p_dataStore.p_Connection = new SqlConnection(p_dataStore.ConnectionString);
                    p_dataStore.p_Command = p_dataStore.p_Connection.CreateCommand();
                }

                if (p_dataStore.p_Command == null)
                {
                    p_dataStore.p_Command = p_dataStore.p_Connection.CreateCommand();
                }

                return p_dataStore;
            }
        }

        /// <summary>
        /// 이건 주소 열때만 사용한다.
        /// </summary>
        public static DataStore Zip_Instance
        {
            get
            {

                if (Z_dataStore == null)
                {
                    Z_dataStore = new DataStore("", "");
                }

                if (Z_dataStore.Z_Connection == null)
                {
                    // 주소가 유효하지 않을경우 오류 발생
                    Z_dataStore.Z_Connection = new SqlConnection(Z_dataStore.ZipConnectionString);
                    Z_dataStore.Z_Command = Z_dataStore.Z_Connection.CreateCommand();
                }

                if (Z_dataStore.Z_Command == null)
                {
                    Z_dataStore.Z_Command = Z_dataStore.Z_Connection.CreateCommand();
                }

                return Z_dataStore;
            }
        }

        public SqlCommand Command
        {
            get { return p_Command; }
        }

        public DataStore()
        {
            p_Connection = new SqlConnection(ConnectionString);
            p_Command = p_Connection.CreateCommand();
            p_Command.CommandTimeout = 0;
        }

        public DataStore(string strID, string strPW)
        {
            p_Connection = new SqlConnection(ZipConnectionString);
            p_Command = p_Connection.CreateCommand();
            p_Command.CommandTimeout = 0;
        }

        public void SetConnectionString(string ipAddress, string id, string password, string catalog)
        {
            StringBuilder sb = new StringBuilder(ConnectionStringDataSource);
            sb.Append(ipAddress);
            sb.Append(catalog);
            sb.Append(id);
            sb.Append(ConnectionStringPWD);
            sb.Append(password);
            sb.Append(ConnectionStringTimeout);

            ConnectionString = sb.ToString();

            if (string.IsNullOrEmpty(p_Connection.ConnectionString) == false)
            {
                p_Connection.Close();
            }

            p_Connection.ConnectionString = ConnectionString;

            if (p_Connection.State == ConnectionState.Closed)
            {
                p_Connection.Open();
            }
        }

        public void CloseConnection()
        {
            if (p_Connection.State != ConnectionState.Closed)
            {
                if (p_Command.Transaction != null)
                {
                    p_Command.Transaction.Rollback();
                }

                p_Connection.Close();
            }
        }

#region Base Query

        /// <summary>
        /// 주소 검색시 사용
        /// </summary>
        /// <param name="queryString"></param>
        /// <returns></returns>
        public DataSet QueryToDataSetByZip(string queryString)
        {
            try
            {
                if (Z_Connection.State == ConnectionState.Closed)
                {
                    Z_Connection.Open();
                }

                Z_Command.CommandText = queryString;
                Z_Command.CommandType = CommandType.Text;

                SqlDataAdapter adapter = new SqlDataAdapter(Z_Command);
                DataSet dataset = new DataSet();

                adapter.Fill(dataset);

                adapter.Dispose();

                return dataset;
            }
            catch (Exception e)
            {
                throw e;
            }
            //finally
            //{
            //    if (p_Connection.State != ConnectionState.Closed)
            //    {
            //        p_Connection.Close();
            //    }
            //}
        }

        public DataSet QueryToDataSet(string queryString)
        {
            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                p_Command.CommandText = queryString;
                p_Command.CommandType = CommandType.Text;

                SqlDataAdapter adapter = new SqlDataAdapter(p_Command);
                DataSet dataset = new DataSet();

                adapter.Fill(dataset);

                adapter.Dispose();

                return dataset;
            }
            catch (Exception e)
            {
                throw e;
            }
            //finally
            //{
            //    if (p_Connection.State != ConnectionState.Closed)
            //    {
            //        p_Connection.Close();
            //    }
            //}
        }

        public int QueryToInt32(string queryString)
        {
            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                p_Command.CommandText = queryString;
                int retVal = ((Int32?)p_Command.ExecuteScalar()) ?? 0;

                return retVal;
            }
            catch (Exception e)
            {
                throw e;
            }
            //finally
            //{
            //    if (p_Connection.State != ConnectionState.Closed)
            //    {
            //        p_Connection.Close();
            //    }
            //}
        }

        public object QueryToScalar(string queryString)
        {
            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                p_Command.CommandText = queryString;
                object retVal = p_Command.ExecuteScalar();

                return retVal;
            }
            catch (Exception e)
            {
                throw e;
            }
            //finally
            //{
            //    if (p_Connection.State != ConnectionState.Closed)
            //    {
            //        p_Connection.Close();
            //    }
            //}
        }

#endregion



#region 쿼리공통

        /// <summary>
        /// ProcedureToDataSet
        /// </summary>
        /// <param name="procedureName"></param>
        /// <param name="sqlParameter"></param>
        /// <returns></returns>
        public DataSet ProcedureToDataSetByZip(string procedureName, Dictionary<string, object> sqlParameter, bool logOn)
        {
            try
            {
                //Cursor.Current = Cursors.WaitCursor;

                if (Z_Connection.State == ConnectionState.Closed)
                {
                    Z_Connection.Open();
                }

                if (logOn == true)
                {
                    // DB Log를 남긴다.
                    StringBuilder trxCommand = new StringBuilder(procedureName);

                    if (p_Command.Parameters.Count > 0)
                    {
                        trxCommand.Append(" ");

                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            trxCommand.Append(kvp.Key + " = " + kvp.Value.ToString());
                            trxCommand.Append(", ");
                        }

                        trxCommand.Remove(trxCommand.Length - 2, 2);
                    }
                    //InsertTrxLogByUserID(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod());
                    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), trxCommand.ToString());
                }

                Z_Command.CommandText = procedureName;
                Z_Command.CommandType = CommandType.StoredProcedure;
                Z_Command.Parameters.Clear();


                if (sqlParameter != null)
                {
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        Z_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }
                }

                SqlDataAdapter adapter = new SqlDataAdapter(Z_Command);
                DataSet dataset = new DataSet();
                adapter.Fill(dataset);
                adapter.Dispose();


                return dataset;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, Resources.MSG_CAPTION_ERROR);
                //MessageBox.Show(e.Message, Resources.MSG_CAPTION_ERROR, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                //Cursor.Current = Cursors.Default;
                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }
        }

        /// <summary>
        /// ProcedureToDataSet
        /// </summary>
        /// <param name="procedureName"></param>
        /// <param name="sqlParameter"></param>
        /// <returns></returns>
        public DataSet ProcedureToDataSet(string procedureName, Dictionary<string, object> sqlParameter, bool logOn)
        {
            try
            {
                //Cursor.Current = Cursors.WaitCursor;

                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                if (logOn == true)
                {
                    // DB Log를 남긴다.
                    StringBuilder trxCommand = new StringBuilder(procedureName);

                    if (p_Command.Parameters.Count > 0)
                    {
                        trxCommand.Append(" ");

                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            trxCommand.Append(kvp.Key + " = " + kvp.Value.ToString());
                            trxCommand.Append(", ");
                        }

                        trxCommand.Remove(trxCommand.Length - 2, 2);
                    }

                    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), trxCommand.ToString());
                }

                p_Command.CommandText = procedureName;
                p_Command.CommandType = CommandType.StoredProcedure;
                p_Command.Parameters.Clear();


                if (sqlParameter != null)
                {
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }
                }

                SqlDataAdapter adapter = new SqlDataAdapter(p_Command);
                DataSet dataset = new DataSet();
                adapter.Fill(dataset);
                adapter.Dispose();


                return dataset;
            }
            catch (Exception e)

            {
                MessageBox.Show(e.Message, Resources.MSG_CAPTION_ERROR);
                //MessageBox.Show(e.Message, Resources.MSG_CAPTION_ERROR, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                //Cursor.Current = Cursors.Default;
                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }
        }

        /// <summary>
        /// 실행쿼리
        /// </summary>
        /// <param name="userId"></param>
        /// <returns></returns>
        public int ExecSQL(string sql, bool logOn)
        {
            if (logOn == true)
            {
                // DB Log를 남긴다.
                InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), sql);
            }

            int value = QueryToInt32(sql);

            return value;
        }

        public object ExecuteScalar(string sql, bool logOn)
        {
            if (logOn == true)
            {
                // DB Log를 남긴다.
                InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), sql);
            }

            object value = QueryToScalar(sql);

            return value;
        }

        public string[] ExecuteQuery(string queryString, bool logOn)
        {
            SqlTransaction transaction = null;

            try
            {

                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }



                if (logOn == true)
                {
                    // DB Log를 남긴다.
                    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), queryString);
                }


                p_Command.CommandText = queryString;
                p_Command.CommandType = CommandType.Text;

                transaction = p_Connection.BeginTransaction();
                p_Command.Transaction = transaction;

                string value = p_Command.ExecuteScalar().ToString();
                transaction.Commit();

                return new string[] { Resources.success, value };
            }
            catch (NullReferenceException)
            {
                return new String[] { Resources.success, "NullReferenceException" };
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                return new string[] { Resources.failure, ex.Message };
            }
            finally
            {
                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }

        }

        public string[] ExecuteProcedure(string procedureName, Dictionary<string, object> sqlParameter, bool logOn)
        {
            SqlTransaction transaction = null;

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }



                if (logOn == true)
                {
                    // DB Log를 남긴다.
                    StringBuilder trxCommand = new StringBuilder(procedureName);

                    if (p_Command.Parameters.Count > 0)
                    {
                        trxCommand.Append(" ");

                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            trxCommand.Append(kvp.Key + " = " + kvp.Value.ToString());
                            trxCommand.Append(", ");
                        }

                        trxCommand.Remove(trxCommand.Length - 2, 2);
                    }
                    InsertTrxLogByUserID(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod()); //2021-10-27 로그 남기는 함수 변경
                    //InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), trxCommand.ToString());
                }



                transaction = p_Connection.BeginTransaction();
                p_Command.Transaction = transaction;

                p_Command.CommandText = procedureName;
                p_Command.CommandType = CommandType.StoredProcedure;
                p_Command.Parameters.Clear();


                if (sqlParameter != null)
                {
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }
                }

                string value = Convert.ToString(p_Command.ExecuteScalar());

#region 출하처리에 프로시저 예외처리 0일경우 commit 나머지 rollback
                string[] valueSplit = value.Split(',');
                if (valueSplit.Length > 1)
                {
                    if (valueSplit[0] == "0")
                    {
                        transaction.Commit();
                        return new string[] { "success", value };
                    }
                    else
                    {
                        throw new Exception(value);
                    }
                }
#endregion

                transaction.Commit();

                return new String[] { Resources.success, value };   //성공! 쿼리에서 리턴값이 있을경우
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                if (transaction != null)
                {
                    transaction.Commit();
                }

                return new String[] { Resources.success, "NullReferenceException" };
            }
            catch (Exception ex)
            {
                try
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }

                    return new string[] { Resources.failure, ex.Message };
                }
                catch (Exception ex1)
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }

                    return new string[] { Resources.failure, ex.Message + "/" + ex1.Message };
                }
            }
            finally
            {
                if (p_Connection.State != ConnectionState.Closed)
                {
                    p_Connection.Close();
                }
            }
        }

        public string[] ExecuteProcedureAll(string[] procedureNameAll, Dictionary<string, object>[] sqlParameterall, bool logOn)
        {
            SqlTransaction transaction = null;
            string value = "";

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                if (logOn == true)
                {
                    // DB Log를 남긴다.
                    //StringBuilder trxCommand = new StringBuilder(procedureName);

                    //if (p_Command.Parameters.Count > 0)
                    //{
                    //    trxCommand.Append(" ");

                    //    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    //    {
                    //        trxCommand.Append(kvp.Key + " = " + kvp.Value.ToString());
                    //        trxCommand.Append(", ");
                    //    }

                    //    trxCommand.Remove(trxCommand.Length - 2, 2);
                    //}

                    //InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), trxCommand.ToString());
                }

                transaction = p_Connection.BeginTransaction();
                p_Command.Transaction = transaction;

                foreach (string procedureName in procedureNameAll)
                {
                    p_Command.CommandText = procedureName;
                    p_Command.CommandType = CommandType.StoredProcedure;
                    p_Command.Parameters.Clear();


                    if (sqlParameterall != null)
                    {
                        foreach (Dictionary<string, object> sqlParameter in sqlParameterall)
                        {
                            foreach (KeyValuePair<string, object> kvp in sqlParameter)
                            {
                                p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                            }
                        }
                    }

                    value = Convert.ToString(p_Command.ExecuteScalar());
                }

                transaction.Commit();

                return new String[] { Resources.success, value };   //성공! 쿼리에서 리턴값이 있을경우
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                if (transaction != null)
                {
                    transaction.Commit();
                }

                return new String[] { Resources.success, "NullReferenceException" };
            }
            catch (Exception ex)
            {
                try
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }

                    return new string[] { Resources.failure, ex.Message };
                }
                catch (Exception ex1)
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }

                    return new string[] { Resources.failure, ex.Message + "/" + ex1.Message };
                }
            }
            finally
            {
                if (p_Connection.State != ConnectionState.Closed)
                {
                    p_Connection.Close();
                }
                CloseConnection();
            }
        }

        /// <summary>
        /// Procedure에 output Parameter가 있을 때 사용함, 트랜젝션 포함
        /// </summary>
        /// <param name="procedureName">호출할 Procedure 명</param>
        /// <param name="sqlParameter">Procedure로 전달할 변수들</param>
        /// <param name="outputParameters">Output으로 지정된 변수들</param>
        /// <param name="okValues">Output으로 넘어온 값들이 정상인지 판단하는 기준, 트랜젝션 commit/rollback의 기준이 됨</param>
        /// <returns>outputParameter별 값과 Result, Message Key가 추가됨</returns>
        public Dictionary<string, string> ExecuteProcedureOutput(string procedureName, Dictionary<string, object> sqlParameter, List<string> outputParameters, Dictionary<string, string> okValues, bool logOn)
        {
            // Output 결과 값을 넣을 Dictionary
            Dictionary<string, string> outputResult = new Dictionary<string, string>();
            SqlTransaction transaction = null;

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }



                if (logOn == true)
                {
                    // DB Log를 남긴다.
                    StringBuilder trxCommand = new StringBuilder(procedureName);

                    if (p_Command.Parameters.Count > 0)
                    {
                        trxCommand.Append(" ");

                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            trxCommand.Append(kvp.Key + " = " + kvp.Value.ToString());
                            trxCommand.Append(", ");
                        }

                        trxCommand.Remove(trxCommand.Length - 2, 2);
                    }

                    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), trxCommand.ToString());
                }



                transaction = p_Connection.BeginTransaction();
                p_Command.Transaction = transaction;

                p_Command.CommandText = procedureName;
                p_Command.CommandType = CommandType.StoredProcedure;
                p_Command.Parameters.Clear();


                if (sqlParameter != null)
                {
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }

                    // Output Parameter 지정 및 output 값 받을 Dictionary 준비
                    foreach (string parameter in outputParameters)
                    {
                        p_Command.Parameters[parameter].Direction = ParameterDirection.Output;
                        outputResult.Add(parameter, "");
                    }
                }

                string value = Convert.ToString(p_Command.ExecuteScalar());

                //output 값 Dictionary에 저장
                foreach (string parameter in outputParameters)
                {
                    outputResult[parameter] = p_Command.Parameters[parameter].Value.ToString();
                }

                // 기준값과 비교하여 트랜젝션을 Commit할 것인지 RollBack할 것인지 결정
                bool isOK = true;

                foreach (KeyValuePair<string, string> kvp in okValues)
                {
                    if (outputResult[kvp.Key].Equals(kvp.Value) == false)
                    {
                        isOK = false;
                    }
                }


                if (isOK == true)
                {
                    outputResult.Add("Result", "success");
                    outputResult.Add("Message", value);
                    transaction.Commit();
                }
                else
                {
                    outputResult.Add("Result", "failure");
                    outputResult.Add("Message", value);
                    transaction.Rollback();
                }


                return outputResult;
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                //output 값 Dictionary에 저장
                foreach (string parameter in outputParameters)
                {
                    outputResult[parameter] = p_Command.Parameters[parameter].Value.ToString();
                }

                // 기준값과 비교하여 트랜젝션을 Commit할 것인지 RollBack할 것인지 결정
                bool isOK = true;

                foreach (KeyValuePair<string, string> kvp in okValues)
                {
                    if (outputResult[kvp.Key].Equals(kvp.Value) == false)
                    {
                        isOK = false;
                    }
                }


                if (isOK == true)
                {
                    outputResult.Add("Result", "success");
                    outputResult.Add("Message", "NullReferenceException");
                    transaction.Commit();
                }
                else
                {
                    outputResult.Add("Result", "failure");
                    outputResult.Add("Message", "NullReferenceException");
                    transaction.Rollback();
                }


                return outputResult;
            }
            catch (Exception ex)
            {
                transaction.Rollback();

                try
                {
                    outputResult.Clear();
                    outputResult.Add("Result", "failure");
                    outputResult.Add("Message", ex.Message);
                    return outputResult;
                }
                catch (Exception ex1)
                {
                    outputResult.Clear();
                    outputResult.Add("Result", "failure");
                    outputResult.Add("Message", ex1.Message);
                    return outputResult;
                }
            }
            finally
            {
                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }
        }
        public Dictionary<string, string> ExecuteProcedureOutputNoTran(string procedureName, Dictionary<string, object> sqlParameter, Dictionary<string, int> outputParameters, bool logOn)
        {
            // Output 결과 값을 넣을 Dictionary
            Dictionary<string, string> outputResult = new Dictionary<string, string>();

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }



                if (logOn == true)
                {
                    // DB Log를 남긴다.
                    StringBuilder trxCommand = new StringBuilder(procedureName);

                    if (p_Command.Parameters.Count > 0)
                    {
                        trxCommand.Append(" ");

                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            trxCommand.Append(kvp.Key + " = " + kvp.Value.ToString());
                            trxCommand.Append(", ");
                        }

                        trxCommand.Remove(trxCommand.Length - 2, 2);
                    }

                    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), trxCommand.ToString());
                }



                p_Command.CommandText = procedureName;
                p_Command.CommandType = CommandType.StoredProcedure;
                p_Command.Parameters.Clear();


                if (sqlParameter != null)
                {
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }

                    // Output Parameter 지정 및 output 값 받을 Dictionary 준비
                    foreach (KeyValuePair<string, int> kvp in outputParameters)
                    {
                        p_Command.Parameters[kvp.Key].Direction = ParameterDirection.Output;
                        p_Command.Parameters[kvp.Key].Size = kvp.Value;
                        outputResult.Add(kvp.Key, "");
                    }
                }

                string value = Convert.ToString(p_Command.ExecuteScalar());

                //output 값 Dictionary에 저장
                foreach (KeyValuePair<string, int> kvp in outputParameters)
                {
                    outputResult[kvp.Key] = p_Command.Parameters[kvp.Key].Value.ToString();
                }

                return outputResult;
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                //output 값 Dictionary에 저장
                foreach (KeyValuePair<string, int> kvp in outputParameters)
                {
                    outputResult[kvp.Key] = p_Command.Parameters[kvp.Key].Value.ToString();
                }


                return outputResult;
                //return null;
            }
            catch (Exception ex)
            {
                try
                {
                    outputResult.Clear();
                    List<string> result = new List<string>();
                    result.Add("9999");
                    result.Add(ex.Message);
                    result.Add(ex.StackTrace);

                    int i = 0;

                    foreach (KeyValuePair<string, int> kvp in outputParameters)
                    {
                        outputResult[kvp.Key] = result.Count > i ? result[i++] : "";
                    }

                    return outputResult;
                }
                catch (Exception ex1)
                {
                    outputResult.Clear();
                    List<string> result = new List<string>();
                    result.Add("9998");
                    result.Add(ex1.Message);
                    result.Add(ex1.StackTrace);

                    int i = 0;

                    foreach (KeyValuePair<string, int> kvp in outputParameters)
                    {
                        outputResult[kvp.Key] = result.Count > i ? result[i++] : "";
                    }

                    return outputResult;
                }
            }
            finally
            {
                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }
        }


        public List<KeyValue> ExecuteAllProcedureOutputGetCS(List<Procedure> AllProcedure, List<Dictionary<string, object>> sqlParameterall)
        {
            // Output 결과 값을 넣을 List
            List<KeyValue> outputVal = new List<KeyValue>();
            SqlTransaction transaction = null;
            string value = "";
            bool complete = false;
            List<KeyValue> Success_List = new List<KeyValue>();//추가/180427

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                transaction = p_Connection.BeginTransaction();
                p_Command.Transaction = transaction;

                foreach (Procedure Procedure in AllProcedure)
                {
                    if (sqlParameterall[AllProcedure.IndexOf(Procedure)] != null)
                    {
                        Dictionary<string, object> sqlParameter = sqlParameterall[AllProcedure.IndexOf(Procedure)];

                        if (Procedure.OutputUseYN == "Y")//리턴받는 output값이 있을때
                        {
                            if (outputVal.Count > 0)
                            {
                                //해당 프로시저의 output으로 리턴받는 값과 동일한 값이 있을 경우 output값 리스트에서 삭제
                                for (int i = outputVal.Count - 1; i >= 0; i--)
                                {
                                    KeyValue kvp = outputVal[i];
                                    if (kvp.key.ToLower().ToString() == Procedure.OutputName.ToLower().ToString())
                                    {
                                        outputVal.Remove(kvp);
                                    }
                                }
                            }
                            else
                            {
                                //output값 리스트에 추가
                                KeyValue kvp = new KeyValue();
                                kvp.key = Procedure.OutputName;
                                kvp.value = "";
                                outputVal.Add(kvp);
                            }
                        }
                        ///2018.02.02 로그용 파라미터 값 세팅
                        p_Command.CommandText = Procedure.Name;               //프로시저명 입력
                        p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
                        p_Command.Parameters.Clear();                           //이전 파라미터 클리어
                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            complete = false;

                            foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
                            {
                                if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
                                {
                                    if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
                                    {
                                        p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
                                        complete = true;
                                    }
                                }
                            }

                            if (!complete)
                            {
                                p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                            }
                        }
                        //로그 메서드

                        InsertTrxLogByUserID(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod());

                        p_Command.CommandText = Procedure.Name;                 //프로시저 이름 셋팅
                        p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
                        p_Command.Parameters.Clear();                           //로그용 파라미터 클리어

                        //입력할 데이터 파라미터 셋팅
                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            complete = false;

                            foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
                            {
                                if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
                                {
                                    if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
                                    {
                                        p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
                                        complete = true;
                                    }
                                }
                            }

                            if (!complete)
                            {
                                p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                            }
                        }

                        if (Procedure.OutputUseYN == "Y")
                        {
                            p_Command.Parameters[Procedure.OutputName].Direction = ParameterDirection.Output;
                            p_Command.Parameters[Procedure.OutputName].Size = int.Parse(Procedure.OutputLength);
                        }
                    }

                    value = Convert.ToString(p_Command.ExecuteScalar());

                    if (Procedure.OutputUseYN == "Y")
                    {
                        complete = false;                                                                       //완료여부

                        foreach (KeyValue mKeyValue in outputVal)                                               //output값 리스트중에서
                        {
                            if (mKeyValue.key == Procedure.OutputName)                                              //같은 이름을 가진 Key값이 리스트에 있을경우
                            {
                                mKeyValue.value = p_Command.Parameters[Procedure.OutputName].Value.ToString();      //해당 리스트에 값 추가
                                complete = true;
                                Success_List.Add(mKeyValue);//추가/180427
                                break;
                            }
                        }
                        if (!complete)
                        {
                            KeyValue kvp = new KeyValue();
                            kvp.key = Procedure.OutputName;                                                     //새로운 output값 이름의 리스트 생성
                            kvp.value = p_Command.Parameters[Procedure.OutputName].Value.ToString();            //새로운 output값 밸류 추가
                            outputVal.Add(kvp);                                                             //output값 리스트에 추가
                            Success_List.Add(kvp);//추가/180427
                        }
                    }
                }
                //추가 /180427
                KeyValue suc_kv = new KeyValue();
                suc_kv.key = "Success";
                suc_kv.value = "";
                Success_List.Insert(0, suc_kv);

                transaction.Commit();

                return Success_List; //추가 /180427
                //return new String[] { "success", value };   //성공! 쿼리에서 리턴값이 있을경우
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                if (transaction != null)
                {
                    transaction.Commit();
                }

                //추가 /180427
                KeyValue suc_kv = new KeyValue();
                suc_kv.key = "Success";
                suc_kv.value = "NullReferenceException";
                Success_List.Add(suc_kv);
                //Success_List.AddRange(outputVal);

                return Success_List;


            }
            catch (Exception ex)
            {
                try
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }
                    //exception용 로그 메서드
                    InsertTrxLogByUserIDErrLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), ex.Message);

                    KeyValue suc_kv = new KeyValue();
                    suc_kv.key = "failure";
                    suc_kv.value = ex.Message;
                    Success_List.Add(suc_kv);

                    return Success_List;
                }
                catch (Exception ex1)
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }
                    //exception용 로그 메서드2
                    InsertTrxLogByUserIDErrLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), ex1.Message);

                    KeyValue suc_kv = new KeyValue();
                    suc_kv.key = "failure";
                    suc_kv.value = ex1.Message;
                    Success_List.Add(suc_kv);

                    return Success_List;
                }
            }
            finally
            {
                if (p_Connection.State != ConnectionState.Closed)
                {
                    p_Connection.Close();
                }
                CloseConnection();
            }
        }

        public List<KeyValue> ExecuteAllProcedureOutputListGetCS(List<Procedure> AllProcedure, List<Dictionary<string, object>> sqlParameterall)
        {
            // Output 결과 값을 넣을 List
            List<KeyValue> outputVal = new List<KeyValue>();
            SqlTransaction transaction = null;
            string value = "";
            bool complete = false;
            List<KeyValue> Success_List = new List<KeyValue>();//추가/180427

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                transaction = p_Connection.BeginTransaction();
                p_Command.Transaction = transaction;

                foreach (Procedure Procedure in AllProcedure)
                {
                    if (sqlParameterall[AllProcedure.IndexOf(Procedure)] != null)
                    {
                        Dictionary<string, object> sqlParameter = sqlParameterall[AllProcedure.IndexOf(Procedure)];

                        if (Procedure.OutputUseYN == "Y")//리턴받는 output값이 있을때
                        {
                            if (outputVal.Count > 0)
                            {
                                //해당 프로시저의 output으로 리턴받는 값과 동일한 값이 있을 경우 output값 리스트에서 삭제
                                for (int i = outputVal.Count - 1; i >= 0; i--)
                                {
                                    KeyValue kvp = outputVal[i];
                                    if (kvp.key.ToLower().ToString() == Procedure.OutputName.ToLower().ToString())
                                    {
                                        outputVal.Remove(kvp);
                                    }
                                }
                            }
                            else
                            {
                                //output값 리스트에 추가
                                if (Procedure.list_OutputName.Count > 0)
                                {
                                    foreach (string str in Procedure.list_OutputName)
                                    {
                                        KeyValue kvp = new KeyValue();
                                        kvp.key = str;
                                        kvp.value = "";
                                        outputVal.Add(kvp);
                                    }
                                }

                                //KeyValue kvp = new KeyValue();
                                //kvp.key = Procedure.OutputName;
                                //kvp.value = "";
                                //outputVal.Add(kvp);
                            }
                        }
                        ///2018.02.02 로그용 파라미터 값 세팅
                        p_Command.CommandText = Procedure.Name;               //프로시저명 입력
                        p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
                        p_Command.Parameters.Clear();                           //이전 파라미터 클리어
                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            complete = false;

                            foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
                            {
                                if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
                                {
                                    if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
                                    {
                                        p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
                                        complete = true;
                                    }
                                }
                            }

                            if (!complete)
                            {
                                p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                            }
                        }
                        //로그 메서드

                        InsertTrxLogByUserID(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod());

                        p_Command.CommandText = Procedure.Name;                 //프로시저 이름 셋팅
                        p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
                        p_Command.Parameters.Clear();                           //로그용 파라미터 클리어

                        //입력할 데이터 파라미터 셋팅
                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            complete = false;

                            foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
                            {
                                if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
                                {
                                    if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
                                    {
                                        p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
                                        complete = true;
                                    }
                                }
                            }

                            if (!complete)
                            {
                                p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                            }
                        }

                        if (Procedure.OutputUseYN == "Y")
                        {
                            //p_Command.Parameters[Procedure.OutputName].Direction = ParameterDirection.Output;
                            //p_Command.Parameters[Procedure.OutputName].Size = int.Parse(Procedure.OutputLength);
                            //foreach (string pro_outputname in Procedure.list_OutputName)
                            //{
                            //    p_Command.Parameters[pro_outputname].Direction = ParameterDirection.Output;
                            //    p_Command.Parameters[pro_outputname].Size = int.Parse(Procedure.OutputLength);
                            //}

                            for (int i = 0; i < Procedure.list_OutputName.Count; i++)
                            {
                                p_Command.Parameters[Procedure.list_OutputName[i]].Direction = ParameterDirection.Output;
                                p_Command.Parameters[Procedure.list_OutputName[i]].Size = int.Parse(Procedure.list_OutputLength[i]);
                            }
                        }
                    }

                    value = Convert.ToString(p_Command.ExecuteScalar());

                    if (Procedure.OutputUseYN == "Y")
                    {
                        complete = false;                                                                       //완료여부

                        foreach (KeyValue mKeyValue in outputVal)                                               //output값 리스트중에서
                        {
                            foreach (string pro_outputname in Procedure.list_OutputName)
                            {
                                if (mKeyValue.key == pro_outputname)
                                {
                                    mKeyValue.value = p_Command.Parameters[pro_outputname].Value.ToString();      //해당 리스트에 값 추가
                                    complete = true;
                                    Success_List.Add(mKeyValue);//추가/180427
                                }
                            }

                            //if (mKeyValue.key == Procedure.OutputName)                                              //같은 이름을 가진 Key값이 리스트에 있을경우
                            //{
                            //    mKeyValue.value = p_Command.Parameters[Procedure.OutputName].Value.ToString();      //해당 리스트에 값 추가
                            //    complete = true;
                            //    Success_List.Add(mKeyValue);//추가/180427
                            //    break;
                            //}
                        }
                        if (!complete)
                        {
                            foreach (string pro_outputname in Procedure.list_OutputName)
                            {
                                KeyValue kvp = new KeyValue();
                                kvp.key = pro_outputname;                                                     //새로운 output값 이름의 리스트 생성
                                kvp.value = p_Command.Parameters[pro_outputname].Value.ToString();            //새로운 output값 밸류 추가
                                outputVal.Add(kvp);                                                             //output값 리스트에 추가
                                Success_List.Add(kvp);//추가/180427
                            }
                            //KeyValue kvp = new KeyValue();
                            //kvp.key = Procedure.OutputName;                                                     //새로운 output값 이름의 리스트 생성
                            //kvp.value = p_Command.Parameters[Procedure.OutputName].Value.ToString();            //새로운 output값 밸류 추가
                            //outputVal.Add(kvp);                                                             //output값 리스트에 추가
                            //Success_List.Add(kvp);//추가/180427
                        }
                    }
                }
                //추가 /180427
                KeyValue suc_kv = new KeyValue();
                suc_kv.key = "Success";
                suc_kv.value = "";
                Success_List.Insert(0, suc_kv);

                transaction.Commit();

                return Success_List; //추가 /180427
                //return new String[] { "success", value };   //성공! 쿼리에서 리턴값이 있을경우
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                if (transaction != null)
                {
                    transaction.Commit();
                }

                //추가 /180427
                KeyValue suc_kv = new KeyValue();
                suc_kv.key = "Success";
                suc_kv.value = "NullReferenceException";
                Success_List.Insert(0, suc_kv);
                //Success_List.AddRange(outputVal);

                return Success_List;


            }
            catch (Exception ex)
            {
                try
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }
                    //exception용 로그 메서드
                    InsertTrxLogByUserIDErrLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), ex.Message);

                    KeyValue suc_kv = new KeyValue();
                    suc_kv.key = "failure";
                    suc_kv.value = ex.Message;
                    Success_List.Insert(0, suc_kv);

                    return Success_List;
                }
                catch (Exception ex1)
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }
                    //exception용 로그 메서드2
                    InsertTrxLogByUserIDErrLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), ex1.Message);

                    KeyValue suc_kv = new KeyValue();
                    suc_kv.key = "failure";
                    suc_kv.value = ex1.Message;
                    Success_List.Insert(0, suc_kv);

                    return Success_List;
                }
            }
            finally
            {
                if (p_Connection.State != ConnectionState.Closed)
                {
                    p_Connection.Close();
                }
                CloseConnection();
            }
        }

        /// <summary>
        /// Transaction 없이 프로시져 실행
        /// Parent/Child 함께 적용되어야 할 때 사용
        /// </summary>
        /// <param name="procedureName"></param>
        /// <param name="sqlParameter"></param>
        /// <param name="logOn"></param>
        /// <returns></returns>
        public string[] ExecuteProcedureWithoutTransaction(string procedureName, Dictionary<string, object> sqlParameter, bool logOn)
        {
            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }




                if (logOn == true)
                {
                    // DB Log를 남긴다.
                    StringBuilder trxCommand = new StringBuilder(procedureName);

                    if (p_Command.Parameters.Count > 0)
                    {
                        trxCommand.Append(" ");

                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            trxCommand.Append(kvp.Key + " = " + kvp.Value.ToString());
                            trxCommand.Append(", ");
                        }

                        trxCommand.Remove(trxCommand.Length - 2, 2);
                    }

                    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), trxCommand.ToString());
                }




                p_Command.CommandText = procedureName;
                p_Command.CommandType = CommandType.StoredProcedure;
                p_Command.Parameters.Clear();


                if (sqlParameter != null)
                {
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }
                }

                string value = Convert.ToString(p_Command.ExecuteScalar());

                return new String[] { Resources.success, value };   //성공! 쿼리에서 리턴값이 있을경우
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                return new String[] { Resources.success, "NullReferenceException" };
            }
            catch (Exception ex)
            {
                try
                {
                    return new string[] { Resources.failure, ex.Message };
                }
                catch (Exception ex1)
                {
                    return new string[] { Resources.failure, ex.Message + "/" + ex1.Message };
                }
            }
        }

        /// <summary>
        /// 트랜잭션 단위로 프로시저 실행
        /// </summary>
        /// <param name="procedureName"> 파라메타 이름</param>
        /// <param name="sqlParameter"> 파라메타 변수</param>
        /// 실패하면 롤백이후 에러메세지
        /// 성공하면 성공 메세지
        public string[] ExecuteTranProcedure(string procedureName, Dictionary<string, object> sqlParameter, Boolean logOn)
        {
            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }




                if (logOn == true)
                {
                    // DB Log를 남긴다.
                    StringBuilder trxCommand = new StringBuilder(procedureName);

                    if (p_Command.Parameters.Count > 0)
                    {
                        trxCommand.Append(" ");

                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            trxCommand.Append(kvp.Key + " = " + kvp.Value.ToString());
                            trxCommand.Append(", ");
                        }

                        trxCommand.Remove(trxCommand.Length - 2, 2);
                    }

                    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), trxCommand.ToString());
                }




                //트랜잭션 단위로 실행
                p_Command.Transaction = p_Command.Connection.BeginTransaction();

                p_Command.CommandText = procedureName;
                p_Command.CommandType = CommandType.StoredProcedure;
                p_Command.Parameters.Clear();


                if (sqlParameter != null)
                {
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }
                }

                string value = p_Command.ExecuteNonQuery().ToString();
                //string value = p_Command.ExecuteScalar().ToString();

                //트랜잭션 commit
                p_Command.Transaction.Commit();

                return new string[] { Resources.success, value };
            }
            catch (NullReferenceException)
            {
                if (p_Command.Transaction != null)
                {
                    p_Command.Transaction.Commit();
                }

                return new String[] { Resources.success, "NullReferenceException" };
            }
            catch (Exception ex)
            {
                if (p_Command.Transaction != null)
                {
                    //오류 발생시 Rollback
                    p_Command.Transaction.Rollback();
                }
                return new String[] { Resources.failure, ex.Message };
            }
            finally
            {
                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }
        }

        /// <summary>
        /// 스토어프로시져의 이름을 이용하여 파라메터 배열을 만든다
        /// </summary>
        /// <param name="spName">스토어프로시져이름</param>
        /// <param name="includeReturnValueParameter"> RETURN_VALUE 를 파라메터에 포함시킬것인가?</param>
        /// <returns>파라메테 배열</returns>
        public DbParameter[] DiscoverSpParameterSet(string spName, bool includeReturnValueParameter)
        {

            if (spName == null || spName.Length == 0)
            {
                throw new ArgumentNullException("프로시져이름이 없습니다.");
            }

            if (p_Connection.State == ConnectionState.Closed)
            {
                p_Connection.Open();
            }

            //OleDB관련 처리 필요할 수도 있음
            //if (p_Connection is System.Data.OleDb.OleDbConnection) 
            //{
            //    System.Data.OleDb.OleDbCommandBuilder.DeriveParameters((System.Data.OleDb.OleDbCommand)p_Command);
            //}

            // 파라메터에 RETURN_VALUE 를 포함시키지 않는다면 삭제를 한다.
            if (!includeReturnValueParameter)
            {
                p_Command.Parameters.RemoveAt(0);
            }

            DbParameter[] discoveredParameters = new DbParameter[this.p_Command.Parameters.Count];
            p_Command.Parameters.CopyTo(discoveredParameters, 0);

            // 파라메터값을 초기화 한다. DBNull value
            foreach (DbParameter discoveredParameter in discoveredParameters)
            {
                switch (discoveredParameter.DbType)
                {
                    case DbType.String:
                        {
                            discoveredParameter.Value = string.Empty;
                        }
                        break;
                    case DbType.Int16:
                    case DbType.Int32:
                    case DbType.UInt16:
                    case DbType.UInt64:
                        {
                            discoveredParameter.Value = 0;
                        }
                        break;
                    default:
                        {
                            discoveredParameter.Value = DBNull.Value;
                        }
                        break;
                }

            }

            return discoveredParameters;
        }

        public string AssignParameterValues(DataRow dataRow)
        {
            if ((this.p_Command.Parameters == null) || (dataRow == null))
            {
                return "";
            }

            string rval = "";

            int i = 0;

            foreach (DbParameter commandParameter in this.p_Command.Parameters)
            {
                if (commandParameter.ParameterName == null || commandParameter.ParameterName.Length <= 1)
                    throw new Exception(
                        string.Format(
                            "Please provide a valid parameter name on the parameter #{0}, the ParameterName property has the following value: '{1}'.",
                            i, commandParameter.ParameterName));

                if (dataRow.Table.Columns.IndexOf(commandParameter.ParameterName) != -1)
                {
                    //MessageBox.Show(commandParameter.ParameterName + " = " + dataRow[commandParameter.ParameterName].ToString());
                    commandParameter.Value = dataRow[commandParameter.ParameterName];
                    rval = rval + commandParameter.ParameterName + " = " + dataRow[commandParameter.ParameterName].ToString() + ", ";
                }
                i++;
            }
            return rval;
        }

        public int ExecuteNonQuery(string spName, DataRow row, bool logOn)
        {
            if (spName == null || spName.Length == 0)
            {
                throw new ArgumentNullException("spName");
            }


            if (p_Connection.State == ConnectionState.Closed)
            {
                p_Connection.Open();
            }

            this.p_Command.CommandText = spName;
            this.p_Command.CommandType = CommandType.StoredProcedure;


            if (row != null)
            {
                DiscoverSpParameterSet(spName, true);
                AssignParameterValues(row);

                int result = this.p_Command.ExecuteNonQuery();

                if (logOn == true)
                {
                    // DB Log를 남긴다.
                    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod());
                }

                return result;
            }
            else
            {
                return -1;
            }
        }
        /// <summary>
        /// 스토어프로시져를 실행 시킨 결과를 DataSet으로 돌려 준다.
        /// </summary>
        /// <param name="spName">스토어프로시져이름</param>
        /// <param name="param">파라메터배열</param>
        /// <returns></returns>
        public DataSet ExecuteDataSet(string procedureName, DbParameter[] param, bool logOn)
        {
            if (procedureName == null || procedureName.Length == 0) throw new ArgumentNullException("Stored Procedure Name이 필요합니다.");


            if (p_Connection.State == ConnectionState.Closed)
            {
                p_Connection.Open();
            }

            p_Command.Parameters.Clear();

            p_Command.CommandText = String.Format("{0}", procedureName);
            p_Command.CommandType = CommandType.StoredProcedure;

            if (param != null)
            {
                Array.ForEach(param, commandParameter => p_Command.Parameters.Add(commandParameter));
            }
            //foreach (DbParameter commandParameter in param)
            //{
            //    _mCmd.Parameters.Add(commandParameter);
            //}

            DataSet ds = null;

            try
            {
                IDbDataAdapter adapter = new SqlDataAdapter((SqlCommand)p_Command);

                ds = new DataSet();

                adapter.Fill(ds);

                if (logOn == true)
                {
                    // DB Log를 남긴다.
                    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod());
                }
            }
            catch (SystemException e)
            {
                throw e;
            }

            catch (Exception e)
            {
                throw e;
            }
            //finally
            //{
            //    p_Connection.Close();
            //}

            //ds.Tables.Add(result);
            return ds;

        }

        /// <summary>
        /// 테이블과 파라메터 동기화 후 output 파라메터를 테이블에 넣어 준다.
        /// </summary>
        /// <param name="spName"></param>
        /// <param name="row"></param>
        /// <param name="flag">다른것과 비교 할려고 만든것 의미 없음</param>
        /// <returns></returns>
        public int ExecuteAsInOk(string spName, DataRow row, string inDate, bool logOn)
        {
            string workDate = string.Empty;

            if (spName == null || spName.Length == 0)
            {
                throw new ArgumentNullException("spName");
            }

            if (p_Connection.State == ConnectionState.Closed)
            {
                p_Connection.Open();
            }

            this.p_Command.CommandText = spName;
            this.p_Command.CommandType = CommandType.StoredProcedure;


            int rval = 1;

            if (row != null)
            {
                try
                {
                    if (logOn == true)
                    {
                        // DB Log를 남긴다.
                        InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod());
                    }


                    DiscoverSpParameterSet(spName, true);
                    AssignParameterValues(row);

                    rval = this.p_Command.ExecuteNonQuery();

                    foreach (DbParameter commandParameter in this.p_Command.Parameters)
                    {
                        if (commandParameter.Direction == ParameterDirection.InputOutput)
                        {
                            //MessageBox.Show(commandParameter.ParameterName + " : " + commandParameter.Direction.ToString());
                            if (row.Table.Columns.IndexOf(commandParameter.ParameterName) != -1)
                            {
                                row[commandParameter.ParameterName] = commandParameter.Value;
                            }
                        }
                    }

                    return rval;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, Resources.MSG_CAPTION_ERROR);
                    //MessageBox.Show(e.Message, Resources.MSG_CAPTION_ERROR, MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return -1;
                }
            }
            else
            {
                return -1;
            }
        }

#endregion



#region Transcation 별도 처리

        public void TransactionBegin()
        {
            if (p_Connection.State == ConnectionState.Closed)
            {
                p_Connection.Open();
            }

            SqlTransaction transaction = p_Connection.BeginTransaction();
            p_Command.Transaction = transaction;
        }

        public void TransactionCommit()
        {
            if (p_Command.Transaction != null)
            {
                p_Command.Transaction.Commit();
            }
#if TransationCheck
			else
			{
				MessageBox.Show("트랜젝션이 없습니다. (Commit)");
			}
#endif
        }

        public void TransactionRollBack()
        {
            if (p_Command.Transaction != null)
            {
                p_Command.Transaction.Rollback();
            }
#if TransationCheck
			else
			{
				MessageBox.Show("트랜젝션이 없습니다. (RollBack)");
			}
#endif
        }

#endregion



#region TxnLog
        private string[] InsertTrxLog(System.Reflection.MethodBase baseInfo)
        {
            //p_Connection Opne/Close를 하지 않는다.
            try
            {
#if UseTxnLog
                string formName = baseInfo.ReflectedType.Name;
                string functionName = baseInfo.Name;
                StringBuilder trxCommand = new StringBuilder(p_Command.CommandText);

                if (p_Command.Parameters.Count > 0)
                {
                    trxCommand.Append(" ");

                    foreach (SqlParameter param in p_Command.Parameters)
                    {
                        trxCommand.Append(param.ParameterName + " = " + param.Value.ToString());
                        trxCommand.Append(", ");
                    }

                    trxCommand.Remove(trxCommand.Length - 2, 2);
                }


                p_Command.CommandText = "xp_com_TxnLog_i";
                p_Command.CommandType = CommandType.StoredProcedure;
                p_Command.Parameters.Clear();

                p_Command.Parameters.AddWithValue("@TxnYear", DateTime.Today.ToString("yyyy"));
                p_Command.Parameters.AddWithValue("@TxnModule", functionName);
                p_Command.Parameters.AddWithValue("@TxnSource", trxCommand.ToString());
                //p_Command.Parameters.AddWithValue("@CreateDate", DateTime.Now);
                p_Command.Parameters.AddWithValue("@CreatePersonID", Globals.Settings.GetString(Resources.PersonID));
                p_Command.Parameters.AddWithValue("@CreateForm", formName);
                p_Command.Parameters.AddWithValue("@CreateIP", Globals.Settings.GetString(Resources.UserIP));


                string value = p_Command.ExecuteScalar().ToString();
                return new String[] { "success", value };
#else
                return new string[] { string.Empty, string.Empty };
#endif
            }
            catch (NullReferenceException)
            {
                return new String[] { Resources.success, "NullReferenceException" };
            }
            catch (Exception ex)
            {
                return new String[] { Resources.failure, ex.Message };
            }
            //finally
            //{
            //    if (p_Connection.State != ConnectionState.Closed)
            //    {
            //        p_Connection.Close();
            //    }
            //}

        }

        private string[] InsertTrxLog(System.Reflection.MethodBase baseInfo, string sql)
        {
            //p_Connection Opne/Close를 하지 않는다.
            try
            {
                //#if UseTxnLog
                //string formName = baseInfo.ReflectedType.Name;
                //string functionName = baseInfo.Name;
                //string userid = "";

                //p_Command.CommandText = "xp_com_TxnLog_i";
                //p_Command.CommandType = CommandType.StoredProcedure;
                //p_Command.Parameters.Clear();

                //p_Command.Parameters.AddWithValue("@TxnYear", DateTime.Today.ToString("yyyy"));
                //p_Command.Parameters.AddWithValue("@TxnModule", functionName);
                //p_Command.Parameters.AddWithValue("@TxnSource", sql);
                //p_Command.Parameters.AddWithValue("@CreatePersonID", userid);
                //p_Command.Parameters.AddWithValue("@CreateForm", formName);
                //p_Command.Parameters.AddWithValue("@CreateIP", lib.UserIPAddress);

                //result = p_Command.ExecuteNonQuery();

                ////20170901 김종영 추가 WizLog DB에 Log를 남기기위해 추가

                //string strVal = "";

                //Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                //sqlParameter.Add("@ComputerID", System.Environment.MachineName); //SystemInformation.ComputerName);수정요망
                //sqlParameter.Add("@UserID", userid);
                //sqlParameter.Add("@LogData", trxCommand.ToString());

                //DataStore.Log_Instance.ExecuteProcedureByErrorLog("xp_iLog", sqlParameter, false);

                //return new String[] { "success", strVal };//20170901 김종영 변경


                //string value = p_Command.ExecuteScalar().ToString();
                //return new String[] { "success", value };
                //#else
                return new string[] { string.Empty, string.Empty };
                //#endif
            }
            catch (NullReferenceException)
            {
                return new String[] { Resources.success, "NullReferenceException" };
            }
            catch (Exception ex)
            {
                return new String[] { Resources.failure, ex.Message };
            }
            //finally
            //{
            //    if (p_Connection.State != ConnectionState.Closed)
            //    {
            //        p_Connection.Close();
            //    }
            //}

        }

#endregion
        public string[] ExecuteAllProcedureOutputNew(List<Procedure> AllProcedure, List<Dictionary<string, object>> sqlParameterall)
        {
            // Output 결과 값을 넣을 List
            List<KeyValue> outputVal = new List<KeyValue>();
            SqlTransaction transaction = null;
            string value = "";
            bool complete = false;

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                transaction = p_Connection.BeginTransaction();
                p_Command.Transaction = transaction;

                foreach (Procedure Procedure in AllProcedure)
                {
                    if (sqlParameterall[AllProcedure.IndexOf(Procedure)] != null)
                    {
                        Dictionary<string, object> sqlParameter = sqlParameterall[AllProcedure.IndexOf(Procedure)];

                        if (Procedure.OutputUseYN == "Y")//리턴받는 output값이 있을때
                        {
                            if (outputVal.Count > 0)
                            {
                                //해당 프로시저의 output으로 리턴받는 값과 동일한 값이 있을 경우 output값 리스트에서 삭제
                                for (int i = outputVal.Count - 1; i >= 0; i--)
                                {
                                    KeyValue kvp = outputVal[i];
                                    if (kvp.key.ToLower().ToString() == Procedure.OutputName.ToLower().ToString())
                                    {
                                        outputVal.Remove(kvp);
                                    }
                                }
                            }
                            else
                            {
                                //output값 리스트에 추가
                                KeyValue kvp = new KeyValue();
                                kvp.key = Procedure.OutputName;
                                kvp.value = "";
                                outputVal.Add(kvp);
                            }
                        }
                        ///2018.02.02 로그용 파라미터 값 세팅
                        p_Command.CommandText = Procedure.Name;               //프로시저명 입력
                        p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
                        p_Command.Parameters.Clear();                           //이전 파라미터 클리어
                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            complete = false;

                            foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
                            {
                                if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
                                {
                                    if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
                                    {
                                        p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
                                        complete = true;
                                    }
                                }
                            }

                            if (!complete)
                            {
                                p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                            }
                        }
                        //로그 메서드

                        InsertTrxLogByUserID(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod());

                        p_Command.CommandText = Procedure.Name;                 //프로시저 이름 셋팅
                        p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
                        p_Command.Parameters.Clear();                           //로그용 파라미터 클리어

                        //입력할 데이터 파라미터 셋팅
                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            complete = false;

                            foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
                            {
                                if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
                                {
                                    if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
                                    {
                                        p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
                                        complete = true;
                                    }
                                }
                            }

                            if (!complete)
                            {
                                p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                            }
                        }

                        if (Procedure.OutputUseYN == "Y")
                        {
                            p_Command.Parameters[Procedure.OutputName].Direction = ParameterDirection.Output;
                            p_Command.Parameters[Procedure.OutputName].Size = int.Parse(Procedure.OutputLength);
                        }
                    }

                    value = Convert.ToString(p_Command.ExecuteScalar());

                    if (Procedure.OutputUseYN == "Y")
                    {
                        complete = false;                                                                       //완료여부

                        foreach (KeyValue mKeyValue in outputVal)                                               //output값 리스트중에서
                        {
                            if (mKeyValue.key == Procedure.OutputName)                                              //같은 이름을 가진 Key값이 리스트에 있을경우
                            {
                                mKeyValue.value = p_Command.Parameters[Procedure.OutputName].Value.ToString();      //해당 리스트에 값 추가
                                complete = true;

                                break;
                            }
                        }
                        if (!complete)
                        {
                            KeyValue kvp = new KeyValue();
                            kvp.key = Procedure.OutputName;                                                     //새로운 output값 이름의 리스트 생성
                            kvp.value = p_Command.Parameters[Procedure.OutputName].Value.ToString();            //새로운 output값 밸류 추가
                            outputVal.Add(kvp);                                                             //output값 리스트에 추가

                        }
                    }
                }


                transaction.Commit();


                return new String[] { "success", value };   //성공! 쿼리에서 리턴값이 있을경우
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                if (transaction != null)
                {
                    transaction.Commit();
                }

                return new String[] { "success", "NullReferenceException" };
            }
            catch (Exception ex)
            {
                try
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }
                    //exception용 로그 메서드
                    InsertTrxLogByUserIDErrLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), ex.Message);

                    return new string[] { "failure", ex.Message };
                }
                catch (Exception ex1)
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }
                    //exception용 로그 메서드2
                    InsertTrxLogByUserIDErrLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), ex1.Message);
                    return new string[] { "failure", ex.Message + "/" + ex1.Message };
                }
            }
            finally
            {
                if (p_Connection.State != ConnectionState.Closed)
                {
                    p_Connection.Close();
                }
                CloseConnection();
            }
        }
#region 주석
        ////public string[] ExecuteAllProcedureOutputNew(List<Procedure> AllProcedure, List<Dictionary<string, object>> sqlParameterall)
        //{
        //    // Output 결과 값을 넣을 Dictionary
        //    List<KeyValue> outputVal = new List<KeyValue>();
        //    SqlTransaction transaction = null;
        //    string value = "";
        //    bool complete = false;

        //    try
        //    {
        //        if (p_Connection.State == ConnectionState.Closed)
        //        {
        //            p_Connection.Open();
        //        }

        //        transaction = p_Connection.BeginTransaction();
        //        p_Command.Transaction = transaction;

        //        foreach (Procedure Procedure in AllProcedure)
        //        {
        //            if (sqlParameterall[AllProcedure.IndexOf(Procedure)] != null)
        //            {
        //                Dictionary<string, object> sqlParameter = sqlParameterall[AllProcedure.IndexOf(Procedure)];

        //                if (Procedure.OutputUseYN == "Y")//리턴받는 output값이 있을때
        //                {
        //                    if (outputVal.Count > 0)
        //                    {
        //                        //해당 프로시저의 output으로 리턴받는 값과 동일한 값이 있을 경우 output값 리스트에서 삭제
        //                        for (int i = outputVal.Count - 1; i >= 0; i--)
        //                        {
        //                            KeyValue kvp = outputVal[i];
        //                            if (kvp.key.ToLower().ToString() == Procedure.OutputName.ToLower().ToString())
        //                            {
        //                                outputVal.Remove(kvp);
        //                            }
        //                        }
        //                    }
        //                    else
        //                    {
        //                        //output값 리스트에 추가
        //                        KeyValue kvp = new KeyValue();
        //                        kvp.key = Procedure.OutputName;
        //                        kvp.value = "";
        //                        outputVal.Add(kvp);
        //                    }
        //                }
        //                ///2018.02.02 로그용 파라미터 값 세팅
        //                p_Command.CommandText = Procedure.Name;               //프로시저명 입력
        //                p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
        //                p_Command.Parameters.Clear();                           //이전 파라미터 클리어
        //                foreach (KeyValuePair<string, object> kvp in sqlParameter)
        //                {
        //                    complete = false;

        //                    foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
        //                    {
        //                        if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
        //                        {
        //                            if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
        //                            {
        //                                p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
        //                                complete = true;
        //                            }
        //                        }
        //                    }

        //                    if (!complete)
        //                    {
        //                        p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
        //                    }
        //                }
        //                //로그 메서드

        //                InsertTrxLogByUserID(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod());

        //                p_Command.CommandText = Procedure.Name;               //프로시저 이름 셋팅
        //                p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
        //                p_Command.Parameters.Clear();                           //로그용 파라미터 클리어


        //                //입력할 데이터 파라미터 셋팅
        //                foreach (KeyValuePair<string, object> kvp in sqlParameter)
        //                {
        //                    complete = false;

        //                    foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
        //                    {
        //                        if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
        //                        {
        //                            if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
        //                            {
        //                                p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
        //                                complete = true;
        //                            }
        //                        }
        //                    }

        //                    if (!complete)
        //                    {
        //                        p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
        //                    }
        //                }

        //                if (Procedure.OutputUseYN == "Y")
        //                {
        //                    p_Command.Parameters[Procedure.Name].Direction = ParameterDirection.Output;
        //                    p_Command.Parameters[Procedure.Name].Size = int.Parse(Procedure.OutputLength);
        //                }

        //            }

        //            value = Convert.ToString(p_Command.ExecuteScalar());

        //            if (Procedure.OutputUseYN == "Y")
        //            {
        //                complete = false;                                                                       //완료여부

        //                foreach (KeyValue mKeyValue in outputVal)                                               //output값 리스트중에서
        //                {
        //                    if (mKeyValue.key == Procedure.OutputName)                                              //같은 이름을 가진 리스트가 있을경우
        //                    {
        //                        mKeyValue.value = p_Command.Parameters[Procedure.OutputName].Value.ToString();      //해당 리스트에 값 추가
        //                        complete = true;
        //                        break;
        //                    }
        //                }
        //                if (!complete)
        //                {
        //                    KeyValue kvp = new KeyValue();
        //                    kvp.key = Procedure.OutputName;                                                     //새로운 output값 이름의 리스트 생성
        //                    kvp.value = p_Command.Parameters[Procedure.OutputName].Value.ToString();            //새로운 output값 밸류 추가
        //                    outputVal.Add(kvp);                                                             //output값 리스트에 추가
        //                }
        //            }
        //        }

        //        transaction.Commit();

        //        return new String[] { "success", value };   //성공! 쿼리에서 리턴값이 있을경우
        //    }
        //    catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
        //    {
        //        if (transaction != null)
        //        {
        //            transaction.Commit();
        //        }

        //        return new String[] { "success", "NullReferenceException" };
        //    }
        //    catch (Exception ex)
        //    {
        //        try
        //        {
        //            if (transaction != null)
        //            {
        //                transaction.Rollback();
        //            }
        //            //exception용 로그 메서드
        //            InsertTrxLogByUserIDErrLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), ex.Message);

        //            return new string[] { "failure", ex.Message };
        //        }
        //        catch (Exception ex1)
        //        {
        //            if (transaction != null)
        //            {
        //                transaction.Rollback();
        //            }
        //            //exception용 로그 메서드2
        //            InsertTrxLogByUserIDErrLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), ex1.Message);
        //            return new string[] { "failure", ex.Message + "/" + ex1.Message };
        //        }
        //    }
        //    finally
        //    {
        //        if (p_Connection.State != ConnectionState.Closed)
        //        {
        //            p_Connection.Close();
        //        }
        //        CloseConnection();
        //    }
        //}
#endregion

        private string[] InsertTrxLogByUserIDErrLog(System.Reflection.MethodBase baseInfo, string exMsg)
        {
            try
            {
                string formName = baseInfo.ReflectedType.Name;
                string functionName = baseInfo.Name;
                StringBuilder trxCommand = new StringBuilder(p_Command.CommandText);
                string userid = "";

                if (p_Command.Parameters.Count > 0)
                {
                    trxCommand.Append(" ");

                    foreach (SqlParameter param in p_Command.Parameters)
                    {
                        if (param.Value == null)
                        {
                            param.Value = "";
                        }

                        trxCommand.Append(param.ParameterName + " = " + param.Value.ToString());
                        trxCommand.Append(", ");

                        if (param.ParameterName.ToLower().Contains("userid"))
                        {
                            userid = param.Value.ToString();
                        }

                    }

                    trxCommand.Remove(trxCommand.Length - 2, 2);
                }

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                sqlParameter.Add("nErrID", "");
                sqlParameter.Add("sComputer", "");//SystemInformation.ComputerName);//수정요망
                sqlParameter.Add("sUserID", userid);
                sqlParameter.Add("nErrNO", 0);
                sqlParameter.Add("nErrIndex", 0);
                sqlParameter.Add("sErrMsg", exMsg);

                Dictionary<string, int> outputParam = new Dictionary<string, int>();
                outputParam.Add("nErrID", 10);
                Dictionary<string, string> dicResult = DataStore.Log_Instance.ExecuteProcedureOutputNoTranByErrorLog("xp_iErrLog", sqlParameter, outputParam, false);
                string ErrID = string.Empty;
                ErrID = dicResult["nErrID"];

                Dictionary<string, object> sqlParameter2 = new Dictionary<string, object>();

                sqlParameter2.Add("nErrID", ErrID);
                sqlParameter2.Add("nErrSeq", 0);
                sqlParameter2.Add("sErrData", trxCommand.ToString());

                DataStore.Log_Instance.ExecuteProcedureByErrorLog("xp_iErrLogSub", sqlParameter2, false);

                return new String[] { "success", "success" };//20170901 김종영 변경
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.ToString());
                return new String[] { "success", "NullReferenceException" };
            }
            catch (Exception ex)
            {
                return new String[] { "failure", ex.Message };
            }
            finally
            {

                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }

        }

        private string[] InsertTrxLogByUserID(System.Reflection.MethodBase baseInfo)
        {
            //p_Connection Opne/Close를 하지 않는다.
            try
            {
                //#if UseTxnLog
                int result = 0;
                string formName = baseInfo.ReflectedType.Name;
                string functionName = baseInfo.Name;
                StringBuilder trxCommand = new StringBuilder(p_Command.CommandText);
                string userid = "";

                if (p_Command.Parameters.Count > 0)
                {
                    trxCommand.Append(" ");

                    foreach (SqlParameter param in p_Command.Parameters)
                    {
                        if (param.Value == null)
                        {
                            param.Value = "";
                        }

                        trxCommand.Append(param.ParameterName + " = " + param.Value.ToString());
                        trxCommand.Append(", ");

                        if (param.ParameterName.ToLower().Contains("userid"))
                        {
                            userid = param.Value.ToString();
                        }
                    }
                    trxCommand.Remove(trxCommand.Length - 2, 2);

                    //MessageBox.Show(trxCommand.ToString());
                }


                p_Command.CommandText = "xp_com_TxnLog_i";
                p_Command.CommandType = CommandType.StoredProcedure;
                p_Command.Parameters.Clear();

                p_Command.Parameters.AddWithValue("@TxnYear", DateTime.Today.ToString("yyyyMMdd"));
                p_Command.Parameters.AddWithValue("@TxnModule", functionName);
                p_Command.Parameters.AddWithValue("@TxnSource", trxCommand.ToString());
                p_Command.Parameters.AddWithValue("@CreatePersonID", userid);// "ADMIN");//Globals.Settings.GetString(Resources.PersonID));//UPDATE김종영
                p_Command.Parameters.AddWithValue("@CreateForm", formName);
                p_Command.Parameters.AddWithValue("@CreateIP", lib.UserIPAddress);
                //p_Command.Parameters.AddWithValue("@SuccessYN", QuerySuccessYN);

                result = p_Command.ExecuteNonQuery();

                //20170901 김종영 추가 WizLog DB에 Log를 남기기위해 추가

                string strVal = "";

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

                sqlParameter.Add("@ComputerID", System.Environment.MachineName); //SystemInformation.ComputerName);수정요망
                sqlParameter.Add("@UserID", userid);
                sqlParameter.Add("@LogData", trxCommand.ToString());

                DataStore.Log_Instance.ExecuteProcedureByErrorLog("xp_iLog", sqlParameter, false);

                return new String[] { "success", strVal };//20170901 김종영 변경
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.ToString());
                return new String[] { "success", "NullReferenceException" };
            }
            catch (Exception ex)
            {
                return new String[] { "failure", ex.Message };
            }
            finally
            {

                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }

        }

        public Dictionary<string, string> ExecuteProcedureOutputNoTranByErrorLog(string procedureName, Dictionary<string, object> sqlParameter, Dictionary<string, int> outputParameters, bool logOn)
        {

            bool QuerySuccessYN = false;
            string exMsg = "";

            // Output 결과 값을 넣을 Dictionary
            Dictionary<string, string> outputResult = new Dictionary<string, string>();

            try
            {
                if (L_Connection.State == ConnectionState.Closed)
                {
                    L_Connection.Open();
                }



                //if (logOn == true)
                //{
                //    // DB Log를 남긴다.
                //    StringBuilder trxCommand = new StringBuilder(procedureName);

                //    if (p_Command.Parameters.Count > 0)
                //    {
                //        trxCommand.Append(" ");

                //        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                //        {
                //            trxCommand.Append(kvp.Key + " = " + kvp.Value.ToString());
                //            trxCommand.Append(", ");
                //        }

                //        trxCommand.Remove(trxCommand.Length - 2, 2);
                //    }

                //    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), trxCommand.ToString()); 
                //}



                L_Command.CommandText = procedureName;
                L_Command.CommandType = CommandType.StoredProcedure;
                L_Command.Parameters.Clear();


                if (sqlParameter != null)
                {
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        L_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }

                    // Output Parameter 지정 및 output 값 받을 Dictionary 준비
                    foreach (KeyValuePair<string, int> kvp in outputParameters)
                    {
                        L_Command.Parameters[kvp.Key].Direction = ParameterDirection.Output;
                        L_Command.Parameters[kvp.Key].Size = kvp.Value;
                        outputResult.Add(kvp.Key, "");
                    }
                }
                //string value = Convert.ToString(L_Command.ExecuteScalar());
                int result = L_Command.ExecuteNonQuery();


                //output 값 Dictionary에 저장
                foreach (KeyValuePair<string, int> kvp in outputParameters)
                {
                    outputResult[kvp.Key] = L_Command.Parameters[kvp.Key].Value.ToString();
                }

                QuerySuccessYN = true;

                return outputResult;
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                //output 값 Dictionary에 저장
                foreach (KeyValuePair<string, int> kvp in outputParameters)
                {
                    outputResult[kvp.Key] = L_Command.Parameters[kvp.Key].Value.ToString();
                }
                QuerySuccessYN = true;
                exMsg = "NullReferenceException";

                return outputResult;
            }
            catch (Exception ex)
            {
                try
                {
                    outputResult.Clear();
                    List<string> result = new List<string>();
                    result.Add("9999");
                    result.Add(ex.Message);
                    result.Add(ex.StackTrace);

                    int i = 0;

                    foreach (KeyValuePair<string, int> kvp in outputParameters)
                    {
                        outputResult[kvp.Key] = result.Count > i ? result[i++] : "";
                    }
                    QuerySuccessYN = false;
                    exMsg = ex.Message;

                    return outputResult;
                }
                catch (Exception ex1)
                {
                    outputResult.Clear();
                    List<string> result = new List<string>();
                    result.Add("9998");
                    result.Add(ex1.Message);
                    result.Add(ex1.StackTrace);

                    int i = 0;

                    foreach (KeyValuePair<string, int> kvp in outputParameters)
                    {
                        outputResult[kvp.Key] = result.Count > i ? result[i++] : "";
                    }

                    QuerySuccessYN = false;
                    exMsg = ex1.Message;

                    return outputResult;
                }
            }
            finally
            {
                if (L_Connection.State != ConnectionState.Closed)
                {
                    p_Connection.Close();
                }
            }
        }


        public string[] ExecuteProcedureByErrorLog(string procedureName, Dictionary<string, object> sqlParameter, bool logOn)
        {
            bool IsQueryOK = false;
            string ExMessage = "";

            SqlTransaction transaction = null;

            try
            {
                if (L_Connection.State == ConnectionState.Closed)
                {
                    L_Connection.Open();
                }

                transaction = L_Connection.BeginTransaction();
                L_Command.Transaction = transaction;

                L_Command.CommandText = procedureName;
                L_Command.CommandType = CommandType.StoredProcedure;
                L_Command.Parameters.Clear();


                if (sqlParameter != null)
                {
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        L_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }
                }
                int result = L_Command.ExecuteNonQuery();
                //string value = Convert.ToString(L_Command.ExecuteScalar());

                //#region 출하처리에 프로시저 예외처리 0일경우 commit 나머지 rollback
                //string[] valueSplit = value.Split(',');
                //if (valueSplit.Length > 1)
                //{
                //    if (valueSplit[0] == "0")
                //    {
                //        transaction.Commit();
                //        return new string[] { "success", value };
                //    }
                //    else
                //    {
                //        throw new Exception(value);
                //    }
                //}


                transaction.Commit();

                IsQueryOK = true;

                if (result > 0)
                {
                    return new String[] { "success", "success" };
                }
                else
                {
                    return new String[] { "failure", "결과값이 없습니다." };
                }
                //return new String[] { "success", value };   //성공! 쿼리에서 리턴값이 있을경우
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                if (transaction != null)
                {
                    transaction.Commit();
                }

                IsQueryOK = true;
                ExMessage = "NullReferenceException";

                return new String[] { "success", "NullReferenceException" };
            }
            catch (Exception ex)
            {
                try
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }

                    IsQueryOK = false;

                    ExMessage = ex.Message;

                    return new string[] { "failure", ex.Message };
                }
                catch (Exception ex1)
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }

                    IsQueryOK = false;

                    ExMessage = ex1.Message;

                    return new string[] { "failure", ex.Message + "/" + ex1.Message };
                }
            }
            finally
            {
                if (L_Connection.State != ConnectionState.Closed)
                {
                    L_Connection.Close();
                }
            }
        }

        /// <summary>
        /// 2021-06-01, GDU 생성
        /// 실행 프로시저중 하나라도 return 값이 11, ErrorMessage 인경우 예외 발생 및 롤백 되도록 
        /// </summary>
        /// <param name="AllProcedure"></param>
        /// <param name="sqlParameterall"></param>
        /// <returns></returns>
        public Dictionary<string, object> ExecuteAllProcedureOutput2(List<Procedure> lstProcedure, List<Dictionary<string, object>> lstSqlParameter)
        {
            // Output 값 list
            Dictionary<string, object> lstOutputVal = new Dictionary<string, object>();

            // 결과값
            Dictionary<string, object> lstResult = new Dictionary<string, object>();

            // 매개변수 세팅
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

            // Output 매개변수용 세팅
            Dictionary<SqlParameter, object> outParameter = new Dictionary<SqlParameter, object>();

            SqlTransaction transaction = null; // 트랜잭션

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                transaction = p_Connection.BeginTransaction();
                p_Command.Transaction = transaction;

                for (int i = 0; i < lstProcedure.Count; i++)
                {
                    Procedure procedure = lstProcedure[i];

                    if (lstSqlParameter[i] == null)
                    {
                        throw new Exception(procedure.Name + " 해당 프로시저에 매개변수 세팅 안됨.");
                    }

                    sqlParameter = lstSqlParameter[i];
                    outParameter = new Dictionary<SqlParameter, object>();

                    //== 로그용 ==//
                    StringBuilder trxCommand = new StringBuilder(procedure.Name);

                    p_Command.CommandText = procedure.Name;                         // 프로시저 이름 셋팅
                    p_Command.CommandType = CommandType.StoredProcedure;    // 명령타입 입력
                    p_Command.Parameters.Clear();                                            // 파라미터 클리어

                    // 아웃풋 값 세팅.
                    if (procedure.dicOutputList != null)
                    {
                        outParameter = procedure.dicOutputList.ToDictionary(x => x.Key, x => x.Value);


                        foreach (SqlParameter outParam in outParameter.Keys)
                        {
                            // 아웃풋 값 리스트에 Key가 있는경우, 해당 값을 전달
                            object value = lstOutputVal.Keys.Contains(outParam.ParameterName) ? lstOutputVal[outParam.ParameterName] : outParameter[outParam];

                            // 파라미터 세팅
                            outParam.Direction = ParameterDirection.InputOutput;
                            outParam.Value = value;

                            p_Command.Parameters.Add(outParam);

                            //== 로그용 ==//
                            trxCommand.Append(" ");
                            trxCommand.Append(outParam.ParameterName.ToString() + " = " + value.ToString());
                            trxCommand.Append(", ");
                        }
                    }

                    // 그 외 파라미터 세팅
                    foreach (string key in sqlParameter.Keys)
                    {
                        // 아웃풋 값 리스트에 Key가 있는경우, 해당 값을 전달
                        object value = lstOutputVal.Keys.Contains(key) ? lstOutputVal[key] : sqlParameter[key];

                        p_Command.Parameters.AddWithValue(key, value);

                        //== 로그용 ==//
                        trxCommand.Append(" ");
                        trxCommand.Append(key.ToString() + " = " + value.ToString());
                        trxCommand.Append(", ");
                    }

                    // 로그 저장!
                    trxCommand.Remove(trxCommand.Length - 2, 2);
                    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), trxCommand.ToString());

                    //InsertTrxLogByUserID(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod());

                    // 실행
                    var result = p_Command.ExecuteScalar() ?? ""; // Null 값도 success 로 취급.

                    if (string.IsNullOrEmpty(result.ToString()) == false)
                    {
                        throw new Exception(result.ToString().Replace("|", "\r\n"));
                    }

                    // 실행 후, output 값 세팅
                    if (outParameter.Count > 0)
                    {
                        foreach (SqlParameter outParam in outParameter.Keys)
                        {
                            var val = p_Command.Parameters[outParam.ParameterName].Value;

                            if (lstOutputVal.Keys.Contains(outParam.ParameterName))
                            {
                                lstOutputVal[outParam.ParameterName] = val;
                            }
                            else
                            {
                                lstOutputVal.Add(outParam.ParameterName, val);
                            }
                        }
                    }
                }

                // 아웃풋 값 세팅
                lstResult.Add("output", lstOutputVal);
                lstResult.Add("result", "success");

                transaction.Commit();

                return lstResult;
            }
            catch (Exception ex)
            {
                if (transaction != null)
                {
                    transaction.Rollback();
                }

                lstResult.Add("message", ex.Message);
                lstResult.Add("result", "failure");

                return lstResult;
            }
            finally
            {
                if (p_Connection.State != ConnectionState.Closed)
                {
                    p_Connection.Close();
                }
                CloseConnection();
            }


        }
        #region AS집중기간 로그기록 (화면별 CRUD 사용횟수, 사용시간)

        /// <summary>
        /// 기존버전에서 로그기록 남기기(AS집중기간 로그기록) 추가
        /// </summary>
        /// <param name="procedureName"></param>
        /// <param name="sqlParameter"></param>
        /// <param name="crudGubn"></param>
        /// <returns></returns>
        public string[] ExecuteProcedure_NewLog(string procedureName, Dictionary<string, object> sqlParameter, string crudGubn)
        {
            SqlTransaction transaction = null;

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                //로그 메서드 2022-06-21 추가
                InsertLogByForm(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod().ReflectedType.Name, crudGubn);

                transaction = p_Connection.BeginTransaction();
                p_Command.Transaction = transaction;

                p_Command.CommandText = procedureName;
                p_Command.CommandType = CommandType.StoredProcedure;
                p_Command.Parameters.Clear();


                if (sqlParameter != null)
                {
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }
                }

                string value = Convert.ToString(p_Command.ExecuteScalar());

                #region 출하처리에 프로시저 예외처리 0일경우 commit 나머지 rollback
                string[] valueSplit = value.Split(',');
                if (valueSplit.Length > 1)
                {
                    if (valueSplit[0] == "0")
                    {
                        transaction.Commit();
                        return new string[] { "success", value };
                    }
                    else
                    {
                        throw new Exception(value);
                    }
                }
                #endregion

                transaction.Commit();

                return new String[] { Resources.success, value };   //성공! 쿼리에서 리턴값이 있을경우
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                if (transaction != null)
                {
                    transaction.Commit();
                }

                return new String[] { Resources.success, "NullReferenceException" };
            }
            catch (Exception ex)
            {
                try
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }

                    return new string[] { Resources.failure, ex.Message };
                }
                catch (Exception ex1)
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }

                    return new string[] { Resources.failure, ex.Message + "/" + ex1.Message };
                }
            }
            finally
            {
                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }
        }

        /// <summary>
        /// 기존버전에서 로그기록 남기기(AS집중기간 로그기록) 추가
        /// </summary>
        /// <param name="AllProcedure"></param>
        /// <param name="sqlParameterall"></param>
        /// <returns></returns>
        public List<KeyValue> ExecuteAllProcedureOutputGetCS_NewLog(List<Procedure> AllProcedure, List<Dictionary<string, object>> sqlParameterall, string crudGubn)
        {
            // Output 결과 값을 넣을 List
            List<KeyValue> outputVal = new List<KeyValue>();
            SqlTransaction transaction = null;
            string value = "";
            bool complete = false;
            List<KeyValue> Success_List = new List<KeyValue>();//추가/180427

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                transaction = p_Connection.BeginTransaction();
                p_Command.Transaction = transaction;

                //로그 메서드
                InsertLogByForm(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod().ReflectedType.Name, crudGubn);

                foreach (Procedure Procedure in AllProcedure)
                {
                    if (sqlParameterall[AllProcedure.IndexOf(Procedure)] != null)
                    {
                        Dictionary<string, object> sqlParameter = sqlParameterall[AllProcedure.IndexOf(Procedure)];

                        if (Procedure.OutputUseYN == "Y")//리턴받는 output값이 있을때
                        {
                            if (outputVal.Count > 0)
                            {
                                //해당 프로시저의 output으로 리턴받는 값과 동일한 값이 있을 경우 output값 리스트에서 삭제
                                for (int i = outputVal.Count - 1; i >= 0; i--)
                                {
                                    KeyValue kvp = outputVal[i];
                                    if (kvp.key.ToLower().ToString() == Procedure.OutputName.ToLower().ToString())
                                    {
                                        outputVal.Remove(kvp);
                                    }
                                }
                            }
                            else
                            {
                                //output값 리스트에 추가
                                KeyValue kvp = new KeyValue();
                                kvp.key = Procedure.OutputName;
                                kvp.value = "";
                                outputVal.Add(kvp);
                            }
                        }

                        ///2018.02.02 로그용 파라미터 값 세팅
                        p_Command.CommandText = Procedure.Name;               //프로시저명 입력
                        p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
                        p_Command.Parameters.Clear();                           //이전 파라미터 클리어]

                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            complete = false;

                            foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
                            {
                                if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
                                {
                                    if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
                                    {
                                        p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
                                        complete = true;
                                    }
                                }
                            }

                            if (!complete)
                            {
                                p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                            }
                        }
                        //로그 메서드
                        InsertTrxLogByUserID(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod());


                        p_Command.CommandText = Procedure.Name;                 //프로시저 이름 셋팅
                        p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
                        p_Command.Parameters.Clear();                           //로그용 파라미터 클리어

                        //입력할 데이터 파라미터 셋팅
                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            complete = false;

                            foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
                            {
                                if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
                                {
                                    if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
                                    {
                                        p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
                                        complete = true;
                                    }
                                }
                            }

                            if (!complete)
                            {
                                p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                            }
                        }

                        if (Procedure.OutputUseYN == "Y")
                        {
                            p_Command.Parameters[Procedure.OutputName].Direction = ParameterDirection.Output;
                            p_Command.Parameters[Procedure.OutputName].Size = int.Parse(Procedure.OutputLength);
                        }
                    }

                    value = Convert.ToString(p_Command.ExecuteScalar());

                    if (Procedure.OutputUseYN == "Y")
                    {
                        complete = false;                                                                       //완료여부

                        foreach (KeyValue mKeyValue in outputVal)                                               //output값 리스트중에서
                        {
                            if (mKeyValue.key == Procedure.OutputName)                                              //같은 이름을 가진 Key값이 리스트에 있을경우
                            {
                                mKeyValue.value = p_Command.Parameters[Procedure.OutputName].Value.ToString();      //해당 리스트에 값 추가
                                complete = true;
                                Success_List.Add(mKeyValue);//추가/180427
                                break;
                            }
                        }
                        if (!complete)
                        {
                            KeyValue kvp = new KeyValue();
                            kvp.key = Procedure.OutputName;                                                     //새로운 output값 이름의 리스트 생성
                            kvp.value = p_Command.Parameters[Procedure.OutputName].Value.ToString();            //새로운 output값 밸류 추가
                            outputVal.Add(kvp);                                                             //output값 리스트에 추가
                            Success_List.Add(kvp);//추가/180427
                        }
                    }
                }
                //추가 /180427
                KeyValue suc_kv = new KeyValue();
                suc_kv.key = "Success";
                suc_kv.value = "";
                Success_List.Insert(0, suc_kv);

                transaction.Commit();

                return Success_List; //추가 /180427
                //return new String[] { "success", value };   //성공! 쿼리에서 리턴값이 있을경우
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                if (transaction != null)
                {
                    transaction.Commit();
                }

                //추가 /180427
                KeyValue suc_kv = new KeyValue();
                suc_kv.key = "Success";
                suc_kv.value = "NullReferenceException";
                Success_List.Add(suc_kv);
                //Success_List.AddRange(outputVal);

                return Success_List;


            }
            catch (Exception ex)
            {
                try
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }
                    //exception용 로그 메서드
                    InsertTrxLogByUserIDErrLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), ex.Message);

                    KeyValue suc_kv = new KeyValue();
                    suc_kv.key = "failure";
                    suc_kv.value = ex.Message;
                    Success_List.Add(suc_kv);

                    return Success_List;
                }
                catch (Exception ex1)
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }
                    //exception용 로그 메서드2
                    InsertTrxLogByUserIDErrLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), ex1.Message);

                    KeyValue suc_kv = new KeyValue();
                    suc_kv.key = "failure";
                    suc_kv.value = ex1.Message;
                    Success_List.Add(suc_kv);

                    return Success_List;
                }
            }
            finally
            {
                if (p_Connection.State != ConnectionState.Closed)
                {
                    p_Connection.Close();
                }
                CloseConnection();
            }
        }

        /// <summary>
        /// 기존버전에서 로그기록 남기기(AS집중기간 로그기록) 추가
        /// </summary>
        /// <param name="AllProcedure"></param>
        /// <param name="sqlParameterall"></param>
        /// <param name="crudGubun"></param>
        /// <returns></returns>
        public string[] ExecuteAllProcedureOutputNew_NewLog(List<Procedure> AllProcedure, List<Dictionary<string, object>> sqlParameterall, string crudGubun)
        {
            // Output 결과 값을 넣을 List
            List<KeyValue> outputVal = new List<KeyValue>();
            SqlTransaction transaction = null;
            string value = "";
            bool complete = false;

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                transaction = p_Connection.BeginTransaction();
                p_Command.Transaction = transaction;

                //로그 메서드 2022-06-21 추가
                InsertLogByForm(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod().ReflectedType.Name, crudGubun);

                foreach (Procedure Procedure in AllProcedure)
                {
                    if (sqlParameterall[AllProcedure.IndexOf(Procedure)] != null)
                    {
                        Dictionary<string, object> sqlParameter = sqlParameterall[AllProcedure.IndexOf(Procedure)];

                        if (Procedure.OutputUseYN == "Y")//리턴받는 output값이 있을때
                        {
                            if (outputVal.Count > 0)
                            {
                                //해당 프로시저의 output으로 리턴받는 값과 동일한 값이 있을 경우 output값 리스트에서 삭제
                                for (int i = outputVal.Count - 1; i >= 0; i--)
                                {
                                    KeyValue kvp = outputVal[i];
                                    if (kvp.key.ToLower().ToString() == Procedure.OutputName.ToLower().ToString())
                                    {
                                        outputVal.Remove(kvp);
                                    }
                                }
                            }
                            else
                            {
                                //output값 리스트에 추가
                                KeyValue kvp = new KeyValue();
                                kvp.key = Procedure.OutputName;
                                kvp.value = "";
                                outputVal.Add(kvp);
                            }
                        }
                        ///2018.02.02 로그용 파라미터 값 세팅
                        p_Command.CommandText = Procedure.Name;               //프로시저명 입력
                        p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
                        p_Command.Parameters.Clear();                           //이전 파라미터 클리어
                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            complete = false;

                            foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
                            {
                                if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
                                {
                                    if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
                                    {
                                        p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
                                        complete = true;
                                    }
                                }
                            }

                            if (!complete)
                            {
                                p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                            }
                        }
                        //로그 메서드

                        InsertTrxLogByUserID(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod());

                        p_Command.CommandText = Procedure.Name;                 //프로시저 이름 셋팅
                        p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
                        p_Command.Parameters.Clear();                           //로그용 파라미터 클리어

                        //입력할 데이터 파라미터 셋팅
                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            complete = false;

                            foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
                            {
                                if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
                                {
                                    if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
                                    {
                                        p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
                                        complete = true;
                                    }
                                }
                            }

                            if (!complete)
                            {
                                p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                            }
                        }

                        if (Procedure.OutputUseYN == "Y")
                        {
                            p_Command.Parameters[Procedure.OutputName].Direction = ParameterDirection.Output;
                            p_Command.Parameters[Procedure.OutputName].Size = int.Parse(Procedure.OutputLength);
                        }
                    }

                    value = Convert.ToString(p_Command.ExecuteScalar());

                    if (Procedure.OutputUseYN == "Y")
                    {
                        complete = false;                                                                       //완료여부

                        foreach (KeyValue mKeyValue in outputVal)                                               //output값 리스트중에서
                        {
                            if (mKeyValue.key == Procedure.OutputName)                                              //같은 이름을 가진 Key값이 리스트에 있을경우
                            {
                                mKeyValue.value = p_Command.Parameters[Procedure.OutputName].Value.ToString();      //해당 리스트에 값 추가
                                complete = true;

                                break;
                            }
                        }
                        if (!complete)
                        {
                            KeyValue kvp = new KeyValue();
                            kvp.key = Procedure.OutputName;                                                     //새로운 output값 이름의 리스트 생성
                            kvp.value = p_Command.Parameters[Procedure.OutputName].Value.ToString();            //새로운 output값 밸류 추가
                            outputVal.Add(kvp);                                                             //output값 리스트에 추가

                        }
                    }
                }


                transaction.Commit();


                return new String[] { "success", value };   //성공! 쿼리에서 리턴값이 있을경우
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                if (transaction != null)
                {
                    transaction.Commit();
                }

                return new String[] { "success", "NullReferenceException" };
            }
            catch (Exception ex)
            {
                try
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }
                    //exception용 로그 메서드
                    InsertTrxLogByUserIDErrLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), ex.Message);

                    return new string[] { "failure", ex.Message };
                }
                catch (Exception ex1)
                {
                    if (transaction != null)
                    {
                        transaction.Rollback();
                    }
                    //exception용 로그 메서드2
                    InsertTrxLogByUserIDErrLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), ex1.Message);
                    return new string[] { "failure", ex.Message + "/" + ex1.Message };
                }
            }
            finally
            {
                if (p_Connection.State != ConnectionState.Closed)
                {
                    p_Connection.Close();
                }
                CloseConnection();
            }
        }

        /// <summary>
        /// 로그기록남기기(AS집중기간 로그기록) (C, R, U, D, E, P 경우 해당)
        /// </summary>
        /// <param name="baseInfo"></param>
        /// <param name="crudGubun"></param>
        /// <returns></returns>
        public string[] InsertLogByForm(string form, string crudGubun)
        {
            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                int result = 0;
                string formName = form;
                string userid = "";

                if (p_Command.Parameters.Count > 0)
                {
                    foreach (SqlParameter param in p_Command.Parameters)
                    {
                        if (param.Value == null)
                        {
                            param.Value = "";
                        }

                        if (param.ParameterName.ToLower().Contains("userid"))
                        {
                            userid = param.Value.ToString();
                        }
                    }
                }

                if (userid.Equals(""))
                {
                    userid = MainWindow.CurrentUser;
                }

                p_Command.CommandText = "xp_iWorkLogWPF_New";
                p_Command.CommandType = CommandType.StoredProcedure;
                p_Command.Parameters.Clear();

                p_Command.Parameters.AddWithValue("@sCompanyID", MainWindow.CompanyID);
                p_Command.Parameters.AddWithValue("@sMenuID", "");
                p_Command.Parameters.AddWithValue("@sWorkFlag", crudGubun);
                p_Command.Parameters.AddWithValue("@sWorkDate", DateTime.Now.ToString("yyyyMMdd"));
                p_Command.Parameters.AddWithValue("@sWorkTime", DateTime.Now.ToString("HHmm"));

                p_Command.Parameters.AddWithValue("@sUserID", userid);
                p_Command.Parameters.AddWithValue("@sWorkComputer", System.Environment.MachineName);
                p_Command.Parameters.AddWithValue("@sWorkComputerIP", lib.UserIPAddress);
                p_Command.Parameters.AddWithValue("@sWorkLog", "");
                p_Command.Parameters.AddWithValue("@sProgramID", formName);

                result = p_Command.ExecuteNonQuery();

                return new String[] { "success", "success" };
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.ToString());
                return new String[] { "success", "NullReferenceException" };
            }
            catch (Exception ex)
            {
                return new String[] { "failure", ex.Message };
            }
            finally
            {

                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }

        }

        /// <summary>
        /// 로그기록남기기(AS집중기간 로그기록) (S :화면사용시간 일때)
        /// </summary>
        /// <param name="baseInfo"></param>
        /// <param name="crudGubun"></param>
        /// <returns></returns>
        public string[] InsertLogByFormS(String form, string stDate, string stTime, string seGubun)
        {
            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                int result = 0;
                string formName = form;
                string userid = "";

                if (p_Command.Parameters.Count > 0)
                {
                    foreach (SqlParameter param in p_Command.Parameters)
                    {
                        if (param.Value == null)
                        {
                            param.Value = "";
                        }

                        if (param.ParameterName.ToLower().Contains("userid"))
                        {
                            userid = param.Value.ToString();
                        }
                    }
                }

                if (userid.Equals(""))
                {
                    userid = MainWindow.CurrentUser;
                }

                //시작이면
                if (seGubun.Equals("S"))
                {
                    p_Command.CommandText = "xp_iWorkLogWPF_New_UseTime";
                    p_Command.CommandType = CommandType.StoredProcedure;
                    p_Command.Parameters.Clear();

                    p_Command.Parameters.AddWithValue("@sCompanyID", MainWindow.CompanyID);
                    p_Command.Parameters.AddWithValue("@sMenuID", "");
                    p_Command.Parameters.AddWithValue("@sWorkFlag", "S"); //seGubun과 workFlag는 다름
                    p_Command.Parameters.AddWithValue("@sWorkDate", stDate);
                    p_Command.Parameters.AddWithValue("@sWorkTime", stTime);
                    p_Command.Parameters.AddWithValue("@sStartDate", stDate);
                    p_Command.Parameters.AddWithValue("@sStartTime", stTime);

                    p_Command.Parameters.AddWithValue("@sUserID", userid);
                    p_Command.Parameters.AddWithValue("@sWorkComputer", System.Environment.MachineName);
                    p_Command.Parameters.AddWithValue("@sWorkComputerIP", lib.UserIPAddress);
                    p_Command.Parameters.AddWithValue("@sWorkLog", "");
                    p_Command.Parameters.AddWithValue("@sProgramID", formName);
                }
                else //종료이면
                {
                    p_Command.CommandText = "xp_uWorkLogWPF_New_UseTime";
                    p_Command.CommandType = CommandType.StoredProcedure;
                    p_Command.Parameters.Clear();

                    p_Command.Parameters.AddWithValue("@sWorkFlag", "S"); //seGubun과 workFlag는 다름
                    p_Command.Parameters.AddWithValue("@sStartDate", stDate);
                    p_Command.Parameters.AddWithValue("@sStartTime", stTime);
                    p_Command.Parameters.AddWithValue("@sEndDate", DateTime.Now.ToString("yyyyMMdd"));
                    p_Command.Parameters.AddWithValue("@sEndTime", DateTime.Now.ToString("HHmm"));

                    p_Command.Parameters.AddWithValue("@sUserID", userid);
                    p_Command.Parameters.AddWithValue("@sWorkComputer", System.Environment.MachineName);
                    p_Command.Parameters.AddWithValue("@sWorkComputerIP", lib.UserIPAddress);
                    p_Command.Parameters.AddWithValue("@sProgramID", formName);
                }


                result = p_Command.ExecuteNonQuery();

                return new String[] { "success", "success" };
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.ToString());
                return new String[] { "success", "NullReferenceException" };
            }
            catch (Exception ex)
            {
                return new String[] { "failure", ex.Message };
            }
            finally
            {

                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }

        }

        /// <summary>
        /// 로그기록남기기(AS집중기간 로그기록) (S :화면사용시간 일때 전체 업데이트)
        /// </summary>
        /// <param name="baseInfo"></param>
        /// <param name="crudGubun"></param>
        /// <returns></returns>
        public string[] InsertLogByFormAllUpdate(string stDate, string stTime)
        {
            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                int result = 0;
                string userid = "";

                if (p_Command.Parameters.Count > 0)
                {
                    foreach (SqlParameter param in p_Command.Parameters)
                    {
                        if (param.Value == null)
                        {
                            param.Value = "";
                        }

                        if (param.ParameterName.ToLower().Contains("userid"))
                        {
                            userid = param.Value.ToString();
                        }
                    }
                }

                if (userid.Equals(""))
                {
                    userid = MainWindow.CurrentUser;
                }

                p_Command.CommandText = "xp_uWorkLogWPF_New_UseTime_All";
                p_Command.CommandType = CommandType.StoredProcedure;
                p_Command.Parameters.Clear();

                p_Command.Parameters.AddWithValue("@sStartDate", stDate);
                p_Command.Parameters.AddWithValue("@sStartTime", stTime);
                p_Command.Parameters.AddWithValue("@sEndDate", DateTime.Now.ToString("yyyyMMdd"));
                p_Command.Parameters.AddWithValue("@sEndTime", DateTime.Now.ToString("HHmm"));

                p_Command.Parameters.AddWithValue("@sUserID", userid);
                p_Command.Parameters.AddWithValue("@sWorkComputer", System.Environment.MachineName);
                p_Command.Parameters.AddWithValue("@sWorkComputerIP", lib.UserIPAddress);


                result = p_Command.ExecuteNonQuery();

                return new String[] { "success", "success" };
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show(ex.ToString());
                return new String[] { "success", "NullReferenceException" };
            }
            catch (Exception ex)
            {
                return new String[] { "failure", ex.Message };
            }
            finally
            {

                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }

        }

        public Dictionary<string, object> ExecuteAllProcedureOutput2_NewLog(List<Procedure> lstProcedure, List<Dictionary<string, object>> lstSqlParameter, string crudGubun)
        {
            // Output 값 list
            Dictionary<string, object> lstOutputVal = new Dictionary<string, object>();

            // 결과값
            Dictionary<string, object> lstResult = new Dictionary<string, object>();

            // 매개변수 세팅
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

            // Output 매개변수용 세팅
            Dictionary<SqlParameter, object> outParameter = new Dictionary<SqlParameter, object>();

            SqlTransaction transaction = null; // 트랜잭션

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }

                transaction = p_Connection.BeginTransaction();
                p_Command.Transaction = transaction;

                InsertLogByForm(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod().ReflectedType.Name, crudGubun);

                for (int i = 0; i < lstProcedure.Count; i++)
                {
                    Procedure procedure = lstProcedure[i];

                    if (lstSqlParameter[i] == null)
                    {
                        throw new Exception(procedure.Name + " 해당 프로시저에 매개변수 세팅 안됨.");
                    }

                    sqlParameter = lstSqlParameter[i];
                    outParameter = new Dictionary<SqlParameter, object>();

                    //== 로그용 ==//
                    StringBuilder trxCommand = new StringBuilder(procedure.Name);

                    p_Command.CommandText = procedure.Name;                         // 프로시저 이름 셋팅
                    p_Command.CommandType = CommandType.StoredProcedure;    // 명령타입 입력
                    p_Command.Parameters.Clear();                                            // 파라미터 클리어

                    // 아웃풋 값 세팅.
                    if (procedure.dicOutputList != null)
                    {
                        outParameter = procedure.dicOutputList.ToDictionary(x => x.Key, x => x.Value);


                        foreach (SqlParameter outParam in outParameter.Keys)
                        {
                            // 아웃풋 값 리스트에 Key가 있는경우, 해당 값을 전달
                            object value = lstOutputVal.Keys.Contains(outParam.ParameterName) ? lstOutputVal[outParam.ParameterName] : outParameter[outParam];

                            // 파라미터 세팅
                            outParam.Direction = ParameterDirection.InputOutput;
                            outParam.Value = value;

                            p_Command.Parameters.Add(outParam);

                            //== 로그용 ==//
                            trxCommand.Append(" ");
                            trxCommand.Append(outParam.ParameterName.ToString() + " = " + value.ToString());
                            trxCommand.Append(", ");
                        }
                    }

                    // 그 외 파라미터 세팅
                    foreach (string key in sqlParameter.Keys)
                    {
                        // 아웃풋 값 리스트에 Key가 있는경우, 해당 값을 전달
                        object value = lstOutputVal.Keys.Contains(key) ? lstOutputVal[key] : sqlParameter[key];

                        p_Command.Parameters.AddWithValue(key, value);

                        //== 로그용 ==//
                        trxCommand.Append(" ");
                        trxCommand.Append(key.ToString() + " = " + value.ToString());
                        trxCommand.Append(", ");
                    }

                    // 로그 저장!
                    trxCommand.Remove(trxCommand.Length - 2, 2);
                    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), trxCommand.ToString());

                    //InsertTrxLogByUserID(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod());

                    // 실행
                    var result = p_Command.ExecuteScalar() ?? ""; // Null 값도 success 로 취급.

                    if (string.IsNullOrEmpty(result.ToString()) == false)
                    {
                        throw new Exception(result.ToString().Replace("|", "\r\n"));
                    }

                    // 실행 후, output 값 세팅
                    if (outParameter.Count > 0)
                    {
                        foreach (SqlParameter outParam in outParameter.Keys)
                        {
                            var val = p_Command.Parameters[outParam.ParameterName].Value;

                            if (lstOutputVal.Keys.Contains(outParam.ParameterName))
                            {
                                lstOutputVal[outParam.ParameterName] = val;
                            }
                            else
                            {
                                lstOutputVal.Add(outParam.ParameterName, val);
                            }
                        }
                    }
                }

                // 아웃풋 값 세팅
                lstResult.Add("output", lstOutputVal);
                lstResult.Add("result", "success");

                transaction.Commit();

                return lstResult;
            }
            catch (Exception ex)
            {
                if (transaction != null)
                {
                    transaction.Rollback();
                }

                lstResult.Add("message", ex.Message);
                lstResult.Add("result", "failure");

                return lstResult;
            }
            finally
            {
                if (p_Connection.State != ConnectionState.Closed)
                {
                    p_Connection.Close();
                }
                CloseConnection();
            }
        }

        public Dictionary<string, string> ExecuteProcedureOutputNoTran_NewLog(string procedureName, Dictionary<string, object> sqlParameter, Dictionary<string, int> outputParameters, bool logOn, string crudGubun)
        {
            // Output 결과 값을 넣을 Dictionary
            Dictionary<string, string> outputResult = new Dictionary<string, string>();

            try
            {
                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();
                }



                if (logOn == true)
                {
                    // DB Log를 남긴다.
                    StringBuilder trxCommand = new StringBuilder(procedureName);

                    if (p_Command.Parameters.Count > 0)
                    {
                        trxCommand.Append(" ");

                        foreach (KeyValuePair<string, object> kvp in sqlParameter)
                        {
                            trxCommand.Append(kvp.Key + " = " + kvp.Value.ToString());
                            trxCommand.Append(", ");
                        }

                        trxCommand.Remove(trxCommand.Length - 2, 2);
                    }

                    InsertTrxLog(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod(), trxCommand.ToString());
                    //로그 남기기
                    InsertLogByForm(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod().ReflectedType.Name, crudGubun);
                    p_Command.CommandText = procedureName;
                    p_Command.CommandType = CommandType.StoredProcedure;
                    p_Command.Parameters.Clear();
                }




                if (sqlParameter != null)
                {
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }

                    // Output Parameter 지정 및 output 값 받을 Dictionary 준비
                    foreach (KeyValuePair<string, int> kvp in outputParameters)
                    {
                        p_Command.Parameters[kvp.Key].Direction = ParameterDirection.Output;
                        p_Command.Parameters[kvp.Key].Size = kvp.Value;
                        outputResult.Add(kvp.Key, "");
                    }
                }

                string value = Convert.ToString(p_Command.ExecuteScalar());

                //output 값 Dictionary에 저장
                foreach (KeyValuePair<string, int> kvp in outputParameters)
                {
                    outputResult[kvp.Key] = p_Command.Parameters[kvp.Key].Value.ToString();
                }

                return outputResult;
            }
            catch (NullReferenceException)  //성공! 쿼리에서 리턴값이 없을경우
            {
                //output 값 Dictionary에 저장
                foreach (KeyValuePair<string, int> kvp in outputParameters)
                {
                    outputResult[kvp.Key] = p_Command.Parameters[kvp.Key].Value.ToString();
                }


                return outputResult;
                //return null;
            }
            catch (Exception ex)
            {
                try
                {
                    outputResult.Clear();
                    List<string> result = new List<string>();
                    result.Add("9999");
                    result.Add(ex.Message);
                    result.Add(ex.StackTrace);

                    int i = 0;

                    foreach (KeyValuePair<string, int> kvp in outputParameters)
                    {
                        outputResult[kvp.Key] = result.Count > i ? result[i++] : "";
                    }

                    return outputResult;
                }
                catch (Exception ex1)
                {
                    outputResult.Clear();
                    List<string> result = new List<string>();
                    result.Add("9998");
                    result.Add(ex1.Message);
                    result.Add(ex1.StackTrace);

                    int i = 0;

                    foreach (KeyValuePair<string, int> kvp in outputParameters)
                    {
                        outputResult[kvp.Key] = result.Count > i ? result[i++] : "";
                    }

                    return outputResult;
                }
            }
            finally
            {
                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }
        }

        #endregion
        /// <summary>
        /// ProcedureToDataSet 화면 조회, 로그남기기
        /// </summary>
        /// <param name="procedureName"></param>
        /// <param name="sqlParameter"></param>
        /// <returns></returns>
        public DataSet ProcedureToDataSet_LogWrite(string procedureName, Dictionary<string, object> sqlParameter, bool logOn, string crudGubun)
        {
            // Output 결과 값을 넣을 List
            List<KeyValue> outputVal = new List<KeyValue>();
            bool complete = false;

            try
            {
                //Cursor.Current = Cursors.WaitCursor;

                if (p_Connection.State == ConnectionState.Closed)
                {
                    p_Connection.Open();

                }

                if (logOn == true)
                {
                    ///2018.02.02 로그용 파라미터 값 세팅
                    p_Command.CommandText = procedureName;               //프로시저명 입력
                    p_Command.CommandType = CommandType.StoredProcedure;    //명령타입 입력
                    p_Command.Parameters.Clear();                           //이전 파라미터 클리어
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        complete = false;

                        foreach (KeyValue mKeyValue in outputVal)   //outputVal list에 KeyValue 클래스가 1개이상 있을때
                        {
                            if (kvp.Key == mKeyValue.key)           //KeyValue 객체의 key값(output값의 컬럼명)과 sql파라미터의 key값이 같을때
                            {
                                if (mKeyValue.value != "")          //KeyValue 객체의 value값이 빈 값이 아닐때 
                                {
                                    p_Command.Parameters.AddWithValue(kvp.Key, mKeyValue.value);//해당 KeyValue객체의 Value값을 sql 파라미터의 value에 넣어준다.
                                    complete = true;
                                }
                            }
                        }

                        if (!complete)
                        {
                            p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                        }
                    }
                    //로그 메서드
                    InsertLogByForm(new System.Diagnostics.StackTrace(1, false).GetFrame(0).GetMethod().ReflectedType.Name, crudGubun);
                    p_Command.CommandText = procedureName;
                    p_Command.CommandType = CommandType.StoredProcedure;
                    p_Command.Parameters.Clear();
                }

                //value = Convert.ToString(p_Command.ExecuteScalar());

                if (sqlParameter != null)
                {
                    foreach (KeyValuePair<string, object> kvp in sqlParameter)
                    {
                        p_Command.Parameters.AddWithValue(kvp.Key, kvp.Value);
                    }
                }

                SqlDataAdapter adapter = new SqlDataAdapter(p_Command);
                DataSet dataset = new DataSet();
                adapter.Fill(dataset);
                adapter.Dispose();


                return dataset;
            }
            catch (Exception e)

            {
                MessageBox.Show(e.Message, Resources.MSG_CAPTION_ERROR);
                //MessageBox.Show(e.Message, Resources.MSG_CAPTION_ERROR, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                //Cursor.Current = Cursors.Default;
                //if (p_Connection.State != ConnectionState.Closed)
                //{
                //    p_Connection.Close();
                //}
            }
        }
    }


}

