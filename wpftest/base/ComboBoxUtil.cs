using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Windows;

namespace WizMes_WellMade
{
    class ComboBoxUtil
    {
        private static ComboBoxUtil mComboBoxUtil = null;

        public static ComboBoxUtil Instance
        {
            get
            {
                if (mComboBoxUtil == null)
                {
                    mComboBoxUtil = new ComboBoxUtil();
                }

                return mComboBoxUtil;
            }
        }


        /// <summary>
        /// 금형 콤보박스
        /// </summary>
        /// <returns></returns>
        public ObservableCollection<CodeView> Get_MOLDID()
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            string sql = " select dm.MoldID, dm.MoldNo ";
            sql += " from dvl_Mold dm ";
            sql += " order by dm.MoldNo ";

            try
            {
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {
                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item[0].ToString().Trim(),
                                code_name = item[1].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }

        /// <summary>
        /// 설비 콤보박스
        /// </summary>
        /// <returns></returns>
        public ObservableCollection<CodeView> Get_MCID()
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            string sql = " select mm.MCID,mm.MCNAME ";
            sql += " from mt_Mc  mm ";
            sql += " where  UseClss <> '*' order by MCNAME ";

            try
            {
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {
                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item[0].ToString().Trim(),
                                code_name = item[1].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }
        /// <summary>
        /// 거래구분 코드 가져오기
        /// </summary>
        /// <param name="prs"></param>
        /// <param name="psCodeGroup"></param>
        /// <param name="psUseYN"></param>
        /// <param name="psParentID"></param>
        /// <param name="psRelation"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> Gf_DB_CM_GetComCodeDataset(string prs, string psCodeGroup,
                                              string psUseYN, string psParentID, string psRelation = "")
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            string sql = "SELECT CODE_GBN,CODE_ID,PARENT_ID,CODE_NAME,COMMENTS                              ";
            sql += "        ,LEVEL,RELATION,SEQ,CODE_SIZE,USE_YN                                            ";
            sql += "        ,CreateDate= CONVERT( VARCHAR(30), CreateDate, 120) , CreateUserID              ";
            sql += "        ,LastUpdateDate= CONVERT( VARCHAR(30), LastUpdateDate, 120) , LastUpdateUserID  ";
            sql += "    FROM CM_CODE                                                                        ";
            sql += "   WHERE 1          = 1                                                                 ";

            if (!(string.IsNullOrEmpty(psCodeGroup)))
            {
                sql += "   AND CODE_GBN                                  =   '" + psCodeGroup + "'             ";
            }
            if (!(string.IsNullOrEmpty(psUseYN)))
            {
                sql += "     AND USE_YN                                    =   '" + psUseYN + "'           ";
            }
            /*
            * Parent ID 구분
            */
            if (!(string.IsNullOrEmpty(psParentID)))
            {
                sql += "     AND PARENT_ID                                 =  '" + psParentID + "'          ";
            }
            /*
            * Relation 참조
            */
            if (!(string.IsNullOrEmpty(psRelation)))
            {
                sql += "     AND RELATION                                  =  '" + psRelation + "'         ";
            }

            sql += "  ORDER BY SEQ                                                                      ";

            try
            {
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {

                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {
                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item["code_id"].ToString().Trim(),
                                code_name = item["code_name"].ToString().Trim(),
                                code_id_plus_code_name = item["code_id"].ToString().Trim() + "." + item["code_name"].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }


        /// <summary>
        /// 거래구분 코드 가져오기
        /// </summary>
        /// <param name="prs"></param>
        /// <param name="psCodeGroup"></param>
        /// <param name="psUseYN"></param>
        /// <param name="psParentID"></param>
        /// <param name="psRelation"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> Gf_DB_CM_GetComCodeDatasetPlusAll(string prs, string psCodeGroup,
                                              string psUseYN, string psParentID, string psRelation = "")
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            string sql = "SELECT CODE_GBN,CODE_ID,PARENT_ID,CODE_NAME,COMMENTS                              ";
            sql += "        ,LEVEL,RELATION,SEQ,CODE_SIZE,USE_YN                                            ";
            sql += "        ,CreateDate= CONVERT( VARCHAR(30), CreateDate, 120) , CreateUserID              ";
            sql += "        ,LastUpdateDate= CONVERT( VARCHAR(30), LastUpdateDate, 120) , LastUpdateUserID  ";
            sql += "    FROM CM_CODE                                                                        ";
            sql += "   WHERE 1          = 1                                                                 ";

            if (!(string.IsNullOrEmpty(psCodeGroup)))
            {
                sql += "   AND CODE_GBN                                  =   '" + psCodeGroup + "'             ";
            }
            if (!(string.IsNullOrEmpty(psUseYN)))
            {
                sql += "     AND USE_YN                                    =   '" + psUseYN + "'           ";
            }
            /*
            * Parent ID 구분
            */
            if (!(string.IsNullOrEmpty(psParentID)))
            {
                sql += "     AND PARENT_ID                                 =  '" + psParentID + "'          ";
            }
            /*
            * Relation 참조
            */
            if (!(string.IsNullOrEmpty(psRelation)))
            {
                sql += "     AND RELATION                                  =  '" + psRelation + "'         ";
            }

            sql += "  ORDER BY SEQ                                                                      ";

            try
            {
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {

                    }
                    else
                    {
                        retunCollection.Add(new CodeView()
                        {
                            code_id = "",
                            code_name = "(전체)"
                        });

                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {
                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item["code_id"].ToString().Trim(),
                                code_name = item["code_name"].ToString().Trim(),
                                code_id_plus_code_name = item["code_id"].ToString().Trim() + "." + item["code_name"].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }



        /// <summary>
        /// 자회사ID와 자회사명을 가져온다.
        /// </summary>
        /// <returns></returns>
        public ObservableCollection<CodeView> Get_CompanyID()
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            string sql = "SELECT CompanyID,KCompany  ";
            sql += "FROM mt_SetCompany         ";
            sql += "WHERE 1          = 1       ";

            try
            {
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0) { }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {

                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item["CompanyID"].ToString().Trim(),
                                code_name = item["KCompany"].ToString().Trim(),
                                code_id_plus_code_name = item["CompanyID"].ToString().Trim() + "." + item["KCompany"].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }

        /// <summary>
        /// 기초코드 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <param name="basisID"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetCode_SetComboBox(string value, string basisID)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

            if (value.Equals("Defect"))
            {
                sqlParameter.Add("@sBasisID", basisID);
            }

            try
            {
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_s" + value, sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {

                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {

                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item[0].ToString().Trim(),
                                code_name = item[1].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }

        /// <summary>
        /// 기초코드 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <param name="basisID"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetCode_SetComboBoxPlusAll(string value, string basisID)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

            if (value.Equals("Defect"))
            {
                sqlParameter.Add("@sBasisID", basisID);
            }

            try
            {
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_s" + value, sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {

                    }
                    else
                    {
                        retunCollection.Add(new CodeView()
                        {
                            code_id = "",
                            code_name = "(전체)"
                        });

                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {

                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item[0].ToString().Trim(),
                                code_name = item[1].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }




        /// <summary>
        /// 기초코드 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <param name="basisID"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetCMCode_SetComboBox(string strCodeGBN, string strRelation)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("CodeGbn", strCodeGBN);
                sqlParameter.Add("sRelation", strRelation);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sCmCode", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {

                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {

                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item[0].ToString().Trim(),
                                code_name = item[1].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }

        /// <summary>
        /// 품명ID 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <param name="num"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetArticleCode_SetComboBox(string value, int num)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("@sArticleGrp", value);
            sqlParameter.Add("@iIncNotUse", num);

            try
            {
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_ArticleGrp_sArticleGrp", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {

                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {

                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item[0].ToString().Trim(),
                                code_name = item[1].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }

        /// <summary>
        /// 품명ID 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <param name="num"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetBuyerArticleNo_SetComboBox(string value)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("ArticleID", value);

            try
            {
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_BuyerArticle_sGetBuyerArticleNo", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {
                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item["BuyerArticle"].ToString().Trim(),
                                code_name = item["BuyerArticleNo"].ToString()
                            };

                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }

        /// <summary>
        /// 차종ID 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <param name="num"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetModelID_SetComboBox(string value)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("ArticleID", value);

            try
            {
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_BuyerArticle_sGetModelID", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {
                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item["ModelID"].ToString().Trim(),
                                code_name = item["Model"].ToString()
                            };

                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }

        //전체 빼고
        public ObservableCollection<CodeView> GetWorkProcess_only(int num, string value)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("@nchkProc", num);
            sqlParameter.Add("@ProcessID", value);

            try
            {
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Work_sProcess", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {

                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        foreach (DataRow item in drc)
                        {
                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item[0].ToString().Trim(),
                                code_name = item[1].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }

        /// <summary>
        /// 공정ID +  전체 가져오기
        /// </summary>
        /// <param name="num"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetWorkProcess(int num, string value)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("@nchkProc", num);
            sqlParameter.Add("@ProcessID", value);

            try
            {
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Work_sProcess", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {

                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        CodeView mCodeView = new CodeView()
                        {
                            code_id = "",
                            code_name = "전체"
                        };
                        retunCollection.Add(mCodeView);

                        foreach (DataRow item in drc)
                        {

                            mCodeView = new CodeView()
                            {
                                code_id = item[0].ToString().Trim(),
                                code_name = item[1].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }

        /// <summary>
        /// 공정 가져오기
        /// </summary>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetProcessByAutoMC()
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sProcessByAutoMC", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {

                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item[0].ToString().Trim(),
                                code_name = item[1].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }

        /// <summary>
        /// 공정그룹가져오기
        /// </summary>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetProcessGroup()
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("sProcessID", "");
            sqlParameter.Add("sArticleGrpID", "");
            sqlParameter.Add("sIncNotUseYN", "");

            try
            {
                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_sProcessGroup", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {

                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {

                            var mCodeView = new CodeView()
                            {
                                code_id = item["ParentProcessID"].ToString().Replace(" ", "").Trim(),
                                code_name = item["ParentProcessName"].ToString().Replace(" ", "").Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }


            return retunCollection;
        }

        /// <summary>
        /// 호기ID 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetMachine(string value)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("@sProcessID", value);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Process_sMachine", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count == 0)
                {

                }
                else
                {
                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow item in drc)
                    {

                        CodeView mCodeView = new CodeView()
                        {
                            code_id = item[0].ToString().Trim(),
                            code_name = item[1].ToString().Trim()
                        };
                        retunCollection.Add(mCodeView);
                    }
                }
            }

            return retunCollection;
        }

        /// <summary>
        /// 호기ID 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetMachinePlusAll(string value)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("@sProcessID", value);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Process_sMachine", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count == 0)
                {

                }
                else
                {
                    CodeView AllCodeView = new CodeView();
                    AllCodeView.code_id = "";
                    AllCodeView.code_name = "전체";
                    retunCollection.Add(AllCodeView);

                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow item in drc)
                    {

                        CodeView mCodeView = new CodeView()
                        {
                            code_id = item[0].ToString().Trim(),
                            code_name = item[1].ToString().Trim()
                        };
                        retunCollection.Add(mCodeView);
                    }
                }
            }

            return retunCollection;
        }

        /// <summary>
        /// 호기ID 가져오기
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetMachineCodeByAutoMC(string value)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("@sProcessID", value);

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Code_SMachineByAutoMC", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow item in drc)
                    {

                        CodeView mCodeView = new CodeView()
                        {
                            code_id = item[2].ToString().Trim(),
                            code_name = item[0].ToString().Trim()
                        };
                        retunCollection.Add(mCodeView);
                    }
                }
            }

            return retunCollection;
        }

        /// <summary>
        /// 제품출고 명세서 / 제품 or 상품만 조회하기. (품명그룹)
        /// </summary>
        /// <returns></returns>
        public ObservableCollection<CodeView> Gf_DB_MT_sArticleGrp()
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            string sql = " select  ArticleGrpID , ArticleGrp ";
            sql += " From   mt_ArticleGrp ";
            sql += " where  UseClss <> '*' ";

            try
            {
                DataSet ds = DataStore.Instance.QueryToDataSet(sql);
                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    if (dt.Rows.Count == 0)
                    {
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow item in drc)
                        {

                            CodeView mCodeView = new CodeView()
                            {
                                code_id = item[0].ToString().Trim(),
                                code_name = item[1].ToString().Trim()
                            };
                            retunCollection.Add(mCodeView);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("콤보박스 생성 중 오류 발생 : " + ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return retunCollection;
        }

        /// <summary>
        /// 직접 콤보박스에 아이템을 넣을때 사용( 단순 문자열만 필요할때)
        /// </summary>
        /// <param name="strValue"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> Direct_SetComboBox(List<string> strValue)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();

            DataTable dt = new DataTable();
            dt.Columns.Add("code_id");
            dt.Columns.Add("code_name");

            for (int i = 0; i < strValue.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr["code_id"] = i;
                dr["code_name"] = strValue[i];

                dt.Rows.Add(dr);
            }

            DataRowCollection drc = dt.Rows;

            foreach (DataRow item in drc)
            {
                CodeView mCodeView = new CodeView()
                {
                    code_id = item[0].ToString().Trim(),
                    code_name = item[1].ToString().Trim()
                };
                retunCollection.Add(mCodeView);
            }

            return retunCollection;
        }

        /// <summary>
        /// 직접 콤보박스에 아이템을 넣을때 사용(문자열에 숨겨진 값이 필요할때)
        /// </summary>
        /// <param name="strValue"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> Direct_SetComboBox(List<string[]> strValue)
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();

            DataTable dt = new DataTable();
            dt.Columns.Add("code_id");
            dt.Columns.Add("code_name");

            foreach (string[] str in strValue)
            {
                DataRow dr = dt.NewRow();
                dr["code_id"] = str[0];
                dr["code_name"] = str[1];

                dt.Rows.Add(dr);
            }

            DataRowCollection drc = dt.Rows;

            foreach (DataRow item in drc)
            {
                CodeView mCodeView = new CodeView()
                {
                    code_id = item[0].ToString().Trim(),
                    code_name = item[1].ToString().Trim()
                };
                retunCollection.Add(mCodeView);
            }

            return retunCollection;
        }

        /// <summary>
        /// 회사정보 가져오기
        /// </summary>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetInfo()
        {
            ObservableCollection<CodeView> retunCollection = new ObservableCollection<CodeView>();
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("nChkCompany", 0);
            sqlParameter.Add("sCompanyID", "");
            sqlParameter.Add("sKCompany", "");
            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Info_GetCompanyInfo", sqlParameter, false);

            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count == 0)
                {
                }
                else
                {
                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow item in drc)
                    {
                        CodeView mCodeView = new CodeView()
                        {
                            code_id = item["CompanyID"].ToString().Trim(),
                            code_name = item["KCompany"].ToString().Trim()
                        };
                        retunCollection.Add(mCodeView);
                    }
                }
            }

            return retunCollection;
        }

        /// <summary>
        /// 광역시,특별시,도 콤보박스
        /// </summary>
        /// <param name="strUseYN"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetSido(string strUseYN)
        {
            ObservableCollection<CodeView> returnCollection = new ObservableCollection<CodeView>();

            string strSqlMSG = "";
            strSqlMSG += " SELECT Sido_Code As Code_ID,Sido_Name AS CODE_NAME,Sido_Eng_Name,Seq ";
            strSqlMSG += " FROM ZipSido ";
            strSqlMSG += " WHERE 1 = 1 ";
            strSqlMSG += " AND Sido_Code <>'00' ";

            if (!strUseYN.Equals(""))
            {
                strSqlMSG += " AND USE_YN = '" + strUseYN + "'";
            }

            DataSet ds = DataStore.Zip_Instance.QueryToDataSetByZip(strSqlMSG);
            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count == 0)
                {
                }
                else
                {
                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow dr in drc)
                    {
                        CodeView codeView = new CodeView()
                        {
                            code_id = dr["Code_ID"].ToString().Trim(),
                            code_name = dr["CODE_NAME"].ToString().Trim()
                        };
                        returnCollection.Add(codeView);
                    }
                }
            }

            return returnCollection;
        }

        /// <summary>
        /// 시.군.구 콤보박스
        /// </summary>
        /// <param name="strGroupCode"></param>
        /// <param name="strUSE_YN"></param>
        /// <returns></returns>
        public ObservableCollection<CodeView> GetSiGunGu(string strGroupCode, string strUSE_YN)
        {
            ObservableCollection<CodeView> returnCollection = new ObservableCollection<CodeView>();

            string strSqlMSG = "";
            strSqlMSG += " SELECT SiGunGu_Code As Code_ID,SiGunGu_Name AS CODE_NAME,SiGunGu_Eng_Name,Seq ";
            strSqlMSG += " FROM ZipSiGunGu ";
            strSqlMSG += " WHERE 1 = 1 ";

            if (!strGroupCode.Equals("") && !strGroupCode.Equals("00"))
            {
                strSqlMSG += "And SubString(SiGunGu_Code,1,2) = '" + strGroupCode + "' ";
            }

            if (!strUSE_YN.Equals(""))
            {
                strSqlMSG += " AND USE_YN = '" + strUSE_YN + "'";
            }

            strSqlMSG += " ORDER BY SEq ";

            DataSet ds = DataStore.Zip_Instance.QueryToDataSetByZip(strSqlMSG);
            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count == 0)
                {
                }
                else
                {
                    CodeView view = new CodeView();
                    view.code_id = "00000";
                    view.code_name = "전체";
                    returnCollection.Add(view);

                    DataRowCollection drc = dt.Rows;

                    foreach (DataRow dr in drc)
                    {
                        CodeView codeView = new CodeView()
                        {
                            code_id = dr["Code_ID"].ToString().Trim(),
                            code_name = dr["CODE_NAME"].ToString().Trim()
                        };
                        returnCollection.Add(codeView);
                    }
                }
            }

            return returnCollection;
        }

        //공정패턴 만들기
        public ObservableCollection<CodeView> GetProcessPattern(string strArticleGrpID)
        {
            ObservableCollection<CodeView> returnCollection = new ObservableCollection<CodeView>();

            List<string> CbView = new List<string>();
            List<string> PatternID = new List<string>();

            string strCompare1 = string.Empty;
            string strCompare2 = string.Empty;
            string TheView = string.Empty;

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sArticleGrpID", strArticleGrpID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_PlanInput_sPatternByArticleGrpID", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("공정 데이터가 없습니다. 먼저 등록해주세요.");
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;
                        int i = 0;

                        foreach (DataRow item in drc)
                        {
                            CodeView codeView = new CodeView();

                            strCompare1 = item["PatternID"].ToString().Trim();
                            strCompare2 = item["Pattern"].ToString().Trim();

                            TheView = strCompare1 + "." + strCompare2 + " : ";

                            foreach (DataRow items in drc)
                            {
                                if (items["PatternID"].ToString().Equals(strCompare1))
                                {
                                    TheView += " [" + items["Process"].ToString() + "] →";
                                }
                            }
                            if (TheView != null && !TheView.Equals(""))
                            {
                                TheView = TheView.Substring(0, TheView.Length - 1);
                            }

                            if (CbView.Count > 0)
                            {
                                if (!CbView[i].Substring(0, 2).Equals(strCompare1))
                                {
                                    codeView.code_id = strCompare1;
                                    codeView.code_name = TheView;

                                    CbView.Add(TheView);
                                    returnCollection.Add(codeView);
                                    i++;
                                }
                            }
                            else
                            {
                                codeView.code_id = strCompare1;
                                codeView.code_name = TheView;

                                CbView.Add(TheView);
                                returnCollection.Add(codeView);
                            }
                        }
                        drc.Clear();
                    }
                    dt.Clear();
                }
                ds.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }

            return returnCollection;
        }

        /// <summary>
        /// 자회사ID와 자회사명을 가져온다.
        /// </summary>
        /// <returns></returns>
        public string Get_DrawNo(string strArticleID)
        {
            string strDrawNo = string.Empty;
            string sql = "SELECT DrawNo,ArticleID  ";
            sql += "FROM dvl_Draw         ";
            sql += "WHERE 1          = 1       ";
            sql += "AND ArticleID     =  '" + strArticleID + "'         ";

            DataSet ds = DataStore.Instance.QueryToDataSet(sql);
            if (ds != null && ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count == 0) { }
                else
                {
                    DataRowCollection drc = dt.Rows;
                    strDrawNo = drc[0]["DrawNo"].ToString().Trim();
                }
            }

            return strDrawNo;
        }
    }

}
