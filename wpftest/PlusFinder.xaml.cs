using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

/**************************************************************************************************
'** 프로그램명 : PlusFinder
'** 설명       : 공통팝업
'** 작성일자   : 
'** 작성자     : 장시영
'**------------------------------------------------------------------------------------------------
'**************************************************************************************************
' 변경일자  , 변경자, 요청자    , 요구사항ID      , 요청 및 작업내용
'**************************************************************************************************
' 2023.03.30, 장시영, 삼익SDT에서 가져옴
'**************************************************************************************************/

namespace WizMes_WellMade
{
    /// <summary>
    /// PlusFinder.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class PlusFinder : Window
    {
        DataTable rs_dt = null;
        delegate void FHideWindow();
        TextBox txtBox;
        TextBox txtLot; //2021-11-09 Lotid를 위해 추가
        TextBox txtBoxName;
        Lib Lib = new Lib();

        // string 값을 받을 수 있는 이벤트
        public delegate void RefEventHandler(string msg);
        public event RefEventHandler refEvent;

        public PlusFinder()
        {
            InitializeComponent();
            LangKor.IsChecked = true;
            SetLangBtn();
            SetDataGrid();
            SetButton();
        }

        protected override void OnClosing(CancelEventArgs e1)
        {
            e1.Cancel = true;
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new FHideWindow(_HideThisWindow));
        }

        //한글, 영어 버튼 이벤트 셋팅
        private void SetLangBtn()
        {
            List<Button> BtnHangle = Lib.FindVisualChildByContainName<Button>(GridKor, "btn");
            List<Button> BtnEnglish = Lib.FindVisualChildByContainName<Button>(GridEng, "btn");
            foreach (var btn in BtnHangle)
            {
                btn.Click += btnLanguage_Click;
            }

            foreach (var btn in BtnEnglish)
            {
                btn.Click += btnLanguage_Click;
            }
        }
        //그리드뷰 셋팅
        private void SetDataGrid()
        {
            mDataGrid.MaxHeight = SystemParameters.WorkArea.Height;
            mDataGrid.VerticalScrollBarVisibility = ScrollBarVisibility.Visible;
            mDataGrid.MouseLeftButtonDown += SelectItem;
            mDataGrid.SelectedIndex = 0;

        }
        //검색,선택,취소버튼
        private void SetButton()
        {
            btnSearch.Click += SelectDataTable;
            btnChoice.Click += SelectItem;
            btnClose.Click += btnClose_Click;
        }

        public void _HideThisWindow()
        {
            this.Hide();
        }

        //rs_dt에서 검색하기. 플러스파인더의 코드,명칭 텍스트박스 사용
        private void SelectDataTable(object sender, RoutedEventArgs e)
        {
            string Data = "";
            DataTable dtCodeTemp = new DataTable();//dt.select where절 조건건 
            DataTable dtNameTemp = new DataTable();//dt.select where절 조건건 

            dtCodeTemp = rs_dt.Clone();//테이블 구조 복사(컬럼 복사를 위해서, 컬럼이 똑같이 없으면 오류남)
            dtNameTemp = rs_dt.Clone();//테이블 구조 복사(컬럼 복사를 위해서, 컬럼이 똑같이 없으면 오류남)
            string ColID = rs_dt.Columns[0].Caption;//명칭
            string ColName = rs_dt.Columns[1].Caption;//코드
            string ColArticle = "";
            if (ColID == "사번")
            {
                ColArticle = rs_dt.Columns[2].Caption;//사번 검색시에만 사용, Sabun/ArticleID/Article 동시 검색을 위한 변수
            }
            string sql = "";


            //1. 코드, 명칭 두개 다 텍스트박스에 값이 있을 경우
            if (txtCode.Text.Trim().Length > 0 && txtName.Text.Trim().Length > 0)
            {
                SearchCode();
                SearchName();
                //JoinDataTable(); → 이해가 안되니 일단 안쓰기로 함 2020.05.13
            }
            //2. 코드로 찾기
            else if (txtCode.Text.Trim().Length > 0)
            {
                SearchCode();
                mDataGrid.ItemsSource = dtCodeTemp.DefaultView;
            }
            //3. 명칭으로 찾기
            else if (txtName.Text.Trim().Length > 0)
            {
                SearchName();
                mDataGrid.ItemsSource = dtNameTemp.DefaultView;
            }
            //4. 둘다 비워져 있으면 전체 세팅
            else if (txtCode.Text.Trim().Length == 0
                        && txtName.Text.Trim().Length == 0)
            {
                SearchName();
                mDataGrid.ItemsSource = dtNameTemp.DefaultView;
            }

            //[1]코드로 검색
            void SearchCode()
            {
                Data = txtCode.Text.Trim();

                //if (Data != "")
                //{
                #region 봉인 2020.05.13 → 특수문자 끼면 프로그램 꺼짐 현상 발생 : .Select() 의 문제
                //if (ColName == "수주번호")
                //{
                //    sql = ColID + " = '" + Data + "' OR 거래처코드 = '" + Data + "'";
                //}
                //else if (ColID == "사번")
                //{
                //    sql = ColID + " LIKE '%" + Data + "%' OR " + ColName + " LIKE '%" + Data + "%' OR " + ColArticle + " LIKE '%" + Data + "%'";
                //}
                //else
                //{
                //    sql = ColID + " = '" + Data + "'";
                //    //sql = ColID + " Like '%" + Data + "%'";
                //}
                //foreach (DataRow dr in rs_dt.Select(sql))
                //{
                //    dtCodeTemp.Rows.Add(dr.ItemArray);
                //}
                #endregion

                // 위의 셀렉트 안쓰고 텍스트로 찾으면?
                foreach (DataRow dr in rs_dt.Rows)
                {
                    if (dr[ColID].ToString().ToUpper().Replace(" ", "").Contains(Data.ToUpper().Replace(" ", "")))
                    {
                        dtCodeTemp.Rows.Add(dr.ItemArray);
                    }
                }
                //}
            }
            //[2]명칭으로 검색
            void SearchName()
            {
                Data = txtName.Text.Trim();
                //if (Data != "")
                //{
                #region 봉인 2020.05.13 → 특수문자 끼면 프로그램 꺼짐 현상 발생 : .Select() 의 문제
                //if (ColName == "수주번호")
                //{
                //    sql = ColName + " LIKE '%" + Data + "%' OR " +
                //        //ColName + " LIKE '(주)" + Data + "%;" + //vb소스상에 있는데 왜 했는지 모르겠음. 그래서 뺌.
                //        " OR 품목명 LIKE '" + Data + "%'" +
                //        " OR 거래처코드 = '" + Data + "'";
                //}
                //else if (ColID == "사번")
                //{
                //    sql = ColID + " LIKE '%" + Data + "%' OR " + ColName + " LIKE '%" + Data + "%' OR " + ColArticle + " LIKE '%" + Data + "%'";
                //}
                //else
                //{
                //    sql = ColName + " LIKE '%" + Data + "%'";
                //}

                //foreach (DataRow dr in rs_dt.Select(sql))
                //{
                //    dtNameTemp.Rows.Add(dr.ItemArray);
                //}
                #endregion

                // 위의 셀렉트 안쓰고 텍스트로 찾으면?
                foreach (DataRow dr in rs_dt.Rows)
                {
                    Console.WriteLine(dr[ColName].ToString());
                    if (dr[ColName].ToString().ToUpper().Replace(" ", "").Contains(Data.ToUpper().Replace(" ", "")))
                    {
                        dtNameTemp.Rows.Add(dr.ItemArray);
                    }
                }
                //}
            }
            //[3]코드,명칭 InnerJoin 결과
            void JoinDataTable()
            {

                DataTable JoinDT = new DataTable();
                JoinDT = rs_dt.Clone();//테이블 구조 복사(컬럼 복사를 위해서, 컬럼이 똑같이 없으면 오류남)
                //Linq 결과값 담을 List
                List<string> list_ID = new List<string>();

                var innerJoin = from tb1 in dtCodeTemp.AsEnumerable()
                                join tb2 in dtNameTemp.AsEnumerable()
                                on tb1.Field<string>(0) equals tb2.Field<string>(0)
                                select tb1.Field<string>(0);

                //rs_dt에 조건을 걸 where변수
                string where = "";

                //Linq절로 가져온 값 쓰기 편하게 리스트에 담기
                foreach (var ID in innerJoin)
                {
                    list_ID.Add(ID);
                }
                //select된게 없을때 종료
                if (list_ID.Count == 0)
                {
                    return;
                }
                //where절 셋팅
                for (int i = 0; i < list_ID.Count; i++)
                {
                    if (i == 0)
                    {
                        where = ColID + " LIKE '%"/*" = '"*/ + list_ID[i] + "%'";
                    }
                    else
                    {
                        where = where + " OR " + ColID + " LIKE '%"/*" = '"*/ + list_ID[i] + "%'";
                    }
                }


                foreach (DataRow drw in rs_dt.Select(where))
                {
                    JoinDT.Rows.Add(drw.ItemArray);
                }
                mDataGrid.ItemsSource = JoinDT.DefaultView;
            }
        }

        //조회 후 전역 DataTable인 rs_dt에 넣어둠.
        private void ProcQuery(int large, string smiddle)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("@nLarge", large);    // 데이터 그리드1의 Code_ID == 데이터 그리드2의 Code_GBN
                sqlParameter.Add("@sMiddle", smiddle);   //입력된 코드를 추가

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Common_PlusFinder", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    rs_dt = ds.Tables[0];

                    if (rs_dt != null)
                    {
                        //그리드의 1번컬럼의 컬럼명을 lblName에 넣고, 컬럼명의 문자열 중간중간에 스페이스바를 삽입해줌
                        if (rs_dt.Columns[1].Caption.ToString() == "코드")
                            lblName.Content = Lib.SetStringSpace(rs_dt.Columns[2].Caption);
                        else
                            lblName.Content = Lib.SetStringSpace(rs_dt.Columns[1].Caption);
                    }
                    
                }

                mDataGrid.Columns.Clear();
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
        //2021-11-09 자재입고반품에서 CustomID가 필요하여 생성
        private void ProcQuery(int large, string smiddle, string CustomID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("@nLarge", large);    // 데이터 그리드1의 Code_ID == 데이터 그리드2의 Code_GBN
                sqlParameter.Add("@sMiddle", smiddle);   //입력된 코드를 추가
                sqlParameter.Add("@sCustomID", CustomID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Common_PlusFinder_MTR", sqlParameter, false); //2021-11-09
                                                                                                                     //자재입고반품을 위해 생성(거래처, ARTICLEID 2개 가지고 가야 됨)

                if (ds != null && ds.Tables.Count > 0)
                {
                    rs_dt = ds.Tables[0];
                }

                //그리드의 1번컬럼의 컬럼명을 lblName에 넣고, 컬럼명의 문자열 중간중간에 스페이스바를 삽입해줌
                if (rs_dt.Columns[1].Caption.ToString() == "코드")
                {
                    lblName.Content = Lib.SetStringSpace(rs_dt.Columns[2].Caption);
                }
                else
                {   //2021-11-09
                    lblName2.Content = Lib.SetStringSpace(rs_dt.Columns[0].Caption);
                    lblName.Content = Lib.SetStringSpace(rs_dt.Columns[1].Caption);
                }

                mDataGrid.Columns.Clear();
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



        private void DataClear()
        {
            txtCode.Text = string.Empty;
            txtName.Text = string.Empty;
        }

        private void DataTableByWhere(TextBox _txtBox, int large)
        {
            try
            {
                //매개변수 텍스트박스의 텍스트에 값이 있을때 rs_dt에서 검색할것.
                if (_txtBox.Text.Trim().Length > 0)
                {
                    string ColID = rs_dt.Columns[0].Caption;//명칭
                    string ColName = rs_dt.Columns[1].Caption;//코드                   
                    //검색어 입력후 엔터 눌렀을때 다른 열의 데이터 값을 넣기 위한 변수
                    string colID2 = string.Empty;
                    string colName2 = string.Empty;

                    string ColArticle = "";
                    if (ColID == "사번")
                    {
                        ColArticle = rs_dt.Columns[2].Caption;//사번 검색시에만 사용, Sabun/ArticleID/Article 동시 검색을 위한 변수
                    }
                    string sql = "";
                    string Data = _txtBox.Text.Trim();

            

                    //[1]코드로 찾기
                    if (Data != "")
                    {
                        //if (ColName == "수주번호")
                        //{
                        //    sql = ColID + " = '" + Data + "' OR 거래처코드 = '" + Data + "'";
                        //}
                        ////else if (ColName == "품명")
                        ////{
                        ////    sql = ColID + " LIKE '%" + Data + "%' OR " + ColName + " LIKE '%" + Data + "%'";
                        ////}
                        //else if (ColID == "사번")
                        //{
                        //    sql = ColID + " LIKE '%" + Data + "%' OR " + ColName + " LIKE '%" + Data + "%' OR " + ColArticle + " LIKE '%" + Data + "%'";
                        //}
                        //else
                        //{
                        //    sql = ColID + " = '" + Data + "' OR " + ColName + " LIKE '%" + Data + "%'";
                        //}

                        DataTable dtCodeTemp = new DataTable();//dt.select where절 조건건 
                        dtCodeTemp = rs_dt.Clone();//테이블 구조 복사(컬럼 복사를 위해서, 컬럼이 똑같이 없으면 오류남)

                        //foreach (DataRow dr in rs_dt.Select(sql))
                        //{
                        //    dtCodeTemp.Rows.Add(dr.ItemArray);
                        //}

                        // 위의 셀렉트 안쓰고 텍스트로 찾으면?
                        foreach (DataRow dr in rs_dt.Rows)
                        {
                            Console.WriteLine(dr.ToString());

                            string str1 = dr[ColName].ToString();
                            string str2 = dr[ColID].ToString();

                            bool compare1 = str1.ToString().ToUpper().Replace(" ", "").Contains(Data.ToUpper().Replace(" ", ""));
                            bool compare2 = str2.ToString().ToUpper().Replace(" ", "").Contains(Data.ToUpper().Replace(" ", ""));
                            if (large == 77)
                                compare1 = compare1 || compare2;

                            if (compare1)
                            {
                                dtCodeTemp.Rows.Add(dr.ItemArray);
                                continue;
                            }                              //20240314 최대현 compare2 추가
                            else if (compare1 || compare2) //plusfinder 프로시저로 찾은 데이터들 중 입력한 값이 있는지 compare1, compare2로 여부를 판단하는데
                                                           //품번검색시 품번이외 텍스트로 검색하면 compare1이 항상 false로 되기때문에 compare2도 추가해줌                                                           
                            {
                                dtCodeTemp.Rows.Add(dr.ItemArray);
                            }
                        }

                        if (dtCodeTemp.Rows.Count == 1)
                        {
                            string col_ID = dtCodeTemp.Rows[0].ItemArray[0].ToString();
                            string col_Name = dtCodeTemp.Rows[0].ItemArray[1].ToString();

                            txtBox.Text = col_Name;
                            txtBox.Tag = col_ID;

                            //검사기준등록에서 쓰기 위해 추가한 부분 2024-05-08 최대현
                            //검사기준등록에서 품번을 직접 일부 또는 전부 입력후 매칭되는 값이 있을 때 
                            //공정, 공정코드를 가지고 오지 못하는 것을 수정하기 위해 추가

                            //동작과정은
                            //빈 텍스트박스에 엔터를 눌러 팝업창에서 직접 선택시 plusFinder.xaml.cs에서 selectedItem()
                            //Win_Qul_InspectAutoBasis_U의 SetBuyerArticleNo()쪽으로 가고
                            //직접 텍스트를 입력하면 이쪽 함수를 거치게 되어있는 것 같습니다.

                            //공정은 품목코드 등록화는 화면에서 선택한 공정을 가지고 오는데 하나의 상품에
                            //두 가지 이상 공정이 있으면 플러스파인더가 팝업으로 뜹니다.
                            if (ColID == "품목코드")
                            {
                                if (rs_dt.Columns[3] != null && rs_dt.Columns[3].Caption == "공정")
                                {
                                    colName2 = dtCodeTemp.Rows[0].ItemArray[3].ToString();
                                    colID2 = dtCodeTemp.Rows[0].ItemArray[4].ToString();
                                    
                                    txtLot.Text = colName2;
                                    txtLot.Tag = colID2;                              
                                }

                            }
                        }
                        //dt count가 0일때
                        //[2]명칭으로 찾기
                        if (dtCodeTemp.Rows.Count == 0)
                        {
                            #region 봉인

                            //if (ColName == "수주번호")
                            //{
                            //    sql = ColName + " LIKE '%" + Data + "%' OR " +
                            //        //ColName + " LIKE '(주)" + Data + "%;" + //vb소스상에 있는데 왜 했는지 모르겠음. 그래서 뺌.
                            //        " OR 품목명 LIKE '" + Data + "%'" +
                            //        " OR 거래처코드 = '" + Data + "'";
                            //}
                            ////else if (ColName == "품명")
                            ////{
                            ////    sql = ColID + " LIKE '%" + Data + "%' OR " + ColName + " LIKE '%" + Data + "%'";
                            ////}
                            //else if (ColID == "사번")
                            //{
                            //    sql = ColID + " LIKE '%" + Data + "%' OR " + ColName + " LIKE '%" + Data + "%' OR " + ColArticle + " LIKE '%" + Data + "%'";
                            //}
                            //else
                            //{
                            //    sql = ColName + " LIKE '%" + Data + "%'";
                            //}

                            //foreach (DataRow dr in rs_dt.Select(sql))
                            //{
                            //    dtCodeTemp.Rows.Add(dr.ItemArray);
                            //}

                            #endregion

                            // 위의 셀렉트 안쓰고 텍스트로 찾으면?
                            foreach (DataRow dr in rs_dt.Rows)
                            {
                                Console.WriteLine(dr.ToString());

                                //string str1 = dr[ColName].ToString();
                                //string str2 = dr[ColID].ToString();

                                if (dr[ColName].ToString().ToUpper().Replace(" ", "").Contains(Data.ToUpper().Replace(" ", "")))
                                {
                                    dtCodeTemp.Rows.Add(dr.ItemArray);
                                }
                            }

                            if (dtCodeTemp.Rows.Count == 1)
                            {
                                string col_ID = dtCodeTemp.Rows[0].ItemArray[0].ToString();
                                string col_Name = dtCodeTemp.Rows[0].ItemArray[1].ToString();

                                txtBox.Text = col_Name;
                                txtBox.Tag = col_ID;
                            }
                        }
                        if (dtCodeTemp.Rows.Count > 1)
                        {
                            mDataGrid.ItemsSource = dtCodeTemp.DefaultView;
                            txtName.Text = txtBox.Text;
                            txtName.Focus();
                            this.ShowDialog();
                            mDataGrid.SelectedIndex = 0;

                        }
                        else if (dtCodeTemp.Rows.Count == 0)
                        {
                            MessageBox.Show("검색결과가 없습니다. 다시 검색해주세요.");
                        }
                    }
                    else
                    {
                        DataTable dtCodeTemp = new DataTable();
                        dtCodeTemp = rs_dt.Clone();

                        foreach (DataRow dr in rs_dt.Rows)
                        {
                            dtCodeTemp.Rows.Add(dr.ItemArray);
                        }
                    }

                }
                else
                {
                    mDataGrid.ItemsSource = rs_dt.DefaultView;
                    txtCode.Text = txtBox.Text;
                    txtCode.Focus();
                    mDataGrid.SelectedIndex = 0;
                    this.ShowDialog();
                  

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 공정코드까지 가져오기
        /// </summary>
        /// <param name="_txtBox"></param>
        /// <param name="_txtProcess"></param>
        /// <param name="large"></param>
        /// <param name="smiddle"></param>
        public void ReturnCode(TextBox _txtBox, TextBox _txtProcess, int large, string smiddle)
        {
            DataClear();
            txtBox = _txtBox;
            txtLot = _txtProcess; //Qty
            ProcQuery(large, smiddle);
            if (rs_dt.Rows.Count > 0)
            {
                DataTableByWhere(txtBox, large);
                mDataGrid.SelectedIndex = 0;
                mDataGrid.SelectedIndex = -1;
            }
            else
            {
                MessageBox.Show("검색결과가 없습니다.");
            }

        }
        /// <summary>
        /// 품번, 품명, 공정 가져오기
        /// </summary>
        /// <param name="_txtBox"></param>
        /// <param name="_txtProcess"></param>
        /// <param name="_txtName"></param>
        /// <param name="large"></param>
        /// <param name="smiddle"></param>
        public void ReturnCode(TextBox _txtBox, TextBox _txtProcess, TextBox _txtBoxName, int large, string smiddle)
        {
            DataClear();
            txtBox = _txtBox;
            txtLot = _txtProcess;
            txtBoxName = _txtBoxName;
            ProcQuery(large, smiddle);
            if (rs_dt.Rows.Count > 0)
            {
                DataTableByWhere(txtBox, large);
            }
            else
            {
                MessageBox.Show("검색결과가 없습니다.");
            }

        }

        public void ReturnCode(TextBox _txtBox, int large, string smiddle)
        {
            DataClear();
            txtBox = _txtBox;
            ProcQuery(large, smiddle);
            if (rs_dt.Rows.Count > 0)
            {
                DataTableByWhere(txtBox, large);
            }
            else
            {
                MessageBox.Show("검색결과가 없습니다.");
            }

        }
        //2021-11-09 자재입고에서 자재입고반품을 위해 Lotid 가져오는 plusfinder 생성
        public void ReturnCodeMTR(TextBox _txtBox, TextBox _txtLot, int large, string smiddle, string CustomID)
        {
            DataClear();
            txtBox = _txtBox; //Lotid
            txtLot = _txtLot; //Qty
            ProcQuery(large, smiddle, CustomID);
            if (rs_dt.Rows.Count > 0)
            {
                DataTableByWhere(txtBox, large);
            }
            else
            {
                MessageBox.Show("검색결과가 없습니다.");
            }

        }



        //GLS전용 플러스파인더 
        public void ReturnCodeGLS(TextBox _txtBox, int large, string smiddle)
        {
            DataClear();
            txtBox = _txtBox;
            ProcQuery(large, smiddle);
            if (rs_dt.Rows.Count > 0)
            {
                DataTableByWhere(txtBox, large);
            }
            else
            {
                MessageBox.Show("거래처별 등록 품목에 등록된 품번이 없습니다.");
            }
        }

        private void SelectItem(object sender, RoutedEventArgs e)
        {

            DataRowView dataRow = (DataRowView)mDataGrid.SelectedItem;
            if (dataRow != null)
            {
                string colID = dataRow.Row.ItemArray[0].ToString();
                string colName = dataRow.Row.ItemArray[1].ToString();
                //2021-11-09 자재입고반품일 경우, LOTID와 수량을 가져가야되서 조건 추가
                if (mDataGrid.Columns[0].Header.ToString() == "거래처")
                {
                    colID = dataRow.Row.ItemArray[1].ToString();
                    colName = dataRow.Row.ItemArray[2].ToString();

                    txtBox.Text = colID;
                    txtLot.Text = colName;

                    this.DialogResult = DialogResult.HasValue;
                    this.Close();


                }
                else if (mDataGrid.Columns[0].Header.ToString() == "품목코드")
                {
                    string colID2 = dataRow.Row.ItemArray[4].ToString();
                    string colName2 = dataRow.Row.ItemArray[3].ToString();

                    txtBox.Text = colName;
                    txtBox.Tag = colID;
                    txtLot.Text = colName2;
                    txtLot.Tag = colID2;

                    this.DialogResult = DialogResult.HasValue;
                    this.Close();
                }
                else
                {
                    if (mDataGrid.Columns[0].Header.ToString() == "사번")
                    {
                        colID = dataRow.Row.ItemArray[1].ToString();
                        colName = dataRow.Row.ItemArray[2].ToString();
                    }
                    if (mDataGrid.Columns[0].Header.ToString() == "품명"
                         && mDataGrid.Columns[5].Header.ToString() == "공정명")
                    {
                        string Process = dataRow.Row.ItemArray[5].ToString();
                        string ProcessID = dataRow.Row.ItemArray[6].ToString();
                        refEvent?.Invoke($"{Process},{ProcessID}");
                        //refEvent?.Invoke(dataRow.Row.ItemArray[6].ToString());
                        //refEvent?.Invoke(dataRow.Row.ItemArray[11].ToString());
                    }
                    if (mDataGrid.Columns[1].Header.ToString() == "출고지시번호") //검색조건 Null오류 피하기 위해서 단계별 조건..
                    {
                        if (mDataGrid.Columns.Count == 9)
                        {
                            if (mDataGrid.Columns[9].Header.ToString() == "출고잔량")
                            {
                                string article = dataRow.Row.ItemArray[4].ToString();
                                string orderQty = dataRow.Row.ItemArray[6].ToString();
                                string reqQty = dataRow.Row.ItemArray[7].ToString();
                                string reqRemainQty = dataRow.Row.ItemArray[9].ToString();
                                refEvent?.Invoke($"{article},{orderQty},{reqQty},{reqRemainQty}");
                            }
                        }
                        #region 관리번호 // 지시번호 나누기 전
                        //refEvent?.Invoke(dataRow.Row.ItemArray[6].ToString());
                        //refEvent?.Invoke(dataRow.Row.ItemArray[11].ToString());


                        //// Large가 98번인 경우 품명(Article) callback
                        //if (mDataGrid.Columns[1].Header.ToString() == "출고지시번호")
                        //{
                        //    refEvent?.Invoke(dataRow.Row.ItemArray[4].ToString());
                        //}
                        //else
                        //{
                        //    refEvent?.Invoke(dataRow.Row.ItemArray[2].ToString());
                        //}
                        #endregion

                    }           
                    if (mDataGrid.Columns[1].Header.ToString() == "관리번호" &&
                        mDataGrid.Columns[2].Header.ToString() == "거래처"   &&
                        mDataGrid.Columns[3].Header.ToString() == "품명")
                    {
                        string orderQty1 = dataRow.Row.ItemArray[4].ToString();

                        refEvent?.Invoke($"{orderQty1},");
                    }




                    txtBox.Text = colName;
                    txtBox.Tag = colID;                    

                    this.DialogResult = DialogResult.HasValue;
                    this.Close();
                }
            }
        }



        private void Lang_Checked(object sender, RoutedEventArgs e)
        {
            var button = sender as RadioButton;
            if (button.Name is "LangKor")
            {
                GridKor.Visibility = Visibility.Visible;
                GridEng.Visibility = Visibility.Hidden;
            }
            else if (button.Name is "LangEng")
            {
                GridEng.Visibility = Visibility.Visible;
                GridKor.Visibility = Visibility.Hidden;
            }

        }

        private void btnLanguage_Click(object sender, RoutedEventArgs e)
        {
            if (rs_dt is null)//결과 DataTable이 Null이면 Return
            {
                return;
            }

            var btnLanguage = sender as Button;
            string Lang = btnLanguage.Content.ToString();
            string ColID = "";
            string ColName = "";
            string strCode = "";
            string strName = "";
            string sql = "";

            if (rs_dt.Rows.Count > 0)//결과 DataTable이 Row가 0개 이상일때
            {
                if (rs_dt.Columns.Count >= 2)//결과 DataTable의 Column 갯수가 2개 이상일때(Code_ID,Code_Name 둘다 사용하기위해서)
                {
                    ColID = rs_dt.Columns[0].ToString();    //코드
                    ColName = rs_dt.Columns[1].ToString();  //명칭
                    strCode = txtCode.Text.Trim();          //코드 텍스트박스 
                    strName = txtName.Text.Trim();          //명칭 텍스트박스

                    //한글 or 영어 초성 버튼 클릭으로 인한 이벤트
                    //한글
                    //ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅍㅎㅋ
                    if (LangKor.IsChecked == true)
                    {
                        string hangle_jaeum = "ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ힣";
                        string startchar = Lang;
                        string endchar = hangle_jaeum.Substring(hangle_jaeum.IndexOf(startchar) + 1, 1);

                        if (Lang == "기타")
                        {
                            sql = ColName + " < 'ㄱ' OR " + ColName + "> '힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣'";
                        }
                        else
                        {
                            if (endchar.Equals("힣"))
                            {
                                endchar = "힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣힣";
                            }
                            //string sql = ColName + " Between '" + startchar + "' And '" + endchar + "'";
                            sql = ColName + " >= '" + startchar + "' And " + ColName + " < '" + endchar + "'";
                        }
                    }
                    //영어
                    else if (LangEng.IsChecked == true)
                    {
                        if (Lang.ToLower() == "else")
                        {
                            sql = ColName + " < 'A' OR " + ColName + "> 'zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz'";
                        }
                        else
                        {
                            sql = ColName + " Like '" + Lang + "%'";
                        }

                    }
                    //mDataGrid.Columns.Clear();
                    mDataGrid.ItemsSource = null; // 초기화
                    DataTable temp_dt = new DataTable();//dt.select where절 조건건 
                    temp_dt = rs_dt.Clone();//테이블 구조 복사

                    foreach (DataRow dr in rs_dt.Select(sql))
                    {
                        temp_dt.Rows.Add(dr.ItemArray);
                    }

                    mDataGrid.ItemsSource = temp_dt.DefaultView;// dt중 조건 걸어서 검색

                }

            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void mDataGrid_KeyPress(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                SelectItem(sender, e);
            }
        }

        // 코드 엔터 → 검색
        private void txtCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                SelectDataTable(null, null);
            }
        }
        // 이름 엔터 → 검색
        private void txtName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                SelectDataTable(null, null);
            }
        }
    }
}


