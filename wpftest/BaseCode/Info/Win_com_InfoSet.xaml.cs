using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_com_Infoset_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_com_InfoSet : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        string strFlag = string.Empty;
        int rowNum = 0;

        int rowNumAll = 0;
        int rowNumPerson = 0;

        public static List<PersonViewModel> mMenulist = new List<PersonViewModel>();
        public static List<PersonViewModel> mMenuSublist = new List<PersonViewModel>();

        string Left_Click_Person_name = string.Empty;        //전역변수 > 클릭된 개인명
        string Left_Click_Person_ID = string.Empty;          // 왼쪽변          
        string Right_Click_Person_Name = string.Empty;
        string Right_Click_Person_ID = string.Empty;        //전역변수 > 클릭된 개인ID   오른쪽 변.

        // FTP 활용모음.
        string strImagePath = string.Empty;
        string strFullPath = string.Empty;

        List<string[]> listFtpFile = new List<string[]>();
        List<string[]> deleteListFtpFile = new List<string[]>(); // 삭제할 파일 리스트
        private FTP_EX _ftp = null;

        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData";
        //string FTP_ADDRESS = "ftp://wizis.iptime.org/ImageData/Draw";
        string FTP_ADDRESS = "ftp://" + LoadINI.FileSvr + ":" + LoadINI.FTPPort + LoadINI.FtpImagePath + "/Info";
        private const string FTP_ID = "wizuser";
        private const string FTP_PASS = "wiz9999";
        private const string LOCAL_DOWN_PATH = "C:\\Temp";


        //string FTP_ADDRESS = "ftp://192.168.0.4/Info";
        //string FTP_ADDRESS = "ftp://192.168.0.120";


        public Win_com_InfoSet()
        {
            InitializeComponent();
        }


        //
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            Lib.Instance.UiLoading(sender);

            // 기존에 있던 공지사항 사원 리스트 제거
            mMenulist.Clear();
            mMenuSublist.Clear();

            FillPersonInfo();       // 개별 사용자 DB정보 가져와서 리스트에 집어넣기.
            MakeTree();             // 리스트 정보를 바탕으로 트리 만들어서 출력

            chkSearchDay.IsChecked = true;
            //FromDateSearch.SelectedDate = DateTime.Today;
            //ToDateSearch.SelectedDate = DateTime.Today;

            // 금월 세팅
            btnThisMonth_Click(null, null);

            // 보기 버튼체크
            btnImgSeeCheckAndSetting();
        }


        #region 추가 수정시 / 취소 저장완료시 메서드

        // 추가, 수정시
        private void SaveUpdateMode()
        {

            // 상단바
            btnSave.IsEnabled = true;
            btnSearch.IsEnabled = false;

            if (strFlag.Equals("IA") || strFlag.Equals("UA"))
            {
                // 전체 공지사항 추가시

                // 개별 공지사항 그룹박스 사용제한
                gbxNotice_Person.IsEnabled = false;

                // 전체 공지사항 데이터 그리드
                dgdAll.IsEnabled = false;

                // 공지사항 설정
                FromDate_All.IsHitTestVisible = true;
                ToDate_All.IsHitTestVisible = true;

                btnAddAll.IsEnabled = false;
                btnUpdateAll.IsEnabled = false;
                btnDeleteAll.IsEnabled = false;

                btnCancelAll.IsEnabled = true;

                txtContent_All.IsHitTestVisible = true;

                // 첨부파일
                btn_AddAttatch_M.IsEnabled = true;
                btn_DelAttatch_M.IsEnabled = true;

                btn_AddAttatch_D1.IsEnabled = true;
                btn_DelAttatch_D1.IsEnabled = true;

                btn_AddAttatch_D2.IsEnabled = true;
                btn_DelAttatch_D2.IsEnabled = true;

                // 보기 버튼체크
                btnImgSeeCheckAndSetting();
            }
            else if (strFlag.Equals("IP") || strFlag.Equals("UP"))
            {
                // 개별 공지사항 추가시

                // 전체 공지사항 그룹박스 제한
                gbxNotice_All.IsEnabled = false;

                // 개별 공지사항 데이터 그리드
                dgdPerson.IsEnabled = false;

                // 공지사항 설정
                FromDate_Person.IsHitTestVisible = true;
                ToDate_Person.IsHitTestVisible = true;

                btnAddPerson.IsEnabled = false;
                btnUpdatePerson.IsEnabled = false;
                btnDeletePerson.IsEnabled = false;

                btnCancelPerson.IsEnabled = true;

                txtContent_Person.IsHitTestVisible = true;

                // 공지 대상 사원 등록
                // 공지사항 트리뷰 배경색 설정
                treeview_test.Background = Brushes.White;

                treeview_test.IsEnabled = true;
                btnAddSelectItem.IsEnabled = true;
                btnDelSelectItem.IsEnabled = true;
                dgdTargetPerson.IsEnabled = true;

                // 보기 버튼체크
                btnImgSeeCheckAndSetting();
            }

        }

        // 저장 완료, 취소시
        private void CompleteCancelMode()
        {
            // 상단바
            btnSave.IsEnabled = false;
            btnSearch.IsEnabled = true;

            // 전체, 개별 그룹박스 활성화
            gbxNotice_Person.IsEnabled = true;
            gbxNotice_All.IsEnabled = true;

            // 전체
            // 전체 공지사항 데이터 그리드
            dgdAll.IsEnabled = true;

            // 공지사항 설정
            FromDate_All.IsHitTestVisible = false;
            ToDate_All.IsHitTestVisible = false;

            btnAddAll.IsEnabled = true;
            btnUpdateAll.IsEnabled = true;
            btnDeleteAll.IsEnabled = true;

            btnCancelAll.IsEnabled = false;

            txtContent_All.IsHitTestVisible = false;

            // 첨부파일
            btn_AddAttatch_M.IsEnabled = false;
            btn_DelAttatch_M.IsEnabled = false;

            btn_AddAttatch_D1.IsEnabled = false;
            btn_DelAttatch_D1.IsEnabled = false;

            btn_AddAttatch_D2.IsEnabled = false;
            btn_DelAttatch_D2.IsEnabled = false;

            // 개별
            // 개별 공지사항 데이터 그리드
            dgdPerson.IsEnabled = true;

            // 공지사항 설정
            FromDate_Person.IsHitTestVisible = false;
            ToDate_Person.IsHitTestVisible = false;

            btnAddPerson.IsEnabled = true;
            btnUpdatePerson.IsEnabled = true;
            btnDeletePerson.IsEnabled = true;

            btnCancelPerson.IsEnabled = false;

            txtContent_Person.IsHitTestVisible = false;

            // 공지 대상 사원 등록
            // 공지사항 트리뷰 배경색 설정
            var bc = new BrushConverter();
            treeview_test.Background = (Brush)bc.ConvertFrom("#ededed");

            treeview_test.IsEnabled = false;
            btnAddSelectItem.IsEnabled = false;
            btnDelSelectItem.IsEnabled = false;
            dgdTargetPerson.IsEnabled = false;

            // 전체, 개별 텍스트박스 비우기
            txtContent_All.Text = "";
            txtContent_Person.Text = "";
        }

        #endregion // 추가 수정시 / 취소 저장완료시 메서드



        #region Header 부분 - 검색조건

        private void lblSearchDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkSearchDay.IsChecked == true)
            {
                chkSearchDay.IsChecked = false;
            }
            else
            {
                chkSearchDay.IsChecked = true;
            }
        }

        // 검색일자
        // 검색일자 체크박스 체크 이벤트
        private void chkSearchDay_Checked(object sender, RoutedEventArgs e)
        {
            chkSearchDay.IsChecked = true;
            FromDateSearch.IsEnabled = true;
            ToDateSearch.IsEnabled = true;
        }
        // 검색일자 체크박스 언체크 이벤트
        private void chkSearchDay_Unchecked(object sender, RoutedEventArgs e)
        {
            chkSearchDay.IsChecked = false;
            FromDateSearch.IsEnabled = false;
            ToDateSearch.IsEnabled = false;
        }

        //전일
        private void btnYesterDay_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastDayDateTimeContinue(ToDateSearch.SelectedDate.Value);

            FromDateSearch.SelectedDate = SearchDate[0];
            ToDateSearch.SelectedDate = SearchDate[1];
        }

        //금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            FromDateSearch.SelectedDate = DateTime.Today;
            ToDateSearch.SelectedDate = DateTime.Today;
        }

        // 전월 버튼 클릭 이벤트
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            DateTime[] SearchDate = lib.BringLastMonthContinue(FromDateSearch.SelectedDate.Value);

            FromDateSearch.SelectedDate = SearchDate[0];
            ToDateSearch.SelectedDate = SearchDate[1];
        }

        // 금월 버튼 클릭 이벤트
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            FromDateSearch.SelectedDate = lib.BringThisMonthDatetimeList()[0];
            ToDateSearch.SelectedDate = lib.BringThisMonthDatetimeList()[1];
        }

        #endregion // Header 부분 - 검색조건

        #region Header 부분 - 상단 오른쪽 버튼

        // 오른쪽 버튼 이벤트 : 저장, 검색, 닫기
        // 저장 버튼 클릭 이벤트
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData(strFlag))
            {
                CompleteCancelMode();

                if (strFlag.Trim().Equals("IA") || strFlag.Trim().Equals("UA"))
                {
                    rowNum = 0;
                    re_Search(rowNumAll, rowNum);
                }
                else if (strFlag.Trim().Equals("IP") || strFlag.Trim().Equals("UP"))
                {
                    rowNum = 0;
                    re_Search(rowNum, rowNumPerson);
                }
                else
                {
                    rowNum = 0;
                    re_Search(rowNum, rowNum);
                }

                strFlag = string.Empty;
            }


        }
        // 닫기 버튼 클릭 이벤트
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }
        // 검색 버튼 클릭 이벤트
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            //검색버튼 비활성화
            btnSearch.IsEnabled = false;

            Dispatcher.BeginInvoke(new Action(() =>

            {
                Thread.Sleep(2000);

                //로직
                rowNum = 0;
                re_Search(rowNum, rowNum);

            }), System.Windows.Threading.DispatcherPriority.Background);



            Dispatcher.BeginInvoke(new Action(() =>

            {
                btnSearch.IsEnabled = true;

            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        #region 저장 메서드

        // 저장
        public bool SaveData(string strFlag)
        {
            bool flag = false;
            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();

            string GetKey = "";

            try
            {
                if (CheckData())
                {
                    // All Or Person  전체 공지사항, 개별 공지사항 인지
                    string aop = strFlag.Substring(1, 1);

                    if (aop.Equals("A")) // 전체공지사항 저장
                    {
                        sqlParameter.Clear();

                        sqlParameter.Add("sCompanyID", "");
                        sqlParameter.Add("sFromDate", FromDate_All.SelectedDate.Value.ToString("yyyyMMdd"));
                        sqlParameter.Add("sToDate", ToDate_All.SelectedDate.Value.ToString("yyyyMMdd"));
                        sqlParameter.Add("Info", txtContent_All.Text);
                        sqlParameter.Add("UserID", MainWindow.CurrentUser);

                        if (strFlag.Equals("IA")) // 전체 공지사항 추가
                        {
                            sqlParameter.Add("sInfoID", "");

                            Dictionary<string, int> outputParam = new Dictionary<string, int>();
                            outputParam.Add("sInfoID", 10);
                            Dictionary<string, string> dicResult = DataStore.Instance.ExecuteProcedureOutputNoTran("xp_Info_iInfo", sqlParameter, outputParam, true);

                            GetKey = dicResult["sInfoID"];

                            if ((GetKey != string.Empty) && (GetKey != "9999"))
                            {
                                flag = true;
                            }
                            else
                            {
                                MessageBox.Show("[저장실패]\r\n" + GetKey);
                                flag = false;
                                return false;
                            }
                        }
                        else if (strFlag.Equals("UA")) // 전체 공지사항 수정
                        {
                            var NoticeAll = dgdAll.SelectedItem as Win_info_Infoset_U_CodeView_All;

                            if (NoticeAll != null)
                            {
                                GetKey = NoticeAll.InfoID;

                                sqlParameter.Add("sInfoID", NoticeAll.InfoID);

                                Procedure pro1 = new Procedure();
                                pro1.Name = "xp_Info_uInfo";
                                pro1.OutputUseYN = "N";
                                pro1.OutputName = "sArticleID";
                                pro1.OutputLength = "10";

                                Prolist.Add(pro1);
                                ListParameter.Add(sqlParameter);

                                string[] Confirm = new string[2];
                                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                                if (Confirm[0] != "success")
                                {
                                    MessageBox.Show("[저장실패]\r\n" + Confirm[1].ToString());
                                    flag = false;
                                    return false;
                                }
                                else
                                {
                                    flag = true;

                                }
                            }
                        }

                        // FTP 파일 업로드 AttachFileUpdate
                        // 파일을 올리자 : GetKey != "" 라면 파일을 올려보자
                        if (!GetKey.Trim().Equals(""))
                        {
                            if (deleteListFtpFile.Count > 0)
                            {
                                foreach (string[] str in deleteListFtpFile)
                                {
                                    FTP_RemoveFile(GetKey + "/" + str[0]);
                                }
                            }

                            if (listFtpFile.Count > 0)
                            {
                                FTP_Save_File(listFtpFile, GetKey);
                            }
                            AttachFileUpdate(GetKey);
                        }

                        // 파일 List 비워주기
                        listFtpFile.Clear();
                        deleteListFtpFile.Clear();
                    }
                    else if (aop.Equals("P")) // 개별공지사항 저장
                    {
                        sqlParameter.Clear();

                        sqlParameter.Add("sCompanyID", "");
                        sqlParameter.Add("sFromDate", FromDate_Person.SelectedDate.Value.ToString("yyyyMMdd"));
                        sqlParameter.Add("sToDate", ToDate_Person.SelectedDate.Value.ToString("yyyyMMdd"));
                        sqlParameter.Add("Info", txtContent_Person.Text);
                        sqlParameter.Add("UserID", MainWindow.CurrentUser);

                        // 개별 공지사항 저장 → 생성된 InfoID 를 output
                        if (strFlag.Equals("IP"))
                        {
                            sqlParameter.Add("sInfoID", "");

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Info_iInfoUser";
                            pro1.OutputUseYN = "Y";
                            pro1.OutputName = "sInfoID";
                            pro1.OutputLength = "10";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);

                            List<KeyValue> list_Result = new List<KeyValue>();
                            list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS(Prolist, ListParameter);
                            string sInfoID = string.Empty;

                            if (list_Result[0].key.ToLower() == "success")
                            {
                                list_Result.RemoveAt(0);
                                for (int i = 0; i < list_Result.Count; i++)
                                {
                                    KeyValue kv = list_Result[i];
                                    if (kv.key == "sInfoID")
                                    {
                                        sInfoID = kv.value;
                                        flag = true;
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("[저장실패]\r\n" + list_Result[0].value.ToString());
                                flag = false;
                                //return false;
                            }

                            // 프로시저를 실행했으니 초기화
                            Prolist.Clear();
                            ListParameter.Clear();

                            // output 된 InfoID 를 가지고 공지대상 저장
                            for (int i = 0; i < dgdTargetPerson.Items.Count; i++)
                            {
                                Dictionary<string, object> sqlParameterSub = new Dictionary<string, object>();
                                sqlParameterSub.Clear();

                                var TargetPerson = dgdTargetPerson.Items[i] as PersonViewModel;
                                string sPersonID = TargetPerson.PersonID;

                                sqlParameterSub.Add("sInfoID", sInfoID);
                                sqlParameterSub.Add("sPersonID", sPersonID);

                                Procedure pro2 = new Procedure();
                                pro2.Name = "xp_Info_iInfoUserList";
                                pro2.OutputUseYN = "N";
                                pro2.OutputName = "sInfoID";
                                pro2.OutputLength = "10";

                                Prolist.Add(pro2);
                                ListParameter.Add(sqlParameterSub);
                            }
                        }
                        else if (strFlag.Equals("UP")) // 개별공지사항 내용은 수정(InfoUser 테이블) / 공지 대상(테이블 : InfoUserSub) 은 전부 삭제 후에 / 다시 등록
                        {
                            // 개별 공지사항 수정 + 개별 공지대상 모두 삭제
                            var NoticePerson = dgdPerson.SelectedItem as Win_info_Infoset_U_CodeView_Person;

                            sqlParameter.Add("sInfoID", NoticePerson.per_InfoID);

                            Procedure pro1 = new Procedure();
                            pro1.Name = "xp_Info_uInfoUser";
                            pro1.OutputUseYN = "N";
                            pro1.OutputName = "sInfoID";
                            pro1.OutputLength = "10";

                            Prolist.Add(pro1);
                            ListParameter.Add(sqlParameter);

                            string[] result = new string[2];
                            result = DataStore.Instance.ExecuteAllProcedureOutputNew(Prolist, ListParameter);
                            if (result[0] != "success")
                            {
                                MessageBox.Show("[저장실패]\r\n" + result[1].ToString());
                                flag = false;
                                //return false;
                            }
                            else
                            {
                                flag = true;
                            }

                            // 프로시저를 실행했으니 초기화
                            Prolist.Clear();
                            ListParameter.Clear();

                            // 개별공지대상 수정
                            for (int i = 0; i < dgdTargetPerson.Items.Count; i++)
                            {
                                Dictionary<string, object> sqlParameterSub = new Dictionary<string, object>();
                                sqlParameterSub.Clear();

                                var TargetPerson = dgdTargetPerson.Items[i] as PersonViewModel;
                                string sPersonID = TargetPerson.PersonID;

                                sqlParameterSub.Add("sInfoID", NoticePerson.per_InfoID);
                                sqlParameterSub.Add("sPersonID", sPersonID);

                                Procedure pro2 = new Procedure();
                                pro2.Name = "xp_Info_iInfoUserList";
                                pro2.OutputUseYN = "N";
                                pro2.OutputName = "sInfoID";
                                pro2.OutputLength = "10";

                                Prolist.Add(pro2);
                                ListParameter.Add(sqlParameterSub);
                            }
                        }

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
                            flag = true;
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

            return flag;
        }

        #endregion // 저장 메서드


        #region 유효성 검사

        private bool CheckData()
        {
            bool flag = true;

            if (strFlag.Trim().Equals("IA") || strFlag.Trim().Equals("UA"))
            {
                if (FromDate_All.SelectedDate == null
                    || FromDate_All.SelectedDate.ToString().Trim().Equals("")
                    || ToDate_All.SelectedDate == null
                    || ToDate_All.SelectedDate.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("공지기간을 선택해주세요.");
                    flag = false;
                    return flag;
                }
            }

            if (strFlag.Trim().Equals("IP") || strFlag.Trim().Equals("UP"))
            {
                if (FromDate_Person.SelectedDate == null
                   || FromDate_Person.SelectedDate.ToString().Trim().Equals("")
                   || ToDate_Person.SelectedDate == null
                   || ToDate_Person.SelectedDate.ToString().Trim().Equals(""))
                {
                    MessageBox.Show("공지기간을 선택해주세요.");
                    flag = false;
                    return flag;
                }

                // 선택된 사람이 아무도 없을경우!
                if (dgdTargetPerson.Items.Count < 1)
                {
                    MessageBox.Show("공지할 대상을 선택해주세요.");
                    flag = false;
                    return flag;
                }
            }

            return flag;
        }

        #endregion // 유효성 검사

        #endregion // Header 부분 - 상단 오른쪽 버튼

        #region 전체공지사항

        // 전체 공지사항 그리드 선택 이벤트
        private void dgdAll_SelectChanged(object sender, SelectionChangedEventArgs e)
        {
            var NoticeAll = dgdAll.SelectedItem as Win_info_Infoset_U_CodeView_All;

            if (NoticeAll != null)
            {
                FromDate_All.SelectedDate = ConvertDateTime(NoticeAll.FromDate);
                ToDate_All.SelectedDate = ConvertDateTime(NoticeAll.ToDate);
                txtContent_All.Text = NoticeAll.Info;

                //txtFileName1.Text = NoticeAll.PartFile;
                //txtFileName1.Text = NoticeAll.AttachFile2;
                //txtFileName3.Text = NoticeAll.AttachFile3;

                //txtFilePath1.Text = NoticeAll.PartPath;
                //txtFilePath2.Text = NoticeAll.AttachPath2;
                //txtFilePath3.Text = NoticeAll.AttachPath2;

                this.DataContext = NoticeAll;
            }

            // 보기 버튼체크
            btnImgSeeCheckAndSetting();
        }

        // 20190919 형식의 문자열을 DateTime으로 변환 : 2019-09-19 로 변환후에 DateTime 으로 변환
        private DateTime StringToDateTime(string str)
        {
            DateTime result = DateTime.Now.Date;
            DateTime chkDT = DateTime.Now.Date;

            if (str.Length == 8)
            {
                string year = str.Substring(0, 4);
                string month = str.Substring(4, 2);
                string day = str.Substring(6, 2);

                string date = year + "-" + month + "-" + day;

                if (DateTime.TryParse(date, out chkDT) == true)
                {
                    result = DateTime.Parse(date);
                }
            }


            return result;
        }

        // 2019-09-19 형식의 문자열을 DateTime으로 변환
        private DateTime ConvertDateTime(string str)
        {
            DateTime result = DateTime.Now.Date;
            DateTime chkDT = DateTime.Now.Date;

            if (DateTime.TryParse(str, out chkDT) == true)
            {
                result = DateTime.Parse(str);
            }

            return result;
        }


        // 전체 공지사항 추가 버튼 클릭 이벤트
        private void btnAddAll_Click(object sender, RoutedEventArgs e)
        {
            // 공지사항 지정에 오늘날짜가 선택되도록
            FromDate_All.SelectedDate = DateTime.Now.Date;
            ToDate_All.SelectedDate = DateTime.Now.Date;

            strFlag = "IA";
            SaveUpdateMode();

            txtContent_All.Text = "";

            // 첨부파일 비우기
            txtFileName1.Text = "";
            txtFileName2.Text = "";
            txtFileName3.Text = "";

            txtFilePath1.Text = "";
            txtFilePath2.Text = "";
            txtFilePath3.Text = "";

            rowNumAll = dgdAll.SelectedIndex;
        }

        // 전체 공지사항 수정 버튼 클릭 이벤트
        private void btnUpdateAll_Click(object sender, RoutedEventArgs e)
        {
            var NoticeAll = dgdAll.SelectedItem as Win_info_Infoset_U_CodeView_All;

            if (NoticeAll != null)
            {
                strFlag = "UA";
                SaveUpdateMode();
            }
            else
            {
                MessageBox.Show("수정할 전체공지사항을 선택해주세요.");
                return;
            }

            rowNumAll = dgdAll.SelectedIndex;
        }

        // 전체 공지사항 삭제 버튼 클릭 이벤트
        private void btnDeleteAll_Click(object sender, RoutedEventArgs e)
        {
            var NoticeAll = dgdAll.SelectedItem as Win_info_Infoset_U_CodeView_All;

            if (NoticeAll != null)
            {
                MessageBoxResult msgresult = MessageBox.Show("해당 공지사항을 삭제하시겠습니까?", "확인", MessageBoxButton.OKCancel);
                if (msgresult == MessageBoxResult.OK)
                {
                    strFlag = "DA";

                    if (DeleteData(strFlag, NoticeAll.InfoID))
                    {
                        rowNum = 0;
                        re_Search(rowNum, rowNumPerson);
                    }
                }
            }
            else
            {
                MessageBox.Show("삭제할 전체공지사항을 선택해주세요.");
                return;
            }

        }

        // 전체 공지사항 취소 버튼 클릭 이벤트
        private void btnCancelAll_Click(object sender, RoutedEventArgs e)
        {
            CompleteCancelMode();

            rowNum = 0;
            re_Search(rowNumAll, rowNumPerson);
        }

        #region 전체공지사항 첨부파일 메서드

        // 첨부파일 찾기 버튼 클릭 이벤트
        private void btn_AddAttach_Click(object sender, RoutedEventArgs e)
        {
            // 1 : 모니터링 / 2 : 첨부파일1 / 3 : 첨부파일2
            string buttonIndex = ((Button)sender).Tag.ToString();

            if (buttonIndex.Trim().Equals("1")) { FTP_Upload_TextBox(txtFileName1); }
            else if (buttonIndex.Trim().Equals("2")) { FTP_Upload_TextBox(txtFileName2); }
            else if (buttonIndex.Trim().Equals("3")) { FTP_Upload_TextBox(txtFileName3); }
        }

        #region FTP_Upload_TextBox - 파일 경로, 이름 텍스트박스에 올림 + 리스트에 ADD

        private void FTP_Upload_TextBox(TextBox textBox)
        {
            if (!textBox.Text.Equals(string.Empty) && strFlag.Equals("U"))
            {
                MessageBox.Show("먼저 해당파일의 삭제를 진행 후 진행해주세요.");
                return;
            }
            else
            {
                Microsoft.Win32.OpenFileDialog OFdlg = new Microsoft.Win32.OpenFileDialog();
                OFdlg.Filter =
                    "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png, *.pcx, *.pdf) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png; *.pcx; *.pdf | All Files|*.*";

                Nullable<bool> result = OFdlg.ShowDialog();
                if (result == true)
                {
                    strFullPath = OFdlg.FileName;

                    string ImageFileName = OFdlg.SafeFileName;  //명.
                    string ImageFilePath = string.Empty;       // 경로

                    ImageFilePath = strFullPath.Replace(ImageFileName, "");

                    StreamReader sr = new StreamReader(OFdlg.FileName);
                    long FileSize = sr.BaseStream.Length;
                    if (sr.BaseStream.Length > (2048 * 1000))
                    {
                        //업로드 파일 사이즈범위 초과
                        MessageBox.Show("이미지의 파일사이즈가 2M byte를 초과하였습니다.");
                        sr.Close();
                        return;
                    }
                    else
                    {
                        textBox.Text = ImageFileName;
                        textBox.Tag = ImageFilePath;

                        string[] strTemp = new string[] { ImageFileName, ImageFilePath.ToString() };
                        listFtpFile.Add(strTemp);
                    }
                }
            }
        }

        #endregion // FTP_Upload_TextBox - 파일 경로, 이름 텍스트박스에 올림 + 리스트에 ADD

        #region FTP_Save_File - 파일 저장, 폴더 생성

        /// <summary>
        /// 해당영역에 파일 있는지 확인
        /// </summary>
        bool FileInfoAndFlag(string[] strFileList, string FileName)
        {
            bool flag = false;
            foreach (string FileList in strFileList)
            {
                if (FileList == FileName)
                {
                    flag = true;
                    break;
                }
            }
            return flag;
        }

        /// <summary>
        /// 해당영역에 폴더가 있는지 확인
        /// </summary>
        bool FolderInfoAndFlag(string[] strFolderList, string FolderName)
        {
            bool flag = false;
            foreach (string FolderList in strFolderList)
            {
                if (FolderList == FolderName)
                {
                    flag = true;
                    break;
                }
            }
            return flag;
        }

        // 파일 저장하기.
        private void FTP_Save_File(List<string[]> listStrArrayFileInfo, string MakeFolderName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

            List<string[]> UpdateFilesInfo = new List<string[]>();
            string[] fileListSimple;
            string[] fileListDetail = null;
            fileListSimple = _ftp.directoryListSimple("", Encoding.Default);

            // 기존 폴더 확인작업.
            bool MakeFolder = false;
            MakeFolder = FolderInfoAndFlag(fileListSimple, MakeFolderName);

            if (MakeFolder == false)        // 같은 아이를 찾지 못한경우,
            {
                //MIL 폴더에 InspectionID로 저장
                if (_ftp.createDirectory(MakeFolderName) == false)
                {
                    MessageBox.Show("업로드를 위한 폴더를 생성할 수 없습니다.");
                    return;
                }
            }
            else
            {
                fileListDetail = _ftp.directoryListSimple(MakeFolderName, Encoding.Default);
            }

            for (int i = 0; i < listStrArrayFileInfo.Count; i++)
            {
                bool flag = true;

                if (fileListDetail != null)
                {
                    foreach (string compare in fileListDetail)
                    {
                        if (compare.Equals(listStrArrayFileInfo[i][0]))
                        {
                            flag = false;
                            break;
                        }
                    }
                }

                if (flag)
                {
                    listStrArrayFileInfo[i][0] = MakeFolderName + "/" + listStrArrayFileInfo[i][0];
                    UpdateFilesInfo.Add(listStrArrayFileInfo[i]);
                }
            }

            if (!_ftp.UploadTempFilesToFTP(UpdateFilesInfo))
            {
                MessageBox.Show("파일업로드에 실패하였습니다.");
                return;
            }
        }

        #endregion // FTP_Save_File - 파일 저장, 폴더 생성

        #region 이미지 파일이 있으면 보기버튼 활성화, 아니면 비활성화

        private void btnImgSeeCheckAndSetting()
        {
            if (!txtFileName1.Text.Trim().Equals(""))
            {
                btn_DownAttatch_M.IsEnabled = true;
                //if (strFlag.Trim().Equals("UA") || strFlag.Trim().Equals("IA")) { btn_DelAttatch_M.IsEnabled = true; }
            }
            else
            {
                btn_DownAttatch_M.IsEnabled = false;
                //if (strFlag.Trim().Equals("UA") || strFlag.Trim().Equals("IA")) { btn_DelAttatch_M.IsEnabled = false; }
            }

            if (!txtFileName2.Text.Trim().Equals(""))
            {
                btn_DownAttatch_D1.IsEnabled = true;
                //if (strFlag.Trim().Equals("UA") || strFlag.Trim().Equals("IA")) { btn_DelAttatch_D1.IsEnabled = true; }
            }
            else
            {
                btn_DownAttatch_D1.IsEnabled = false;
                //if (strFlag.Trim().Equals("UA") || strFlag.Trim().Equals("IA")) { btn_DelAttatch_D1.IsEnabled = false; }
            }

            if (!txtFileName3.Text.Trim().Equals(""))
            {
                btn_DownAttatch_D2.IsEnabled = true;
                //if (strFlag.Trim().Equals("UA") || strFlag.Trim().Equals("IA")) { btn_DelAttatch_D2.IsEnabled = true; }
            }
            else
            {
                btn_DownAttatch_D2.IsEnabled = false;
                //if (strFlag.Trim().Equals("UA") || strFlag.Trim().Equals("IA")) { btn_DelAttatch_D2.IsEnabled = false; }
            }
        }

        #endregion // 이미지 파일이 있으면 보기버튼 활성화, 아니면 비활성화

        #region FTP 파일 삭제


        //파일만 삭제 - 버튼에 Tag로 구분
        private void btn_DelAttach_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 삭제 하시겠습니까?", "삭제 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                string buttonIndex = ((Button)sender).Tag.ToString();

                if (buttonIndex.Trim().Equals("1") && (txtFileName1.Text != string.Empty)) { FileDeleteAndTextBoxEmpty(txtFileName1); }
                else if (buttonIndex.Trim().Equals("2") && (txtFileName2.Text != string.Empty)) { FileDeleteAndTextBoxEmpty(txtFileName2); }
                else if (buttonIndex.Trim().Equals("3") && (txtFileName3.Text != string.Empty)) { FileDeleteAndTextBoxEmpty(txtFileName3); }
            }

            // 보기 버튼체크
            btnImgSeeCheckAndSetting();
        }
        private void FileDeleteAndTextBoxEmpty(TextBox txt)
        {
            if (strFlag.Equals("UA"))
            {
                var Info = dgdAll.SelectedItem as Win_info_Infoset_U_CodeView_All;

                if (Info != null)
                {
                    //FTP_RemoveFile(Article.ArticleID + "/" + txt.Text);

                    // 파일이름, 파일경로
                    string[] strFtp = { txt.Text, txt.Tag != null ? txt.Tag.ToString() : "" };

                    deleteListFtpFile.Add(strFtp);
                }
            }

            txt.Text = "";
            txt.Tag = "";
        }
        //파일 삭제
        private bool FTP_RemoveFile(string strSaveName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            if (_ftp.delete(strSaveName) == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //폴더 삭제(내부 파일 자동 삭제)
        private bool FTP_RemoveDir(string strSaveName)
        {
            _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);
            if (_ftp.removeDir(strSaveName) == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        #endregion // FTP 파일 삭제

        // 첨부파일 다운로드 버튼 클릭 이벤트
        private void btn_DownAttach_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult msgresult = MessageBox.Show("파일을 보시겠습니까?", "보기 확인", MessageBoxButton.YesNo);
            if (msgresult == MessageBoxResult.Yes)
            {
                // 1 : 모니터링 / 2 : 첨부파일1 / 3 : 첨부파일2
                string buttonIndex = ((Button)sender).Tag.ToString();

                if ((buttonIndex.Trim().Equals("1") && txtFileName1.Text == string.Empty)
                        || (buttonIndex.Trim().Equals("2") && txtFileName2.Text == string.Empty)
                        || (buttonIndex.Trim().Equals("3") && txtFileName3.Text == string.Empty))
                {
                    MessageBox.Show("파일이 없습니다.");
                    return;
                }

                try
                {
                    var Info = dgdAll.SelectedItem as Win_info_Infoset_U_CodeView_All;

                    if (Info != null)
                    {
                        // 접속 경로
                        _ftp = new FTP_EX(FTP_ADDRESS, FTP_ID, FTP_PASS);

                        string str_path = string.Empty;
                        str_path = FTP_ADDRESS + '/' + Info.InfoID;
                        _ftp = new FTP_EX(str_path, FTP_ID, FTP_PASS);

                        string str_remotepath = string.Empty;
                        string str_localpath = string.Empty;

                        if (buttonIndex.Trim().Equals("1")) { str_remotepath = txtFileName1.Text; }
                        else if (buttonIndex.Trim().Equals("2")) { str_remotepath = txtFileName2.Text; }
                        else if (buttonIndex.Trim().Equals("3")) { str_remotepath = txtFileName3.Text; }

                        if (buttonIndex.Trim().Equals("1")) { str_localpath = LOCAL_DOWN_PATH + "\\" + txtFileName1.Text; }
                        else if (buttonIndex.Trim().Equals("2")) { str_localpath = LOCAL_DOWN_PATH + "\\" + txtFileName2.Text; }
                        else if (buttonIndex.Trim().Equals("3")) { str_localpath = LOCAL_DOWN_PATH + "\\" + txtFileName2.Text; }

                        DirectoryInfo DI = new DirectoryInfo(LOCAL_DOWN_PATH);      // Temp 폴더가 없는 컴터라면, 만들어 줘야지.
                        if (DI.Exists == false)
                        {
                            DI.Create();
                        }

                        FileInfo file = new FileInfo(str_localpath);
                        if (file.Exists)
                        {
                            file.Delete();
                        }

                        _ftp.download(str_remotepath, str_localpath);

                        ProcessStartInfo proc = new ProcessStartInfo(str_localpath);
                        proc.UseShellExecute = true;
                        Process.Start(proc);
                    }


                }
                catch (Exception ex) // 뭐든 간에 파일 없다고 하자
                {
                    MessageBox.Show("파일이 존재하지 않습니다.\r관리자에게 문의해주세요.");
                    return;
                }
            }

        }


        // 1) 첨부문서가 있을경우, 2) FTP에 정상적으로 업로드가 완료된 경우.  >> DB에 정보 업데이트 
        private void AttachFileUpdate(string InfoID)
        {
            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("sPartFile", txtFileName1.Text);
                sqlParameter.Add("sPartPath", !txtFileName1.Text.Trim().Equals("") ? "/Info/" + InfoID : "");
                sqlParameter.Add("sAttachFile1", txtFileName2.Text);
                sqlParameter.Add("sAttachPath1", !txtFileName2.Text.Trim().Equals("") ? "/Info/" + InfoID : "");
                sqlParameter.Add("sAttachFile2", txtFileName3.Text);
                sqlParameter.Add("sAttachPath2", !txtFileName3.Text.Trim().Equals("") ? "/Info/" + InfoID : "");
                sqlParameter.Add("sInfoID", InfoID);

                sqlParameter.Add("UserID", MainWindow.CurrentUser);

                string[] result = DataStore.Instance.ExecuteProcedure("xp_Info_uInfo_Attach", sqlParameter, false);
                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("이상발생, 관리자에게 문의하세요");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString());
            }
            finally
            {
                DataStore.Instance.CloseConnection(); //2021-09-13 
            }
        }

        #endregion

        #endregion

        #region 개별공지사항

        // 개별 공지사항 그리드 선택 이벤트
        private void dgdPerson_SelectChanged(object sender, SelectionChangedEventArgs e)
        {
            var NoticePerson = dgdPerson.SelectedItem as Win_info_Infoset_U_CodeView_Person;

            if (NoticePerson != null)
            {
                txtContent_Person.Text = NoticePerson.per_Info;

                FromDate_Person.SelectedDate = StringToDateTime(NoticePerson.per_FromDate);
                ToDate_Person.SelectedDate = StringToDateTime(NoticePerson.per_ToDate);

                // 해당하는 사원 리스트를 조회 해야 됨여~!
                try
                {
                    dgdTargetPerson.Items.Clear();

                    // 개별 공지사항 내용
                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();
                    sqlParameter.Add("sInfoID", NoticePerson.per_InfoID);

                    DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Info_sInfoUserID", sqlParameter, false);

                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];
                        int i = 0;

                        if (dt.Rows.Count == 0)
                        {
                            // MessageBox.Show("조회된 데이터가 없습니다.");
                            return;
                        }
                        else
                        {
                            DataRowCollection drc = dt.Rows;

                            foreach (DataRow dr in drc)
                            {
                                var TargetPerson = new PersonViewModel()
                                {
                                    PersonID = dr["PersonID"].ToString(),
                                    Name = dr["Name"].ToString()
                                };

                                dgdTargetPerson.Items.Add(TargetPerson);
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
        }
        // 개별 공지사항 추가 버튼 클릭 이벤트
        private void btnAddPerson_Click(object sender, RoutedEventArgs e)
        {


            // 공지사항 지정에 오늘날짜가 선택되도록
            FromDate_Person.SelectedDate = DateTime.Today.Date;
            ToDate_Person.SelectedDate = DateTime.Today.Date;

            strFlag = "IP";
            SaveUpdateMode();

            txtContent_Person.Text = "";

            // 개별 공지대상 데이터그리드 초기화
            dgdTargetPerson.Items.Clear();

            rowNumAll = dgdPerson.SelectedIndex;

        }
        // 개별 공지사항 수정 버튼 클릭 이벤트
        private void btnUpdatePerson_Click(object sender, RoutedEventArgs e)
        {
            var NoticePerson = dgdPerson.SelectedItem as Win_info_Infoset_U_CodeView_Person;

            if (NoticePerson != null)
            {
                strFlag = "UP";
                SaveUpdateMode();
            }
            else
            {
                MessageBox.Show("수정할 개별 공지사항을 선택해주세요.");
                return;
            }

            rowNumAll = dgdPerson.SelectedIndex;

        }
        // 개별 공지사항 삭제 버튼 클릭 이벤트
        private void btnDeletePerson_Click(object sender, RoutedEventArgs e)
        {
            var NoticePerson = dgdPerson.SelectedItem as Win_info_Infoset_U_CodeView_Person;

            if (NoticePerson != null)
            {
                MessageBoxResult msgresult = MessageBox.Show("해당 공지사항을 삭제하시겠습니까?", "확인", MessageBoxButton.OKCancel);
                if (msgresult == MessageBoxResult.OK)
                {
                    strFlag = "DP";

                    if (DeleteData(strFlag, NoticePerson.per_InfoID))
                    {
                        rowNum = 0;
                        re_Search(rowNumAll, rowNum);
                    }
                }
            }
            else
            {
                MessageBox.Show("삭제할 공지사항을 선택해주세요.");
                return;
            }
        }
        // 개별 공지사항 취소 버튼 클릭 이벤트
        private void btnCancelPerson_Click(object sender, RoutedEventArgs e)
        {
            CompleteCancelMode();

            rowNum = 0;
            re_Search(rowNumAll, rowNumPerson);
        }



        #endregion

        // 재검색
        private void re_Search(int selectedIndexAll, int selectedIndexPerson)
        {
            FillGridAll();
            FillGridPerson();

            if (dgdAll.Items.Count > 0)
            {
                dgdAll.SelectedIndex = selectedIndexAll;
            }
            if (dgdPerson.Items.Count > 0)
            {
                dgdPerson.SelectedIndex = selectedIndexPerson;
            }

            // 검색건수 작성
            int noticeAll_count = dgdAll.Items.Count;
            int noticePerson_count = dgdPerson.Items.Count;
            tblNoticeCount.Text = "▶ 검색 결과 : 전체 공지 : " + noticeAll_count
                + "건 / 개별 공지 : " + noticePerson_count + "건";
        }

        // 전체 공지사항 검색
        private void FillGridAll()
        {
            if (dgdAll.Items.Count > 0)
            {
                dgdAll.Items.Clear();
            }

            try
            {

                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sCompanyID", "");
                sqlParameter.Add("SDate", chkSearchDay.IsChecked == true ? FromDateSearch.SelectedDate.Value.ToString("yyyyMMdd") : "20000101");
                sqlParameter.Add("EDate", chkSearchDay.IsChecked == true ? ToDateSearch.SelectedDate.Value.ToString("yyyyMMdd") : "29000101");

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_Info_sInfoByDate", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        // MessageBox.Show("조회된 데이터가 없습니다.");
                        return;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var NoticeAll = new Win_info_Infoset_U_CodeView_All()
                            {
                                FromDate = DateTime.ParseExact(dr["FromDate"].ToString(), "yyyyMMdd", null).ToString("yyyy-MM-dd"),
                                ToDate = DateTime.ParseExact(dr["ToDate"].ToString(), "yyyyMMdd", null).ToString("yyyy-MM-dd"),
                                Info = dr["Info"] as string,
                                InfoID = dr["InfoID"].ToString(),
                                PartFile = dr["PartFile"].ToString(),
                                PartPath = dr["PartPath"].ToString(),
                                AttachFile1 = dr["AttachFile1"].ToString(),
                                AttachPath1 = dr["AttachPath1"].ToString(),
                                AttachFile2 = dr["AttachFile2"].ToString(),
                                AttachPath2 = dr["AttachPath2"].ToString(),
                                AttachFile3 = dr["AttachFile3"].ToString(),
                                AttachPath3 = dr["AttachPath3"].ToString()
                            };

                            dgdAll.Items.Add(NoticeAll);
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
        }

        // 개별 공지사항 검색
        private void FillGridPerson()
        {
            if (dgdPerson.Items.Count > 0)
            {
                dgdPerson.Items.Clear();
            }

            try
            {

                // 개별 공지사항 내용
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sCompanyID", "");
                sqlParameter.Add("FDate", chkSearchDay.IsChecked == true ? FromDateSearch.SelectedDate.Value.ToString("yyyyMMdd") : "20000101");
                sqlParameter.Add("TDate", chkSearchDay.IsChecked == true ? ToDateSearch.SelectedDate.Value.ToString("yyyyMMdd") : "29000101");
                sqlParameter.Add("PersonID", MainWindow.CurrentPersonID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Info_sInfoUserByPersonID", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];
                    int i = 0;

                    if (dt.Rows.Count == 0)
                    {
                        // MessageBox.Show("조회된 데이터가 없습니다.");
                        return;
                    }
                    else
                    {
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            var NoticePerson = new Win_info_Infoset_U_CodeView_Person()
                            {
                                per_FromDate = dr["FromDate"].ToString(),
                                per_ToDate = dr["ToDate"].ToString(),
                                per_Info = dr["Info"].ToString(),
                                per_InfoID = dr["InfoID"].ToString()
                            };

                            dgdPerson.Items.Add(NoticePerson);
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
        }
        // 삭제 메서드
        private bool DeleteData(string strFlag, string sInfoID)
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();
                sqlParameter.Add("sInfoID", sInfoID);

                if (strFlag.Equals("DA")) // 전체 공지사항 삭제
                {
                    // 전체 공지사항 삭제
                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_Info_dInfo";
                    pro1.OutputUseYN = "N";
                    pro1.OutputName = "sArticleID";
                    pro1.OutputLength = "10";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);

                }
                else if (strFlag.Equals("DP")) // 개별 공지사항 삭제
                {
                    // 개별 공지사항 삭제
                    Procedure pro1 = new Procedure();
                    pro1.Name = "xp_Info_dInfoUser";
                    pro1.OutputUseYN = "N";
                    pro1.OutputName = "sArticleID";
                    pro1.OutputLength = "10";

                    Prolist.Add(pro1);
                    ListParameter.Add(sqlParameter);

                    // 개별 공지 대상 삭제
                    Procedure pro2 = new Procedure();
                    pro2.Name = "xp_Info_dInfoUserList";
                    pro2.OutputUseYN = "N";
                    pro2.OutputName = "sArticleID";
                    pro2.OutputLength = "10";

                    Prolist.Add(pro2);
                    ListParameter.Add(sqlParameter);
                }

                string[] Confirm = new string[2];
                Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"D");
                if (Confirm[0] != "success")
                {
                    MessageBox.Show("[삭제 실패]\r\n" + Confirm[1].ToString());
                    flag = false;
                    //return false;
                }
                else
                {
                    flag = true;
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

        #region 개별 공지사항 - 공지대상 사원코드 관련 메서드

        // 개별 사용자 DB정보 가져와서 리스트에 집어넣기.
        private void FillPersonInfo()
        {
            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Add("nChkDepartID", 0);
            sqlParameter.Add("sDepartID", "");

            DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Person_sPerson_Treeview_HYG", sqlParameter, false);
            DataTable dt_RootList = null;
            DataTable dt_NodeList = null;

            if (ds != null && ds.Tables.Count > 0)
            {
                dt_RootList = ds.Tables[1];
                dt_NodeList = ds.Tables[0];

                if (dt_RootList.Rows.Count == 0)
                {
                    MessageBox.Show("조회된 데이터가 없습니다.");
                    return;
                }

                DataRowCollection drcRoot = dt_RootList.Rows;

                foreach (DataRow item in drcRoot)
                {
                    var mMenuviewModel_ROOT = new PersonViewModel()
                    {
                        Depart = item["Depart"] as string
                    };
                    mMenulist.Add(mMenuviewModel_ROOT);
                }

                DataRowCollection drcNode = dt_NodeList.Rows;

                foreach (DataRow item in drcNode)
                {
                    var mMenuviewModel_Node = new PersonViewModel()
                    {
                        PersonID = item["PersonID"] as string,
                        Name = item["Name"] as string,
                        UserID = item["UserID"] as string,
                        DepartID = item["DepartID"] as string,
                        Depart = item["Depart"] as string,
                    };
                    mMenuSublist.Add(mMenuviewModel_Node);
                }
            }
        }

        // 리스트 정보를 바탕으로 트리 만들어서 출력
        private void MakeTree()                 //매뉴 생성 + 클릭시 기억해야 할 정보 각기 기억.  root와 nodesms 
        {
            TreeViewItem mTreeViewItem_Root = null;
            TreeViewItem mTreeViewItem_Node = null;

            foreach (PersonViewModel mvm1 in mMenulist)
            {
                mTreeViewItem_Root = new TreeViewItem() { Header = mvm1.Depart };
                treeview_test.Items.Add(mTreeViewItem_Root);
                //mTreeViewItem_Root.MouseLeftButtonUp += (s, e) => { fmenu_click(s, null); };

                for (int i = 0; i < mMenuSublist.Count; i++)
                {
                    if (mvm1.Depart == mMenuSublist[i].Depart.ToString())
                    {
                        mTreeViewItem_Node = new TreeViewItem() { Header = mMenuSublist[i].Name.ToString(), Tag = mMenuSublist[i].PersonID.ToString() };
                        mTreeViewItem_Root.Items.Add(mTreeViewItem_Node);
                        mTreeViewItem_Node.MouseLeftButtonUp += (s, e) => { All_Area_Node_Person_Click(s, null); };
                    }
                }
            }
        }

        // 개별 공지 대상 리스트 추가하기!!!!!!!!!~!!
        // 공지대상 사원 추가 버튼 클릭 이벤트
        private void btn_MoveInsert_Click(object sender, RoutedEventArgs e)
        {
            //1. 기존 그리드에 넘기려고 하는 애랑 똑같은 애가 있으면 못넘기게 해야 할 거 아냐.
            var data = new PersonViewModel
            {
                PersonID = Left_Click_Person_ID,
                Name = Left_Click_Person_name
            };

            // 부서(관리부, 품질부) 를 선택후에 오른쪽 버튼을 눌렀을때, 넘어가는 데이터가 없도록
            TreeViewItem tviDepart = (TreeViewItem)treeview_test.SelectedItem;

            // 빈칸등록을 막기 위해서 추가 
            // && 제일 하위 노드는 items.Count : 0 → 선택한 것이 제일 하위 노드일때만 데이터가 넘어가도록 추가
            if (data.PersonID.Trim().Equals("") == false && tviDepart.Items.Count == 0)
            {
                DataGridRow CheckRow = new DataGridRow();
                TextBlock TX = new TextBlock();

                for (int i = 0; i < dgdTargetPerson.Items.Count; i++)
                {
                    CheckRow = (DataGridRow)dgdTargetPerson.ItemContainerGenerator.ContainerFromIndex(i);
                    TX = dgdTargetPerson.Columns[1].GetCellContent(CheckRow) as TextBlock;

                    if (data.PersonID == TX.Text.ToString())
                    {
                        return;         // 동일 아이디 차단.
                    }
                }
                dgdTargetPerson.Items.Add(data);     // 인서트.

                // 데이터를 넣고 난 뒤에는 전역변수 클리어 시킬 것.
                Left_Click_Person_name = string.Empty;
                Left_Click_Person_ID = string.Empty;
            }

        }
        // 공지대상 사원 삭제 버튼 클릭 이벤트
        private void btn_MoveDelete_Click(object sender, RoutedEventArgs e)
        {
            var Del_data = dgdTargetPerson.SelectedItem;
            if (Del_data != null)
            {
                dgdTargetPerson.Items.Remove(Del_data);
            }
        }

        // 사원 데이터그리드(오른쪽) 클릭 이벤트
        private void All_Area_Node_Person_Click(object sender, MouseButtonEventArgs e)
        {
            //전역변수에 담아 기억해야 할 녀석은 1. 사원명.// 2. personID. // 

            if (sender is TreeViewItem)     //트리뷰에서 클릭된 경우, LEFT SIDE.
            {
                Left_Click_Person_name = ((sender as TreeViewItem).Header as string).Replace(" ", "");
                Left_Click_Person_ID = ((sender as TreeViewItem).Tag as string).Replace(" ", "");
            }
            //else if (sender is DataGrid)    // 데이터그리드에서 클릭된 경우, RIGHT SIDE.
            //{
            //    object OBJ = dgdTargetPerson.SelectedItem;
            //    Right_Click_Person_Name = (dgdTargetPerson.SelectedCells[0].Column.GetCellContent(OBJ) as TextBlock).Text;
            //    Right_Click_Person_ID = (dgdTargetPerson.SelectedCells[1].Column.GetCellContent(OBJ) as TextBlock).Text;
            //}
        }

        #endregion



        // 테스트
        private void btnTest_Click(object sender, RoutedEventArgs e)
        {
            hoit(txtFileName1);
        }

        private void hoit(TextBox tb)
        {
            tb.Text = "흐아ㅓㄹ나ㅣ어리ㅏㄴ";
        }


    }

    class Win_info_Infoset_U_CodeView_All : BaseView
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }

        // 전체공지용 조회 Fill그리드 정보.
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public string Info { get; set; }
        public string InfoID { get; set; }
        public string PartFile { get; set; }
        public string PartPath { get; set; }
        public string AttachFile1 { get; set; }
        public string AttachPath1 { get; set; }
        public string AttachFile2 { get; set; }
        public string AttachPath2 { get; set; }
        public string AttachFile3 { get; set; }
        public string AttachPath3 { get; set; }

    }

    class Win_info_Infoset_U_CodeView_Person : BaseView
    {
        // 개별공지용 조회 Fill 그리드 정보
        public string per_FromDate { get; set; }
        public string per_ToDate { get; set; }
        public string per_Info { get; set; }
        public string per_InfoID { get; set; }
    }

    public class PersonViewModel
    {
        public override string ToString()
        {
            return (this.ReportAllProperties());
        }
        public string PersonID { get; set; }
        public string Name { get; set; }
        public string UserID { get; set; }
        public string DepartID { get; set; }
        public string Depart { get; set; }

    }
}
