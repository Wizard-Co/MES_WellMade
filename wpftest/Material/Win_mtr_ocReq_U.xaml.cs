using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WPF.MDI;

namespace WizMes_WellMade
{
    /// <summary>
    /// Win_mt_ocReq_U.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Win_mtr_ocReq_U : UserControl
    {
        string stDate = string.Empty;
        string stTime = string.Empty;

        Lib lib = new Lib();
        string strBasisID = string.Empty;
        string InspectName = string.Empty;
        string AASS = string.Empty;


        string strFlag = string.Empty;
        int rowNum = 0;

        // 인쇄 활용 객체
        private Microsoft.Office.Interop.Excel.Application excelapp;
        private Microsoft.Office.Interop.Excel.Workbook workbook;
        private Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private Microsoft.Office.Interop.Excel.Range workrange;
        private Microsoft.Office.Interop.Excel.Worksheet copysheet;
        private Microsoft.Office.Interop.Excel.Worksheet pastesheet;
        WizMes_WellMade.PopUp.NoticeMessage msg = new WizMes_WellMade.PopUp.NoticeMessage();

        List<Win_mtr_OCReq_U_CodeView_Sub> deleteOcReqSub = new List<Win_mtr_OCReq_U_CodeView_Sub>();


        public Win_mtr_ocReq_U()
        {
            InitializeComponent();
        }

        // 폼 로드 됬을때
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stDate = DateTime.Now.ToString("yyyyMMdd");
            stTime = DateTime.Now.ToString("HHmm");

            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "S");

            // 콤보박스 세팅
            comboBoxSetting();

            // 발주일자 체크박스 IsChecked = true
            chkReqDateSearch.IsChecked = true;
        }

        // 콤보박스 세팅
        private void comboBoxSetting()
        {
            cboApproval.Items.Clear();

            //승인여부
            cboApproval.Items.Add("Y");
            cboApproval.Items.Add("N");

            cboApproval.SelectedIndex = 1;
        }

        #region 저장, 수정시 / 저장완료, 취소시 메서드

        // 저장, 수정
        private void SaveUpdateMode()
        {
            // Header
            // 발주일자
            chkReqDateSearch.IsEnabled = false;
            dtpFromDateSearch.IsEnabled = false;
            dtpToDateSearch.IsEnabled = false;
            btnYesterday.IsEnabled = false;
            btnToday.IsEnabled = false;
            btnLastMonth.IsEnabled = false;
            btnThisMonth.IsEnabled = false;
            // 거래처
            chkCustom.IsEnabled = false;
            txtCustomSearch.IsEnabled = false;
            btnCustomSearch.IsEnabled = false;
            // 승인여부
            chkApproval.IsEnabled = false;
            cboApproval.IsEnabled = false;

            // 상단 오른쪽 버튼
            btnAdd.IsEnabled = false;
            btnUpdate.IsEnabled = false;
            btnDelete.IsEnabled = false;
            btnSearch.IsEnabled = false;
            btnPrint.IsEnabled = false;

            btnSave.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;

            // Content
            dgdMain.IsEnabled = false;

            // 오른쪽 부분
            lblMsg.Visibility = Visibility.Visible;
            if (strFlag.Equals("I"))
            { tbkMsg.Text = "자료 추가 중"; }
            else
            { tbkMsg.Text = "자료 수정 중"; }

            dtpReqDate.IsEnabled = true; // 발주일자
            dtpDueDate.IsEnabled = true; // 납기일자
            txtSupplierName.IsReadOnly = false; // 거래처
            btnSupplierName.IsEnabled = true;
            txtPayCondition.IsReadOnly = false;
            txtSupplierCharge.IsReadOnly = false; // 업체담당자
            txtReqCharge.IsReadOnly = false; // 발주처리자
            txtComments.IsReadOnly = false; // 비고

            // 서브 그리드 추가, 삭제
            btnAddSub.IsEnabled = true;
            btnDeleteSub.IsEnabled = true;

            dgdSub_CanChecked(); // 서브 그리드 - 입고마감 활성화(체크박스)

            // 승인, 승인취소 버튼은 추가, 저장인 상태에서는 비활성화 시키자...
            btnApproval_Y.IsEnabled = false;
            btnApproval_N.IsEnabled = false;
        }

        // 취소, 저장 후
        private void CompleteCancelMode()
        {
            // Header
            // 발주일자
            chkReqDateSearch.IsEnabled = true;
            if (chkReqDateSearch.IsChecked == true)
            {
                dtpFromDateSearch.IsEnabled = true;
                dtpToDateSearch.IsEnabled = true;
            }
            btnYesterday.IsEnabled = true;
            btnToday.IsEnabled = true;
            btnLastMonth.IsEnabled = true;
            btnThisMonth.IsEnabled = true;
            // 거래처
            chkCustom.IsEnabled = true;
            if (chkCustom.IsChecked == true)
            {
                txtCustomSearch.IsEnabled = true;
                btnCustomSearch.IsEnabled = true;
            }
            // 승인여부
            chkApproval.IsEnabled = true;
            if (chkApproval.IsChecked == true)
                cboApproval.IsEnabled = true;

            // 상단 오른쪽 버튼
            btnAdd.IsEnabled = true;
            btnUpdate.IsEnabled = true;
            btnDelete.IsEnabled = true;
            btnSearch.IsEnabled = true;
            btnPrint.IsEnabled = true;

            btnSave.Visibility = Visibility.Hidden;
            btnCancel.Visibility = Visibility.Hidden;

            // Content
            dgdMain.IsEnabled = true;

            // 오른쪽 부분
            lblMsg.Visibility = Visibility.Hidden;

            dtpReqDate.IsEnabled = false; // 발주일자
            dtpDueDate.IsEnabled = false; // 납기일자
            txtSupplierName.IsReadOnly = true; // 거래처
            btnSupplierName.IsEnabled = false;
            txtPayCondition.IsReadOnly = true;
            txtSupplierCharge.IsReadOnly = true; // 업체담당자
            txtReqCharge.IsReadOnly = true; // 발주처리자
            txtComments.IsReadOnly = true; // 비고

            // 서브 그리드 추가, 삭제
            btnAddSub.IsEnabled = false;
            btnDeleteSub.IsEnabled = false;

            // 승인, 승인취소 버튼은 추가, 저장인 상태에서는 비활성화 시키자...
            btnApproval_Y.IsEnabled = true;
            btnApproval_N.IsEnabled = true;
        }

        #endregion // 저장, 수정시 / 저장완료, 취소시 메서드

        #region Header 부분 - 검색조건

        // 전일
        private void btnYesterday_Click(object sender, RoutedEventArgs e)
        {
            //dtpFromDateSearch.SelectedDate = DateTime.Today.AddDays(-1);
            //dtpToDateSearch.SelectedDate = DateTime.Today.AddDays(-1);

            try
            {
                if (dtpFromDateSearch.SelectedDate != null)
                {
                    dtpFromDateSearch.SelectedDate = dtpFromDateSearch.SelectedDate.Value.AddDays(-1);
                    dtpToDateSearch.SelectedDate = dtpFromDateSearch.SelectedDate;
                }
                else
                {
                    dtpFromDateSearch.SelectedDate = DateTime.Today.AddDays(-1);
                    dtpToDateSearch.SelectedDate = DateTime.Today.AddDays(-1);
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnYesterday_Click : " + ee.ToString());
            }

        }
        // 금일
        private void btnToday_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDateSearch.SelectedDate = DateTime.Today;
            dtpToDateSearch.SelectedDate = DateTime.Today;
        }
        // 전월
        private void btnLastMonth_Click(object sender, RoutedEventArgs e)
        {
            //dtpFromDateSearch.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[0];
            //dtpToDateSearch.SelectedDate = Lib.Instance.BringLastMonthDatetimeList()[1];

            try
            {
                if (dtpFromDateSearch.SelectedDate != null)
                {
                    DateTime ThatMonth1 = dtpFromDateSearch.SelectedDate.Value.AddDays(-(dtpFromDateSearch.SelectedDate.Value.Day - 1)); // 선택한 일자 달의 1일!

                    DateTime LastMonth1 = ThatMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThatMonth1.AddDays(-1); // 저번달 말일

                    dtpFromDateSearch.SelectedDate = LastMonth1;
                    dtpToDateSearch.SelectedDate = LastMonth31;
                }
                else
                {
                    DateTime ThisMonth1 = DateTime.Today.AddDays(-(DateTime.Today.Day - 1)); // 이번달 1일

                    DateTime LastMonth1 = ThisMonth1.AddMonths(-1); // 저번달 1일
                    DateTime LastMonth31 = ThisMonth1.AddDays(-1); // 저번달 말일

                    dtpFromDateSearch.SelectedDate = LastMonth1;
                    dtpToDateSearch.SelectedDate = LastMonth31;
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show("오류지점 - btnLastMonth_Click : " + ee.ToString());
            }

        }
        // 금월
        private void btnThisMonth_Click(object sender, RoutedEventArgs e)
        {
            dtpFromDateSearch.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[0];
            dtpToDateSearch.SelectedDate = Lib.Instance.BringThisMonthDatetimeList()[1];
        }

        // 검색 발주일자 라벨 왼쪽 클릭 이벤트
        private void lblDate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkReqDateSearch.IsChecked == true)
            {
                chkReqDateSearch.IsChecked = false;
            }
            else
            {
                chkReqDateSearch.IsChecked = true;
            }
        }
        // 발주일자 체크박스 이벤트
        private void chkReqDate_Checked(object sender, RoutedEventArgs e)
        {
            chkReqDateSearch.IsChecked = true;
            dtpFromDateSearch.IsEnabled = true;
            dtpToDateSearch.IsEnabled = true;

            // 오늘날짜 세팅하기
            dtpFromDateSearch.SelectedDate = DateTime.Today;
            dtpToDateSearch.SelectedDate = DateTime.Today;
        }
        private void chkReqDate_UnChecked(object sender, RoutedEventArgs e)
        {
            chkReqDateSearch.IsChecked = false;
            dtpFromDateSearch.IsEnabled = false;
            dtpToDateSearch.IsEnabled = false;
        }


        // 검색 거래처 라벨 왼쪽 클릭 이벤트
        private void lblCustom_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkCustom.IsChecked == true)
            {
                chkCustom.IsChecked = false;
            }
            else
            {
                chkCustom.IsChecked = true;
            }
        }
        // 거래처 체크박스 이벤트
        private void chkCustom_Checked(object sender, RoutedEventArgs e)
        {
            chkCustom.IsChecked = true;
            txtCustomSearch.IsEnabled = true;
            btnCustomSearch.IsEnabled = true;
        }
        private void chkCustom_UnChecked(object sender, RoutedEventArgs e)
        {
            chkCustom.IsChecked = false;
            txtCustomSearch.IsEnabled = false;
            btnCustomSearch.IsEnabled = false;
        }
        // 거래처 텍스트박스 엔터 → 플러스파인더
        private void txtCustomSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtCustomSearch, 0, "");
            }
        }
        // 거래처 플러스파인더 이벤트
        private void btnCustomSearch_Click(object sender, RoutedEventArgs e)
        {
            // 거래처 : 0
            MainWindow.pf.ReturnCode(txtCustomSearch, 0, "");
        }

        // 승인여부 라벨 왼쪽 클릭 이벤트
        private void lblApproval_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkApproval.IsChecked == true)
            {
                chkApproval.IsChecked = false;
            }
            else
            {
                chkApproval.IsChecked = true;
            }
        }
        // 승인여부 체크박스 이벤트
        private void chkApproval_Checked(object sender, RoutedEventArgs e)
        {
            chkApproval.IsChecked = true;
            cboApproval.IsEnabled = true;
        }
        private void chkApproval_UnChecked(object sender, RoutedEventArgs e)
        {
            chkApproval.IsChecked = false;
            cboApproval.IsEnabled = false;
        }

        #endregion // Header 부분 - 검색조건

        #region 상단 오른쪽 버튼 이벤트 (추가, 삭제 등등)

        // 추가버튼 클릭 이벤트
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            if (dgdSub.Items.Count > 0)
                dgdSub.Items.Clear();

            this.DataContext = null;
            strFlag = "I";
            SaveUpdateMode();

            // 발주일자, 납기일자 오늘로 세팅하기
            dtpReqDate.SelectedDate = DateTime.Today;
            dtpDueDate.SelectedDate = DateTime.Today;

            rowNum = dgdMain.SelectedIndex;
        }
        // 수정버튼 클릭 이벤트
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var OcReq = dgdMain.SelectedItem as Win_mtr_OCReq_U_CodeView;

            if (OcReq != null)
            {
                strFlag = "U";
                SaveUpdateMode();

                rowNum = dgdMain.SelectedIndex;
            }

        }
        // 삭제버튼 클릭 이벤트
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            var OcReq = dgdMain.SelectedItem as Win_mtr_OCReq_U_CodeView;

            if (OcReq != null)
            {
                if (MessageBox.Show("선택한 항목을 삭제하시겠습니까?", "삭제 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {

                    if (DeleteData(OcReq.REQ_ID))
                    {
                        rowNum = 0;
                        re_Search(rowNum);
                    }
                }
            }
            else
            {
                MessageBox.Show("삭제할 데이터를 선택해주세요.");
            }
        }
        // 닫기버튼 클릭 이벤트
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            DataStore.Instance.InsertLogByFormS(this.GetType().Name, stDate, stTime, "E");
            Lib.Instance.ChildMenuClose(this.ToString());
        }
        // 검색버튼 클릭 이벤트
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            rowNum = 0;
            re_Search(rowNum);
        }


        //// 저장버튼 클릭 이벤트
        //private void btnSave_Click(object sender, RoutedEventArgs e)
        //{
        //    if (SaveData(strFlag))
        //    {
        //        CompleteCancelMode();

        //        if (!strFlag.Trim().Equals("U"))
        //        {
        //            rowNum = 0;
        //        }

        //        strFlag = string.Empty;

        //        re_Search(rowNum);
        //    }
        //}

        //저장
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SaveData(strFlag))
            {
                CompleteCancelMode();
                strBasisID = string.Empty;
                lblMsg.Visibility = Visibility.Hidden;

                if (strFlag.Equals("I"))
                {
                    InspectName = txtReqID.ToString();
                    //InspectName = txtKCustom.ToString();
                    //InspectDate = dtpInspectDate.SelectedDate.ToString().Substring(0, 10);

                    //rowNum = 0;
                    //re_Search(rowNum);
                    int i = 0;

                    foreach (Win_mtr_OCReq_U_CodeView WMRIC in dgdMain.Items)
                    {

                        string a = WMRIC.REQ_ID.ToString();
                        string b = AASS;


                        if (a == b)
                        {
                            System.Diagnostics.Debug.WriteLine("데이터 같음");

                            break;
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine("다름");
                        }

                        i++;
                    }

                    rowNum = i;
                    re_Search(rowNum);
                }
                else
                {
                    rowNum = dgdMain.SelectedIndex;
                    re_Search(rowNum);
                }
            }


        }




        // 취소버튼 클릭 이벤트
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            strFlag = string.Empty;
            CompleteCancelMode();

            //rowNum = 0;
            re_Search(rowNum);
        }
        // 발주서 인쇄버튼 클릭 이벤트
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            // 발주 승인처리가 되어있지 않으면, 발주서 인쇄는 불가능
            var DyeReq = dgdMain.SelectedItem as Win_mtr_OCReq_U_CodeView;

            if (DyeReq != null)
            {
                if (!DyeReq.ApprovalYN.Trim().Equals("Y"))
                {
                    MessageBox.Show("승인되지 않은 발주건은 발주서 인쇄가 불가능 합니다.\r[승인] 버튼을 통해 승인처리를 먼저 해주세요.");
                    return;
                }

                // 인쇄 메서드
                ContextMenu menu = btnPrint.ContextMenu;
                menu.StaysOpen = true;
                menu.IsOpen = true;
            }
            else
            {

            }
        }

        #region 인쇄 메서드

        // 미리보기
        private void menuSeeAhead_Click(object sender, RoutedEventArgs e)
        {

            if (dgdSub.Items.Count < 1)
            {
                MessageBox.Show("해당 자료가 존재하지 않습니다.");
                return;
            }

            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            Lib.Instance.Delay(1000);

            PrintWork(true);

            msg.Visibility = Visibility.Hidden;
        }
        // 바로 인쇄
        private void menuRighPrint_Click(object sender, RoutedEventArgs e)
        {
            if (dgdSub.Items.Count < 1)
            {
                MessageBox.Show("해당 자료가 존재하지 않습니다.");
                return;
            }
            DataStore.Instance.InsertLogByForm(this.GetType().Name, "P");
            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            Lib.Instance.Delay(1000);

            PrintWork(false);

            msg.Visibility = Visibility.Hidden;
        }
        // 닫기
        private void menuClose_Click(object sender, RoutedEventArgs e)
        {
            ContextMenu menu = btnPrint.ContextMenu;
            menu.StaysOpen = false;
            menu.IsOpen = false;
        }

        // 프린터 엑셀 작업
        // 실제 엑셀작업
        private void PrintWork(bool previewYN)
        {
            // 메인 그리드 필요한 정보 변수
            // 수신, 참조??, 수신 Tel / Fax, 발주일자, 담당자
            string supplierName = ""; // 수신
            string comment = ""; // 비고
            string supplierTel = ""; // 수신 Tel / Fax
            string reqDate = ""; // 발주일자
            string reqCharge = ""; // 담당자

            var OcReq = dgdMain.SelectedItem as Win_mtr_OCReq_U_CodeView;

            if (OcReq != null)
            {
                supplierName = OcReq.SUPPLIER_NAME;
                comment = OcReq.COMMENTS; // 참조
                supplierTel = OcReq.SUPPLIER_PhoneNo + " / " + OcReq.SUPPLIER_FaxNo; // 수신 Tel / Fax
                reqDate = OcReq.REQ_DATE; // 발주일자
                reqCharge = OcReq.REQ_CHARGE; // 담당자
            }

            excelapp = new Microsoft.Office.Interop.Excel.Application();

            string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\Report\\자재발주서.xls";
            workbook = excelapp.Workbooks.Add(MyBookPath);
            worksheet = workbook.Sheets["Form"];
            pastesheet = workbook.Sheets["Print"];

            //// 워크시트 이름 변경 : 에너지-2018,2019
            ////int year = ConvertInt(txtYear.Text);
            //worksheet.Name = "에너지-" + (year - 1) + "," + year;

            // 페이지 계산 등
            int rowCount = dgdSub.Items.Count;
            int excelStartRow = 12;

            int copyLine = 0;
            int Page = 0;
            int PageAll = (int)Math.Ceiling(rowCount / 21.0);
            int DataCount = 0;


            // 총 금액 계산하기
            double SumAmount = 0;

            // 수신
            workrange = worksheet.get_Range("E4");
            workrange.Value2 = supplierName;

            // 참조???

            // 수신 Tel / Fax
            if (supplierTel.Trim().Equals("/"))
                supplierTel = "";
            workrange = worksheet.get_Range("E6");
            workrange.Value2 = supplierTel;

            // 발주일자
            workrange = worksheet.get_Range("E7");
            workrange.Value2 = reqDate;
            //workrange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            //workrange.Font.Size = 11;

            // 담당자 - 발주담당자 - Req_Charge
            workrange = worksheet.get_Range("W7");
            workrange.Value2 = reqCharge;

            // 비고 C34
            workrange = worksheet.get_Range("C34");
            workrange.Value2 = comment;

            for (int k = 0; k < PageAll; k++)
            {
                Page++;
                copyLine = ((Page - 1) * 40);

                int excelNum = 0;

                // 기존에 있는 데이터 지우기 "A7", "W41"
                worksheet.Range["A12", "AJ32"].EntireRow.ClearContents();

                for (int i = DataCount; i < rowCount; i++)
                {
                    if (i == 21 * Page)
                    {
                        break;
                    }

                    var OcReqSub = dgdSub.Items[i] as Win_mtr_OCReq_U_CodeView_Sub;
                    int excelRow = excelStartRow + excelNum;

                    if (OcReqSub != null)
                    {
                        // 순번
                        workrange = worksheet.get_Range("A" + excelRow);
                        workrange.Value2 = OcReqSub.ReqItemNo;

                        // 상품명
                        workrange = worksheet.get_Range("C" + excelRow);
                        workrange.Value2 = OcReqSub.Item_Name;

                        // 용도
                        workrange = worksheet.get_Range("K" + excelRow);
                        workrange.Value2 = OcReqSub.Item_For_Useing;

                        // 수량 - 2020.01.07 수량을 kg 로 변경 → 수량*중량/1000 (kg 이므로 나눠야됨) 
                        // 2020.05.07 - 중량 DB 값 0 이여 출력시 0 으로 표기
                        workrange = worksheet.get_Range("S" + excelRow);
                        workrange.Value2 = stringFormatN2(ConvertDouble(OcReqSub.Qty) /* * ConvertDouble(OcReqSub.Weight) */ );

                        // 납품일자
                        workrange = worksheet.get_Range("W" + excelRow);
                        workrange.Value2 = DatePickerFormat(OcReqSub.Ddate);

                        // 비고
                        workrange = worksheet.get_Range("AC" + excelRow);
                        workrange.Value2 = OcReqSub.COMMENTSITEM;

                        SumAmount += ConvertDouble(OcReqSub.Amount);

                        excelNum++;
                        DataCount = i;
                    }
                }

                // 2장 이상 넘어가면 페이지 넘버 입력
                if (PageAll > 1)
                {
                    pastesheet.PageSetup.CenterFooter = "&P / &N";
                }

                // Form 시트 내용 Print 시트에 복사 붙여넣기
                worksheet.Select();
                worksheet.UsedRange.EntireRow.Copy();
                pastesheet.Select();
                workrange = pastesheet.Cells[copyLine + 1, 1];
                workrange.Select();
                pastesheet.Paste();

                DataCount++;
            }

            // 총금액 입력하기 : 10, 50, 90
            for (int i = 0; i < PageAll; i++)
            {
                int sumAmount_Index = 10 + (40 * i);

                workrange = pastesheet.get_Range("E" + sumAmount_Index);
                workrange.Value2 = SumAmount;
            }

            pastesheet.UsedRange.EntireRow.Select();

            //
            excelapp.Visible = true;
            msg.Hide();

            if (previewYN == true)
            {
                pastesheet.PrintPreview();
            }
            else
            {
                pastesheet.PrintOutEx();
            }
        }

        #endregion // 인쇄 메서드


        #endregion // 오른쪽 상단 버튼 이벤트 (추가, 삭제 등등)

        #region Content 부분

        private void dgdMain_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var OcReq = dgdMain.SelectedItem as Win_mtr_OCReq_U_CodeView;

            if (OcReq != null)
            {
                this.DataContext = OcReq;
                FillGridSub(OcReq.REQ_ID);

                if (dgdSub.Items.Count > 0)
                {
                    dgdSub.SelectedIndex = 0;
                }

                // 승인 처리가 되어 있다면, 승인취소처리만 활성화 / 그 이외에는 승인 버튼만 활성화
                if (OcReq.ApprovalYN.Trim().Equals("Y"))
                {
                    btnApproval_Y.IsEnabled = false;
                    btnApproval_N.IsEnabled = true;
                }
                else
                {
                    btnApproval_Y.IsEnabled = true;
                    btnApproval_N.IsEnabled = false;
                }
            }
        }

        // 발주정보 - 거래처 플러스파인더 클릭 이벤트
        private void txtSupplierName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                MainWindow.pf.ReturnCode(txtSupplierName, 0, "");
            }
        }
        private void btnSupplierName_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtSupplierName, 0, "");
        }


        #region 서브 그리드 모음 (키 이벤트 제외)

        #region 재고조회 이동 버튼, 발주 승인, 발주 승인 취소 이벤트
        // 재고 조회 페이지로 이동하는 버튼 이벤트
        private void btnGoStock_Click(object sender, RoutedEventArgs e)
        {
            // 재고현황(제품포함)
            int i = 0;
            foreach (MenuViewModel mvm in MainWindow.mMenulist)
            {
                if (mvm.Menu.Equals("재고현황 조회(제품포함)"))
                {
                    break;
                }
                i++;
            }
            try
            {
                if (MainWindow.MainMdiContainer.Children.Contains(MainWindow.mMenulist[i].subProgramID as MdiChild))
                {
                    (MainWindow.mMenulist[i].subProgramID as MdiChild).Focus();
                }
                else
                {
                    Type type = Type.GetType("WizMes_WellMade." + MainWindow.mMenulist[i].ProgramID.Trim(), true);
                    object uie = Activator.CreateInstance(type);

                    MainWindow.mMenulist[i].subProgramID = new MdiChild()
                    {
                        Title = "WizMes_WellMade [" + MainWindow.mMenulist[i].MenuID.Trim() + "] " + MainWindow.mMenulist[i].Menu.Trim() +
                                " (→" + MainWindow.mMenulist[i].ProgramID + ")",
                        Height = SystemParameters.PrimaryScreenHeight * 0.8,
                        MaxHeight = SystemParameters.PrimaryScreenHeight * 0.85,
                        Width = SystemParameters.WorkArea.Width * 0.85,
                        MaxWidth = SystemParameters.WorkArea.Width,
                        Content = uie as UIElement,
                        Tag = MainWindow.mMenulist[i]
                    };
                    Lib.Instance.AllMenuLogInsert(MainWindow.mMenulist[i].MenuID, MainWindow.mMenulist[i].Menu, MainWindow.mMenulist[i].subProgramID);
                    MainWindow.MainMdiContainer.Children.Add(MainWindow.mMenulist[i].subProgramID as MdiChild);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("해당 화면이 존재하지 않습니다.");
            }
        }
        // 해당 발주건 승인처리 버튼 이벤트 : [DYE_ORDER_REQ] - ApprovalYN = Y
        // 프로시저 : [xp_DyeAuxReq_uDyeAuxReqApprovalYN]
        private void btnApproval_Y_Click(object sender, RoutedEventArgs e)
        {
            var DyeReq = dgdMain.SelectedItem as Win_mtr_OCReq_U_CodeView;

            if (DyeReq != null)
            {
                if (MessageBox.Show("해당 발주건을 승인처리 하시겠습니까?", "승인 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (updateApprovalYN(DyeReq.REQ_ID, "Y"))
                    {
                        MessageBox.Show("해당 발주의 승인 처리가 완료되었습니다.");

                        rowNum = 0;
                        re_Search(rowNum);
                    }
                }
            }

        }
        // 해당 발주건 승인취소 버튼 이벤트 : [DYE_ORDER_REQ] - ApprovalYN = Y
        private void btnApproval_N_Click(object sender, RoutedEventArgs e)
        {
            var DyeReq = dgdMain.SelectedItem as Win_mtr_OCReq_U_CodeView;

            if (DyeReq != null)
            {
                if (MessageBox.Show("해당 발주건을 승인 취소 처리 하시겠습니까?", "승인 취소 전 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    if (updateApprovalYN(DyeReq.REQ_ID, "N"))
                    {
                        MessageBox.Show("해당 발주의 승인 취소 처리가 완료되었습니다.");

                        rowNum = 0;
                        re_Search(rowNum);
                    }
                }
            }
        }

        // 승인처리, 승인취소 하는 메서드
        private bool updateApprovalYN(string Req_ID, string ApprovalYN)
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Clear();

                sqlParameter.Add("REQ_ID", Req_ID);
                sqlParameter.Add("ApprovalYN", ApprovalYN);

                Procedure pro1 = new Procedure();
                pro1.Name = "xp_DyeAuxReq_uDyeAuxReqApprovalYN";
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
                    //return false;
                }
                else
                {
                    flag = true;
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

        #endregion // 재고조회 이동 버튼, 발주 승인, 발주 승인 취소 이벤트

        // 문제 : 체크 된 상태에서 체크박스 바로 옆의 빈 공간을 클릭했을때 체크해제가 됨 - 이 메서드를 탐
        // 해결 : 아마도 데이터그리드 키 이벤트 getFocus 때문인듯? → 입고마감 Column : 7 Index 일 경우에는 키 이벤트가 발동하지 않도록 설정
        private void DataGridCell_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            //DataGridCell dgc = sender as DataGridCell;
            //if (dgc.Column.DisplayIndex == 7)
            //{
            //    var OcReqSub = dgdSub.SelectedItem as Win_mtr_OCReq_U_CodeView_Sub;

            //    if (OcReqSub != null)
            //    {
            //        if (OcReqSub.InwareCloseYN.Equals("Y"))
            //        {
            //            OcReqSub.InwareCloseYN = "N";
            //            OcReqSub.InwareCloseChecked = false;
            //        }
            //        else
            //        {
            //            OcReqSub.InwareCloseYN = "Y";
            //            OcReqSub.InwareCloseChecked = true;
            //        }
            //    }

            //}

        }

        // 서브 그리드 - 발주상세항목 - 추가 이벤트
        private void btnAddSub_Click(object sender, RoutedEventArgs e)
        {
            int i = 1;

            if (dgdSub.Items.Count > 0)
                i = dgdSub.Items.Count + 1;

            var OcReqSub = new Win_mtr_OCReq_U_CodeView_Sub()
            {
                ReqItemNo = i.ToString(),
                Seq = "",
                Item_ID = "",
                Item_Name = "",
                Unit_Price = "0",
                Qty = "0",
                Amount = "0",
                Vat = "0",
                Vat_YN = "",
                COMMENTSITEM = "",
                Item_For_Useing = "",
                Ddate = dtpDueDate.SelectedDate != null ? dtpDueDate.SelectedDate.Value.ToString("yyyyMMdd") : "",
                Ddate_CV = dtpDueDate.SelectedDate != null ? dtpDueDate.SelectedDate.Value.ToString("yyyy-MM-dd") : "",
                InwareCloseYN = "N",
                IsEnabled = true
            };

            dgdSub.Items.Add(OcReqSub);
        }

        // 서브 그리드 - 발주상세항목 - 제거 이벤트
        private void btnDeleteSub_Click(object sender, RoutedEventArgs e)
        {
            var OcReqSub = dgdSub.SelectedItem as Win_mtr_OCReq_U_CodeView_Sub;

            if (OcReqSub != null)
            {
                dgdSub.Items.Remove(OcReqSub);
            }
            else
            {
                if (dgdSub.Items.Count > 0)
                {
                    dgdSub.Items.Remove(dgdSub.Items[dgdSub.Items.Count - 1]);
                }
            }
        }

        private void InwareCloseYN_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {

        }

        // 문제 : 서브 그리드 체크 상태에서 다른곳을 클릭 → 체크박스 클릭시, 언체크 메서드를 타지 않는다~!!!!!!!!~!!!!! - 요 메서드를 탐
        // 해결 : 아마도 데이터그리드 키 이벤트 getFocus 때문인듯? → 입고마감 Column : 7 Index 일 경우에는 키 이벤트가 발동하지 않도록 설정
        private void InwareCloseYN_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            //CheckBox chkSender = sender as CheckBox;
            //Win_mtr_OCReq_U_CodeView_Sub senderOCReqSub = chkSender.DataContext as Win_mtr_OCReq_U_CodeView_Sub;

            //if (senderOCReqSub.InwareCloseYN.Equals("Y"))
            //{
            //    senderOCReqSub.InwareCloseYN = "N";
            //    senderOCReqSub.InwareCloseChecked = false;
            //}
            //else
            //{
            //    senderOCReqSub.InwareCloseYN = "Y";
            //    senderOCReqSub.InwareCloseChecked = true;
            //}
        }

        // 서브 그리드 입고마감 체크 이벤트
        private void chkInwareCloseYN_Checked(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            Win_mtr_OCReq_U_CodeView_Sub senderOCReqSub = chkSender.DataContext as Win_mtr_OCReq_U_CodeView_Sub;
            senderOCReqSub.InwareCloseYN = "Y";

            senderOCReqSub.InwareCloseChecked = true;
        }

        private void chkInwareCloseYN_Unchecked(object sender, RoutedEventArgs e)
        {
            CheckBox chkSender = sender as CheckBox;
            Win_mtr_OCReq_U_CodeView_Sub senderOCReqSub = chkSender.DataContext as Win_mtr_OCReq_U_CodeView_Sub;
            senderOCReqSub.InwareCloseYN = "N";

            senderOCReqSub.InwareCloseChecked = false;
        }

        // 서브 그리드 
        private void dgdSub_CanChecked()
        {
            if (dgdSub.Items.Count > 0)
            {
                for (int i = 0; i < dgdSub.Items.Count; i++)
                {
                    var OcReqSub = dgdSub.Items[i] as Win_mtr_OCReq_U_CodeView_Sub;
                    if (OcReqSub != null)
                    {
                        OcReqSub.IsEnabled = true;
                    }
                }
            }
        }

        #endregion // 서브 그리드 모음 (키 이벤트 제외)

        #region 서브 그리드 키 이벤트

        // 2019.08.27 PreviewKeyDown 는 key 다운과 같은것 같음
        private void DataGird_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Down || e.Key == Key.Up || e.Key == Key.Left || e.Key == Key.Right)
                {
                    DataGird_KeyDown(sender, e);
                }
            }
            catch (Exception ex)
            {

            }
        }
        // KeyDown 이벤트
        private void DataGird_KeyDown(object sender, KeyEventArgs e)
        {
            int currRow = dgdSub.Items.IndexOf(dgdSub.CurrentItem);
            int currCol = dgdSub.Columns.IndexOf(dgdSub.CurrentCell.Column);
            int startCol = 1;
            int endCol = 6;

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 열, 마지막 행 아님
                if (endCol == currCol && dgdSub.Items.Count - 1 > currRow)
                {
                    dgdSub.SelectedIndex = currRow + 1; // 이건 한줄 파란색으로 활성화 된 걸 조정하는 것입니다.
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow + 1], dgdSub.Columns[startCol]);

                } // 마지막 열 아님
                else if (endCol > currCol && dgdSub.Items.Count - 1 >= currRow)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);
                } // 마지막 열, 마지막 행
                else if (endCol == currCol && dgdSub.Items.Count - 1 == currRow)
                {
                    //btnSave.Focus();
                }
                else
                {
                    MessageBox.Show("나머지가 있나..");
                }
            }
            else if (e.Key == Key.Down)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 마지막 행 아님
                if (dgdSub.Items.Count - 1 > currRow)
                {
                    dgdSub.SelectedIndex = currRow + 1;
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow + 1], dgdSub.Columns[currCol]);
                } // 마지막 행일때
                else if (dgdSub.Items.Count - 1 == currRow)
                {
                    if (endCol > currCol) // 마지막 열이 아닌 경우, 열을 오른쪽으로 이동
                    {
                        //dgdSub.SelectedIndex = 0;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Up)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                // 첫행 아님
                if (currRow > 0)
                {
                    dgdSub.SelectedIndex = currRow - 1;
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow - 1], dgdSub.Columns[currCol]);
                } // 첫 행
                else if (dgdSub.Items.Count - 1 == currRow)
                {
                    if (0 < currCol) // 첫 열이 아닌 경우, 열을 왼쪽으로 이동
                    {
                        //dgdSub.SelectedIndex = 0;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol - 1]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Left)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (startCol < currCol)
                {
                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol - 1]);
                }
                else if (startCol == currCol)
                {
                    if (0 < currRow)
                    {
                        dgdSub.SelectedIndex = currRow - 1;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow - 1], dgdSub.Columns[endCol]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
            else if (e.Key == Key.Right)
            {
                e.Handled = true;
                (sender as DataGridCell).IsEditing = false;

                if (endCol > currCol)
                {

                    dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow], dgdSub.Columns[currCol + 1]);
                }
                else if (endCol == currCol)
                {
                    if (dgdSub.Items.Count - 1 > currRow)
                    {
                        dgdSub.SelectedIndex = currRow + 1;
                        dgdSub.CurrentCell = new DataGridCellInfo(dgdSub.Items[currRow + 1], dgdSub.Columns[startCol]);
                    }
                    else
                    {
                        //btnSave.Focus();
                    }
                }
            }
        }
        // KeyUp 이벤트
        private void DatagridIn_TextFocus(object sender, KeyEventArgs e)
        {
            // 엔터 → 포커스 = true → cell != null → 해당 텍스트박스가 null이 아니라면 
            // → 해당 텍스트박스가 포커스가 안되있음 SelectAll() or 포커스
            Lib.Instance.DataGridINTextBoxFocus(sender, e);
        }
        // GotFocus 이벤트
        private void DataGridCell_GotFocus(object sender, RoutedEventArgs e)
        {
            DataGridCell cell = sender as DataGridCell;
            if (lblMsg.Visibility == Visibility.Visible)
            {
                if (cell.Column.DisplayIndex != 7)
                {
                    cell.IsEditing = true;

                }
            }
        }
        // 2019.08.27 MouseUp 이벤트
        private void DataGridCell_MouseUp(object sender, MouseButtonEventArgs e)
        {


            DataGridCell cell = sender as DataGridCell;

            if (cell.Column.DisplayIndex != 7)
            {
                Lib.Instance.DataGridINTextBoxFocusByMouseUP(sender, e);

            }
        }
        // 2019.08.27 데이터그리드 Process 플러스파인더 검색
        // 서브그리드 - 발주 상세항목 : 품명 중복되지 않도록!!!!!!!!!!! → 중복이 되면, 자재입고등록화면에서, 중복으로 검색이됨
        private void txtItemName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var OcReqSub = dgdSub.CurrentItem as Win_mtr_OCReq_U_CodeView_Sub;

                if (OcReqSub != null)
                {
                    TextBox tb = new TextBox();

                    MainWindow.pf.ReturnCode(tb, 77, OcReqSub.Item_Name == null ? "" : OcReqSub.Item_Name);

                    e.Handled = true;

                    if (txtSupplierName.Tag == null || txtSupplierName.Tag.ToString().Trim().Equals("")
                 || txtSupplierName.Text.Trim().Equals(""))
                    {
                        MessageBox.Show("거래처를 먼저 선택해주세요.");
                        return;
                    }

                    if (txtSupplierName != null && txtSupplierName.Text != "")
                    {
                        // MainWindow.pf.ReturnCode(tb, 80, OcReqSub.Item_Name == null ? "" : OcReqSub.Item_Name);
                        MainWindow.pf.ReturnCode(tb, 7070, txtSupplierName.Tag.ToString().Trim());
                    }
                    else
                    {
                        MainWindow.pf.ReturnCode(tb, 7071, "");
                    }
                    ////MainWindow.pf.ReturnCode(tb, 80, OcReqSub.Item_Name == null ? "" : OcReqSub.Item_Name);
                    //MainWindow.pf.ReturnCode(tb, 80, txtSupplierName.Tag.ToString().Trim());


                    if (!tb.Text.Equals("") && tb.Tag != null && !tb.Tag.ToString().Equals(""))
                    {

                        getArticleInfo(tb.Tag.ToString());

                        //string[] req = tb.Tag.ToString().Trim().Split('/');
                        //tb.Text = req[0];
                        //tb.Tag = req[0];
                        //if (req.Length == 2)
                        //{
                        //    getCustomArticle(req[0], req[1]);
                        //}

                    }

                    if (tb.Tag != null)
                    {
                        ArticleInfo ai = getArticleInfo(tb.Tag.ToString());

                        if (ai != null)
                        {
                            OcReqSub.BuyerArticleNo = ai.BuyerArticleNo;
                            OcReqSub.Item_ID = ai.ArticleID;
                            OcReqSub.Item_Name = ai.Article;
                        }
                    }
                }
            }
        }

        // 서브 그리드 품번
        private void txtBuyerArticleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var OcReqSub = dgdSub.CurrentItem as Win_mtr_OCReq_U_CodeView_Sub;

                if (OcReqSub != null)
                {
                    TextBox tb = new TextBox();

                    //MainWindow.pf.ReturnCode(tb, 76, OcReqSub.BuyerArticleNo == null ? "" : OcReqSub.BuyerArticleNo);
                    e.Handled = true;

                    if (txtSupplierName.Tag == null || txtSupplierName.Tag.ToString().Trim().Equals("")
                 || txtSupplierName.Text.Trim().Equals(""))
                    {
                        MessageBox.Show("거래처를 먼저 선택해주세요.");
                        return;
                    }

                    if (txtSupplierName != null && txtSupplierName.Text != "")
                    {
                        // MainWindow.pf.ReturnCode(tb, 80, OcReqSub.Item_Name == null ? "" : OcReqSub.Item_Name);
                        MainWindow.pf.ReturnCode(tb, 7070, txtSupplierName.Tag.ToString().Trim());
                    }
                    else
                    {
                        MainWindow.pf.ReturnCode(tb, 7071, "");
                    }
                    ////MainWindow.pf.ReturnCode(tb, 80, OcReqSub.Item_Name == null ? "" : OcReqSub.Item_Name);
                    //MainWindow.pf.ReturnCode(tb, 80, txtSupplierName.Tag.ToString().Trim());


                    if (!tb.Text.Equals("") && tb.Tag != null && !tb.Tag.ToString().Equals(""))
                    {

                        getArticleInfo(tb.Tag.ToString());
                        //string[] req = tb.Tag.ToString().Trim().Split('/');
                        //tb.Text = req[0];
                        //tb.Tag = req[0];
                        //if (req.Length == 2)
                        //{
                        //    getCustomArticle(req[0], req[1]);
                        //}

                    }


                    if (tb.Tag != null)
                    {
                        ArticleInfo ai = getArticleInfo(tb.Tag.ToString());

                        if (ai != null)
                        {
                            OcReqSub.BuyerArticleNo = ai.BuyerArticleNo;
                            OcReqSub.Item_ID = ai.ArticleID;
                            OcReqSub.Item_Name = ai.Article;
                        }
                    }
                }
            }
        }

        // ArticleID 로 Article 정보 가져오기
        private ArticleInfo getArticleInfo(string setArticleID)
        {
            var getArticleInfo = new ArticleInfo();

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ArticleID", setArticleID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_Order_sArticleData", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRow dr = dt.Rows[0];

                        getArticleInfo = new ArticleInfo
                        {
                            ArticleGrpID = dr["ArticleGrpID"].ToString(),
                            UnitPrice = dr["UnitPrice"].ToString(),
                            UnitPriceClss = dr["UnitPriceClss"].ToString(),
                            UnitClss = dr["UnitClss"].ToString(),
                            PartGBNID = dr["PartGBNID"].ToString(),
                            BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            Article = dr["Article"].ToString(),
                            ArticleID = dr["ArticleID"].ToString(),
                        };
                    }
                }

                return getArticleInfo;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                DataStore.Instance.CloseConnection();
            }
        }

        #endregion // 서브 그리드 키 이벤트

        #endregion // Content 부분

        #region 메서드 모음

        private void re_Search(int rowNum)
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

        // 메인 그리드 검색 메서드
        private void FillGrid()
        {
            if (dgdMain.Items.Count > 0)
            {
                dgdMain.Items.Clear();
                dgdSub.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("ChkDate", chkReqDateSearch.IsChecked == true ? 1 : 0);
                sqlParameter.Add("FromDate", chkReqDateSearch.IsChecked == true ? dtpFromDateSearch.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("ToDate", chkReqDateSearch.IsChecked == true ? dtpToDateSearch.SelectedDate.Value.ToString("yyyyMMdd") : "");
                sqlParameter.Add("SUPPLIER_ID", chkCustom.IsChecked == true && txtCustomSearch.Tag != null ? txtCustomSearch.Tag.ToString() : ""); //거래처
                sqlParameter.Add("nChkApproval", chkApproval.IsChecked == true ? 1 : 0); //승인여부 체크
                sqlParameter.Add("sApproval", chkApproval.IsChecked == true && cboApproval.SelectedValue != null ? cboApproval.SelectedValue.ToString() : ""); //승인여부
                sqlParameter.Add("ChkArticleID", chkArticleSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("ArticleID", chkArticleSrh.IsChecked == true ? (txtArticleSrh.Tag == null ? "" : txtArticleSrh.Tag.ToString()) : "");
                sqlParameter.Add("ChkBuyArticleID", chkBuyArticleSrh.IsChecked == true ? 1 : 0);
                sqlParameter.Add("BuyArticleID", chkBuyArticleSrh.IsChecked == true ? (txtBuyArticleSrh.Tag == null ? "" : txtBuyArticleSrh.Tag.ToString()) : "");

                DataSet ds = DataStore.Instance.ProcedureToDataSet_LogWrite("xp_DyeAuxReq_sDyeAuxReq", sqlParameter, true, "R");

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        int i = 0;
                        DataRowCollection drc = dt.Rows;

                        foreach (DataRow dr in drc)
                        {
                            i++;
                            var OcReq = new Win_mtr_OCReq_U_CodeView()
                            {
                                ReqNo = i.ToString(),
                                REQ_ID = dr["REQ_ID"].ToString(),
                                REQ_DATE = ConvertDatePickerValue(dr["REQ_DATE"].ToString()), // ConvertDatePickerValue : 20190929 → 2019-09-29 로 변환
                                DUE_DATE = ConvertDatePickerValue(dr["DUE_DATE"].ToString()),
                                REQ_CHARGE = dr["REQ_CHARGE"].ToString(),
                                SUPPLIER_ID = dr["SUPPLIER_ID"].ToString(),
                                SUPPLIER_NAME = dr["SUPPLIER_NAME"].ToString(),
                                SUPPLIER_CHARGE = dr["SUPPLIER_CHARGE"].ToString(),
                                PAY_CONDITION = dr["PAY_CONDITION"].ToString(),
                                COMMENTS = dr["COMMENTS"].ToString(),
                                USE_CLSS = dr["USE_CLSS"].ToString(),
                                ApprovalYN = dr["ApprovalYN"].ToString(),
                                SUPPLIER_PhoneNo = dr["SUPPLIER_PhoneNo"].ToString(),
                                SUPPLIER_FaxNo = dr["SUPPLIER_FaxNo"].ToString(),
                                InspectReportAbleYN = dr["InspectReportAbleYN"].ToString()
                            };

                            dgdMain.Items.Add(OcReq);
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

        // 서브 그리드 검색 메서드
        private void FillGridSub(string ReqID)
        {
            if (dgdSub.Items.Count > 0)
            {
                dgdSub.Items.Clear();
            }

            try
            {
                Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                sqlParameter.Add("REQ_ID", ReqID);

                DataSet ds = DataStore.Instance.ProcedureToDataSet("xp_DyeAuxReq_sDyeAuxReqSub", sqlParameter, false);

                if (ds != null && ds.Tables.Count > 0)
                {
                    DataTable dt = ds.Tables[0];

                    if (dt.Rows.Count > 0)
                    {
                        DataRowCollection drc = dt.Rows;
                        int j = 0;

                        for (int i = 0; i < drc.Count; i++)
                        {
                            j = i;
                            DataRow dr = drc[i];

                            var OcReqSub = new Win_mtr_OCReq_U_CodeView_Sub()
                            {
                                ReqItemNo = (i + 1).ToString(),
                                Seq = dr["Seq"].ToString(),
                                Item_ID = dr["Item_ID"].ToString(),
                                Item_Name = dr["Item_Name"].ToString(),
                                Unit_Price = dr["Unit_Price"].ToString(),
                                Qty = stringFormatN0(dr["Qty"]),
                                Amount = stringFormatN0(dr["Amount"]),
                                Vat = dr["Vat"].ToString(),
                                Vat_YN = dr["Vat_YN"].ToString(),
                                COMMENTSITEM = dr["COMMENTS"].ToString(),
                                Item_For_Useing = dr["Item_For_Useing"].ToString().Trim().Equals("") ? "" : dr["Item_For_Useing"].ToString(),
                                Ddate = dr["Ddate"].ToString(),
                                Ddate_CV = DatePickerFormat(dr["Ddate"].ToString()),
                                //InspectReportAbleYN = dr["InspectReportAbleYN"].ToString() // 검사??
                                InwareCloseYN = dr["InwareCloseYN"].ToString(),
                                Weight = dr["Weight"].ToString(),
                                BuyerArticleNo = dr["BuyerArticleNo"].ToString(),
                            };

                            OcReqSub.InwareCloseChecked = OcReqSub.InwareCloseYN.Equals("Y") == true ? true : false;

                            dgdSub.Items.Add(OcReqSub);

                        }
                    } // for문 끝
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

        // 삭제 메서드
        private bool DeleteData(string req_ID)
        {
            bool flag = false;

            Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
            sqlParameter.Clear();
            sqlParameter.Add("REQ_ID", req_ID);

            try
            {
                string[] result = DataStore.Instance.ExecuteProcedure_NewLog("xp_DyeAuxReq_dDyeAuxReq", sqlParameter, "D");

                if (!result[0].Equals("success"))
                {
                    MessageBox.Show("삭제 실패");
                    flag = false;
                }
                else
                {
                    //MessageBox.Show("성공적으로 삭제되었습니다.");
                    flag = true;
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

        // 유효성 검사
        private bool CheckData()
        {
            bool flag = true;

            if (txtSupplierName.Tag == null
                || txtSupplierName.Tag.ToString().Trim().Equals(""))
            {
                MessageBox.Show("거래처가 입력되지 않았습니다.");
                flag = false;
            }

            // 서브 그리드 (발주 상세항목) : 필수 입력 (품명) 입력 안했을시
            if (dgdSub.Items.Count > 0)
            {
                bool dgdSubFlag = true;

                for (int i = 0; i < dgdSub.Items.Count; i++)
                {
                    var OcReqSub = dgdSub.Items[i] as Win_mtr_OCReq_U_CodeView_Sub;

                    if (OcReqSub != null)
                    {
                        if (OcReqSub.Item_ID == null
                            || OcReqSub.Item_ID.Trim().Equals(""))
                        {
                            dgdSubFlag = false;
                            break;
                        }

                        // 납기일자 8자리 날짜 20191112
                        if (OcReqSub.Ddate != null
                            && !OcReqSub.Ddate.Trim().Equals("")
                            && (CheckConvertInt(OcReqSub.Ddate) == false
                            || OcReqSub.Ddate.Trim().Length != 8))
                        {
                            dgdSub.SelectedIndex = i;

                            MessageBox.Show("납기일자는 8자리의 숫자로 입력해주세요.\r(ex 2019년 11월 12일 → 20191112)");
                            flag = false;
                            return flag;
                        }
                    }
                }

                if (dgdSubFlag == false)
                {
                    MessageBox.Show("발주상세항목에 품명이 입력되지 않았습니다.");
                    flag = false;
                    return flag;
                }
            }
            else
            {
                MessageBox.Show("발주상세항목에 품목을 추가해 주세요.");
                flag = false;
                return flag;
            }

            return flag;
        }

        // 저장
        private bool SaveData(string strFlag)
        {
            bool flag = false;

            List<Procedure> Prolist = new List<Procedure>();
            List<Dictionary<string, object>> ListParameter = new List<Dictionary<string, object>>();

            try
            {
                if (CheckData())
                {

                    Dictionary<string, object> sqlParameter = new Dictionary<string, object>();
                    sqlParameter.Clear();

                    sqlParameter.Add("REQ_ID", txtReqID.Text != null ? txtReqID.Text : "");
                    sqlParameter.Add("Req_Date", dtpReqDate.SelectedDate != null ? dtpReqDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("Due_Date", dtpDueDate.SelectedDate != null ? dtpDueDate.SelectedDate.Value.ToString("yyyyMMdd") : "");
                    sqlParameter.Add("Req_Charge", txtReqCharge.Text);
                    sqlParameter.Add("Supplier_ID", txtSupplierName.Tag != null ? txtSupplierName.Tag.ToString() : "");

                    sqlParameter.Add("Supplier_Charge", txtSupplierCharge.Text);
                    sqlParameter.Add("Pay_Condition", txtPayCondition.Text);
                    sqlParameter.Add("Comments", txtComments.Text);
                    sqlParameter.Add("Use_Clss", "");


                    sqlParameter.Add("InspectReportAbleYN", "");
                    sqlParameter.Add("OutwareReportAbleYN", "");

                    #region 추가

                    if (strFlag.Equals("I"))
                    {
                        sqlParameter.Add("ApprovalYN", "");
                        sqlParameter.Add("Create_User_ID", MainWindow.CurrentUser);

                        Procedure pro1 = new Procedure();
                        pro1.Name = "xp_DyeAuxReq_iDyeAuxReq";
                        pro1.OutputUseYN = "Y";
                        pro1.OutputName = "REQ_ID";
                        pro1.OutputLength = "10";

                        Prolist.Add(pro1);
                        ListParameter.Add(sqlParameter);

                        List<KeyValue> list_Result = new List<KeyValue>();
                        list_Result = DataStore.Instance.ExecuteAllProcedureOutputGetCS_NewLog(Prolist, ListParameter,"C");
                        string getReqID = string.Empty;

                        if (list_Result[0].key.ToLower() == "success")
                        {
                            list_Result.RemoveAt(0);
                            for (int i = 0; i < list_Result.Count; i++)
                            {
                                KeyValue kv = list_Result[i];
                                if (kv.key == "REQ_ID")
                                {
                                    getReqID = kv.value;

                                    AASS = kv.value;

                                    flag = true;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("저장 실패!! 에러 : " + list_Result[1]);
                            return false;
                        }

                        // 위에서 실행한 저장 메인그리드 저장 프로시저 삭제
                        Prolist.Clear();
                        ListParameter.Clear();

                        for (int i = 0; i < dgdSub.Items.Count; i++)
                        {

                            var OcReqSub = dgdSub.Items[i] as Win_mtr_OCReq_U_CodeView_Sub;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();

                            sqlParameter.Add("REQ_ID", getReqID);
                            sqlParameter.Add("Seq", i + 1);
                            sqlParameter.Add("Item_ID", OcReqSub.Item_ID);
                            sqlParameter.Add("item_Name", OcReqSub.Item_Name);
                            sqlParameter.Add("Item_For_Useing", OcReqSub.Item_For_Useing);

                            sqlParameter.Add("Unit_Price", OcReqSub.Unit_Price);
                            sqlParameter.Add("Qty", OcReqSub.Qty);
                            sqlParameter.Add("Amount", OcReqSub.Amount);
                            sqlParameter.Add("Vat", 0);
                            sqlParameter.Add("Vat_YN", "N");

                            sqlParameter.Add("Comments", OcReqSub.COMMENTSITEM);
                            sqlParameter.Add("DDate", OcReqSub.Ddate);
                            sqlParameter.Add("InwareCloseYN", OcReqSub.InwareCloseYN);
                            sqlParameter.Add("Create_User_ID", MainWindow.CurrentUser);

                            //System.Diagnostics.Debug.WriteLine("가격1" + ocReqItem.Unit_Price);
                            //System.Diagnostics.Debug.WriteLine("가격2" + ocReqItem.Qty);
                            //System.Diagnostics.Debug.WriteLine("일자1" + DatePickerStart.SelectedDate.Value.ToString("yyyyMMdd"));

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_DyeAuxReq_iDyeAuxReqSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "REQ_ID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);

                        }
                    }

                    #endregion

                    #region 수정

                    else if (strFlag.Equals("U"))
                    {

                        string ApprovalYN = "";
                        var OcReq = dgdMain.SelectedItem as Win_mtr_OCReq_U_CodeView;
                        if (OcReq != null)
                            ApprovalYN = OcReq.ApprovalYN;

                        sqlParameter.Add("ApprovalYN", ApprovalYN);
                        sqlParameter.Add("Update_User_ID", MainWindow.CurrentUser);

                        Procedure pro3 = new Procedure();
                        pro3.Name = "xp_DyeAuxReq_uDyeAuxReq";
                        pro3.OutputUseYN = "N";
                        pro3.OutputName = "REQ_ID";
                        pro3.OutputLength = "10";

                        Prolist.Add(pro3);
                        ListParameter.Add(sqlParameter);

                        // 모든것을 삭제한 후에, 새롭게 추가
                        sqlParameter = new Dictionary<string, object>();
                        sqlParameter.Clear();
                        sqlParameter.Add("REQ_ID", txtReqID.Text);
                        //sqlParameter.Add("Seq", "");

                        Procedure pro4 = new Procedure();
                        pro4.Name = "xp_DyeAuxReq_dDyeAuxReqSubAll";
                        pro4.OutputUseYN = "N";
                        pro4.OutputName = "REQ_ID";
                        pro4.OutputLength = "10";

                        Prolist.Add(pro4);
                        ListParameter.Add(sqlParameter);

                        for (int i = 0; i < dgdSub.Items.Count; i++)
                        {

                            var OcReqSub = dgdSub.Items[i] as Win_mtr_OCReq_U_CodeView_Sub;

                            sqlParameter = new Dictionary<string, object>();
                            sqlParameter.Clear();
                            sqlParameter.Add("REQ_ID", txtReqID.Text);
                            sqlParameter.Add("Seq", i + 1);
                            sqlParameter.Add("Item_ID", OcReqSub.Item_ID);
                            sqlParameter.Add("item_Name", OcReqSub.Item_Name);
                            sqlParameter.Add("Item_For_Useing", OcReqSub.Item_For_Useing);
                            sqlParameter.Add("Unit_Price", OcReqSub.Unit_Price);
                            sqlParameter.Add("Qty", ConvertDouble(OcReqSub.Qty));
                            sqlParameter.Add("Amount", ConvertDouble(OcReqSub.Amount));
                            sqlParameter.Add("Vat", 0);
                            sqlParameter.Add("Vat_YN", "N");
                            sqlParameter.Add("Comments", OcReqSub.COMMENTSITEM);
                            sqlParameter.Add("DDate", OcReqSub.Ddate);
                            sqlParameter.Add("InwareCloseYN", OcReqSub.InwareCloseYN);
                            sqlParameter.Add("Create_User_ID", MainWindow.CurrentUser);

                            //System.Diagnostics.Debug.WriteLine("가격1" + ocReqItem.Unit_Price);
                            //System.Diagnostics.Debug.WriteLine("가격2" + ocReqItem.Qty);
                            //System.Diagnostics.Debug.WriteLine("일자1" + DatePickerStart.SelectedDate.Value.ToString("yyyyMMdd"));

                            Procedure pro2 = new Procedure();
                            pro2.Name = "xp_DyeAuxReq_iDyeAuxReqSub";
                            pro2.OutputUseYN = "N";
                            pro2.OutputName = "REQ_ID";
                            pro2.OutputLength = "10";

                            Prolist.Add(pro2);
                            ListParameter.Add(sqlParameter);

                        }
                    }

                    #endregion

                    string[] Confirm = new string[2];
                    Confirm = DataStore.Instance.ExecuteAllProcedureOutputNew_NewLog(Prolist, ListParameter,"U");
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

        #endregion // 메서드 모음

        #region 기타 메서드 모음

        private string ConvertDatePickerValue(string str)
        {
            string result = "";

            if (!str.Trim().Equals("") && str.Length == 8)
            {
                result = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
            }

            return result;
        }

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN0(object obj)
        {
            return string.Format("{0:N0}", obj);
        }

        // 천마리 콤마, 소수점 버리기
        private string stringFormatN2(object obj)
        {
            return string.Format("{0:N2}", obj);
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

        // Int로 변환 가능한지 체크 이벤트
        private bool CheckConvertInt(string str)
        {
            bool flag = false;
            double chkInt = 0;

            if (!str.Trim().Equals(""))
            {
                str = str.Replace(",", "");
                str = str.Replace(".", "");
                str = str.Replace("-", "");

                if (Double.TryParse(str, out chkInt) == true)
                {
                    flag = true;
                }
            }

            return flag;
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

        #endregion // 기타 메서드 모음


        // 테스트
        private void btnTest1_Click(object sender, RoutedEventArgs e)
        {
            var OcReqSub = dgdSub.SelectedItem as Win_mtr_OCReq_U_CodeView_Sub;

            if (OcReqSub != null)
            {
                string msg = "";
                msg = "ReqItemNo : " + OcReqSub.ReqItemNo + "\r"
                    + "REQ_ID : " + OcReqSub.REQ_ID + "\r"
                    + "Seq : " + OcReqSub.Seq + "\r"
                    + "Item_ID : " + OcReqSub.Item_ID + "\r"
                    + "Item_Name : " + OcReqSub.Item_Name + "\r"
                    + "Unit_Price : " + OcReqSub.Unit_Price + "\r"
                    + "Qty : " + OcReqSub.Qty + "\r"
                    + "Amount : " + OcReqSub.Amount + "\r"
                    + "Vat : " + OcReqSub.Vat + "\r"
                    + "Vat_YN : " + OcReqSub.Vat_YN + "\r"
                    + "COMMENTSITEM : " + OcReqSub.COMMENTSITEM + "\r"
                    + "Item_For_Useing : " + OcReqSub.Item_For_Useing + "\r"
                    + "Ddate : " + OcReqSub.Ddate + "\r"
                    + "InspectReportAbleYN : " + OcReqSub.InspectReportAbleYN + "\r"
                    + "InwareCloseYN : " + OcReqSub.InwareCloseYN + "\r"
                    + "InwareCloseChecked : " + OcReqSub.InwareCloseChecked + "\r"
                    + "IsEnabled : " + OcReqSub.IsEnabled;

                MessageBox.Show(msg);
            }
        }

        private void btnTest2_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < dgdSub.Items.Count; i++)
            {
                var OcReqSub = dgdSub.Items[i] as Win_mtr_OCReq_U_CodeView_Sub;
                if (OcReqSub != null)
                {
                    OcReqSub.IsEnabled = true;
                }
            }
        }

        private void dtpDdate_CalendarClosed(object sender, RoutedEventArgs e)
        {
            DatePicker dtpSender = sender as DatePicker;

            var OcReqSub = dtpSender.DataContext as Win_mtr_OCReq_U_CodeView_Sub;

            if (OcReqSub != null)
            {
                OcReqSub.Ddate_CV = dtpSender.SelectedDate.Value.ToString("yyyy-MM-dd");
                OcReqSub.Ddate = dtpSender.SelectedDate.Value.ToString("yyyyMMdd");
            }
        }

        // 품명
        // 품명 검색 라벨 왼쪽 클릭 이벤트
        private void lblArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkArticleSrh.IsChecked == true)
            {
                chkArticleSrh.IsChecked = false;
            }
            else
            {
                chkArticleSrh.IsChecked = true;
            }
        }
        // 품명 검색 체크박스 이벤트
        private void chkArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkArticleSrh.IsChecked = true;

            txtArticleSrh.IsEnabled = true;
            btnPfArticleSrh.IsEnabled = true;
        }
        private void chkArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkArticleSrh.IsChecked = false;

            txtArticleSrh.IsEnabled = false;
            btnPfArticleSrh.IsEnabled = false;
        }
        // 품명 검색 엔터 → 플러스 파인더 이벤트
        private void txtArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {

            //if (e.Key == Key.Enter)
            //{
            //    rowNum = 0;
            //    re_Search();
            //}

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtArticleSrh, 76, "");
            }
        }
        // 품명 검색 플러스파인더 이벤트
        private void btnPfArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtArticleSrh, 76, "");
        }

        //

        // 품명
        // 품명 검색 라벨 왼쪽 클릭 이벤트
        private void lblBuyArticleSrh_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (chkBuyArticleSrh.IsChecked == true)
            {
                chkBuyArticleSrh.IsChecked = false;
            }
            else
            {
                chkBuyArticleSrh.IsChecked = true;
            }
        }
        // 품명 검색 체크박스 이벤트
        private void chkBuyArticleSrh_Checked(object sender, RoutedEventArgs e)
        {
            chkBuyArticleSrh.IsChecked = true;

            txtBuyArticleSrh.IsEnabled = true;
            btnPfBuyArticleSrh.IsEnabled = true;
        }
        private void chkBuyArticleSrh_Unchecked(object sender, RoutedEventArgs e)
        {
            chkBuyArticleSrh.IsChecked = false;

            txtBuyArticleSrh.IsEnabled = false;
            btnPfBuyArticleSrh.IsEnabled = false;
        }
        // 품명 검색 엔터 → 플러스 파인더 이벤트
        private void txtBuyArticleSrh_KeyDown(object sender, KeyEventArgs e)
        {

            //if (e.Key == Key.Enter)
            //{
            //    rowNum = 0;
            //    re_Search();
            //}

            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                MainWindow.pf.ReturnCode(txtBuyArticleSrh, 77, "");
            }
        }
        // 품명 검색 플러스파인더 이벤트
        private void btnPfBuyArticleSrh_Click(object sender, RoutedEventArgs e)
        {
            MainWindow.pf.ReturnCode(txtBuyArticleSrh, 77, "");
        }
    }

    #region CodeView(코드뷰)

    class Win_mtr_OCReq_U_CodeView : BaseView
    {
        public string ReqNo { get; set; }                  // 순번
        public string REQ_ID { get; set; }                // 발주번호
        public string REQ_DATE { get; set; }            // 발주일자
        public string DUE_DATE { get; set; }            // 납기일자
        public string REQ_CHARGE { get; set; }        // 발주처리자    
        public string SUPPLIER_ID { get; set; }
        public string SUPPLIER_NAME { get; set; }
        public string SUPPLIER_CHARGE { get; set; }
        public string PAY_CONDITION { get; set; }
        public string COMMENTS { get; set; }
        public string USE_CLSS { get; set; }
        public string ApprovalYN { get; set; } // 승인여부
        public string SUPPLIER_PhoneNo { get; set; }
        public string SUPPLIER_FaxNo { get; set; }
        public string InspectReportAbleYN { get; set; }
    }

    class Win_mtr_OCReq_U_CodeView_Sub : BaseView
    {
        public string ReqItemNo { get; set; }
        public string REQ_ID { get; set; }
        public string Seq { get; set; }
        public string Item_ID { get; set; }
        public string Item_Name { get; set; }
        public string Unit_Price { get; set; }
        public string Qty { get; set; }
        public string Amount { get; set; }
        public string Vat { get; set; }
        public string Vat_YN { get; set; }
        public string COMMENTSITEM { get; set; } // 비고
        public string Item_For_Useing { get; set; } // 용도
        public string Ddate { get; set; }
        public string Ddate_CV { get; set; }
        public string InspectReportAbleYN { get; set; }
        public string InwareCloseYN { get; set; } // 입고마감여부
        public string Weight { get; set; } // 중량 추가 : 발주서 인쇄 시에는 - 중량 * 갯수 / 1000 (kg 으로 변경해야 하므로)
        public bool InwareCloseChecked { get; set; }
        public string BuyerArticleNo { get; set; }

        public bool IsEnabled { get; set; }
    }
    #endregion
}