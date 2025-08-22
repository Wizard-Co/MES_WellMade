using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.ObjectModel;
using System.Data;

namespace WizMes_WellMade
{
    class PrintWorkAll
    {
        WizMes_WellMade.PopUp.NoticeMessage msg = new PopUp.NoticeMessage();

        private Application excelapp;
        private Workbook workbook;
        private Worksheet worksheet;
        private Range workrange;
        private Worksheet copysheet;
        private Worksheet pastesheet;

        /// <summary>
        /// 검교정 등록에서 계획서
        /// </summary>
        public void PrintWorkMeasureMachinePlan(bool preview_click,
            ObservableCollection<Win_Qul_MeasureMachine_U_CodeView> ovcMeasureMachine)
        {
            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            excelapp = new Application();

            string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
                + "\\Report\\계측기등록검교정계획서(품질관리).xls";
            workbook = excelapp.Workbooks.Add(MyBookPath);
            worksheet = workbook.Sheets["Form"];

            string strToday = "작성일자:" + DateTime.Today.ToString("yyyy년 MM월 dd일");
            //string strToday = DateTime.Today.ToString("yyyy년 MM월 dd일");
            string strYear = DateTime.Today.Year.ToString();

            workrange = worksheet.get_Range("A4", "C4");
            workrange.Value2 = strToday;
            workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            workrange = worksheet.get_Range("A3", "T3");
            workrange.Value2 = strYear + "년도 계측기 검교정 계획서";
            workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            int Page = 0;
            int DataCount = 0;
            int copyLine = 0;
            int ThisPageAllCount = (ovcMeasureMachine.Count - 1) / 10;
            if (((ovcMeasureMachine.Count - 1) % 10) > 0)
            {
                ThisPageAllCount++;
            }

            copysheet = workbook.Sheets["Form"];
            pastesheet = workbook.Sheets["Print"];

            while (ovcMeasureMachine.Count - 1 > DataCount)
            {
                Page++;
                if (Page != 1) { DataCount++; }  //+1
                copyLine = (Page - 1) * 27;
                copysheet.Select();
                //copysheet.UsedRange.Copy();
                copysheet.Cells.Range["A1", "T27"].Copy();
                pastesheet.Select();
                workrange = pastesheet.Cells[copyLine + 1, 1];
                //workrange = pastesheet.get_Range("A" + (copyLine +1), "A" + (copyLine+1));
                workrange.Select();
                pastesheet.Paste();

                workrange = pastesheet.get_Range("T" + (copyLine + 4), "T" + (copyLine + 4));    //페이지넘버
                workrange.Value2 = Page + "/" + ThisPageAllCount;
                workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                int j = 0;
                for (int i = DataCount; i < ovcMeasureMachine.Count; i++)
                {
                    if (j == 20) { break; }
                    int insertline = copyLine + 8 + j;

                    var MeasureMachine = ovcMeasureMachine[i] as Win_Qul_MeasureMachine_U_CodeView;

                    workrange = pastesheet.get_Range("A" + (insertline - 1), "A" + insertline);    //순번
                    workrange.Value2 = MeasureMachine.Num;
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 39;
                    workrange.ColumnWidth = 8.38;

                    workrange = pastesheet.get_Range("B" + (insertline - 1), "B" + insertline);    //계측기명
                    workrange.Value2 = MeasureMachine.MsrMachineName;
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 39;
                    workrange.ColumnWidth = 14.88;

                    workrange = pastesheet.get_Range("C" + (insertline - 1), "C" + insertline);    //NO
                    workrange.Value2 = MeasureMachine.MsrMachineMgrNo;
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 39;
                    workrange.ColumnWidth = 8.38;

                    workrange = pastesheet.get_Range("D" + (insertline - 1), "D" + insertline);    //교정주기
                    workrange.Value2 = MeasureMachine.ProofCycle + "/월";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 39;
                    workrange.ColumnWidth = 8.38;

                    workrange = pastesheet.get_Range("E" + (insertline - 1), "E" + insertline);    //최종검교정일
                    workrange.Value2 = Lib.Instance.StrDateTimeDot(MeasureMachine.LastProofDate_CV);
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 39;
                    workrange.ColumnWidth = 9.5;

                    workrange = pastesheet.get_Range("F" + (insertline - 1), "F" + insertline);    //차기검교정일
                    workrange.Value2 = Lib.Instance.StrDateTimeDot(MeasureMachine.NextProofDate_CV);
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 39;
                    workrange.ColumnWidth = 9.5;

                    workrange = pastesheet.get_Range("G" + (insertline - 1), "G" + (insertline - 1));    //넓이 및 높이 지정
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 6.75;

                    workrange = pastesheet.get_Range("G" + insertline, "G" + insertline);    //넓이 및 높이 지정
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 6.75;

                    workrange = pastesheet.get_Range("H" + (insertline - 1), "H" + (insertline - 1));  //검사일정 계획1
                    workrange.Value2 = MeasureMachine.PR1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("H" + insertline, "H" + insertline);    //검사일정 실적1
                    workrange.Value2 = MeasureMachine.R1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("H" + (insertline - 1), "H" + (insertline - 1));  //검사일정 계획2
                    workrange.Value2 = MeasureMachine.PR1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("H" + insertline, "H" + insertline);    //검사일정 실적2
                    workrange.Value2 = MeasureMachine.R1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("I" + (insertline - 1), "I" + (insertline - 1));  //검사일정 계획3
                    workrange.Value2 = MeasureMachine.PR1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("I" + insertline, "I" + insertline);    //검사일정 실적3
                    workrange.Value2 = MeasureMachine.R1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("J" + (insertline - 1), "J" + (insertline - 1));  //검사일정 계획4
                    workrange.Value2 = MeasureMachine.PR1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("J" + insertline, "J" + insertline);    //검사일정 실적4
                    workrange.Value2 = MeasureMachine.R1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("K" + (insertline - 1), "K" + (insertline - 1));  //검사일정 계획5
                    workrange.Value2 = MeasureMachine.PR1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("K" + insertline, "K" + insertline);    //검사일정 실적5
                    workrange.Value2 = MeasureMachine.R1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("L" + (insertline - 1), "L" + (insertline - 1));  //검사일정 계획6
                    workrange.Value2 = MeasureMachine.PR1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("L" + insertline, "L" + insertline);    //검사일정 실적6
                    workrange.Value2 = MeasureMachine.R1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("M" + (insertline - 1), "M" + (insertline - 1));  //검사일정 계획7
                    workrange.Value2 = MeasureMachine.PR1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("M" + insertline, "M" + insertline);    //검사일정 실적7
                    workrange.Value2 = MeasureMachine.R1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("N" + (insertline - 1), "N" + (insertline - 1));  //검사일정 계획8
                    workrange.Value2 = MeasureMachine.PR1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("N" + insertline, "N" + insertline);    //검사일정 실적8
                    workrange.Value2 = MeasureMachine.R1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("O" + (insertline - 1), "O" + (insertline - 1));  //검사일정 계획9
                    workrange.Value2 = MeasureMachine.PR1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("O" + insertline, "O" + insertline);    //검사일정 실적9
                    workrange.Value2 = MeasureMachine.R1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("P" + (insertline - 1), "P" + (insertline - 1));  //검사일정 계획10
                    workrange.Value2 = MeasureMachine.PR1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("P" + insertline, "P" + insertline);    //검사일정 실적10
                    workrange.Value2 = MeasureMachine.R1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("Q" + (insertline - 1), "Q" + (insertline - 1));  //검사일정 계획11
                    workrange.Value2 = MeasureMachine.PR1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("Q" + insertline, "Q" + insertline);    //검사일정 실적11
                    workrange.Value2 = MeasureMachine.R1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("R" + (insertline - 1), "R" + (insertline - 1));  //검사일정 계획12
                    workrange.Value2 = MeasureMachine.PR1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("R" + insertline, "R" + insertline);    //검사일정 실적12
                    workrange.Value2 = MeasureMachine.R1.Equals("1") ? "○" : "";
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("T" + (insertline - 1), "T" + insertline);    //비고
                    workrange.Value2 = MeasureMachine.Comments;
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 19.5;
                    workrange.ColumnWidth = 7.25;

                    DataCount = i;
                    j += 2;
                }
                //DataCount++;
            }

            msg.Hide();

            if (preview_click == true)
            {
                excelapp.Visible = true;
                pastesheet.PrintPreview();
            }
            else
            {
                excelapp.Visible = true;
                pastesheet.PrintOutEx();
            }
        }

        /// <summary>
        /// 검교정 등록에서 등록대장
        /// </summary>
        public void PrintWorkMeasureMachineRecordDocument(bool preview_click,
            ObservableCollection<Win_Qul_MeasureMachine_U_CodeView> ovcMeasureMachine)
        {
            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            excelapp = new Application();

            string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
                + "\\Report\\계측기이력등록대장(품질관리).xls";
            workbook = excelapp.Workbooks.Add(MyBookPath);
            worksheet = workbook.Sheets["Form"];

            string strToday = "작성일자:" + DateTime.Today.ToString("yyyy년 MM월 dd일");
            //string strToday = DateTime.Today.ToString("yyyy년 MM월 dd일");
            string strYear = DateTime.Today.Year.ToString();

            workrange = worksheet.get_Range("A4", "C4");
            workrange.Value2 = strToday;
            workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            workrange.RowHeight = 15.00;


            int Page = 0;
            int DataCount = 0;
            int copyLine = 0;
            int ThisPageAllCount = (ovcMeasureMachine.Count - 1) / 10;
            if (((ovcMeasureMachine.Count - 1) % 10) > 0)
            {
                ThisPageAllCount++;
            }

            copysheet = workbook.Sheets["Form"];
            pastesheet = workbook.Sheets["Print"];

            int ModCount = 0;

            if (ovcMeasureMachine.Count % 10 > 0)
            {
                ModCount = 10 - (ovcMeasureMachine.Count % 10);
            }

            while (ovcMeasureMachine.Count - 1 > DataCount)
            {
                Page++;
                if (Page != 1) { DataCount++; }  //+1
                copyLine = (Page - 1) * 16;
                copysheet.Select();
                //copysheet.UsedRange.Copy();
                copysheet.Cells.Range["A1", "O16"].Copy();
                pastesheet.Select();
                workrange = pastesheet.Cells[copyLine + 1, 1];
                workrange.Select();
                pastesheet.Paste();

                workrange = pastesheet.get_Range("A" + (copyLine + 1), "A" + (copyLine + 1));    //페이지넘버
                workrange.RowHeight = 13.5;
                workrange = pastesheet.get_Range("A" + (copyLine + 2), "A" + (copyLine + 2));    //페이지넘버
                workrange.RowHeight = 13.5;
                workrange = pastesheet.get_Range("A" + (copyLine + 3), "A" + (copyLine + 3));    //페이지넘버
                workrange.RowHeight = 30.75;
                workrange = pastesheet.get_Range("A" + (copyLine + 5), "A" + (copyLine + 5));    //페이지넘버
                workrange.RowHeight = 41.25;
                workrange = pastesheet.get_Range("A" + (copyLine + 16), "A" + (copyLine + 16));    //페이지넘버
                workrange.RowHeight = 26.25;

                workrange = pastesheet.get_Range("O" + (copyLine + 4), "O" + (copyLine + 4));    //페이지넘버
                workrange.Value2 = Page + "/" + ThisPageAllCount;
                workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                workrange.ColumnWidth = 3.00;
                workrange.RowHeight = 15.00;

                int j = 0;

                for (int i = DataCount; i < ovcMeasureMachine.Count + ModCount; i++)
                {
                    if (j == 10) { break; }
                    int insertline = copyLine + 6 + j;

                    if (ovcMeasureMachine.Count > i)
                    {
                        var MeasureMachine = ovcMeasureMachine[i] as Win_Qul_MeasureMachine_U_CodeView;

                        workrange = pastesheet.get_Range("A" + insertline, "A" + insertline);    //순번
                        workrange.Value2 = MeasureMachine.Num;
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 3.33;

                        workrange = pastesheet.get_Range("B" + insertline, "B" + insertline);    //관리No
                        workrange.Value2 = MeasureMachine.MsrMachineMgrNo;
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 5.67;

                        workrange = pastesheet.get_Range("C" + insertline, "C" + insertline);    //계측기명
                        workrange.Value2 = MeasureMachine.MsrMachineName;
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 14.56;

                        workrange = pastesheet.get_Range("D" + insertline, "D" + insertline);    //계측기명
                        workrange.ColumnWidth = 0;

                        workrange = pastesheet.get_Range("E" + insertline, "E" + insertline);    //제작사
                        workrange.Value2 = MeasureMachine.MsrMachineMsrBuyCustom;
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 7.67;

                        workrange = pastesheet.get_Range("F" + insertline, "F" + insertline);    //규격
                        workrange.Value2 = MeasureMachine.MsrMachineSpec;
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 9.44;

                        workrange = pastesheet.get_Range("G" + insertline, "G" + insertline);    //범위
                        workrange.Value2 = MeasureMachine.MsrMachineRange;
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 6.89;

                        workrange = pastesheet.get_Range("H" + insertline, "H" + insertline);    //기기번호
                        workrange.Value2 = MeasureMachine.MsrMachineNo;
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 7.89;

                        workrange = pastesheet.get_Range("I" + insertline, "I" + insertline);    //구매일자
                        workrange.Value2 = Lib.Instance.StrDateTimeDot(MeasureMachine.MsrMachineBuyDate_CV);
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 10.11;

                        workrange = pastesheet.get_Range("J" + insertline, "J" + insertline);    //교정주기
                        workrange.Value2 = MeasureMachine.ProofCycle + "/월";
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 6.22;

                        workrange = pastesheet.get_Range("K" + insertline, "K" + insertline);    //등록일자
                        workrange.Value2 = Lib.Instance.StrDateTimeDot(MeasureMachine.MsrMachineSetDate_CV);
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 10.11;

                        workrange = pastesheet.get_Range("L" + insertline, "L" + insertline);    //담당자확인
                        workrange.Value2 = MeasureMachine.MsrmachinePerson;
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 5.67;

                        workrange = pastesheet.get_Range("M" + insertline, "M" + insertline);    //확인
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 5.67;

                        workrange = pastesheet.get_Range("N" + insertline, "N" + insertline);    //계측기명
                        workrange.ColumnWidth = 9.89;
                        workrange = pastesheet.get_Range("O" + insertline, "O" + insertline);    //계측기명
                        workrange.ColumnWidth = 3.00;

                        workrange = pastesheet.get_Range("N" + insertline, "O" + insertline);    //비고                    
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.Value2 = MeasureMachine.Comments;
                    }
                    else
                    {
                        workrange = pastesheet.get_Range("A" + insertline, "A" + insertline);    //순번
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 3.33;

                        workrange = pastesheet.get_Range("B" + insertline, "B" + insertline);    //관리No
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 5.67;

                        workrange = pastesheet.get_Range("C" + insertline, "C" + insertline);    //계측기명
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 14.56;

                        workrange = pastesheet.get_Range("D" + insertline, "D" + insertline);    //계측기명
                        workrange.ColumnWidth = 0;

                        workrange = pastesheet.get_Range("E" + insertline, "E" + insertline);    //제작사
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 7.67;

                        workrange = pastesheet.get_Range("F" + insertline, "F" + insertline);    //규격
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 9.44;

                        workrange = pastesheet.get_Range("G" + insertline, "G" + insertline);    //범위
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 6.89;

                        workrange = pastesheet.get_Range("H" + insertline, "H" + insertline);    //기기번호
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 7.89;

                        workrange = pastesheet.get_Range("I" + insertline, "I" + insertline);    //구매일자
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 10.11;

                        workrange = pastesheet.get_Range("J" + insertline, "J" + insertline);    //교정주기
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 6.22;

                        workrange = pastesheet.get_Range("K" + insertline, "K" + insertline);    //등록일자
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 10.11;

                        workrange = pastesheet.get_Range("L" + insertline, "L" + insertline);    //담당자확인
                        workrange.Value2 = "";
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 5.67;

                        workrange = pastesheet.get_Range("M" + insertline, "M" + insertline);    //확인
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.ColumnWidth = 5.67;

                        workrange = pastesheet.get_Range("N" + insertline, "N" + insertline);    //계측기명
                        workrange.ColumnWidth = 9.89;
                        workrange = pastesheet.get_Range("O" + insertline, "O" + insertline);    //계측기명
                        workrange.ColumnWidth = 3.00;

                        workrange = pastesheet.get_Range("N" + insertline, "O" + insertline);    //비고                    
                        workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        workrange.RowHeight = 39.75;
                        workrange.Value2 = "";
                    }

                    DataCount = i;
                    j += 1;
                }
                //DataCount++;
            }

            msg.Hide();

            if (preview_click == true)
            {
                excelapp.Visible = true;
                pastesheet.PrintPreview();
            }
            else
            {
                excelapp.Visible = true;
                pastesheet.PrintOutEx();
            }
        }

        /// <summary>
        /// 검교정 검교정 등록에서 계측기 이력카드
        /// </summary>
        public void PrintWorkMeasureDocument(bool preview_click,
            ObservableCollection<Win_Qul_Measure_U_CodeView> ovcMeasure, DataRow dataRow)
        {
            msg.Show();
            msg.Topmost = true;
            msg.Refresh();

            excelapp = new Application();

            string MyBookPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
                + "\\Report\\계측기이력관리카드(품질관리).xls";
            workbook = excelapp.Workbooks.Add(MyBookPath);
            worksheet = workbook.Sheets["Form"];

            int Page = 0;
            int DataCount = 0;
            int copyLine = 0;
            int ThisPageAllCount = (ovcMeasure.Count - 1) / 10;
            if (((ovcMeasure.Count - 1) % 10) > 0)
            {
                ThisPageAllCount++;
            }

            copysheet = workbook.Sheets["Form"];
            pastesheet = workbook.Sheets["Print"];

            while (ovcMeasure.Count - 1 > DataCount)
            {
                Page++;
                if (Page != 1) { DataCount++; }  //+1
                copyLine = (Page - 1) * 23;
                copysheet.Select();
                //copysheet.UsedRange.Copy();
                copysheet.Cells.Range["A1", "X23"].Copy();
                pastesheet.Select();
                workrange = pastesheet.Cells[copyLine + 1, 1];
                workrange.Select();
                pastesheet.Paste();

                workrange = pastesheet.get_Range("F" + (copyLine + 2), "L" + (copyLine + 2));    //관리번호
                workrange.Value2 = dataRow["MsrMachineMgrNo"];
                workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                workrange.ColumnWidth = 3;
                workrange.RowHeight = 35;

                workrange = pastesheet.get_Range("F" + (copyLine + 3), "L" + (copyLine + 3));    //계측기명
                workrange.Value2 = dataRow["MsrMachineName"];
                workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                workrange.ColumnWidth = 3;
                workrange.RowHeight = 35;

                workrange = pastesheet.get_Range("F" + (copyLine + 4), "L" + (copyLine + 4));    //규격1
                workrange.Value2 = dataRow["MsrMachineSpec"];
                workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                workrange.ColumnWidth = 3;
                workrange.RowHeight = 35;

                workrange = pastesheet.get_Range("F" + (copyLine + 5), "L" + (copyLine + 5));    //규격2
                workrange.Value2 = ovcMeasure[ovcMeasure.Count - 1].MsrMachineSpec;
                workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                workrange.ColumnWidth = 3;
                workrange.RowHeight = 35;

                workrange = pastesheet.get_Range("F" + (copyLine + 6), "L" + (copyLine + 6));    //기기번호
                workrange.Value2 = dataRow["MsrMachineNo"];
                workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                workrange.ColumnWidth = 3;
                workrange.RowHeight = 35;

                workrange = pastesheet.get_Range("F" + (copyLine + 7), "L" + (copyLine + 7));    //제작회사
                workrange.Value2 = dataRow["MsrMachineMsrCustom"];
                workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                workrange.ColumnWidth = 3;
                workrange.RowHeight = 35;

                workrange = pastesheet.get_Range("F" + (copyLine + 8), "L" + (copyLine + 8));    //교정주기
                workrange.Value2 = ovcMeasure[ovcMeasure.Count - 1].ProofCycle;
                workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                workrange.ColumnWidth = 3;
                workrange.RowHeight = 35;

                workrange = pastesheet.get_Range("F" + (copyLine + 9), "L" + (copyLine + 9));    //구입가격
                workrange.Value2 = ovcMeasure[ovcMeasure.Count - 1].MsrMachinePrice;
                workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                workrange.ColumnWidth = 3;
                workrange.RowHeight = 35;

                workrange = pastesheet.get_Range("F" + (copyLine + 10), "L" + (copyLine + 10));    //사용팀
                workrange.Value2 = ovcMeasure[ovcMeasure.Count - 1].MsrMachineUseTeam;
                workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                workrange.ColumnWidth = 3;
                workrange.RowHeight = 35;

                int j = 0;
                for (int i = DataCount; i < ovcMeasure.Count; i++)
                {
                    if (j == 10) { break; }
                    int insertline = copyLine + 13 + j;

                    var Measure = ovcMeasure[i] as Win_Qul_Measure_U_CodeView;

                    workrange = pastesheet.get_Range("B" + insertline, "B" + insertline);    //순번
                    workrange.Value2 = Measure.Num;
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 31.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("C" + insertline, "E" + insertline);    //일자
                    workrange.Value2 = Lib.Instance.StrDateTimeDot(Measure.ProofDate);
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 31.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("F" + insertline, "K" + insertline);    //변경사항
                    workrange.Value2 = Measure.ChangePoint;
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 31.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("P" + insertline, "S" + insertline);    //확인
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 31.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("T" + insertline, "W" + insertline);    //승인
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 31.5;
                    workrange.ColumnWidth = 3;

                    workrange = pastesheet.get_Range("T" + insertline, "W" + insertline);    //비고
                    workrange.Value2 = Measure.Comments;
                    workrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    workrange.RowHeight = 31.5;
                    workrange.ColumnWidth = 3;

                    DataCount = i;
                    j += 1;
                }
                //DataCount++;
            }

            msg.Hide();

            if (preview_click == true)
            {
                excelapp.Visible = true;
                pastesheet.PrintPreview();
            }
            else
            {
                excelapp.Visible = true;
                pastesheet.PrintOutEx();
            }
        }
    }
}
