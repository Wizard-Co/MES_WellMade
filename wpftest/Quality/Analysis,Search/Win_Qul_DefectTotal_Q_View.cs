using System.Collections.ObjectModel;

namespace WizMes_WellMade
{
    class Win_Qul_DefectTotal_Q_View : BaseView
    {
        public Win_Qul_DefectTotal_Q_View()
        {
        }

        public ObservableCollection<CodeView> cboTrade { get; set; }


        // 전체 탭 _ 그리드 로우컬럼 헤더.
        public string TOTAL_RowHeaderColumns { get; set; }
        // 인수,공정, 출하, 고객 탭 _ 그리드 로우컬럼 헤더.
        public string INSU_RowHeaderColumns { get; set; }


        //전체 탭 _ 각 항목별 목표설정 용도.
        public string TOTAL_PreDefectGoal { get; set; }        //전년도 실적
        public string TOTAL_DefectGoal { get; set; }        //당해목표
        public string TOTAL_AvgDefectGoal { get; set; }     // 평균


        public string TOTAL_MM1 { get; set; }
        public string TOTAL_MM2 { get; set; }
        public string TOTAL_MM3 { get; set; }
        public string TOTAL_MM4 { get; set; }
        public string TOTAL_MM5 { get; set; }
        public string TOTAL_MM6 { get; set; }
        public string TOTAL_MM7 { get; set; }
        public string TOTAL_MM8 { get; set; }
        public string TOTAL_MM9 { get; set; }
        public string TOTAL_MM10 { get; set; }
        public string TOTAL_MM11 { get; set; }
        public string TOTAL_MM12 { get; set; }


        //인수 탭 _ 종합현황 그리드.

        public string INSU_PreDefectGoal { get; set; }        //전년도 실적
        public string INSU_DefectGoal { get; set; }        //당해목표
        public string INSU_TotalDefectGoal { get; set; }     // 합계 토탈

        public string INSU_MM1 { get; set; }
        public string INSU_MM2 { get; set; }
        public string INSU_MM3 { get; set; }
        public string INSU_MM4 { get; set; }
        public string INSU_MM5 { get; set; }
        public string INSU_MM6 { get; set; }
        public string INSU_MM7 { get; set; }
        public string INSU_MM8 { get; set; }
        public string INSU_MM9 { get; set; }
        public string INSU_MM10 { get; set; }
        public string INSU_MM11 { get; set; }
        public string INSU_MM12 { get; set; }

        //인수 탭 _ 유형불량 그리드.

        public string INSU_GroupingName { get; set; }
        public string INSU_Minus2qty { get; set; }
        public string INSU_Minus2rate { get; set; }
        public string INSU_Minus1qty { get; set; }
        public string INSU_Minus1rate { get; set; }
        public string INSU_MMqty { get; set; }
        public string INSU_MMrate { get; set; }



        //자주/공정 탭 _ 종합현황 그리드.

        public string PROC_PreDefectGoal { get; set; }        //전년도 실적
        public string PROC_DefectGoal { get; set; }        //당해목표
        public string PROC_TotalDefectGoal { get; set; }     // 합계 토탈

        public string PROC_MM1 { get; set; }
        public string PROC_MM2 { get; set; }
        public string PROC_MM3 { get; set; }
        public string PROC_MM4 { get; set; }
        public string PROC_MM5 { get; set; }
        public string PROC_MM6 { get; set; }
        public string PROC_MM7 { get; set; }
        public string PROC_MM8 { get; set; }
        public string PROC_MM9 { get; set; }
        public string PROC_MM10 { get; set; }
        public string PROC_MM11 { get; set; }
        public string PROC_MM12 { get; set; }

        //자주/공정 탭 _ 유형불량 그리드.

        public string PROC_GroupingName { get; set; }
        public string PROC_Minus2qty { get; set; }
        public string PROC_Minus2rate { get; set; }
        public string PROC_Minus1qty { get; set; }
        public string PROC_Minus1rate { get; set; }
        public string PROC_MMqty { get; set; }
        public string PROC_MMrate { get; set; }



        //출하 탭 _ 종합현황 그리드.

        public string SHIP_PreDefectGoal { get; set; }        //전년도 실적
        public string SHIP_DefectGoal { get; set; }        //당해목표
        public string SHIP_TotalDefectGoal { get; set; }     // 합계 토탈

        public string SHIP_MM1 { get; set; }
        public string SHIP_MM2 { get; set; }
        public string SHIP_MM3 { get; set; }
        public string SHIP_MM4 { get; set; }
        public string SHIP_MM5 { get; set; }
        public string SHIP_MM6 { get; set; }
        public string SHIP_MM7 { get; set; }
        public string SHIP_MM8 { get; set; }
        public string SHIP_MM9 { get; set; }
        public string SHIP_MM10 { get; set; }
        public string SHIP_MM11 { get; set; }
        public string SHIP_MM12 { get; set; }

        //출하 탭 _ 유형불량 그리드.

        public string SHIP_GroupingName { get; set; }
        public string SHIP_Minus2qty { get; set; }
        public string SHIP_Minus2rate { get; set; }
        public string SHIP_Minus1qty { get; set; }
        public string SHIP_Minus1rate { get; set; }
        public string SHIP_MMqty { get; set; }
        public string SHIP_MMrate { get; set; }



        //고객 탭 _ 종합현황 그리드.

        public string CUST_PreDefectGoal { get; set; }        //전년도 실적
        public string CUST_DefectGoal { get; set; }        //당해목표
        public string CUST_TotalDefectGoal { get; set; }     // 합계 토탈

        public string CUST_MM1 { get; set; }
        public string CUST_MM2 { get; set; }
        public string CUST_MM3 { get; set; }
        public string CUST_MM4 { get; set; }
        public string CUST_MM5 { get; set; }
        public string CUST_MM6 { get; set; }
        public string CUST_MM7 { get; set; }
        public string CUST_MM8 { get; set; }
        public string CUST_MM9 { get; set; }
        public string CUST_MM10 { get; set; }
        public string CUST_MM11 { get; set; }
        public string CUST_MM12 { get; set; }

        //고객 탭 _ 유형불량 그리드.

        public string CUST_GroupingName { get; set; }
        public string CUST_Minus2qty { get; set; }
        public string CUST_Minus2rate { get; set; }
        public string CUST_Minus1qty { get; set; }
        public string CUST_Minus1rate { get; set; }
        public string CUST_MMqty { get; set; }
        public string CUST_MMrate { get; set; }



    }
}
