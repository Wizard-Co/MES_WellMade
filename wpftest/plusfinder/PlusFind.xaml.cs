using System.Windows.Controls;
using System.Windows.Data;

namespace WizMes_WellMade
{
    /// <summary>
    /// PlusFind.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class PlusFind : UserControl
    {
        private PlusFinderView pfv = new PlusFinderView();
        public PlusFind(string txtsearch)
        {


            this.DataContext = pfv;
            this.Width = 400;
            this.Height = 300;

            DataGrid datagrid = new DataGrid();
            DataGridTextColumn h1 = new DataGridTextColumn();
            h1.Header = "코드";
            h1.Binding = new Binding("code");
            h1.Width = 120;
            DataGridTextColumn h2 = new DataGridTextColumn();
            h2.Header = "국가";
            h2.Binding = new Binding("name");

            datagrid.Columns.Add(h1);
            datagrid.Columns.Add(h2);

            InitializeComponent();

            mRootGrid.Children.Add(datagrid);
            this.UpdateLayout();
        }

        public void getGrid()
        {
            int large = 0;
            string sql = "";
            string dOrderByStr = "";
            string mmiddle = "";

            string m_sArticleField;
            string lsTemp;

            string mLarge = "";
            string[] pnlName = new string[3];

            //[1] 거래처 코드
            if (large == 0)
            {

                pfv.m_sCodeField = "CustomID";
                pfv.m_sNameField = "KCustom";
                sql =
                    "SELECT CustomID AS [코드], KCustom AS [상호] " +
                    "       , [상호(변경)] = replace(REPLACE(KCustom, '(주)', '') , '㈜', '')           " +
                    "       , [거래]  = case when TradeID = '1' then '1' else '9' end              " +
                    "FROM [mt_Custom] " +
                    "WHERE UseClss = '' ";
                dOrderByStr = " ORDER BY [거래] ,  [상호(변경)]    ";           // S_201505_AFT_06  에 따른 정렬 방식 변경, old :  ORDER BY [상호]

                if (!(string.IsNullOrEmpty(mmiddle)))
                    sql += "AND TradeID in ( select code_id from cm_code where code_gbn = 'CMMTRAD' and Relation = '" + mmiddle + "' ) ";

            }
            // [2] 품명 코드
            else if (large == 1)
            {
                pfv.m_sCodeField = "ArticleID";
                pfv.m_sNameField = "Article";


                sql =
                    "SELECT ArticleID AS [코드], Article AS [품명] " +
                    "FROM [mt_Article] " +
                    "WHERE UseClss = '' ";

                if (!(string.IsNullOrEmpty(mmiddle)))
                    sql += "AND articleid in ( select item_id  from Dye_Order_Req_Sub dor where dor.REQ_ID   = '" + mmiddle.Trim() + "' ) ";

                // S_201505_AFT_06  에 따른 정렬 방식 변경, old :  ORDER BY [코드]
                dOrderByStr = " ORDER BY [품명]  ";

            }
            // [3] 사원코드
            else if (large == 2)
            {
                pfv.m_sCodeField = "A.PersonID";
                pfv.m_sNameField = "A.Name";

                sql =
                    "SELECT A.PersonID AS [코드], A.Name AS [성명], B.Depart AS [부서] " +
                    "FROM [mt_Person] A, [mt_Depart] B " +
                    "WHERE A.DepartID = B.DepartID AND A.UseClss = '' ";

                if (!(string.IsNullOrEmpty(mmiddle)))
                    sql += "AND A.DepartID = '" + mmiddle + "'";

                dOrderByStr = " ORDER BY [코드], [성명], [부서] ";
            }
            //[4] 불량 코드
            else if (large == 3)
            {
                pfv.m_sCodeField = "DefectID";
                pfv.m_sNameField = "KDefect";

                sql =
                    "SELECT DefectID AS [코드], KDefect AS [불량명], TagName AS [Tag 명] " +
                    "FROM [mt_Defect] ";


                if (!(string.IsNullOrEmpty(mmiddle)))
                    sql += "WHERE DefectID like %'" + mmiddle + "'%";


                dOrderByStr = " ORDER BY [코드] ";
            }
            // [5] 수주
            else if (large == 4)
            {
                pfv.m_sCodeField = "A.OrderID";
                pfv.m_sNameField = "A.OrderNo";
                m_sArticleField = "C.BuyerArticleNo"; //S_201704_대한산업_15에 의한 추가

                sql =
            "SELECT A.OrderID AS [ 관리번호 ], A.OrderNo AS [     Order No     ], B.KCustom AS [거래처], " +
                    "C.Article AS [품명], " +
                    "[  차종  ] = (select model from mt_ModelMaster mm where mm.ModelID = A.BuyerModelID  ) , " +
                    "[  품목명] = C.BuyerArticleNo , " +
                    "A.OrderQty AS [수주량], " +
                    "[단위]    = dbo.fn_cm_sCodeInfo( 'CMMUNIT', A.UnitClss, 'N') , " +
                    "A.DvlyDate AS [납기일자], " +
                    "D.WorkName AS [  가공구분  ], " +
                    "A.CustomID AS [거래처코드], A.ArticleID AS [품명코드] " +
            "FROM [Order] A, [mt_Custom] B, [mt_Article] C, [mt_Work] D  " +
            "WHERE A.CustomID = B.CustomID " +
                "AND A.ArticleID = C.ArticleID " +
                "AND A.WorkID = D.WorkID ";

                //mOrderByStr = " ORDER BY [ 관리번호 ] "
                dOrderByStr = " ORDER BY AcptDate DESC ";

            }
            // 염료 코드
            else if (large == 5)
            {

                pfv.m_sCodeField = "DyeAuxID";
                pfv.m_sNameField = "DyeAux";


                sql =
                    "SELECT DyeAuxID AS [코드], DyeAux AS [염료명] " +
                    "FROM [vw_mt_DyeAux] " +
                    "WHERE DyeAuxID LIKE '1%' AND ISNULL(UseClss, '') NOT IN ('*') ";
                dOrderByStr = " ORDER BY [코드] ";
            }
            // 조제코드
            else if (large == 6)
            {
                pfv.m_sCodeField = "DyeAuxID";
                pfv.m_sNameField = "DyeAux";
                sql =
                    "SELECT DyeAuxID AS [코드], DyeAux AS [염료명] " +
                    "FROM [vw_mt_DyeAux] " +
                    "WHERE DyeAuxID LIKE '0%' AND ISNULL(UseClss, '') NOT IN ('*') ";
                dOrderByStr = " ORDER BY [코드] ";
            }
            // 가공코드
            else if (large == 7)
            {
                pfv.m_sCodeField = "WorkID";
                pfv.m_sNameField = "WorkName";

                sql = "SELECT WorkID AS [코드], WorkName AS [가공명] " +
                      "FROM [mt_Work] " +
                "WHERE UseClss = '' ";
                dOrderByStr = " ORDER BY [코드] ";

            }
            // 사종구분
            else if (large == 8)
            {
                pfv.m_sCodeField = "ThreadID";
                pfv.m_sNameField = "Thread";

                sql = "SELECT ThreadID AS [코드], Thread AS [사종명] " +
                        "FROM [mt_Thread] " +
                        "WHERE UseClss = '' ";
                dOrderByStr = " ORDER BY [코드] ";
            }
            // 원단폭
            else if (large == 9)
            {
                //pfv.m_sCodeField = "StuffWidthID"
                //pfv.m_sNameField = "StuffWidth"

                //SQL = "SELECT StuffWidthID AS [코드], StuffWidth AS [원단폭] " & _
                //      "FROM [mt_StuffWidth] " & _
                //      "WHERE UseClss = '' "
                //mOrderByStr = " ORDER BY [코드] "
            }
            // 공정명
            else if (large == 10)
            {
                pfv.m_sCodeField = "ProcessID";
                pfv.m_sNameField = "Process";


                sql = "SELECT ProcessID As [코드], Process AS [공정명]" +
                      "FROM [mt_Process] WHERE  useclss <> '*' and ProcessID NOT LIKE '%00' ";
                dOrderByStr = " ORDER BY [코드] ";

            }
            // 염색패턴
            else if (large == 11)
            {
                pfv.m_sCodeField = "PatternID";
                pfv.m_sNameField = "Pattern";


                sql = "SELECT PatternID AS [코드], Pattern AS [패턴명], WaterRate AS [욕비], Temper AS [온도], Times AS [시간] " +
                        "FROM [mt_DyePattern] " +
                        "WHERE UseClss = '' ";
                dOrderByStr = " ORDER BY [코드] ";
            }
            // 기계코드
            else if (large == 12)
            {
                pfv.m_sCodeField = "MCID";
                pfv.m_sNameField = "MCNAME";


                sql = "SELECT MCID AS [코드], MCNAME AS [기계명] " +
                        "FROM [mt_Mc] " +
                        "WHERE UseClss = '' ";
                dOrderByStr = " ORDER BY [코드] ";
            }
            // 부품코드
            else if (large == 13)
            {
                pfv.m_sCodeField = "MCPartID";
                pfv.m_sNameField = "MCPartName";
                sql = "SELECT MCPartID AS [코드], MCPartName AS [부품명] " +
                        "FROM mt_MCPart  " +
                        "WHERE UseClss = '' ";
                if (!(string.IsNullOrEmpty(mmiddle)))
                    sql += "AND ForUse  in ( " + mmiddle + ") "; //ForUse :용도, 공통, 기계용, 금형용

                dOrderByStr = " ORDER BY [코드] ";
            }
            //점검항목코드
            else if (large == 14)
            {
                pfv.m_sNameField = "checkitem";
                sql = "SELECT DISTINCT '' AS [코드], checkitem AS [점검항목] " +
                    "FROM MCCheckSub ";

                dOrderByStr = " ORDER BY [점검항목] ";
            }
            //점검방법코드
            else if (large == 15)
            {
                pfv.m_sNameField = "checkmethod";
                sql = "SELECT DISTINCT '' AS [코드], checkmethod AS [점검방법] " +
                      "FROM MCCheckSub ";

                dOrderByStr = " ORDER BY [점검방법] ";
            }
            //점검기준코드
            else if (large == 16)
            {
                pfv.m_sNameField = "checkstd";
                sql = "SELECT DISTINCT '' AS [코드], checkstd AS [점검기준] " +
                        "FROM MCCheckSub ";
                dOrderByStr = " ORDER BY [점검기준] ";
            }
            //BT 조회
            else if (large == 17)
            {
                pfv.m_sNameField = "BTID";
                sql = "SELECT A.BTID AS [접수번호], A.BTIDSeq AS [차수], B.KCustom AS [거래처], C.Article AS [품명], A.ColorCnt AS [색상수], " +
                    "A.RecpDate AS [접수일자] , A.SendDate AS [발송일자] " +
                    "FROM [Bt] A, [mt_Custom] B, [mt_Article] C " +
                    "WHERE A.CustomID = B.CustomID And A.ArticleID = C.ArticleID " +
                    "AND A.BTID + dbo.fn_format(A.BTIDSeq,2) IN (SELECT BTID + dbo.fn_format(BTIDSeq,2)  FROM BTSub WHERE ConfClss IS NOT NULL) ";

                dOrderByStr = " ORDER BY [접수번호] ";
            }
            //규격코드
            else if (large == 18)
            {
                pfv.m_sCodeField = "mtrDCode";
                pfv.m_sNameField = "mtrDName";


                sql =
                    "SELECT cast( A.mtrDCode as varchar(20))  AS [코     드], mtrDName as '[규  격]' ,  cast(  (Select mtrSNAME From MT_materialsub where mtrSCode = left(A.mtrDCode,5)) + ' ' + mtrDName  as varchar(50)) AS [품명_규격], " +
                    "  cast( UNITPRICE as varchar(15))  as [단  가],   cast( '('+  isnull(cast(dbo.fn_stockQty('99999999',A.mtrDCode) as varchar(20)),'') +')' + (select max(Remark)  from mt_MaterialUnitprice  where useClss='' and mtrDCode=A.mtrDCode)  as varchar(50)) as [비 고] " +
                    " FROM mt_MaterialUnitprice  A , (Select mtrDCode as mtrDCode1 From MT_MaterialsubDetail WHERE DeleteClss='') B  WHERE  A.UseClss = '' And A.mtrDCode=B.mtrDCode1 ";

                if (!(string.IsNullOrEmpty(mmiddle)))
                    sql += "AND mtrDCode  LIKE   '" + mmiddle + "%' ";
                //txtFind.Visible = True
                //txtFind.Enabled = True
                //pnlName(2).Visible = True
            }
            // [18] 품명 2
            else if (large == 19)
            {
                pfv.m_sCodeField = "mtrSCode";
                pfv.m_sNameField = "mtrSNAME";
                sql =
                    "SELECT mtrSCode AS [코드],mtrSNAME AS [품         명] " +
                    "  FROM [mt_MATERIALSUB] " +
                    " WHERE UseClss = '' ";
                if (!(string.IsNullOrEmpty(mmiddle)))
                    sql += "AND mtrSCode  LIKE   '%" + mmiddle + "%' ";
            }
            //  LG_HSCODE = 20
            else if (large == 20)
            {
                pfv.m_sCodeField = "HsCode";
                pfv.m_sNameField = "HsCode";

                sql =
                    "SELECT  Distinct HsCode AS [코드] ,  HsCode AS [코드명] " +
                    "  FROM [cm_FTAHSCode] " +
                    " WHERE USE_YN = 'Y' ";

                if (!(string.IsNullOrEmpty(mmiddle)))
                    sql += "AND HsCode  LIKE   '%" + mmiddle + "%' ";
            }
            //LG_SDESIGN- 디자인명 , 2011.09.29, 조일 setting 시 추가함
            else if (large == 21)
            {
                pfv.m_sCodeField = "DsnName";
                pfv.m_sNameField = "DsnName";

                sql = "SELECT DISTINCT DsnName AS [코드], DsnName AS [DSN 명] " +
                    " FROM [SOrder] ";

            }

            // --------------------------------------------------------------------------------
            //  매입 매출관련 추가된것
            // --------------------------------------------------------------------------------
            else if (large == 24)      // LG_SALESCHARGE - 영업담당자, 2011.09.29, 조일 setting 시 추가함
            {
                pfv.m_sCodeField = "A.PersonID";
                pfv.m_sNameField = "A.Name";
                //Gf_DB_CM_GetComCodeList
                //If Gf_DB_CM_GetComCodeList(rs, "YNG", "Y", "") = False Then
                //    MsgBox "영업직 직무코드를 가져오는중 오류 발생"
                //    Exit Function
                //End If
                //            }

                //    lsTemp = ""

                //Do While rs.EOF = False
                //    If lsTemp = "" And rs.RecordCount <= 1 Then                         '1 건일 경우 '를 부치지 않는다.
                //        lsTemp = CheckNull(rs!CODE_ID)
                //    ElseIf lsTemp = "" And rs.RecordCount > 1 Then                      '1 건 이상의 경우에만 '를 좌우에 첨부한다.
                //        lsTemp = "'" & CheckNull(rs!CODE_ID) & "'"
                //    Else
                //        lsTemp = lsTemp & "," & "'" & CheckNull(rs!CODE_ID) & "'"       '영업담당자
                //    End If
                //    rs.MoveNext
                //Loop


                //sql = "SELECT A.PersonID AS [코드], A.Name AS [성명] " +
                //     " FROM [MT_Person] A                           " +
                //     " WHERE A.USECLSS <>'*'                        " +
                //     " AND A.DUTYID in  ( " + lsTemp + ")           ";         //직무가 영업인 직원
                dOrderByStr = " ORDER BY [성명] ";
            }
            else if (large == 25)      // LG_SALESCHARGE - 영업담당자, 2011.09.29, 조일 setting 시 추가함
            {
                pfv.m_sCodeField = "CODE_ID";
                pfv.m_sNameField = "CODE_NAME";
                sql = "SELECT CODE_ID AS [코드], CODE_NAME AS [코드명] " +
                      "FROM [CM_CODE] " +
                      "WHERE USE_YN = 'Y' " +
                      " AND  CODE_GBN = 'SIT' ";
                dOrderByStr = " ORDER BY [코드] ";

            }
            else if (large == 26) // LG_BUYITEM - 매입ITEM, 2011.09.29, 조일 setting 시 추가함
            {
                pfv.m_sCodeField = "CODE_ID";
                pfv.m_sNameField = "CODE_NAME";
                sql = "SELECT CODE_ID AS [코드], CODE_NAME AS [코드명] " +
                  "FROM [CM_CODE] " +
                  "WHERE USE_YN = 'Y' " +
                  " AND  CODE_GBN = 'BIT' ";
                dOrderByStr = " ORDER BY [코드] ";
            }
            else if (large == 27) // LG_DYEAUX = 27            '염조제ITEM
            {
                pfv.m_sCodeField = "DyeAuxID";
                pfv.m_sNameField = "DyeAux";
                sql =
                    "SELECT DyeAuxID AS [코드], DyeAux AS [염료명] " +
                    "FROM vw_mt_DyeAux  " +
                    "WHERE ISNULL(UseClss, '') NOT IN ('*') ";
                dOrderByStr = " ORDER BY [코드] ";
            }
            else if (large == 28) //Model  :거래처 모델
            {
                pfv.m_sCodeField = "ModelID";
                pfv.m_sNameField = "Model";

                sql =
                    "SELECT ModelID AS [코드], Model AS [차종명] " +
                    "FROM mt_ModelMaster   " +
                    "WHERE ISNULL(UseClss, '') NOT IN ('*') ";
                dOrderByStr = " ORDER BY [코드] ";
            }
            else if (large == 29) // Model  :LG_REQ '발주
            {
                pfv.m_sCodeField = "REQ_ID";
                pfv.m_sNameField = "Comments";
                sql =
                    "SELECT REQ_ID AS [발주번호],dor.Req_Date  as [발주일 ],  cast( Comments as varchar(40))   AS [발주내용],   cast( (select count(*) from  Dye_Order_Req_Sub dors where dor.Req_ID = dors.REQ_ID ) as varchar(5))   as [건수] " +
                    "FROM Dye_Order_Req  dor  " +
                    "WHERE ISNULL(Use_Clss, '') NOT IN ('*') ";

                if (!(string.IsNullOrEmpty(mmiddle)))
                    sql += " AND dor.Supplier_ID = '" + mmiddle + "'";

                dOrderByStr = " ORDER BY [발주번호] ";
            }
            else if (large == 30) //  Model  :LG_ACCBUYSALECODE '매입매출항목
            {
                pfv.m_sCodeField = "BSItemCode";
                pfv.m_sNameField = "BSItemName";

                sql = "SELECT BSItemCode AS [코드], BSItemName AS [코드명] " +
                      "FROM   ACC_BSITEM_CODE " +
                      "WHERE USEYN = 'Y' " +
                      " AND  ISNULL( BSITEMSCODE ,'') <> '' ";

                dOrderByStr = " ORDER BY [코드] ";

                if (!(string.IsNullOrEmpty(mmiddle)))
                    sql += "AND  BSGBN = " + mmiddle + " ";
            }
            else if (large == 31) //  거래처에 해당하는 발주
            {
                pfv.m_sCodeField = "A.REQ_ID";
                pfv.m_sNameField = "A.REQ_ID";   //20160317 AFT 요청에의한 수정.

                sql = "select A.REQ_ID + '|' + B.Seq as [발주ID], A.Req_Date as [발주일] ,  B.Item_Name as [품명] , A.Comments as [비고] " +
                      "from Dye_Order_Req A inner join Dye_Order_Req_Sub B  " +
                      "on a.Req_ID = b.Req_ID " +
                      "WHERE isnull(A.Use_Clss, '') NOT IN ('*') " +
                      "  AND Supplier_ID = '" + mmiddle + "' ";
                dOrderByStr = " ORDER BY [발주ID] ";


            }
            else if (large == 32) // 입금항목
            {
                pfv.m_sCodeField = "RPItemCode";
                pfv.m_sNameField = "RPItemName";
                pnlName[0] = "Code ID";
                pnlName[1] = "Code 명";

                sql = "SELECT RPItemCode  AS [코드], RPItemName AS [항목명] " +
              "     ,[대분류]    = (select  RPItemName from Acc_RPItem_Code ARCL  where ARC.RPItemLCode = ARCL.RPItemCode  AND ARCL.RPItemMCode =''  AND  ARCL.RPItemSCode = ''  )  " +
              "     ,[중분류]    = (select  RPItemName from Acc_RPItem_Code ARCL  where ARC.RPItemLCode + ARC.RPItemMCode = ARCL.RPItemCode  AND ARCL.RPItemLCode <>'' AND ARCL. RPItemSCode = ''  )  " +
              "FROM   Acc_RPItem_Code       ARC     " +
              "WHERE  RPGBN           =   '1'     " +
              " AND   USEYN            =   'Y'     ";
                if (mLarge == "S")//S = 소분류단계만 조회
                {
                    sql += " AND  RPItemLCode <> ''  AND  RPItemMCode <> ''  AND RPItemSCode <> ''  ";

                }
                dOrderByStr = " ORDER BY [코드] ";

            }
            else if (large == 33) // 출금항목
            {

                pfv.m_sCodeField = "RPItemCode";
                pfv.m_sNameField = "RPItemName";
                pnlName[0] = "Code ID";
                pnlName[1] = "Code 명";


                sql = "SELECT RPItemCode  AS [코드], RPItemName AS [항목명] " +
                      "     ,[대분류] = (select  RPItemName from Acc_RPItem_Code ARCL  where ARC.RPItemLCode = ARCL.RPItemCode  AND ARCL.RPItemMCode =''  AND  ARCL.RPItemSCode = ''  )  " +
                      "     ,[중분류] = (select  RPItemName from Acc_RPItem_Code ARCL  where ARC.RPItemLCode + ARC.RPItemMCode = ARCL.RPItemCode  AND ARCL.RPItemLCode <>'' AND  ARCL.RPItemSCode = ''  )   " +
                      "FROM   Acc_RPItem_Code ARC " +
                      "WHERE  RPGBN = '2' " +
                      " AND   USEYN = 'Y' ";
                if (mLarge.Equals("S"))
                {               // S = 소분류단계만 조회
                    sql += " AND  RPItemLCode <> ''  AND  RPItemMCode <> ''   AND RPItemSCode <> ''  ";
                }
                dOrderByStr = " ORDER BY [코드] ";
            }
            else if (large == 41) // LG_ApplicFTA = 41  FTA협정코드
            {
                pfv.m_sCodeField = "ApplicFTAID";
                pfv.m_sNameField = "ApplicFTAName";
                pnlName[0] = "Code ID";
                pnlName[1] = "Code 명";

                sql = "SELECT ApplicFTAID [코드]  , ApplicFTAName  [항목명] , ApplicCrierion  [원산지결정기준] , Comments   " +
                        "FROM   CM_FTAAPPlic ARC ";

                if (mLarge.Equals("S")) //S = 소분류단계만 조회
                {
                    sql += " AND  RPItemLCode <> ''  AND  RPItemMCode <> ''   AND RPItemSCode <> ''  ";
                }

                dOrderByStr = " ORDER BY [코드] ";
            }
            else if (large == 51) //  LG_ApplicFTA = 51           금형 '2016.07.21 변경
            {
                pfv.m_sCodeField = "MoldID";
                pfv.m_sNameField = "MoldNo";
                pnlName[0] = "Code ID";
                pnlName[1] = "Code 명";
                sql = "SELECT dm.MoldID  [코드] , dm.MoldNo   [금형LotNo],  cc.Code_Name [금형번호]   " +
                      "FROM   dvl_Mold  dm left outer join cm_Code cc on dm.MoldKind = cc.Code_ID and cc.Code_GBN = 'MOLDN' ";

                if (string.IsNullOrEmpty(mmiddle))
                    sql += "WHERE MoldID like '%" + mmiddle + "%' ";

                dOrderByStr = " ORDER BY  [코드] ";

            }
            else if (large == 52)       //공통.기계 부품입고거래처
            {
                pfv.m_sCodeField = "CustomID";
                pfv.m_sNameField = "KCustom";
                sql = " select  mps.CustomID AS [코드] , mc.KCustom  AS [거래처]  , StuffDate = max(mps.StuffDate) , " +
                    "   from mt_McPart      mmp             " +
                    "   ,   mc_PartStuffIN  mps             " +
                    "   ,   mt_Custom       mc              " +
                    " Where mps.MCPartID    = mmp.MCPartID  " +
                    " and   mps.CustomID    = mc.CustomID   " +
                    " and   mmp.foruse in  ( '1', '2')      ";          // ForUse :용도, 1공통, 2기계용, 3금형용

                if (string.IsNullOrEmpty(mmiddle))
                {
                    sql += " and   mps.MCPartID    = '" + mmiddle + "' ";
                    sql += " group by  mps.CustomID , mc.KCustom";
                }
                dOrderByStr = " ORDER BY [코드] ";

            }
            else if (large == 53)
            {
                pfv.m_sCodeField = "CustomID";
                pfv.m_sNameField = "KCustom";

                sql = " select  mps.CustomID AS [코드] , mc.KCustom  AS [거래처]  , StuffDate = max(mps.StuffDate)  ";
                sql += "   from mt_McPart      mmp             ";
                sql += "   ,   mc_PartStuffIN  mps             ";
                sql += "   ,   mt_Custom       mc              ";
                sql += " Where mps.MCPartID    = mmp.MCPartID  ";
                sql += " and   mps.CustomID    = mc.CustomID   ";
                sql += " and   mmp.foruse in  ( '1', '3')      ";           //ForUse :용도, 1공통, 2기계용, 3금형용

                if (string.IsNullOrEmpty(mmiddle))
                {
                    sql += " and   mps.MCPartID    = '" + mmiddle + "' ";
                }
                sql += " group by  mps.CustomID , mc.KCustom";

                dOrderByStr = " ORDER BY [코드] ";

            }
            else if (large == 54)       //계측기
            {
                pfv.m_sCodeField = "MsrMachineID";
                pfv.m_sNameField = "MsrMachineName";

                sql = " select  mm.MsrMachineID AS [코드] , mm.MsrMachineName  AS [계측기명]  , mm.MsrMachineMgrNo AS [관리번호] ";
                sql += "   from Mt_MeasureMachine      mm             ";


                dOrderByStr = " ORDER BY [코드] ";


            }
            else if (large == 55)       //국가 2016.01.21
            {
                pfv.m_sCodeField = "Code_ID";
                pfv.m_sNameField = "Code_Name";

                sql = " select  Code_ID AS [코드] , Code_Name  AS [국가] " +
                      "FROM [CM_Code] " +
                      "WHERE Code_GBN = '311' ";

                if (string.IsNullOrEmpty(mmiddle))
                {
                    sql += "and Code_ID like '%" + mmiddle + "%' ";
                }

                dOrderByStr = " ORDER BY [국가]  ";

            }
            else if (large == 56)       //공통코드 등록된 금형명 검색 2016.05.12
            {
                pfv.m_sCodeField = "Code_ID";
                pfv.m_sNameField = "Code_Name";

                sql = " select  Code_ID AS [코드] , Code_Name  AS [금형명] " +
                      "FROM [CM_Code] " +
                      "WHERE Code_GBN = 'MOLDN' ";
                if (string.IsNullOrEmpty(mmiddle))
                {
                    sql += "and Code_ID like '%" + mmiddle + "%' ";
                }

                dOrderByStr = " ORDER BY [코드]  ";
            }
            else if (large == 57)       //LG_ArticleByCustom = 57               '품명가져옴 - 거래처 기준
            {
                pfv.m_sCodeField = "ArticleID";
                pfv.m_sNameField = "Article";

                sql =
                    "SELECT ArticleID AS [코드], Article AS [품명] " +
                    "FROM [mt_Article] " +
                    "WHERE UseClss = '' ";

                if (string.IsNullOrEmpty(mmiddle))
                {
                    sql += "AND articleid in ( select ArticleID  from mt_CustomArticle  mca  where mca.Customid   = '" + mmiddle + "' ) ";
                }

                dOrderByStr = " ORDER BY [품명]  ";
            }
            else if (large == 58)       //            else if (large == 58)       //LG_ArticleByCustom = 57               '품명가져옴 - 거래처 기준
            {
                pfv.m_sCodeField = "CustomID";
                pfv.m_sNameField = "KCustom";

                sql =
                    "SELECT CustomID AS [코드], KCustom AS [상호] " +
                    "       , [상호(변경)] = replace(REPLACE(KCustom, '(주)', '') , '㈜', '')           " +
                    "       , [거래]  = case when TradeID = '1' then '1' else '9' end              " +
                    "FROM [mt_Custom] " +
                    "WHERE UseClss = '' ";

                if (string.IsNullOrEmpty(mmiddle))
                {
                    sql += "AND Customid in ( select Customid from mt_CustomArticle mca  where mca.ArticleID   = '" + mmiddle + "' ) ";
                }

                dOrderByStr = " ORDER BY [거래] ,  [상호(변경)]    ";
            }
            else if (large == 59)       //  LG_MtrNotInspect = 59         '입고후 미 검사 조회, 2017.01.24 lkm, 추가
            {
                pfv.m_sCodeField = "LOTID";
                pfv.m_sNameField = "LOTID";

                sql = "Select  StuffINID    as [입고번호]                       " +
                "           ,   StuffDate   as [입고일자]                       " +
                "           ,   LOTID       as [로트번호]                       " +
                "           ,   KCustom     as [거래처]                         " +
                "       from (                                                  " +
                "           select  si.StuffINID                                " +
                "               ,   StuffDate   = max(si.StuffDate)             " +
                "               ,   LOTID       = max(sis.LOTID)                " +
                "               ,   KCustom     = max(mc.KCustom )              " +
                "           from    [StuffINsub]  sis                           " +
                "               ,   StuffIN si                                  " +
                "               ,   mt_Custom mc                                " +
                "           Where   si.StuffINID        = sis.StuffINID         " +
                "           and     si.StuffDate        >=   dateadd(mm, -1,   GETDATE() )   " +
                "           and     si.CustomID         = mc.CustomID           " +
                "           and     si.InsStuffinYN     <> 'Y'      -- 생산에 의한 자동 재고 입고 가 아닌 건     " +
                "           group by   si.StuffINID                             " +
                "           )  si2                                              " +
                "       where not exists ( select  1                -- 검사되지 않은 건    " +
                "                           from    dbo.Ins_InspectAuto iia     " +
                "                           where   si2.LOTID = iia.LotID )     ";

                if (string.IsNullOrEmpty(mmiddle))
                {
                    sql += "AND LOTID    = '" + mmiddle + "'  ";
                }

                dOrderByStr = " ORDER BY [입고일자]  , [로트번호]                   ";
            }
            //else if (large == 59)       //  LG_MtrNotInspect = 59         '입고후 미 검사 조회, 2017.01.24 lkm, 추가

            //If IsMissing(NewData) Then  ' 찾고자하는 데이타가 없을 경우 빈 폼 뛰우기
            //    Me.Show vbModal
            //Else
            //    Call SetGrid(FL_BY_CODE, NewData) ' [1] 코드으로 찾기


            //    If grdData.Rows = 1 Then
            //        txtName = NewData
            //        Call SetGrid(FL_BY_NAME, NewData) ' [2] 명칭으로 찾기 (코드검색이 않되었을 경우)
            //    End If

            //    With grdData
            //        If.Rows > .FixedRows Then
            //            If.Rows = .FixedRows + 1 Then
            //               Call SelectData
            //           Else
            //                Me.Show vbModal
            //            End If
            //        End If
            //    End With
            //End If


            //'=================================================================================================='
            //If m_bSelected Then
            //    With grdData
            //        ReDim SelData(.Cols - 1)
            //        For i = 0 To.Cols - 1
            //            SelData(i) = wData(i)
            //        Next i
            //    End With
            //End If


            //SetMsg = m_bSelected

        }
    }
}
