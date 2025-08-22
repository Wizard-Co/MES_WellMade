namespace WizMes_WellMade
{
    enum Defind_CodeFind
    {
        DCF_CUSTOM = 0,
        DCF_Article = 1,
        DCF_PERSON = 2,
        DCF_DEFECT = 3,
        DCF_ORDER = 4,
        DCF_DYE = 5,
        DCF_AUX = 6,
        DCF_WORK = 7,
        DCF_THREAD = 8,
        DCF_STUFFWIDTH = 9,
        DCF_PROCESS = 10,

        DCF_PLANREMARK = 11,
        DCF_MC = 12,
        DCF_PART = 13,        //'설비부품
        DCF_CHECKITEM = 14,
        DCF_CHECKMETHOD = 15,
        DCF_CHECKSTD = 16,
        DCF_BT = 17,
        DCF_articleSize = 18,
        DCF_article2 = 19,
        DCF_HSCODE = 20,

        DCF_21 = 21,
        DCF_22 = 22,
        DCF_SDESIGN = 23,
        //아래 추가된 값에대한 코드 검색 부분 PlusFind2 프로젝트의 frmFind.frm( 검색화면) 내용에 추가 할것
        DCF_SalesCharge = 24,    // '영업사원 S_201101_대안_02에 의한 추가
        DCF_SALESITEM = 25,      // '매출ITEM S_201101_대안_029에 의한 추가
        DCF_BUYITEM = 26,        //'매입ITEM S_201101_대안_02에 의한 추가
                                 //DCF_Depart = 12
        DCF_DYEAUX = 27,         //'염조제ITEM S_201105_대안_03에 의한 추가
                                 //'    DCF_Materials = 11
        DCF_BUYERMODEL = 28,    //        '고객모델,AFT 적용시 추가함
        DCF_REQ = 29,            //       '발주번호,AFT 적용시 추가함'DCF_ACCBUYSALECODE
        DCF_ACCBUYSALECODE = 30,    //    '2014.06.05, lkm, 추가

        DCF_CUSTOMREQ = 31,
        DCF_ACCRCVITEM = 32,    //'2014.10.28, lkm, 추가
        DCF_ACCPAYITEM = 33,  //'2014.10.28, lkm, 추가
        DCF_34 = 34,
        DCF_35 = 35,
        DCF_36 = 36,
        DCF_37 = 37,
        DCF_38 = 38,
        DCF_39 = 39,
        DCF_40 = 40,

        DCF_ApplicFTA = 41,     //'FTA협정코드, 2015.05.01, 추가
        DCF_42 = 42,
        DCF_43 = 43,
        DCF_44 = 44,
        DCF_45 = 45,
        DCF_46 = 46,
        DCF_47 = 47,
        DCF_48 = 48,
        DCF_49 = 49,
        DCF_50 = 50,

        DCF_MOLD = 51,       //'금형번호   , 2015.08.11, lkm, 추가
        DCF_MCPARTINCUSTOM = 52, //'설비부품 입고처
        DCF_MOLDPARTINCUSTOM = 53, //'금형부품 입고처
        DCF_QULMSRMACHINE = 54,  //      '계측기 , 2015.08.12 추가
        DCF_COUNTRY = 55,          //     '국가', 2016.01.21 추가
        LG_MOLDN = 56,      //      '금형명', 2016.05.12 추가
        LG_ArticleByGroup = 57,  //'s_201706_AFT_05 에 의한 추가

        DCF_Hogi = 66, //  누가 언제넣은거지???  // 19_0822 허윤구 확인.
        DCF_InspectGage = 75, //  2020.1.10 HYG 삼주테크 추가

        DCF_BuyerArticleNo = 76, // 지엘에스 품번 추가

        DCF_95 = 95 //2021-11-10 GLS 자재입고반품 LOTID PLUSFINDER

    }
}
