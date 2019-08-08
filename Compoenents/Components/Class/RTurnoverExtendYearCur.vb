Imports System.Text

Public Enum ProductTypeEnum
    FinishedGood = 1
    Components = 2
End Enum

Public Enum ReportTypeEnum
    ALLData = 1
    WMF = 2
    SEBAsia = 3
End Enum
Public Enum TurnoverHistoryActionEnum
    DoNotSavePeriod = 1
    SavePeriod = 2
    DeletePeriod = 3
End Enum
Public Class RTurnoverExtendYearCur
    Implements IDisposable

    Public dbAdapter1 As DbAdapter = DbAdapter.getInstance
    Public errorMessage As String = String.Empty


    Public ProductType As ProductTypeEnum
    Private ReportType As ReportTypeEnum
    Private TurnoverHistoryAction As TurnoverHistoryActionEnum
    Private LastPeriod As Date
    Public BaseItem As String
    Private HistoryDate As Date
    Public myFirstDate As Date
    Public myLastDate As Date

    Public Sub New()
        ProductType = ProductTypeEnum.FinishedGood
        ReportType = ReportTypeEnum.ALLData
    End Sub

    Public Sub New(ByVal productType As ProductTypeEnum, ByVal reportType As ReportTypeEnum, ByVal turnoverHistoryAction As TurnoverHistoryActionEnum, ByVal LastPeriod As Date, ByVal BaseItem As String, ByVal HistoryDate As Date)
        Me.ProductType = productType
        Me.ReportType = reportType
        Me.TurnoverHistoryAction = turnoverHistoryAction
        Me.LastPeriod = LastPeriod
        Me.BaseItem = BaseItem
        Me.HistoryDate = HistoryDate

        If LastPeriod.Month = 12 Then
            myFirstDate = CDate(String.Format("{0}-1-1", LastPeriod.Year))
        Else
            myFirstDate = CDate(String.Format("{0}-1-1", LastPeriod.Year - 1))
        End If

        Dim DateTemp = LastPeriod.AddMonths(1)
        myLastDate = String.Format("{0}-{1}-{2}", DateTemp.Year, DateTemp.Month, Date.DaysInMonth(DateTemp.Year, DateTemp.Month))
    End Sub

    Public Function loadCombobox(ByRef ds As DataSet) As Boolean
        Dim Sqlstr As String = String.Format("select distinct period from turnoverhistory where groupsbuid = {0:d} order by period desc", ProductTypeEnum.FinishedGood)
        Return dbAdapter1.TbgetDataSet(Sqlstr, ds, errorMessage)
    End Function

    Public Function getQueryData() As String
        Dim sb As New StringBuilder
        Dim ReportTypeCriteria As String = String.Empty
        Dim ReportTypeCriteriaForecast As String = String.Empty
        Select Case ReportType
            Case ReportTypeEnum.ALLData
            Case ReportTypeEnum.SEBAsia
                ReportTypeCriteria = " and pg.groupsbuidpg <> 10"
                ReportTypeCriteriaForecast = "where b.brandtype isnull  and (mm.cmmf <= 3200000000 or (mm.cmmf >= 3300000000 and mm.cmmf <= 8400000000 ) or mm.cmmf >= 8500000000)"
            Case ReportTypeEnum.WMF
                ReportTypeCriteria = " and pg.groupsbuidpg = 10"
                ReportTypeCriteriaForecast = "where b.brandtype = 1 and ((mm.cmmf >= 3200000000 and mm.cmmf < 3300000000 ) or (mm.cmmf >= 8400000000 and mm.cmmf < 8500000000 ))"
        End Select
        'FirmOrder
        sb.Append(String.Format("(SELECT 'FIRMORDER'::character varying as ""TYPE"", so.soldtoparty  AS ""Sold To Party SAP Code"", c2.customername AS ""Sold To Party Name"",pd.shiptoparty as ""Ship To Party"",c1.customername as ""Ship To Party Name"", e.vendorcode AS ""SAP Vendor Code"",v.vendorname AS ""Vendor Name"", mm.rri AS ""Research and Industry Responsible"", sbu.sbuname AS ""SBU"", sbu1.sbuname AS ""BU"", sbu2.sbuname AS familysbu, pd.cmmf AS ""CMMF"",mm.familylv1 AS ""Commercial Family CG2 n"", f.familyname, b.brandname AS ""Brand Name CG2 n"", mm.commref AS ""CommercialRef"", mm.materialdesc AS ""Description""," & _
               " (getpurchasevalue(pd.osqty,pd.fob) / getexchangerate(e.currency)) AS ""Purchase Value"",pd.osqty,getyearinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{0:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""YEAR_INQCONF"",(getyearpurchasevalue({2:yyyy}::integer,od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob)/getexchangerate(e.currency) )as ""Year{2:yyyy} Purchase Value""," & _
               " (getyearpurchasevalue({0:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob) / getexchangerate(e.currency)) as ""Year{0:yyyy} Purchase Value"" ,Null::numeric as ""Year{1} Purchase Value"" ,getyearqty({2:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty) as ""Year{2:yyyy} QTY"",getyearqty({0:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty) as ""Year{0:yyyy} QTY"",Null::numeric as ""Year{1} QTY"" ,getmonthinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{0:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""MONTH INQCONF"",ssm.officersebname as ""SPM"" " & _
               " ,adjustedmonth(confirmationstatus,od.currentconfirmedetd,pd.comments) as adjustedmonth, od.currentconfirmedetd,sd.currentinquiryetd,getconfirmstatus(confs.confirmationstatus,pd.comments,od.currentconfirmedetd) as confirmstatus,pd.comments,pd.sebasiapono,pd.polineno,sp.sopdescription as industrialfamily" & _
               " ,e.currency,getpurchasevalue(pd.osqty,pd.fob) as ""Purchase Value (OC)"" , '1'::integer as isreal" & _
               " FROM cxsebodtp od LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" & _
               " LEFT JOIN cxsalesorder so ON so.sebasiasalesorder = sd.sebasiasalesorder" & _
               " LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid" & _
               " LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono" & _
               " LEFT JOIN customer c1 ON c1.customercode = pd.shiptoparty" & _
               " LEFT JOIN customer c2 ON c2.customercode = so.soldtoparty" & _
               " LEFT JOIN cxpoconf pc ON pc.sebasiapono = pd.sebasiapono AND pc.polineno = pd.polineno" & _
               " LEFT JOIN cxpoconfother pco ON pco.sebasiapono = pd.sebasiapono AND pco.polineno = pd.polineno" & _
               " LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid LEFT JOIN cxconfstatus confs ON confs.cxconfid = c.cxconfid Left join ekko e on e.po = pd.sebasiapono Left join vendor v on v.vendorcode = e.vendorcode LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl LEFT JOIN materialmaster mm  ON mm.cmmf = pd.cmmf left join sspcmmfsop sc on sc.cmmf = mm.cmmf" & _
               " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid" & _
               " LEFT JOIN brand b ON b.brandid = mm.brandid" & _
               " LEFT JOIN activity ac ON ac.activitycode = mm.rri" & _
               " LEFT JOIN sbu ON sbu.sbuid = ac.sbuidsp" & _
               " LEFT JOIN sbu sbu1 ON sbu1.sbuid = ac.sbuidlg" & _
               " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
               " LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu" & _
               " LEFT JOIN purchasinggroup pg ON pg.purchasinggroup = e.purchasinggroup::bpchar" & _
               " WHERE od.ordertype = 'Header'::bpchar AND ph.receptiondate <= sd.inquiryetd and pd.osqty > 0  and pg.groupsbuid = {3:d} {4} ORDER BY ph.receptiondate )", myLastDate, myLastDate.Year + 1, myFirstDate, ProductType, ReportTypeCriteria))

        'Forecast
        If ProductType = ProductTypeEnum.FinishedGood Then
            sb.Append(" union all ")
            'sb.Append(String.Format("(with validcmmf as (select distinct cmmf from forecastestimation )," &
            ' " lp as (select distinct cmmf,vendorcode,first_value(pricelistid) over(partition by cmmf,vendorcode order by cmmf,vendorcode,validfrom desc) as pricelistid from pricelist" & _
            ' " where cmmf in (select cmmf from  validcmmf) order by cmmf)," & _
            ' " price as(select lp.cmmf,lp.validfrom,c.rangeid,c.comfam,(pl.amount / pl.perunit) * getexchangerate(pl.currency) as newamount from latestpricelist lp" & _
            ' " LEFT JOIN pricelist pl ON pl.cmmf = lp.cmmf AND pl.validfrom = lp.validfrom LEFT JOIN cmmf c on c.cmmf = lp.cmmf)," & _
            ' " pr as(select rangeid,avg(newamount) as newamount from price group by rangeid order by rangeid)," & _
            ' " pf as(select comfam,avg(newamount) as newamount from price group by comfam order by comfam)" &
            '    " SELECT 'FORECAST'::character varying,cv.sassebasia AS sapcustomercode, c1.customername AS sapcustomername, cv.sassebasia AS sapcustomercode2, c1.customername AS sapcustomername2, fe.vendorcode AS ""SEB Asia SAP Vendor Code"", v2.vendorname AS ""SEB ASIA SAP Vendor Name"",mm.rri as ""Research and Industry Responsible"",  sbu.sbuname, sbu1.sbuname AS bu, sbu2.sbuname AS familysbu,fe.cmmf,mm.familylv1, f.familyname,b.brandname as ""Brand Name CG2 n"",mm.commref,mm.materialdesc,getforecastvalue(fe.qty,pl.amount,pl.perunit,pl.currency,pr.newamount::numeric,pf.newamount::numeric) " & _
            '    " AS value,qty, wm.myyear,getyearpurchasevalue({2:yyyy},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit,pl.currency,pr.newamount::numeric,pf.newamount::numeric)),getyearpurchasevalue({0:yyyy},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit,pl.currency,pr.newamount::numeric,pf.newamount::numeric)),getyearpurchasevalue({1},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit,pl.currency,pr.newamount::numeric,pf.newamount::numeric)) ,getyearqty({2:yyyy},wm.myyear,qty),getyearqty({0:yyyy},wm.myyear,qty),getyearqty({1},wm.myyear,qty), " & _
            '    " wm.mymonth, ssm.officersebname as ""SSM""," &
            '    " null::integer,null::date,null::date,null::character varying,null::character varying,null::bigint,null::integer,sp.sopdescription" & _
            '    " ,getcurrency(pl.currency) as currency,getforecastvalue(fe.qty,pl.amount,pl.perunit,'USD',pr.newamount::numeric,pf.newamount::numeric) as ""Purchase Value (OC)"", isrealprice(pl.currency)::integer FROM forecastestimation fe" & _
            '    " LEFT JOIN customer c ON c.customercode = fe.customercode LEFT JOIN materialmaster mm ON mm.cmmf = fe.cmmf" & _
            '    " LEFT JOIN cmmf on cmmf.cmmf = fe.cmmf left join sspcmmfsop sc on sc.cmmf = fe.cmmf" & _
            '    " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid LEFT JOIN brand b ON b.brandid = mm.brandid" & _
            '    " LEFT JOIN vendor v ON v.vendorcode = fe.vendorcode LEFT JOIN vendor v2 ON v2.vendorcode = fe.vendorcode" & _
            '    " LEFT JOIN officerseb ssm ON ssm.ofsebid = v2.ssmidpl" & _
            '    " LEFT JOIN weektomonth wm ON wm.yearweek = fe.weeketa" & _
            '    " LEFT JOIN activity act ON act.activitycode = mm.rri" & _
            '    " LEFT JOIN sbu ON sbu.sbuid = act.sbuid" & _
            '    " LEFT JOIN sbu sbu1 ON sbu1.sbuid = act.sbuidlg" & _
            '    " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
            '    " LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu" & _
            '    " LEFT JOIN lp ON lp.cmmf = fe.cmmf and lp.vendorcode = fe.vendorcode LEFT JOIN pricelist pl ON pl.pricelistid = lp.pricelistid" & _
            '    " LEFT JOIN pr on pr.rangeid = cmmf.rangeid LEFT JOIN pf on pf.comfam = cmmf.comfam" & _
            '    " LEFT JOIN feexposure fex ON fex.feid = fe.feid" & _
            '    " LEFT JOIN convcustsas cv ON cv.customercode = fe.customercode" & _
            '    " LEFT JOIN customer c1 ON c1.customercode = cv.sassebasia" & _
            '    " ORDER BY fe.feid)", myLastDate, myLastDate.Year + 1, myFirstDate))
            sb.Append(String.Format("(with validcmmf as (select distinct cmmf from forecastestimation )," &
             " lp as (select distinct cmmf,vendorcode,first_value(pricelistid) over(partition by cmmf,vendorcode order by cmmf,vendorcode,validfrom desc) as pricelistid from pricelist" & _
             " where cmmf in (select cmmf from  validcmmf) order by cmmf)," & _
             " price as(select pl.cmmf,pl.validfrom,c.rangeid,c.comfam,(pl.amount / pl.perunit) / getexchangerate(pl.currency) as newamount from lp" & _
             " LEFT JOIN pricelist pl ON pl.pricelistid = lp.pricelistid LEFT JOIN cmmf c on c.cmmf = lp.cmmf)," & _
             " pr as(select rangeid,avg(newamount) as newamount from price group by rangeid order by rangeid)," & _
             " pf as(select comfam,avg(newamount) as newamount from price group by comfam order by comfam)" &
                " SELECT 'FORECAST'::character varying,cv.sassebasia AS sapcustomercode, c1.customername AS sapcustomername, cv.sassebasia AS sapcustomercode2, c1.customername AS sapcustomername2, fe.vendorcode AS ""SEB Asia SAP Vendor Code"", v2.vendorname AS ""SEB ASIA SAP Vendor Name"",mm.rri as ""Research and Industry Responsible"",  sbu.sbuname, sbu1.sbuname AS bu, sbu2.sbuname AS familysbu,fe.cmmf,mm.familylv1, f.familyname,b.brandname as ""Brand Name CG2 n"",mm.commref,mm.materialdesc,getforecastvalue(fe.qty,pl.amount,pl.perunit,pl.currency,pr.newamount::numeric,pf.newamount::numeric) " & _
                " AS value,qty, wm.myyear,getyearpurchasevalue({2:yyyy},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit,pl.currency,pr.newamount::numeric,pf.newamount::numeric)),getyearpurchasevalue({0:yyyy},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit,pl.currency,pr.newamount::numeric,pf.newamount::numeric)),getyearpurchasevalue({1},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit,pl.currency,pr.newamount::numeric,pf.newamount::numeric)) ,getyearqty({2:yyyy},wm.myyear,qty),getyearqty({0:yyyy},wm.myyear,qty),getyearqty({1},wm.myyear,qty), " & _
                " wm.mymonth, ssm.officersebname as ""SSM""," &
                " null::integer,null::date,null::date,null::character varying,null::character varying,null::bigint,null::integer,sp.sopdescription" & _
                " ,getcurrency(pl.currency) as currency,getforecastvalue(fe.qty,pl.amount,pl.perunit,'USD',pr.newamount::numeric,pf.newamount::numeric) as ""Purchase Value (OC)"", isrealprice(pl.currency)::integer FROM forecastestimation fe" & _
                " LEFT JOIN customer c ON c.customercode = fe.customercode LEFT JOIN materialmaster mm ON mm.cmmf = fe.cmmf" & _
                " LEFT JOIN cmmf on cmmf.cmmf = fe.cmmf left join sspcmmfvendorsop sc on sc.cmmf = fe.cmmf and sc.vendorcode = fe.vendorcode" & _
                " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid LEFT JOIN brand b ON b.brandid = mm.brandid" & _
                " LEFT JOIN vendor v ON v.vendorcode = fe.vendorcode LEFT JOIN vendor v2 ON v2.vendorcode = fe.vendorcode" & _
                " LEFT JOIN officerseb ssm ON ssm.ofsebid = v2.ssmidpl" & _
                " LEFT JOIN weektomonth wm ON wm.yearweek = fe.weeketa" & _
                " LEFT JOIN activity act ON act.activitycode = mm.rri" & _
                " LEFT JOIN sbu ON sbu.sbuid = act.sbuid" & _
                " LEFT JOIN sbu sbu1 ON sbu1.sbuid = act.sbuidlg" & _
                " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
                " LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu" & _
                " LEFT JOIN lp ON lp.cmmf = fe.cmmf and lp.vendorcode = fe.vendorcode LEFT JOIN pricelist pl ON pl.pricelistid = lp.pricelistid" & _
                " LEFT JOIN pr on pr.rangeid = cmmf.rangeid LEFT JOIN pf on pf.comfam = cmmf.comfam" & _
                " LEFT JOIN feexposure fex ON fex.feid = fe.feid" & _
                " LEFT JOIN convcustsas cv ON cv.customercode = fe.customercode" & _
                " LEFT JOIN customer c1 ON c1.customercode = cv.sassebasia " & _
                " {3} " &
                " ORDER BY fe.feid)", myLastDate, myLastDate.Year + 1, myFirstDate, ReportTypeCriteriaForecast))
        Else
            sb.Append(" union all ")
            sb.Append(String.Format("(with validcmmf as (select distinct cmmf from forecastestimationcomp )," &
             " lp as (select distinct cmmf,vendorcode,first_value(pricelistid) over(partition by cmmf,vendorcode order by cmmf,vendorcode,validfrom desc) as pricelistid from pricelist" & _
             " where cmmf in (select cmmf from  validcmmf) order by cmmf)," & _
             " price as(select lp.cmmf,pl.validfrom,c.rangeid,c.comfam,(pl.amount / pl.perunit) / getexchangerate(pl.currency) as newamount from lp" & _
             " LEFT JOIN pricelist pl ON pl.pricelistid = lp.pricelistid LEFT JOIN cmmf c on c.cmmf = lp.cmmf)," & _
             " pr as(select rangeid,avg(newamount) as newamount from price group by rangeid order by rangeid)," & _
             " pf as(select comfam,avg(newamount) as newamount from price group by comfam order by comfam) SELECT 'FORECAST'::character varying,fe.customercode AS sapcustomercode, c1.customername AS sapcustomername,fe.customercode AS sapcustomercode2, c1.customername AS sapcustomername2, fe.vendorcode AS ""SEB Asia SAP Vendor Code"", v.vendorname AS ""SEB ASIA SAP Vendor Name"",mm.rri as ""Research and Industry Responsible"",  sbu.sbuname, sbu1.sbuname AS bu, sbu2.sbuname AS familysbu,fe.cmmf,mm.familylv1, f.familyname,b.brandname as ""Brand Name CG2 n"",mm.commref,mm.materialdesc,getforecastvalue(fe.qty,pl.amount,pl.perunit,pl.currency,pr.newamount::numeric,pf.newamount::numeric) " & _
                " AS value,qty, wm.myyear,getyearpurchasevalue({2:yyyy},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit,pl.currency,pr.newamount::numeric,pf.newamount::numeric)),getyearpurchasevalue({0:yyyy},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit,pl.currency,pr.newamount::numeric,pf.newamount::numeric)),getyearpurchasevalue({1},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit,pl.currency,pr.newamount::numeric,pf.newamount::numeric)) ,getyearqty({2:yyyy},wm.myyear,qty),getyearqty({0:yyyy},wm.myyear,qty),getyearqty({1},wm.myyear,qty), " & _
                " wm.mymonth, ssm.officersebname as ""SSM""," & _
                " null::integer,null::date,null::date,null::character varying,null::character varying,null::bigint,null::integer,sp.sopdescription" & _
                "  ,getcurrency(pl.currency)  as currency,getforecastvalue(fe.qty,pl.amount,pl.perunit,'USD',pr.newamount::numeric,pf.newamount::numeric) as ""Purchase Value (OC)"", isrealprice(pl.currency)::integer FROM forecastestimationcomp fe" & _
                " LEFT JOIN customer c ON c.customercode = fe.customercode LEFT JOIN materialmaster mm on mm.cmmf = fe.cmmf" & _
                " LEFT JOIN cmmf on cmmf.cmmf = fe.cmmf left join sspcmmfsop sc on sc.cmmf = fe.cmmf" & _
                " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid LEFT JOIN brand b ON b.brandid = mm.brandid" & _
                " LEFT JOIN vendor v ON v.vendorcode = fe.vendorcode LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl" & _
                " LEFT JOIN weektomonth wm ON wm.yearweek = fe.weeketa LEFT JOIN activity act ON act.activitycode = mm.rri" & _
                " LEFT JOIN sbu ON sbu.sbuid = act.sbuid" & _
                " LEFT JOIN sbu sbu1 ON sbu1.sbuid = act.sbuidlg" & _
                " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
                " LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu" & _
                 " LEFT JOIN lp ON lp.cmmf = fe.cmmf and lp.vendorcode = fe.vendorcode LEFT JOIN pricelist pl ON pl.pricelistid = lp.pricelistid" & _
                " LEFT JOIN pr on pr.rangeid = cmmf.rangeid LEFT JOIN pf on pf.comfam = cmmf.comfam" & _
                " LEFT JOIN feexposure fex ON fex.feid = fe.feid" & _
                " LEFT JOIN customer c1 ON c1.customercode = fe.customercode" & _
                " {3}" &
                " ORDER BY fe.feid)", myLastDate, myLastDate.Year + 1, myFirstDate, ReportTypeCriteriaForecast))
        End If
        'Shipment
        sb.Append(" union all ")
        sb.Append(String.Format("(select 'SHIPMENT'::character varying as type,so.soldtoparty,cu1.customername, cxpo.shiptoparty,cu2.customername, m.vendorcode,v.vendorname,mm.rri,sbu.sbuname,sbu1.sbuname,sbu2.sbuname as familysbu,d.cmmf,mm.familylv1,f.familyname,b.brandname,mm.commref,mm.materialdesc,(p.amount / getexchangerate(e.currency))as amount,p.qty::integer,date_part('Year',miropostingdate)::integer,(getyearpurchasevalue({1:yyyy},miropostingdate,p.amount)/getexchangerate(e.currency)) as lastyear,(getyearpurchasevalue({0:yyyy},miropostingdate,p.amount)/getexchangerate(e.currency)) as thisyear,null::numeric,getyearqty({1:yyyy},miropostingdate,p.qty::integer),getyearqty({0:yyyy},miropostingdate,p.qty::integer),null::numeric,date_part('Month',miropostingdate),ssm.officersebname," & _
                " null::integer,null::date,null::date,null::character varying,null::character varying,hd.pohd,d.polineno,sp.sopdescription," & _
                " e.currency, p.amount, '1'::integer as isreal " & _
                " from pomiro p left join miro m on m.miroid = p.miroid left join podtl d on d.podtlid = p.podtlid left join vendor v on v.vendorcode = m.vendorcode" & _
                " left join cxsebpodtl cxpo on cxpo.sebasiapono = d.pohd and cxpo.polineno = d.polineno left join ekko e on e.po = d.pohd" & _
                " left join pohd hd on hd.pohd = d.pohd left join materialmaster mm on mm.cmmf = d.cmmf" & _
                " left join sspcmmfsop sc on sc.cmmf = d.cmmf left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid" & _
                " LEFT JOIN brand b ON b.brandid = mm.brandid" & _
                " LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl" & _
                " left join cxsebpodtl pd on pd.sebasiapono = d.pohd and pd.polineno = d.polineno " & _
                " left join cxrelsalesdocpo r on r.cxsebpodtlid = pd.cxsebpodtlid" & _
                " left join cxsalesorderdtl sd on sd.cxsalesorderdtlid = r.cxsalesorderdtlid" & _
                " left join cxsalesorder so on so.sebasiasalesorder = sd.sebasiasalesorder" & _
                " left join customer cu1 on cu1.customercode = so.soldtoparty" & _
                " left join customer cu2 on cu2.customercode = cxpo.shiptoparty" & _
                " LEFT JOIN activity ac ON  ac.activitycode = mm.rri" & _
                " LEFT JOIN sbu ON sbu.sbuid = ac.sbuid" & _
                " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
                " LEFT JOIN sbu sbu1 on sbu1.sbuid = ac.sbuidlg" & _
                " LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu Left Join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup" & _
                " where miropostingdate >= '{1:yyyy-MM-dd}' and miropostingdate <= '{0:yyyy-MM-dd}'" &
                " and pg.groupsbuid = {2:d} {3} order by p.pomiroid)", myLastDate, myFirstDate, ProductType, ReportTypeCriteria))
        'Select Case ProductType
        '    Case ProductTypeEnum.FinishedGood

        '    Case ProductTypeEnum.Components

        'End Select
        Return sb.ToString
    End Function
    Public Function getQueryDataSummary() As String
        Dim sqlstr As String
        If ProductType = ProductTypeEnum.FinishedGood Then
            sqlstr = String.Format("select * from turnoverhistoryviewfg where period >= '{0}'" & " and groupsbuid = {1:d}", BaseItem, ProductType)
        Else
            sqlstr = String.Format("select * from turnoverhistoryviewcomp where period >= '{0}'", BaseItem)
        End If
        Return sqlstr
    End Function
    Public Function DeletePeriod() As Boolean
        Dim myret As Boolean
        Dim Sqlstr = String.Format("delete from turnoverhistory where period = '{0:yyyyMMdd}' and groupsbuid = {1:d}", HistoryDate, ProductType)
        myret = dbAdapter1.ExecuteNonQuery(Sqlstr)
        Return myret
    End Function

    Sub AddNewPeriod()
        Dim sb As StringBuilder = New StringBuilder
        Dim myFirstDate = CDate(String.Format("{0}-1-1", Year(myLastDate)))
        Dim mypg As Integer = ProductType
        Dim period As String = String.Format("'{0:yyyyMMdd}'", HistoryDate)

        If ProductType = ProductTypeEnum.FinishedGood Then

            Dim additionalquery = "with validcmmf as (select distinct cmmf from forecastestimation )," & _
                         " lp as (select distinct cmmf,vendorcode,first_value(pricelistid) over(partition by cmmf,vendorcode order by cmmf,vendorcode,validfrom desc) as pricelistid from pricelist" & _
                         " where cmmf in (select cmmf from  validcmmf) order by cmmf) "

            'Firmordergroup
            sb.Append(String.Format("(SELECT {0:yyyy} as myyear,e.vendorcode,null as sbuid,2 as datatype,{1}::character varying as period,getmonthinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{0:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as monthinqconf,(sum(getyearpurchasevalue({0:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob))::real / getexchangerate(e.currency)) as purchasevalue,sum(getyearqty(" & Year(myLastDate) & ", od.currentconfirmedetd, sd.currentinquiryetd, confs.confirmationstatus, PD.Comments, PD.osqty)) As qty" &
                        ",sum(getyearpurchasevalue({0:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob))::real as purchasevalueoc ,e.currency,mm.sbu" &
                        " FROM cxsebodtp od LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid" & _
                        " LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
                        " LEFT JOIN cxsalesorder so ON so.sebasiasalesorder = sd.sebasiasalesorder" & _
                        " LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid" & _
                        " LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono" & _
                        " Left join ekko e on e.po = pd.sebasiapono" & _
                        " Left join vendor v on v.vendorcode = e.vendorcode" & _
                        " LEFT JOIN materialmaster mm ON mm.cmmf = pd.cmmf" & _
                        " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
                        " LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid" & _
                        " LEFT JOIN cxconfstatus confs ON confs.cxconfid = c.cxconfid" & _
                        " left join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup" & _
                        " WHERE od.ordertype = 'Header'::bpchar AND ph.receptiondate <= sd.inquiryetd and pd.osqty > 0 and pg.groupsbuid = {2:d} group by e.vendorcode," & _
                        " mm.sbu,datatype,period,monthinqconf,e.currency)", myLastDate, period, ProductType))
            sb.Append(" union all ")
            'FirmordergroupExtend
            sb.Append(String.Format("(SELECT {3}, e.vendorcode,null as sbuid,2 as datatype,{1}::character varying as period,getmonthinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{0:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as monthinqconf,(sum(getyearpurchasevalue({3},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob))::real / getexchangerate(e.currency)) as purchasevalue,sum(getyearqty({3}, od.currentconfirmedetd, sd.currentinquiryetd, confs.confirmationstatus, PD.Comments, PD.osqty)) As qty" & _
                        ",sum(getyearpurchasevalue({3},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob))::real as purchasevalueoc ,e.currency,mm.sbu" & _
                        " FROM cxsebodtp od LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid" & _
                        " LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" & _
                        " LEFT JOIN cxsalesorder so ON so.sebasiasalesorder = sd.sebasiasalesorder" & _
                        " LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid" & _
                        " LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono" & _
                        " Left join ekko e on e.po = pd.sebasiapono" & _
                        " Left join vendor v on v.vendorcode = e.vendorcode" & _
                        " LEFT JOIN materialmaster mm ON mm.cmmf = pd.cmmf" & _
                        " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
                        " LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid" & _
                        " LEFT JOIN cxconfstatus confs ON confs.cxconfid = c.cxconfid" & _
                        " left join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup" & _
                        " WHERE od.ordertype = 'Header'::bpchar AND ph.receptiondate <= sd.inquiryetd and pd.osqty > 0 and pg.groupsbuid = {2:d} group by e.vendorcode," & _
                        " mm.sbu,datatype,period,monthinqconf,e.currency having sum(getyearqty({3}, od.currentconfirmedetd, sd.currentinquiryetd, confs.confirmationstatus, PD.Comments, PD.osqty)) <> 0)", myLastDate, period, ProductType, Year(myLastDate) + 1))

            sb.Append(" union all ")
            'ForecastGroup
            sb.Append(String.Format("({0} SELECT {1:yyyy},fe.vendorcode,null as sbuid,1 as datatype,{2}::character varying as period,wm.mymonth,sum(getyearpurchasevalue({1:yyyy},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit)/getexchangerate(pl.currency))),sum(getyearqty({1:yyyy},wm.myyear,qty))" & _
                        " ,sum(getyearpurchasevalue({1:yyyy},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit)))::real as purchasevalueoc,pl.currency,mm.sbu" & _
                        " FROM forecastestimation fe LEFT JOIN customer c ON c.customercode = fe.customercode" & _
                        " LEFT JOIN materialmaster mm ON mm.cmmf = fe.cmmf" & _
                        " LEFT JOIN vendor v2 ON v2.vendorcode = fe.vendorcode" & _
                        " LEFT JOIN weektomonth wm ON wm.yearweek = fe.weeketa" & _
                        " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
                        " LEFT JOIN lp ON lp.cmmf = fe.cmmf and lp.vendorcode = v2.vendorcode" & _
                        " LEFT JOIN pricelist pl ON pl.pricelistid = lp.pricelistid " & _
                        " group by fe.vendorcode,mm.sbu,datatype,period,wm.mymonth,pl.currency)", additionalquery, myLastDate, period))
            sb.Append(" union all ")
            'ForecastGroupExtend
            sb.Append(String.Format("({0} SELECT {3}, fe.vendorcode,null as sbuid,1 as datatype,{2}::character varying as period,wm.mymonth,sum(getyearpurchasevalue({3},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit)/getexchangerate(pl.currency))),sum(getyearqty({3},wm.myyear,qty))" & _
                        " ,sum(getyearpurchasevalue({3},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit)))::real as purchasevalueoc,pl.currency,mm.sbu" & _
                        " FROM forecastestimation fe LEFT JOIN customer c ON c.customercode = fe.customercode" & _
                        " LEFT JOIN materialmaster mm ON mm.cmmf = fe.cmmf" & _
                        " LEFT JOIN vendor v2 ON v2.vendorcode = fe.vendorcode" & _
                        " LEFT JOIN weektomonth wm ON wm.yearweek = fe.weeketa" & _
                        " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
                        " LEFT JOIN lp ON lp.cmmf = fe.cmmf and lp.vendorcode = v2.vendorcode" & _
                        " LEFT JOIN pricelist pl ON pl.pricelistid = lp.pricelistid " & _
                        " group by fe.vendorcode,mm.sbu,datatype,period,wm.mymonth,pl.currency having sum(getyearqty({3},wm.myyear,qty))<>0)", additionalquery, myLastDate, period, Year(myLastDate) + 1))
            sb.Append(" union all ")
            'ShipmentGroup
            sb.Append(String.Format("(select {1:yyyy}, m.vendorcode ,null as sbuid,3 as datatype, {2}::character varying as period,date_part('Month',miropostingdate) as monthinq ,(sum(getyearpurchasevalue({1:yyyy},miropostingdate,p.amount))/getexchangerate(e.currency)) as purchase,sum(getyearqty({1:yyyy},miropostingdate,p.qty::integer)) as qty" & _
                        " ,sum(getyearpurchasevalue(" & Year(myLastDate) & ",miropostingdate,p.amount))::real as purchasevalueoc,e.currency,mm.sbu" & _
                        " from miro m" & _
                        " left join pomiro p on p.miroid = m.miroid" & _
                        " left join podtl d on d.podtlid = p.podtlid" & _
                        " left join ekko e on e.po = d.pohd" & _
                        " left join pohd hd on hd.pohd = d.pohd " & _
                        " left join materialmaster mm on mm.cmmf = d.cmmf" & _
                        " left join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup" & _
                        " where m.miropostingdate >= '{0:yyyy-MM-dd}' and m.miropostingdate <= '{1:yyyy-MM-dd}'" & _
                        " and pg.groupsbuid = {3:d} group by m.vendorcode ,mm.sbu,datatype,period,monthinq,e.currency)", myFirstDate, myLastDate, period, ProductType))

        Else
            Dim additionalquery = "with validcmmf as (select distinct cmmf from forecastestimationcomp )," & _
                           " lp as (select distinct cmmf,vendorcode,first_value(pricelistid) over(partition by cmmf,vendorcode order by cmmf,vendorcode,validfrom desc) as pricelistid from pricelist" & _
                           " where cmmf in (select cmmf from  validcmmf) order by cmmf) "
            'FirmorderGroup
            sb.Append(String.Format("(SELECT {0:yyyy} as myyear,e.vendorcode,pd.shiptoparty as sbuid,2 as datatype,{1}::character varying as period,getmonthinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{0:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as monthinqconf,(sum(getyearpurchasevalue({0:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob))::real / getexchangerate(e.currency)) as purchasevalue,sum(getyearqty({0:yyyy}, od.currentconfirmedetd, sd.currentinquiryetd, confs.confirmationstatus, PD.Comments, PD.osqty)) As qty" & _
                       ",sum(getyearpurchasevalue({0:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob))::real as purchasevalueoc ,e.currency" & _
                       " FROM cxsebodtp od LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid" & _
                       " LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" & _
                       " LEFT JOIN cxsalesorder so ON so.sebasiasalesorder = sd.sebasiasalesorder" & _
                       " LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid" & _
                       " LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono" & _
                       " Left join ekko e on e.po = pd.sebasiapono" & _
                       " Left join vendor v on v.vendorcode = e.vendorcode" & _
                       " LEFT JOIN materialmaster mm ON mm.cmmf = pd.cmmf" & _
                       " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
                       " LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid" & _
                       " LEFT JOIN cxconfstatus confs ON confs.cxconfid = c.cxconfid" & _
                       " left join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup" & _
                       " WHERE od.ordertype = 'Header'::bpchar AND ph.receptiondate <= sd.inquiryetd and pd.osqty > 0 and pg.groupsbuid = {2:d} group by e.vendorcode," & _
                       " pd.shiptoparty ,datatype,period,monthinqconf,e.currency)", myLastDate, period, ProductType))
            sb.Append(" union all ")
            'FirmorderGroupExtend
            sb.Append(String.Format("(SELECT {3},e.vendorcode,pd.shiptoparty as sbuid,2 as datatype,{1}::character varying as period,getmonthinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{0:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as monthinqconf,(sum(getyearpurchasevalue({3},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob))::real / getexchangerate(e.currency)) as purchasevalue,sum(getyearqty({3}, od.currentconfirmedetd, sd.currentinquiryetd, confs.confirmationstatus, PD.Comments, PD.osqty)) As qty" & _
                      ",sum(getyearpurchasevalue({3},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob))::real as purchasevalueoc ,e.currency" & _
                      " FROM cxsebodtp od LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid" & _
                      " LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" & _
                      " LEFT JOIN cxsalesorder so ON so.sebasiasalesorder = sd.sebasiasalesorder" & _
                      " LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid" & _
                      " LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono" & _
                      " Left join ekko e on e.po = pd.sebasiapono" & _
                      " Left join vendor v on v.vendorcode = e.vendorcode" & _
                      " LEFT JOIN materialmaster mm ON mm.cmmf = pd.cmmf" & _
                      " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
                      " LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid" & _
                      " LEFT JOIN cxconfstatus confs ON confs.cxconfid = c.cxconfid" & _
                      " left join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup" & _
                      " WHERE od.ordertype = 'Header'::bpchar AND ph.receptiondate <= sd.inquiryetd and pd.osqty > 0 and pg.groupsbuid = {2:d} group by e.vendorcode," & _
                      " pd.shiptoparty ,datatype,period,monthinqconf,e.currency)", myLastDate, period, ProductType, Year(myLastDate) + 1))
            sb.Append(" union all ")
            'ForecastGroup
            sb.Append(String.Format("({0} SELECT {1:yyyy},fe.vendorcode,fe.customercode as sbuid,1 as datatype,{2}::character varying as period,wm.mymonth,sum(getyearpurchasevalue({1:yyyy},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit)/getexchangerate(pl.currency))),sum(getyearqty({1:yyyy},wm.myyear,qty))" & _
                       " ,sum(getyearpurchasevalue({1:yyyy},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit)))::real as purchasevalueoc,pl.currency" & _
                       " FROM forecastestimationcomp fe LEFT JOIN customer c ON c.customercode = fe.customercode" & _
                       " LEFT JOIN materialmaster mm ON mm.cmmf = fe.cmmf" & _
                       " LEFT JOIN vendor v2 ON v2.vendorcode = fe.vendorcode" & _
                       " LEFT JOIN weektomonth wm ON wm.yearweek = fe.weeketa" & _
                       " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
                       " LEFT JOIN lp ON lp.cmmf = fe.cmmf and lp.vendorcode = v2.vendorcode" & _
                       " LEFT JOIN pricelist pl ON pl.pricelistid = lp.pricelistid " & _
                       " group by fe.vendorcode,fe.customercode,datatype,period,wm.mymonth,pl.currency)", additionalquery, myLastDate, period))

            sb.Append(" union all ")
            'ForecastgroupExtend
            sb.Append(String.Format("({0} SELECT {3},fe.vendorcode,fe.customercode as sbuid,1 as datatype,{2}::character varying as period,wm.mymonth,sum(getyearpurchasevalue({3},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit)/getexchangerate(pl.currency))),sum(getyearqty({3},wm.myyear,qty))" & _
                        " ,sum(getyearpurchasevalue({3},wm.myyear,getforecastvalue(fe.qty,pl.amount,pl.perunit)))::real as purchasevalueoc,pl.currency" & _
                        " FROM forecastestimationcomp fe LEFT JOIN customer c ON c.customercode = fe.customercode" & _
                        " LEFT JOIN materialmaster mm ON mm.cmmf = fe.cmmf" & _
                        " LEFT JOIN vendor v2 ON v2.vendorcode = fe.vendorcode" & _
                        " LEFT JOIN weektomonth wm ON wm.yearweek = fe.weeketa" & _
                        " LEFT JOIN family f ON f.familyid = mm.familylv1" & _
                        " LEFT JOIN lp ON lp.cmmf = fe.cmmf and lp.vendorcode = v2.vendorcode" & _
                        " LEFT JOIN pricelist pl ON pl.pricelistid = lp.pricelistid " & _
                        " group by fe.vendorcode,fe.customercode,datatype,period,wm.mymonth,pl.currency)", additionalquery, myLastDate, period, Year(myLastDate) + 1))

            sb.Append(" union all ")
            'Shipmentgroup
            sb.Append(String.Format("(select {1:yyyy}, m.vendorcode ,pd.shiptoparty as sbuid,3 as datatype, {2}::character varying as period,date_part('Month',miropostingdate) as monthinq ,(sum(getyearpurchasevalue({1:yyyy},miropostingdate,p.amount))/getexchangerate(e.currency)) as purchase,sum(getyearqty({1:yyyy},miropostingdate,p.qty::integer)) as qty" & _
                            " ,sum(getyearpurchasevalue({1:yyyy},miropostingdate,p.amount))::real as purchasevalueoc,e.currency" & _
                           " from miro m" & _
                       " left join pomiro p on p.miroid = m.miroid" & _
                       " left join podtl d on d.podtlid = p.podtlid" & _
                       " left join ekko e on e.po = d.pohd" & _
                       " left join pohd hd on hd.pohd = d.pohd " & _
                       " left join materialmaster mm on mm.cmmf = d.cmmf" & _
                       " left join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup" & _
                       " left join cxsebpodtl pd on pd.sebasiapono = d.pohd and pd.polineno = d.polineno" & _
                       " left join cxrelsalesdocpo r on r.cxsebpodtlid = pd.cxsebpodtlid" & _
                       " left join cxsalesorderdtl sd on sd.cxsalesorderdtlid = r.cxsalesorderdtlid" & _
                       " left join cxsalesorder so on so.sebasiasalesorder = sd.sebasiasalesorder" & _
                       " where m.miropostingdate >= '{0:yyyy-MM-dd}' and m.miropostingdate <= '{1:yyyy-MM-dd}'" & _
                       " and pg.groupsbuid = {3:d} group by m.vendorcode ,pd.shiptoparty,datatype,period,monthinq,e.currency)", myFirstDate, myLastDate, period, ProductType))

        End If



        Dim sqlstrgroup = String.Format("select foo.*,{1:d} as mypg from ({0}) as foo", sb.ToString, ProductType)


        If mypg = 1 Then
            sqlstrgroup = String.Format("insert into turnoverhistory(myyear,vendorcode,familysbu,doctype,period,inqconf,purchase,qty,purchaselc,crcy,sbuid,groupsbuid) " &
                                        " select foo.myyear::integer,foo.vendorcode::bigint,sbuid::bigint,datatype::bigint,period::character varying," &
                                        " foo.monthinqconf::integer,foo.purchasevalue::numeric,foo.qty::numeric,purchasevalueoc::numeric," &
                                        " foo.currency::character varying,foo.sbu::character vayring,1 as mypg from ({0}) as foo", sb.ToString)
        Else
            sqlstrgroup = String.Format("insert into turnoverhistory(myyear,vendorcode,familysbu,doctype,period,inqconf,purchase,qty,purchaselc,crcy,groupsbuid) " &
                                        " select foo.myyear::integer,foo.vendorcode::bigint,sbuid::bigint,datatype::bigint,period::character varying,foo.monthinqconf::integer,foo.purchasevalue::numeric,foo.qty::numeric,purchasevalueoc::numeric," &
                                        " foo.currency::character varying,2 as mypg from ({0}) as foo", sb.ToString)
        End If
        'HistoryDate = period
        Call DeletePeriod()
        dbAdapter1.ExecuteNonQuery(sqlstrgroup)
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region



   

End Class
