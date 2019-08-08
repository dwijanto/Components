Imports System.Threading
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Imports System.Text
Imports Components.PublicClass
Public Class FormTurnoverReport
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim sqlstr As New StringBuilder
    Dim myFirstDate As Date
    Dim myLastDate As Date
    Dim myFirstDateShipment As Date
    Dim myQueryWorksheetList As New List(Of QueryWorksheet)


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        myQueryWorksheetList.Clear()


        If DateTimePicker1.Value.Month = 12 Then
            myFirstDate = CDate(String.Format("{0:yyyy}-1-1", DateTimePicker1.Value))
            myFirstDateShipment = CDate(String.Format("{0}-1-1", DateTimePicker1.Value.Year - 1))
        Else
            myFirstDate = CDate(String.Format("{0}-1-1", DateTimePicker1.Value.Year - 1))
            myFirstDateShipment = CDate(String.Format("{0}-1-1", DateTimePicker1.Value.Year - 2))
        End If

        myLastDate = getLastdate(DateTimePicker1.Value.Month, DateTimePicker1.Value.Year)

        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty

        Dim myuser As String = String.Empty
        sqlstr.Clear()
        'sqlstr.Append(String.Format("with firmorder as " &
        '    " (SELECT case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP"",'FIRMORDER'::character varying as ""TYPE"", so.soldtoparty  AS ""Sold To Party SAP Code"", c2.customername AS ""Sold To Party Name""," &
        '    " pd.shiptoparty as ""Ship To Party"",c1.customername as ""Ship To Party Name"", e.vendorcode AS ""SAP Vendor Code"",v.vendorname AS ""Vendor Name"", " &
        '    " mm.rri AS ""Research and Industry Responsible"", sbu.sbuname AS ""SBU"", sbu1.sbuname AS ""BU"", sbu2.sbuname AS familysbu, pd.cmmf AS ""CMMF""," &
        '    " mm.familylv1 AS ""Commercial Family CG2 n"", f.familyname, b.brandname AS ""Brand Name CG2 n"", mm.commref AS ""CommercialRef"", " &
        '    " mm.materialdesc AS ""Description"", (getpurchasevalue(pd.osqty,pd.fob) / getexchangerate(e.currency)) AS ""Purchase Value""," &
        '    " pd.osqty,getyearinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{1:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""YEAR_INQCONF""," &
        '    " (getyearpurchasevalue({0:yyyy}::integer,od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob)/getexchangerate(e.currency) )as ""Year{0:yyyy} Purchase Value"", " &
        '    " (getyearpurchasevalue({1:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob) / getexchangerate(e.currency)) as ""Year{1:yyyy} Purchase Value"" ," &
        '    " Null::numeric as ""Year2018 Purchase Value"" ,getyearqty({0:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus," &
        '    " pd.comments,pd.osqty) as ""Year{0:yyyy} QTY"",getyearqty({1:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty) as ""Year{1:yyyy} QTY""," &
        '    " Null::numeric as ""Year2018 QTY"" ,getmonthinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{1:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""MONTH INQCONF""," &
        '    " ssm.officersebname as ""SPM""  ,adjustedmonth(confirmationstatus,od.currentconfirmedetd,pd.comments) as adjustedmonth, od.currentconfirmedetd,sd.currentinquiryetd,getconfirmstatus(confs.confirmationstatus,pd.comments,od.currentconfirmedetd) as confirmstatus,pd.comments,pd.sebasiapono,pd.polineno,sp.sopdescription as industrialfamily ,e.currency,getpurchasevalue(pd.osqty,pd.fob) as ""Purchase Value (OC)"" , '1'::integer as" &
        '    " isreal,cf.flow,cf.continent,cf.continent_group,cf.continent_group_emea FROM cxsebodtp od LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid " &
        '    " LEFT JOIN cxsalesorder so ON so.sebasiasalesorder = sd.sebasiasalesorder LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono " &
        '    " LEFT JOIN customer c1 ON c1.customercode = pd.shiptoparty LEFT JOIN customer c2 ON c2.customercode = so.soldtoparty LEFT JOIN cxpoconf pc ON pc.sebasiapono = pd.sebasiapono AND pc.polineno = pd.polineno " &
        '    " LEFT JOIN cxpoconfother pco ON pco.sebasiapono = pd.sebasiapono AND pco.polineno = pd.polineno LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid " &
        '    " LEFT JOIN cxconfstatus confs ON confs.cxconfid = c.cxconfid Left join ekko e on e.po = pd.sebasiapono Left join vendor v on v.vendorcode = e.vendorcode " &
        '    " LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl LEFT JOIN materialmaster mm  ON mm.cmmf = pd.cmmf left join sspcmmfsop sc on sc.cmmf = mm.cmmf " &
        '    " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid LEFT JOIN brand b ON b.brandid = mm.brandid " &
        '    " LEFT JOIN activity ac ON ac.activitycode = mm.rri LEFT JOIN sbu ON sbu.sbuid = ac.sbuidsp LEFT JOIN sbu sbu1 ON sbu1.sbuid = ac.sbuidlg " &
        '    " LEFT JOIN family f ON f.familyid = mm.familylv1 LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu LEFT JOIN purchasinggroup pg ON pg.purchasinggroup = e.purchasinggroup::bpchar " &
        '    " Left Join customerflow cf on cf.soldtoparty = so.soldtoparty and cf.shiptoparty = pd.shiptoparty" &
        '    " WHERE od.ordertype = 'Header'::bpchar AND ph.receptiondate <= sd.inquiryetd and pd.osqty > 0  " &
        '    " and pg.groupsbuid <= 2 ORDER BY pg.groupsbuid,ph.receptiondate )," &
        '    " shipment as " &
        '    " (select miropostingdate,case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP"",'SHIPMENT'::character varying as type,so.soldtoparty,cu1.customername, cxpo.shiptoparty,cu2.customername, m.vendorcode,v.vendorname,mm.rri,sbu.sbuname,sbu1.sbuname," &
        '    " sbu2.sbuname as familysbu,d.cmmf,mm.familylv1,f.familyname,b.brandname,mm.commref,mm.materialdesc,(p.amount / getexchangerate(e.currency))as amount,p.qty::integer," &
        '    " date_part('Year',miropostingdate)::integer,(getyearpurchasevalue({0:yyyy},miropostingdate,p.amount)/getexchangerate(e.currency)) as lastyear," &
        '    " (getyearpurchasevalue({1:yyyy},miropostingdate,p.amount)/getexchangerate(e.currency)) as thisyear,null::numeric,getyearqty({0:yyyy},miropostingdate,p.qty::integer)," &
        '    " getyearqty({1:yyyy},miropostingdate,p.qty::integer),null::numeric,date_part('Month',miropostingdate),ssm.officersebname, null::integer,null::date,null::date,null::character varying,null::character varying,hd.pohd,d.polineno,sp.sopdescription, e.currency, p.amount," &
        '    " '1'::integer as isreal,cf.flow,cf.continent,cf.continent_group,cf.continent_group_emea from pomiro " &
        '    " p left join miro m on m.miroid = p.miroid left join podtl d on d.podtlid = p.podtlid left join vendor v on v.vendorcode = m.vendorcode left join cxsebpodtl cxpo on cxpo.sebasiapono = d.pohd and cxpo.polineno = d.polineno " &
        '    " left join ekko e on e.po = d.pohd left join pohd hd on hd.pohd = d.pohd " &
        '    " left join materialmaster mm on mm.cmmf = d.cmmf left join sspcmmfsop sc on sc.cmmf = d.cmmf " &
        '    " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid " &
        '    " LEFT JOIN brand b ON b.brandid = mm.brandid LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl " &
        '    " left join cxsebpodtl pd on pd.sebasiapono = d.pohd and pd.polineno = d.polineno  " &
        '    " left join cxrelsalesdocpo r on r.cxsebpodtlid = pd.cxsebpodtlid " &
        '    " left join cxsalesorderdtl sd on sd.cxsalesorderdtlid = r.cxsalesorderdtlid " &
        '    " left join cxsalesorder so on so.sebasiasalesorder = sd.sebasiasalesorder " &
        '    " left join customer cu1 on cu1.customercode = so.soldtoparty " &
        '    " left join customer cu2 on cu2.customercode = cxpo.shiptoparty " &
        '    " LEFT JOIN activity ac ON  ac.activitycode = mm.rri " &
        '    " LEFT JOIN sbu ON sbu.sbuid = ac.sbuid " &
        '    " LEFT JOIN family f ON f.familyid = mm.familylv1 " &
        '    " LEFT JOIN sbu sbu1 on sbu1.sbuid = ac.sbuidlg " &
        '    " LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu " &
        '    " Left Join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup " &
        '    " Left Join customerflow cf on cf.soldtoparty = so.soldtoparty and cf.shiptoparty = pd.shiptoparty" &
        '    " where miropostingdate >= '{0:yyyy-MM-dd}' and miropostingdate <= '{1:yyyy-MM-dd}' and pg.groupsbuid <= 2" &
        '    " order by pg.groupsbuid,p.pomiroid)" &
        '    " select case ""Sold To Party SAP Code"" when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",case when ""FG/CP"" = 'FG' or ""Sold To Party SAP Code"" = 99005151 then 'FG+GSE' else ""FG/CP"" end as ""Shipment Type (FG+GSE)"" ,case when ""SAP Vendor Code"" < 90000000 then 'non-group' else 'group' end as ""Vendor Type"" , to_date(""YEAR_INQCONF"" || '-' || ""MONTH INQCONF"" || '-1','yyyy-MM-dd' ) as txdate,* from firmorder" &
        '    " union all" &
        '    " select case soldtoparty when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",case when ""FG/CP"" = 'FG' or soldtoparty = 99005151 then 'FG+GSE' else ""FG/CP"" end as ""Shipment Type (FG+GSE)"" ,case when vendorcode < 90000000 then 'non-group' else 'group' end as ""Vendor Type"" ,* from shipment;", myFirstDate, myLastDate))

        'sqlstr.Append(String.Format("with firmorder as " &
        '    " (SELECT case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP"",'FIRMORDER'::character varying as ""TYPE"", so.soldtoparty  AS ""Sold To Party SAP Code"", c2.customername AS ""Sold To Party Name""," &
        '    " pd.shiptoparty as ""Ship To Party"",c1.customername as ""Ship To Party Name"", e.vendorcode AS ""SAP Vendor Code"",v.vendorname AS ""Vendor Name"", " &
        '    " mm.rri AS ""Research and Industry Responsible"", sbu.sbuname AS ""SBU"", sbu1.sbuname AS ""BU"", sbu2.sbuname AS familysbu, pd.cmmf AS ""CMMF""," &
        '    " mm.familylv1 AS ""Commercial Family CG2 n"", f.familyname, b.brandname AS ""Brand Name CG2 n"", mm.commref AS ""CommercialRef"", " &
        '    " mm.materialdesc AS ""Description"", (getpurchasevalue(pd.osqty,pd.fob) / getexchangerate(e.currency)) AS ""Purchase Value""," &
        '    " pd.osqty,getyearinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{1:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""YEAR_INQCONF""," &
        '    " (getyearpurchasevalue({0:yyyy}::integer,od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob)/getexchangerate(e.currency) )as ""Year{0:yyyy} Purchase Value"", " &
        '    " (getyearpurchasevalue({1:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob) / getexchangerate(e.currency)) as ""Year{1:yyyy} Purchase Value"" ," &
        '    " Null::numeric as ""Year{2} Purchase Value"" ,getyearqty({0:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus," &
        '    " pd.comments,pd.osqty) as ""Year{0:yyyy} QTY"",getyearqty({1:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty) as ""Year{1:yyyy} QTY""," &
        '    " Null::numeric as ""Year{2} QTY"" ,getmonthinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{1:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""MONTH INQCONF""," &
        '    " ssm.officersebname as ""SPM""  ,adjustedmonth(confirmationstatus,od.currentconfirmedetd,pd.comments) as adjustedmonth, od.currentconfirmedetd,sd.currentinquiryetd,getconfirmstatus(confs.confirmationstatus,pd.comments,od.currentconfirmedetd) as confirmstatus,pd.comments,pd.sebasiapono,pd.polineno,sp.sopdescription as industrialfamily ,e.currency,getpurchasevalue(pd.osqty,pd.fob) as ""Purchase Value (OC)"" , '1'::integer as" &
        '    " isreal,cf.flow,cf.continent,cf.continent_group,cf.continent_group_emea FROM cxsebodtp od LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid " &
        '    " LEFT JOIN cxsalesorder so ON so.sebasiasalesorder = sd.sebasiasalesorder LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono " &
        '    " LEFT JOIN customer c1 ON c1.customercode = pd.shiptoparty LEFT JOIN customer c2 ON c2.customercode = so.soldtoparty LEFT JOIN cxpoconf pc ON pc.sebasiapono = pd.sebasiapono AND pc.polineno = pd.polineno " &
        '    " LEFT JOIN cxpoconfother pco ON pco.sebasiapono = pd.sebasiapono AND pco.polineno = pd.polineno LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid " &
        '    " LEFT JOIN cxconfstatus confs ON confs.cxconfid = c.cxconfid Left join ekko e on e.po = pd.sebasiapono Left join vendor v on v.vendorcode = e.vendorcode " &
        '    " LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl LEFT JOIN materialmaster mm  ON mm.cmmf = pd.cmmf left join sspcmmfsop sc on sc.cmmf = mm.cmmf " &
        '    " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid LEFT JOIN brand b ON b.brandid = mm.brandid " &
        '    " LEFT JOIN activity ac ON ac.activitycode = mm.rri LEFT JOIN sbu ON sbu.sbuid = ac.sbuidsp LEFT JOIN sbu sbu1 ON sbu1.sbuid = ac.sbuidlg " &
        '    " LEFT JOIN family f ON f.familyid = mm.familylv1 LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu LEFT JOIN purchasinggroup pg ON pg.purchasinggroup = e.purchasinggroup::bpchar " &
        '    " Left Join customerflow cf on cf.soldtoparty = so.soldtoparty and cf.shiptoparty = pd.shiptoparty" &
        '    " WHERE od.ordertype = 'Header'::bpchar AND ph.receptiondate <= sd.inquiryetd and pd.osqty > 0  " &
        '    " and pg.groupsbuid <= 2 ORDER BY pg.groupsbuid,ph.receptiondate )," &
        '    " shipment as " &
        '    " (select miropostingdate,case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP"",'SHIPMENT'::character varying as type,so.soldtoparty,cu1.customername, cxpo.shiptoparty,cu2.customername, m.vendorcode,v.vendorname,mm.rri,sbu.sbuname,sbu1.sbuname," &
        '    " sbu2.sbuname as familysbu,d.cmmf,mm.familylv1,f.familyname,b.brandname,mm.commref,mm.materialdesc,(p.amount / getexchangerate(e.currency))as amount,p.qty::integer," &
        '    " date_part('Year',miropostingdate)::integer,(getyearpurchasevalue({0:yyyy},miropostingdate,p.amount)/getexchangerate(e.currency)) as lastyear," &
        '    " (getyearpurchasevalue({1:yyyy},miropostingdate,p.amount)/getexchangerate(e.currency)) as thisyear,null::numeric,getyearqty({0:yyyy},miropostingdate,p.qty::integer)," &
        '    " getyearqty({1:yyyy},miropostingdate,p.qty::integer),null::numeric,date_part('Month',miropostingdate),ssm.officersebname, null::integer,null::date,null::date,null::character varying,null::character varying,hd.pohd,d.polineno,sp.sopdescription, e.currency, p.amount," &
        '    " '1'::integer as isreal,cf.flow,cf.continent,cf.continent_group,cf.continent_group_emea from pomiro " &
        '    " p left join miro m on m.miroid = p.miroid left join podtl d on d.podtlid = p.podtlid left join vendor v on v.vendorcode = m.vendorcode left join cxsebpodtl cxpo on cxpo.sebasiapono = d.pohd and cxpo.polineno = d.polineno " &
        '    " left join ekko e on e.po = d.pohd left join pohd hd on hd.pohd = d.pohd " &
        '    " left join materialmaster mm on mm.cmmf = d.cmmf left join sspcmmfsop sc on sc.cmmf = d.cmmf " &
        '    " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid " &
        '    " LEFT JOIN brand b ON b.brandid = mm.brandid LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl " &
        '    " left join cxsebpodtl pd on pd.sebasiapono = d.pohd and pd.polineno = d.polineno  " &
        '    " left join cxrelsalesdocpo r on r.cxsebpodtlid = pd.cxsebpodtlid " &
        '    " left join cxsalesorderdtl sd on sd.cxsalesorderdtlid = r.cxsalesorderdtlid " &
        '    " left join cxsalesorder so on so.sebasiasalesorder = sd.sebasiasalesorder " &
        '    " left join customer cu1 on cu1.customercode = so.soldtoparty " &
        '    " left join customer cu2 on cu2.customercode = cxpo.shiptoparty " &
        '    " LEFT JOIN activity ac ON  ac.activitycode = mm.rri " &
        '    " LEFT JOIN sbu ON sbu.sbuid = ac.sbuid " &
        '    " LEFT JOIN family f ON f.familyid = mm.familylv1 " &
        '    " LEFT JOIN sbu sbu1 on sbu1.sbuid = ac.sbuidlg " &
        '    " LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu " &
        '    " Left Join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup " &
        '    " Left Join customerflow cf on cf.soldtoparty = so.soldtoparty and cf.shiptoparty = pd.shiptoparty" &
        '    " where miropostingdate >= '{0:yyyy-MM-dd}' and miropostingdate <= '{1:yyyy-MM-dd}' and pg.groupsbuid <= 2" &
        '    " order by pg.groupsbuid,p.pomiroid)," &
        '    " f1 as (select case ""Sold To Party SAP Code"" when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",case when ""SAP Vendor Code"" < 90000000 then 'non-group' else 'group' end as ""Vendor Type"" , to_date(""YEAR_INQCONF"" || '-' || ""MONTH INQCONF"" || '-1','yyyy-MM-dd' ) as txdate,* from firmorder)," &
        '    " s1 as (select case soldtoparty when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",case when vendorcode < 90000000 then 'non-group' else 'group' end as ""Vendor Type"" ,* from shipment)" &
        '    " select case ""Shipment Type"" when 'FG' then 'FG+GSE' when 'GSE' then 'FG+GSE' else ""Shipment Type"" end as ""Shipment Type (FG+GSE)"",f1.* from f1 " &
        '    " union all select case ""Shipment Type"" when 'FG' then 'FG+GSE' when 'GSE' then 'FG+GSE' else ""Shipment Type"" end as ""Shipment Type (FG+GSE)"",* from s1;", myFirstDate, myLastDate, Year(myLastDate) + 1))

        'sqlstr.Append(String.Format("with firmorder as " &
        '    " (SELECT case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP"",'FIRMORDER'::character varying as ""TYPE"", so.soldtoparty  AS ""Sold To Party SAP Code"", c2.customername AS ""Sold To Party Name""," &
        '    " pd.shiptoparty as ""Ship To Party"",c1.customername as ""Ship To Party Name"", e.vendorcode AS ""SAP Vendor Code"",v.vendorname AS ""Vendor Name"", " &
        '    " mm.rri AS ""Research and Industry Responsible"", sbu.sbuname AS ""SBU"", sbu1.sbuname AS ""BU"", sbu2.sbuname AS familysbu, pd.cmmf AS ""CMMF""," &
        '    " mm.familylv1 AS ""Commercial Family CG2 n"", f.familyname, b.brandname AS ""Brand Name CG2 n"", mm.commref AS ""CommercialRef"", " &
        '    " mm.materialdesc AS ""Description"", (getpurchasevalue(pd.osqty,pd.fob) / getexchangerate(e.currency)) AS ""Purchase Value""," &
        '    " pd.osqty,getyearinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{1:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""YEAR_INQCONF""," &
        '    " (getyearpurchasevalue({0:yyyy}::integer,od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob)/getexchangerate(e.currency) )as ""Year{0:yyyy} Purchase Value"", " &
        '    " (getyearpurchasevalue({1:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob) / getexchangerate(e.currency)) as ""Year{1:yyyy} Purchase Value"" ," &
        '    " getyearqty({0:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus," &
        '    " pd.comments,pd.osqty) as ""Year{0:yyyy} QTY"",getyearqty({1:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty) as ""Year{1:yyyy} QTY""," &
        '    " getmonthinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{1:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""MONTH INQCONF""," &
        '    " ssm.officersebname as ""SPM""  ,adjustedmonth(confirmationstatus,od.currentconfirmedetd,pd.comments) as adjustedmonth, od.currentconfirmedetd,sd.currentinquiryetd,getconfirmstatus(confs.confirmationstatus,pd.comments,od.currentconfirmedetd) as confirmstatus,pd.comments,pd.sebasiapono,pd.polineno,sp.sopdescription as industrialfamily ,e.currency,getpurchasevalue(pd.osqty,pd.fob) as ""Purchase Value (OC)"" , '1'::integer as" &
        '    " isreal,cf.flow,cf.continent,cf.continent_group,cf.continent_group_emea FROM cxsebodtp od LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid " &
        '    " LEFT JOIN cxsalesorder so ON so.sebasiasalesorder = sd.sebasiasalesorder LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono " &
        '    " LEFT JOIN customer c1 ON c1.customercode = pd.shiptoparty LEFT JOIN customer c2 ON c2.customercode = so.soldtoparty " &
        '    " LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid " &
        '    " LEFT JOIN cxconfstatus confs ON confs.cxconfid = c.cxconfid Left join ekko e on e.po = pd.sebasiapono Left join vendor v on v.vendorcode = e.vendorcode " &
        '    " LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl LEFT JOIN materialmaster mm  ON mm.cmmf = pd.cmmf left join sspcmmfsop sc on sc.cmmf = mm.cmmf " &
        '    " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid LEFT JOIN brand b ON b.brandid = mm.brandid " &
        '    " LEFT JOIN activity ac ON ac.activitycode = mm.rri LEFT JOIN sbu ON sbu.sbuid = ac.sbuidsp LEFT JOIN sbu sbu1 ON sbu1.sbuid = ac.sbuidlg " &
        '    " LEFT JOIN family f ON f.familyid = mm.familylv1 LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu LEFT JOIN purchasinggroup pg ON pg.purchasinggroup = e.purchasinggroup::bpchar " &
        '    " Left Join customerflow cf on cf.soldtoparty = so.soldtoparty and cf.shiptoparty = pd.shiptoparty" &
        '    " WHERE od.ordertype = 'Header'::bpchar AND ph.receptiondate <= sd.inquiryetd and pd.osqty > 0  " &
        '    " and pg.groupsbuid <= 2 ORDER BY pg.groupsbuid,ph.receptiondate )," &
        '    " shipment as " &
        '    " (select miropostingdate,case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP"",'SHIPMENT'::character varying as type,so.soldtoparty,cu1.customername, pd.shiptoparty,cu2.customername, m.vendorcode,v.vendorname,mm.rri,sbu.sbuname,sbu1.sbuname," &
        '    " sbu2.sbuname as familysbu,d.cmmf,mm.familylv1,f.familyname,b.brandname,mm.commref,mm.materialdesc,(p.amount / getexchangerate(e.currency))as amount,p.qty::integer," &
        '    " date_part('Year',miropostingdate)::integer,(getyearpurchasevalue({0:yyyy},miropostingdate,p.amount)/getexchangerate(e.currency)) as lastyear," &
        '    " (getyearpurchasevalue({1:yyyy},miropostingdate,p.amount)/getexchangerate(e.currency)) as thisyear,getyearqty({0:yyyy},miropostingdate,p.qty::integer)," &
        '    " getyearqty({1:yyyy},miropostingdate,p.qty::integer),date_part('Month',miropostingdate),ssm.officersebname, null::integer,null::date,null::date,null::character varying,null::character varying,hd.pohd,d.polineno,sp.sopdescription, e.currency, p.amount," &
        '    " '1'::integer as isreal,cf.flow,cf.continent,cf.continent_group,cf.continent_group_emea from pomiro " &
        '    " p left join miro m on m.miroid = p.miroid left join podtl d on d.podtlid = p.podtlid left join vendor v on v.vendorcode = m.vendorcode " &
        '    " left join ekko e on e.po = d.pohd left join pohd hd on hd.pohd = d.pohd " &
        '    " left join materialmaster mm on mm.cmmf = d.cmmf left join sspcmmfsop sc on sc.cmmf = d.cmmf " &
        '    " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid " &
        '    " LEFT JOIN brand b ON b.brandid = mm.brandid LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl " &
        '    " left join cxsebpodtl pd on pd.sebasiapono = d.pohd and pd.polineno = d.polineno  " &
        '    " left join cxrelsalesdocpo r on r.cxsebpodtlid = pd.cxsebpodtlid " &
        '    " left join cxsalesorderdtl sd on sd.cxsalesorderdtlid = r.cxsalesorderdtlid " &
        '    " left join cxsalesorder so on so.sebasiasalesorder = sd.sebasiasalesorder " &
        '    " left join customer cu1 on cu1.customercode = so.soldtoparty " &
        '    " left join customer cu2 on cu2.customercode = pd.shiptoparty " &
        '    " LEFT JOIN activity ac ON  ac.activitycode = mm.rri " &
        '    " LEFT JOIN sbu ON sbu.sbuid = ac.sbuid " &
        '    " LEFT JOIN family f ON f.familyid = mm.familylv1 " &
        '    " LEFT JOIN sbu sbu1 on sbu1.sbuid = ac.sbuidlg " &
        '    " LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu " &
        '    " Left Join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup " &
        '    " Left Join customerflow cf on cf.soldtoparty = so.soldtoparty and cf.shiptoparty = pd.shiptoparty" &
        '    " where miropostingdate >= '{0:yyyy-MM-dd}' and miropostingdate <= '{1:yyyy-MM-dd}' and pg.groupsbuid <= 2" &
        '    " order by pg.groupsbuid,p.pomiroid)," &
        '    " f1 as (select case ""Sold To Party SAP Code"" when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",case when ""SAP Vendor Code"" < 90000000 then 'non-group' else 'group' end as ""Vendor Type"" , to_date(""YEAR_INQCONF"" || '-' || ""MONTH INQCONF"" || '-1','yyyy-MM-dd' ) as txdate,* from firmorder)," &
        '    " s1 as (select case soldtoparty when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",case when vendorcode < 90000000 then 'non-group' else 'group' end as ""Vendor Type"" ,* from shipment)," &
        '    " fs as (select case ""Shipment Type"" when 'FG' then 'FG+GSE' when 'GSE' then 'FG+GSE' else ""Shipment Type"" end as ""Shipment Type (FG+GSE)"",f1.* from f1 " &
        '    " union all select case ""Shipment Type"" when 'FG' then 'FG+GSE' when 'GSE' then 'FG+GSE' else ""Shipment Type"" end as ""Shipment Type (FG+GSE)"",* from s1) " &
        '    " select fs.*,case ""YEAR_INQCONF"" when 2016 then 1 end as count2016,case ""YEAR_INQCONF"" when 2017 then 1 end as count2017 from fs", myFirstDate, myLastDate))
        'sqlstr.Append(String.Format("with firmorder as " &
        '    " (SELECT case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP"",'FIRMORDER'::character varying as ""TYPE"", so.soldtoparty  AS ""Sold To Party SAP Code"", c2.customername AS ""Sold To Party Name""," &
        '    " pd.shiptoparty as ""Ship To Party"",c1.customername as ""Ship To Party Name"", e.vendorcode AS ""SAP Vendor Code"",v.vendorname AS ""Vendor Name"", " &
        '    " mm.rri AS ""Research and Industry Responsible"", sbu.sbuname AS ""SBU"", sbu1.sbuname AS ""BU"", sbu2.sbuname AS familysbu, pd.cmmf AS ""CMMF""," &
        '    " mm.familylv1 AS ""Commercial Family CG2 n"", f.familyname, b.brandname AS ""Brand Name CG2 n"", mm.commref AS ""CommercialRef"", " &
        '    " mm.materialdesc AS ""Description"", (getpurchasevalue(pd.osqty,pd.fob) / getexchangerate(e.currency)) AS ""Purchase Value""," &
        '    " pd.osqty,getyearinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{1:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""YEAR_INQCONF""," &
        '    " Null::numeric as ""Year{0:yyyy} Purchase Value"", " &
        '    " (getyearpurchasevalueinc({1:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob) / getexchangerate(e.currency)) as ""Year{1:yyyy} Purchase Value"" ," &
        '    " getyearqty({0:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus," &
        '    " pd.comments,pd.osqty) as ""Year{0:yyyy} QTY"",getyearqty({1:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty) as ""Year{1:yyyy} QTY""," &
        '    " getmonthinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{1:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""MONTH INQCONF""," &
        '    " ssm.officersebname as ""SPM""  ,adjustedmonth(confirmationstatus,od.currentconfirmedetd,pd.comments) as adjustedmonth, od.currentconfirmedetd,sd.currentinquiryetd,getconfirmstatus(confs.confirmationstatus,pd.comments,od.currentconfirmedetd) as confirmstatus,pd.comments,pd.sebasiapono,pd.polineno,sp.sopdescription as industrialfamily ,e.currency,getpurchasevalue(pd.osqty,pd.fob) as ""Purchase Value (OC)"" , '1'::integer as" &
        '    " isreal,cf.flow,cf.continent,cf.continent_group,cf.continent_group_emea FROM cxsebodtp od LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid " &
        '    " LEFT JOIN cxsalesorder so ON so.sebasiasalesorder = sd.sebasiasalesorder LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono " &
        '    " LEFT JOIN customer c1 ON c1.customercode = pd.shiptoparty LEFT JOIN customer c2 ON c2.customercode = so.soldtoparty " &
        '    " LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid " &
        '    " LEFT JOIN cxconfstatus confs ON confs.cxconfid = c.cxconfid Left join ekko e on e.po = pd.sebasiapono Left join vendor v on v.vendorcode = e.vendorcode " &
        '    " LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl LEFT JOIN materialmaster mm  ON mm.cmmf = pd.cmmf left join sspcmmfsop sc on sc.cmmf = mm.cmmf " &
        '    " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid LEFT JOIN brand b ON b.brandid = mm.brandid " &
        '    " LEFT JOIN activity ac ON ac.activitycode = mm.rri LEFT JOIN sbu ON sbu.sbuid = ac.sbuidsp LEFT JOIN sbu sbu1 ON sbu1.sbuid = ac.sbuidlg " &
        '    " LEFT JOIN family f ON f.familyid = mm.familylv1 LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu LEFT JOIN purchasinggroup pg ON pg.purchasinggroup = e.purchasinggroup::bpchar " &
        '    " Left Join customerflow cf on cf.soldtoparty = so.soldtoparty and cf.shiptoparty = pd.shiptoparty" &
        '    " WHERE od.ordertype = 'Header'::bpchar AND ph.receptiondate <= sd.inquiryetd and pd.osqty > 0  " &
        '    " and pg.groupsbuid <= 2 ORDER BY pg.groupsbuid,ph.receptiondate )," &
        '    " shipment as " &
        '    " (select miropostingdate,case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP"",'SHIPMENT'::character varying as type,so.soldtoparty,cu1.customername, pd.shiptoparty,cu2.customername, m.vendorcode,v.vendorname,mm.rri,sbu.sbuname,sbu1.sbuname," &
        '    " sbu2.sbuname as familysbu,d.cmmf,mm.familylv1,f.familyname,b.brandname,mm.commref,mm.materialdesc,(p.amount / getexchangerate(e.currency))as amount,p.qty::integer," &
        '    " date_part('Year',miropostingdate)::integer,(getyearpurchasevalue({0:yyyy},miropostingdate,p.amount)/getexchangerate(e.currency)) as lastyear," &
        '    " (getyearpurchasevalue({1:yyyy},miropostingdate,p.amount)/getexchangerate(e.currency)) as thisyear,getyearqty({0:yyyy},miropostingdate,p.qty::integer)," &
        '    " getyearqty({1:yyyy},miropostingdate,p.qty::integer),date_part('Month',miropostingdate),ssm.officersebname, null::integer,null::date,null::date,null::character varying,null::character varying,hd.pohd,d.polineno,sp.sopdescription, e.currency, p.amount," &
        '    " '1'::integer as isreal,cf.flow,cf.continent,cf.continent_group,cf.continent_group_emea from pomiro " &
        '    " p left join miro m on m.miroid = p.miroid left join podtl d on d.podtlid = p.podtlid left join vendor v on v.vendorcode = m.vendorcode " &
        '    " left join ekko e on e.po = d.pohd left join pohd hd on hd.pohd = d.pohd " &
        '    " left join materialmaster mm on mm.cmmf = d.cmmf left join sspcmmfsop sc on sc.cmmf = d.cmmf " &
        '    " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid " &
        '    " LEFT JOIN brand b ON b.brandid = mm.brandid LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl " &
        '    " left join cxsebpodtl pd on pd.sebasiapono = d.pohd and pd.polineno = d.polineno  " &
        '    " left join cxrelsalesdocpo r on r.cxsebpodtlid = pd.cxsebpodtlid " &
        '    " left join cxsalesorderdtl sd on sd.cxsalesorderdtlid = r.cxsalesorderdtlid " &
        '    " left join cxsalesorder so on so.sebasiasalesorder = sd.sebasiasalesorder " &
        '    " left join customer cu1 on cu1.customercode = so.soldtoparty " &
        '    " left join customer cu2 on cu2.customercode = pd.shiptoparty " &
        '    " LEFT JOIN activity ac ON  ac.activitycode = mm.rri " &
        '    " LEFT JOIN sbu ON sbu.sbuid = ac.sbuid " &
        '    " LEFT JOIN family f ON f.familyid = mm.familylv1 " &
        '    " LEFT JOIN sbu sbu1 on sbu1.sbuid = ac.sbuidlg " &
        '    " LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu " &
        '    " Left Join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup " &
        '    " Left Join customerflow cf on cf.soldtoparty = so.soldtoparty and cf.shiptoparty = pd.shiptoparty" &
        '    " where miropostingdate >= '{2:yyyy-MM-dd}' and miropostingdate <= '{1:yyyy-MM-dd}' and pg.groupsbuid <= 2" &
        '    " order by pg.groupsbuid,p.pomiroid)," &
        '    " f1 as (select case ""Sold To Party SAP Code"" when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",case when ""SAP Vendor Code"" < 90000000 then 'non-group' else 'group' end as ""Vendor Type"" , to_date(""YEAR_INQCONF"" || '-' || ""MONTH INQCONF"" || '-1','yyyy-MM-dd' ) as txdate,* from firmorder)," &
        '    " s1 as (select case soldtoparty when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",case when vendorcode < 90000000 then 'non-group' else 'group' end as ""Vendor Type"" ,* from shipment)," &
        '    " fs as (select case ""Shipment Type"" when 'FG' then 'FG+GSE' when 'GSE' then 'FG+GSE' else ""Shipment Type"" end as ""Shipment Type (FG+GSE)"",f1.* from f1 " &
        '    " union all select case ""Shipment Type"" when 'FG' then 'FG+GSE' when 'GSE' then 'FG+GSE' else ""Shipment Type"" end as ""Shipment Type (FG+GSE)"",* from s1) " &
        '    " select fs.*,pd.cvalue as market,case ""YEAR_INQCONF"" when 2016 then 1 end as count2016,case ""YEAR_INQCONF"" when 2017 then 1 end as count2017 from fs left join soldtomarket stm on stm.soldtoparty = fs.""Sold To Party SAP Code"" left join paramdt pd on pd.paramdtid = stm.market", myFirstDate, myLastDate, myFirstDateShipment))

        'Dim sqlstr2 As String = String.Format("with firmorder as(select  distinct getyearinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{2:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""YEAR_INQCONF""," &
        '                                      " so.soldtoparty  AS ""Sold To Party SAP Code"", e.vendorcode AS ""SAP Vendor Code"",case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP"",case when e.vendorcode < 90000000 then 'non-group' else 'group' end as vendortype " &
        '                                      " FROM cxsebodtp od LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
        '                                      " LEFT JOIN cxsalesorder so ON so.sebasiasalesorder = sd.sebasiasalesorder LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid " &
        '                                      " LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid LEFT JOIN cxconfstatus confs ON confs.cxconfid = c.cxconfid" &
        '                                      " Left join ekko e on e.po = pd.sebasiapono LEFT JOIN purchasinggroup pg ON pg.purchasinggroup = e.purchasinggroup::bpchar  WHERE od.ordertype = 'Header'::bpchar AND ph.receptiondate <= sd.inquiryetd and pd.osqty > 0  and pg.groupsbuid <= 2 )," &
        '                                      " f1 as(select case ""Sold To Party SAP Code"" when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",f.*" &
        '                                      " from firmorder f), " &
        '                                      " fall as(select distinct ""YEAR_INQCONF"",""Sold To Party SAP Code"",""SAP Vendor Code"",""Shipment Type"",vendortype from f1" &
        '                                      " where ""YEAR_INQCONF"" = {1:yyyy})," &
        '                                      " shipment as (select distinct date_part('Year',miropostingdate)::integer as ""YEAR_INQCONF"",so.soldtoparty as ""Sold To Party SAP Code"", m.vendorcode as ""SAP Vendor Code"",case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP""," &
        '                                      " case when m.vendorcode < 90000000 then 'non-group' else 'group' end as ""Vendor Type""" &
        '                                      " from pomiro p" &
        '                                      " left join miro m on m.miroid = p.miroid left join podtl d on d.podtlid = p.podtlid left join ekko e on e.po = d.pohd " &
        '                                      " Left Join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup left join cxsebpodtl pd on pd.sebasiapono = d.pohd and pd.polineno = d.polineno" &
        '                                      " left join cxrelsalesdocpo r on r.cxsebpodtlid = pd.cxsebpodtlid left join cxsalesorderdtl sd on sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
        '                                      " left join cxsalesorder so on so.sebasiasalesorder = sd.sebasiasalesorder where miropostingdate >= '{2:yyyy-MM-dd}' and miropostingdate <= '{1:yyyy-MM-yy}' and pg.groupsbuid <= 2 )," &
        '                                      " s1 as (select ""YEAR_INQCONF"",""Sold To Party SAP Code"",""SAP Vendor Code"",case ""Sold To Party SAP Code"" when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",""Vendor Type"" from shipment)," &
        '                                      " alldata as (select * from fall  union all select * from s1), " &
        '                                      " alldatadistinct as( select distinct ""YEAR_INQCONF"",""Sold To Party SAP Code"",null::bigint as ""SAP Vendor Code"",""Shipment Type"",null::text as vendortype,'#of market'::text as details from alldata" &
        '                                      " union all " &
        '                                      " select distinct ""YEAR_INQCONF"",null::bigint as ""Sold To Party SAP Code"",""SAP Vendor Code"",""Shipment Type"",vendortype,'#of supplier' from alldata)" &
        '                                      " select *,case when ""YEAR_INQCONF"" = {1:yyyy} then 1 end as countof{1:yyyy},case when ""YEAR_INQCONF"" = {0:yyyy} then 1 end as countof{0:yyyy},case when ""YEAR_INQCONF"" = {2:yyyy} then 1 end as countof{2:yyyy} from alldatadistinct", myFirstDate, myLastDate, myFirstDateShipment)

        'Dim sqlstr3 As String = String.Format("with shipment as (select distinct date_part('Year',miropostingdate)::integer as ""YEAR_INQCONF"",so.soldtoparty as ""Sold To Party SAP Code"", m.vendorcode as ""SAP Vendor Code"",case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP""," &
        '                                      " case when m.vendorcode < 90000000 then 'non-group' else 'group' end as ""Vendor Type""" &
        '                                      " from pomiro p" &
        '                                      " left join miro m on m.miroid = p.miroid left join podtl d on d.podtlid = p.podtlid left join ekko e on e.po = d.pohd " &
        '                                      " Left Join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup left join cxsebpodtl pd on pd.sebasiapono = d.pohd and pd.polineno = d.polineno" &
        '                                      " left join cxrelsalesdocpo r on r.cxsebpodtlid = pd.cxsebpodtlid left join cxsalesorderdtl sd on sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
        '                                      " left join cxsalesorder so on so.sebasiasalesorder = sd.sebasiasalesorder where miropostingdate >= '{2:yyyy-MM-dd}' and miropostingdate <= '{1:yyyy-MM-dd}' and pg.groupsbuid <= 2 )," &
        '                                      " s1 as (select ""YEAR_INQCONF"",""Sold To Party SAP Code"",""SAP Vendor Code"",case ""Sold To Party SAP Code"" when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",""Vendor Type"" from shipment)," &
        '                                      " alldatadistinct as( select distinct ""YEAR_INQCONF"",""Sold To Party SAP Code"",null::bigint as ""SAP Vendor Code"",""Shipment Type"",null::text as ""Vendor Type"",'#of market'::text as details from s1 " &
        '                                      " union all select distinct ""YEAR_INQCONF"",null::bigint as ""Sold To Party SAP Code"",""SAP Vendor Code"",""Shipment Type"",""Vendor Type"",'#of supplier' from s1)" &
        '                                      " select *,case when ""YEAR_INQCONF"" = {1:yyyy} then 1 end as countof{1:yyyy},case when ""YEAR_INQCONF"" = {0:yyyy} then 1 end as countof{0:yyyy},case when ""YEAR_INQCONF"" = {2:yyyy} then 1 end as countof{2:yyyy} from alldatadistinct", myFirstDate, myLastDate, myFirstDateShipment)

        sqlstr.Append(String.Format("with firmorder as " &
            " (SELECT case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP"",'FIRMORDER'::character varying as ""TYPE"", so.soldtoparty  AS ""Sold To Party SAP Code"", c2.customername AS ""Sold To Party Name""," &
            " pd.shiptoparty as ""Ship To Party"",c1.customername as ""Ship To Party Name"", e.vendorcode AS ""SAP Vendor Code"",v.vendorname AS ""Vendor Name"", " &
            " mm.rri AS ""Research and Industry Responsible"", sbu.sbuname AS ""SBU"", sbu1.sbuname AS ""BU"", sbu2.sbuname AS familysbu, pd.cmmf AS ""CMMF""," &
            " mm.familylv1 AS ""Commercial Family CG2 n"", f.familyname, b.brandname AS ""Brand Name CG2 n"", mm.commref AS ""CommercialRef"", " &
            " mm.materialdesc AS ""Description"", (getpurchasevalue(pd.osqty,pd.fob) / getexchangerate(e.currency)) AS ""Purchase Value""," &
            " pd.osqty,getyearinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{1:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""YEAR_INQCONF""," &
            " Null::numeric as ""Year Y-1 Purchase Value"", " &
            " (getyearpurchasevalueinc({1:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty,pd.fob) / getexchangerate(e.currency)) as ""Year Y Purchase Value"" ," &
            " getyearqty({0:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus," &
            " pd.comments,pd.osqty) as ""Year Y-1 QTY"",getyearqty({1:yyyy},od.currentconfirmedetd,sd.currentinquiryetd,confs.confirmationstatus,pd.comments,pd.osqty) as ""Year Y QTY""," &
            " getmonthinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{1:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""MONTH INQCONF""," &
            " ssm.officersebname as ""SPM""  ,adjustedmonth(confirmationstatus,od.currentconfirmedetd,pd.comments) as adjustedmonth, od.currentconfirmedetd,sd.currentinquiryetd,getconfirmstatus(confs.confirmationstatus,pd.comments,od.currentconfirmedetd) as confirmstatus,pd.comments,pd.sebasiapono,pd.polineno,sp.sopdescription as industrialfamily ,e.currency,getpurchasevalue(pd.osqty,pd.fob) as ""Purchase Value (OC)"" , '1'::integer as" &
            " isreal,cf.flow,cf.continent,cf.continent_group,cf.continent_group_emea FROM cxsebodtp od LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid " &
            " LEFT JOIN cxsalesorder so ON so.sebasiasalesorder = sd.sebasiasalesorder LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono " &
            " LEFT JOIN customer c1 ON c1.customercode = pd.shiptoparty LEFT JOIN customer c2 ON c2.customercode = so.soldtoparty " &
            " LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid " &
            " LEFT JOIN cxconfstatus confs ON confs.cxconfid = c.cxconfid Left join ekko e on e.po = pd.sebasiapono Left join vendor v on v.vendorcode = e.vendorcode " &
            " LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl LEFT JOIN materialmaster mm  ON mm.cmmf = pd.cmmf left join sspcmmfsop sc on sc.cmmf = mm.cmmf " &
            " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid LEFT JOIN brand b ON b.brandid = mm.brandid " &
            " LEFT JOIN activity ac ON ac.activitycode = mm.rri LEFT JOIN sbu ON sbu.sbuid = ac.sbuidsp LEFT JOIN sbu sbu1 ON sbu1.sbuid = ac.sbuidlg " &
            " LEFT JOIN family f ON f.familyid = mm.familylv1 LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu LEFT JOIN purchasinggroup pg ON pg.purchasinggroup = e.purchasinggroup::bpchar " &
            " Left Join customerflow cf on cf.soldtoparty = so.soldtoparty and cf.shiptoparty = pd.shiptoparty" &
            " WHERE od.ordertype = 'Header'::bpchar AND ph.receptiondate <= sd.inquiryetd and pd.osqty > 0  " &
            " and pg.groupsbuid <= 2 ORDER BY pg.groupsbuid,ph.receptiondate )," &
            " shipment as " &
            " (select miropostingdate,case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP"",'SHIPMENT'::character varying as type,so.soldtoparty,cu1.customername, pd.shiptoparty,cu2.customername, m.vendorcode,v.vendorname,mm.rri,sbu.sbuname,sbu1.sbuname," &
            " sbu2.sbuname as familysbu,d.cmmf,mm.familylv1,f.familyname,b.brandname,mm.commref,mm.materialdesc,(p.amount / getexchangerate(e.currency))as amount,p.qty::integer," &
            " date_part('Year',miropostingdate)::integer,(getyearpurchasevalue({0:yyyy},miropostingdate,p.amount)/getexchangerate(e.currency)) as lastyear," &
            " (getyearpurchasevalue({1:yyyy},miropostingdate,p.amount)/getexchangerate(e.currency)) as thisyear,getyearqty({0:yyyy},miropostingdate,p.qty::integer)," &
            " getyearqty({1:yyyy},miropostingdate,p.qty::integer),date_part('Month',miropostingdate),ssm.officersebname, null::integer,null::date,null::date,null::character varying,null::character varying,hd.pohd,d.polineno,sp.sopdescription, e.currency, p.amount," &
            " '1'::integer as isreal,cf.flow,cf.continent,cf.continent_group,cf.continent_group_emea from pomiro " &
            " p left join miro m on m.miroid = p.miroid left join podtl d on d.podtlid = p.podtlid left join vendor v on v.vendorcode = m.vendorcode " &
            " left join ekko e on e.po = d.pohd left join pohd hd on hd.pohd = d.pohd " &
            " left join materialmaster mm on mm.cmmf = d.cmmf left join sspcmmfsop sc on sc.cmmf = d.cmmf " &
            " left join sspsopfamilies sp on sp.sspsopfamilyid = sc.sopfamilyid " &
            " LEFT JOIN brand b ON b.brandid = mm.brandid LEFT JOIN officerseb ssm ON ssm.ofsebid = v.ssmidpl " &
            " left join cxsebpodtl pd on pd.sebasiapono = d.pohd and pd.polineno = d.polineno  " &
            " left join cxrelsalesdocpo r on r.cxsebpodtlid = pd.cxsebpodtlid " &
            " left join cxsalesorderdtl sd on sd.cxsalesorderdtlid = r.cxsalesorderdtlid " &
            " left join cxsalesorder so on so.sebasiasalesorder = sd.sebasiasalesorder " &
            " left join customer cu1 on cu1.customercode = so.soldtoparty " &
            " left join customer cu2 on cu2.customercode = pd.shiptoparty " &
            " LEFT JOIN activity ac ON  ac.activitycode = mm.rri " &
            " LEFT JOIN sbu ON sbu.sbuid = ac.sbuid " &
            " LEFT JOIN family f ON f.familyid = mm.familylv1 " &
            " LEFT JOIN sbu sbu1 on sbu1.sbuid = ac.sbuidlg " &
            " LEFT JOIN sbusap sbu2 ON sbu2.sbuid = mm.sbu " &
            " Left Join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup " &
            " Left Join customerflow cf on cf.soldtoparty = so.soldtoparty and cf.shiptoparty = pd.shiptoparty" &
            " where miropostingdate >= '{2:yyyy-MM-dd}' and miropostingdate <= '{1:yyyy-MM-dd}' and pg.groupsbuid <= 2" &
            " order by pg.groupsbuid,p.pomiroid)," &
            " f1 as (select case ""Sold To Party SAP Code"" when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",case when ""SAP Vendor Code"" < 90000000 then 'non-group' else 'group' end as ""Vendor Type"" , to_date(""YEAR_INQCONF"" || '-' || ""MONTH INQCONF"" || '-1','yyyy-MM-dd' ) as txdate,* from firmorder)," &
            " s1 as (select case soldtoparty when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",case when vendorcode < 90000000 then 'non-group' else 'group' end as ""Vendor Type"" ,* from shipment)," &
            " fs as (select case ""Shipment Type"" when 'FG' then 'FG+GSE' when 'GSE' then 'FG+GSE' else ""Shipment Type"" end as ""Shipment Type (FG+GSE)"",f1.* from f1 " &
            " union all select case ""Shipment Type"" when 'FG' then 'FG+GSE' when 'GSE' then 'FG+GSE' else ""Shipment Type"" end as ""Shipment Type (FG+GSE)"",* from s1) " &
            " select fs.*,pd.cvalue as market,case ""YEAR_INQCONF"" when {0:yyyy} then 1 end as ""count Y-1"",case ""YEAR_INQCONF"" when {1:yyyy} then 1 end as ""count Y"" from fs left join soldtomarket stm on stm.soldtoparty = fs.""Sold To Party SAP Code"" left join paramdt pd on pd.paramdtid = stm.market", myFirstDate, myLastDate, myFirstDateShipment))

        Dim sqlstr2 As String = String.Format("with firmorder as(select  distinct getyearinqconf(od.currentconfirmedetd,sd.currentinquiryetd,'{2:yyyy-MM-dd}',confs.confirmationstatus,pd.comments) as ""YEAR_INQCONF""," &
                                              " so.soldtoparty  AS ""Sold To Party SAP Code"", e.vendorcode AS ""SAP Vendor Code"",case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP"",case when e.vendorcode < 90000000 then 'non-group' else 'group' end as vendortype " &
                                              " FROM cxsebodtp od LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
                                              " LEFT JOIN cxsalesorder so ON so.sebasiasalesorder = sd.sebasiasalesorder LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid " &
                                              " LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid LEFT JOIN cxconfstatus confs ON confs.cxconfid = c.cxconfid" &
                                              " Left join ekko e on e.po = pd.sebasiapono LEFT JOIN purchasinggroup pg ON pg.purchasinggroup = e.purchasinggroup::bpchar  WHERE od.ordertype = 'Header'::bpchar AND ph.receptiondate <= sd.inquiryetd and pd.osqty > 0  and pg.groupsbuid <= 2 )," &
                                              " f1 as(select case ""Sold To Party SAP Code"" when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",f.*" &
                                              " from firmorder f), " &
                                              " fall as(select distinct ""YEAR_INQCONF"",""Sold To Party SAP Code"",""SAP Vendor Code"",""Shipment Type"",vendortype from f1" &
                                              " where ""YEAR_INQCONF"" = {1:yyyy})," &
                                              " shipment as (select distinct date_part('Year',miropostingdate)::integer as ""YEAR_INQCONF"",so.soldtoparty as ""Sold To Party SAP Code"", m.vendorcode as ""SAP Vendor Code"",case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP""," &
                                              " case when m.vendorcode < 90000000 then 'non-group' else 'group' end as ""Vendor Type""" &
                                              " from pomiro p" &
                                              " left join miro m on m.miroid = p.miroid left join podtl d on d.podtlid = p.podtlid left join ekko e on e.po = d.pohd " &
                                              " Left Join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup left join cxsebpodtl pd on pd.sebasiapono = d.pohd and pd.polineno = d.polineno" &
                                              " left join cxrelsalesdocpo r on r.cxsebpodtlid = pd.cxsebpodtlid left join cxsalesorderdtl sd on sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
                                              " left join cxsalesorder so on so.sebasiasalesorder = sd.sebasiasalesorder where miropostingdate >= '{2:yyyy-MM-dd}' and miropostingdate <= '{1:yyyy-MM-yy}' and pg.groupsbuid <= 2 )," &
                                              " s1 as (select ""YEAR_INQCONF"",""Sold To Party SAP Code"",""SAP Vendor Code"",case ""Sold To Party SAP Code"" when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",""Vendor Type"" from shipment)," &
                                              " alldata as (select * from fall  union all select * from s1), " &
                                              " alldatadistinct as( select distinct ""YEAR_INQCONF"",""Sold To Party SAP Code"",null::bigint as ""SAP Vendor Code"",""Shipment Type"",null::text as vendortype,'#of market'::text as details from alldata" &
                                              " union all " &
                                              " select distinct ""YEAR_INQCONF"",null::bigint as ""Sold To Party SAP Code"",""SAP Vendor Code"",""Shipment Type"",vendortype,'#of supplier' from alldata)" &
                                              " select *,case when ""YEAR_INQCONF"" = {1:yyyy} then 1 end as countofy,case when ""YEAR_INQCONF"" = {0:yyyy} then 1 end as ""countofy-1"",case when ""YEAR_INQCONF"" = {2:yyyy} then 1 end as ""countofy-2"" from alldatadistinct", myFirstDate, myLastDate, myFirstDateShipment)

        Dim sqlstr3 As String = String.Format("with shipment as (select distinct date_part('Year',miropostingdate)::integer as ""YEAR_INQCONF"",so.soldtoparty as ""Sold To Party SAP Code"", m.vendorcode as ""SAP Vendor Code"",case pg.groupsbuid when 1 then 'FG' when '2' then 'CP' end as ""FG/CP""," &
                                              " case when m.vendorcode < 90000000 then 'non-group' else 'group' end as ""Vendor Type""" &
                                              " from pomiro p" &
                                              " left join miro m on m.miroid = p.miroid left join podtl d on d.podtlid = p.podtlid left join ekko e on e.po = d.pohd " &
                                              " Left Join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup left join cxsebpodtl pd on pd.sebasiapono = d.pohd and pd.polineno = d.polineno" &
                                              " left join cxrelsalesdocpo r on r.cxsebpodtlid = pd.cxsebpodtlid left join cxsalesorderdtl sd on sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
                                              " left join cxsalesorder so on so.sebasiasalesorder = sd.sebasiasalesorder where miropostingdate >= '{2:yyyy-MM-dd}' and miropostingdate <= '{1:yyyy-MM-dd}' and pg.groupsbuid <= 2 )," &
                                              " s1 as (select ""YEAR_INQCONF"",""Sold To Party SAP Code"",""SAP Vendor Code"",case ""Sold To Party SAP Code"" when 99005151 then 'GSE' when 99009500 then 'SIS' else ""FG/CP"" end as ""Shipment Type"",""Vendor Type"" from shipment)," &
                                              " alldatadistinct as( select distinct ""YEAR_INQCONF"",""Sold To Party SAP Code"",null::bigint as ""SAP Vendor Code"",""Shipment Type"",null::text as ""Vendor Type"",'#of market'::text as details from s1 " &
                                              " union all select distinct ""YEAR_INQCONF"",null::bigint as ""Sold To Party SAP Code"",""SAP Vendor Code"",""Shipment Type"",""Vendor Type"",'#of supplier' from s1)" &
                                              " select *,case when ""YEAR_INQCONF"" = {1:yyyy} then 1 end as countofy,case when ""YEAR_INQCONF"" = {0:yyyy} then 1 end as ""countofy-1"",case when ""YEAR_INQCONF"" = {2:yyyy} then 1 end as ""countofy-2"" from alldatadistinct", myFirstDate, myLastDate, myFirstDateShipment)
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = String.Format("Turnover", Date.Today)
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myqueryworksheet = New QueryWorksheet With {.DataSheet = 10,
                                                .SheetName = "RAWDATA",
                                                .Sqlstr = sqlstr.ToString
                                                }
            myQueryWorksheetList.Add(myqueryworksheet)

            myqueryworksheet = New QueryWorksheet With {.DataSheet = 11,
                                                            .SheetName = "header(order & shipment)",
                                                            .Sqlstr = sqlstr2
                                                            }
            myQueryWorksheetList.Add(myqueryworksheet)

            myqueryworksheet = New QueryWorksheet With {.DataSheet = 12,
                                                            .SheetName = "header(shipment)",
                                                            .Sqlstr = sqlstr3
                                                            }
            myQueryWorksheetList.Add(myqueryworksheet)

            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback)
            Dim myreport As New ExportToExcelFile(Me, myQueryWorksheetList, filename, reportname, "\templates\ReportTemplate004.xltx", mycallback, PivotCallback)

            'Dim myreport As New ExportToExcelFile(Me, sqlstr.ToString, filename, reportname, mycallback, PivotCallback)
            myreport.Run(Me, e)
        End If

    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        Dim oWB As Excel.Workbook = DirectCast(sender, Excel.Workbook)

        'Refresh All Pivot Table
        oWB.RefreshAll()




        ''Worksheet 1 chart (by continent)
        'oWB.Worksheets(1).select()
        'Dim osheet As Excel.Worksheet = oWB.Worksheets(1)

        'Dim myChart As Excel.Chart = osheet.ChartObjects("Chart 1").Chart
        'myChart.ChartTitle.Text = String.Format("{0} YTD vs {1} YTD (by TXN)", myFirstDate.Year, myLastDate.Year)

        'myChart = osheet.ChartObjects("Chart 2").Chart
        'myChart.ChartTitle.Text = String.Format("{0} YTD vs {1} YTD (by Value)", myFirstDate.Year, myLastDate.Year)

        'myChart = osheet.ChartObjects("Chart 5").Chart
        'For i = 0 To myChart.SeriesCollection.count - 1
        '    myChart.SeriesCollection(1).delete()
        'Next
        ''Get Count of Market


        'osheet = oWB.Worksheets(5)
        'oWB.Names.Add("DistPCTName", RefersToR1C1:=String.Format("='{0}'!R6C2", osheet.Name))
        'oWB.Names.Add("DistPCTValue", RefersTo:=String.Format("=offset('{0}'!R7C2,0,0,counta('{0}'!C1)-5,1)", osheet.Name))

        'osheet = oWB.Worksheets(4)
        'oWB.Names.Add("ColumnLabel", RefersTo:=String.Format("=counta('{0}'!R7)-1", osheet.Name))
        'Dim obj As Excel.Range
        'obj = osheet.Range("Z1")
        'obj.FormulaR1C1 = "=counta(R7)-1"

        'Dim myCountField = obj.Value

        'Dim myCountRow = osheet.PivotTables("PivotTable1").PivotFields("continent_group_emea").PivotItems.count
        'If osheet.PivotTables("PivotTable1").PivotFields("continent_group_emea").PivotItems(myCountRow).value = "(blank)" Then
        '    myCountRow = myCountRow - 1
        'End If
        'oWB.Names.Add("SeriesXValue", RefersTo:=String.Format("=offset('{0}'!R8C1,0,0,counta('{0}'!C1)-6,1)", osheet.Name))
        'For i = 1 To myCountField
        '    oWB.Names.Add("SeriesName" & i, RefersToR1C1:=String.Format("='{0}'!R7C{1}", osheet.Name, i + 1))
        '    oWB.Names.Add("SeriesValue" & i, RefersTo:=String.Format("=offset('{0}'!R8C{1},0,0,counta('{0}'!C1)-6,1)", osheet.Name, i + 1))
        'Next



        'osheet = oWB.Worksheets(1)
        'Dim myDataLabel As Excel.DataLabels
        'For i = 1 To myCountField
        '    myChart.SeriesCollection.NewSeries()
        '    myChart.SeriesCollection(i).Name = String.Format("='{0}'!SeriesName{1}", osheet.Name, i)
        '    myChart.SeriesCollection(i).Values = String.Format("='{0}'!SeriesValue{1}", osheet.Name, i)
        '    myChart.SeriesCollection(i).XValues = String.Format("='{0}'!SeriesXValue", osheet.Name, i)
        '    myDataLabel = myChart.SeriesCollection(i).DataLabels
        '    myDataLabel.ShowSeriesName = True
        '    myDataLabel.ShowValue = False
        'Next

        'osheet = oWB.Worksheets(5)

        'myChart.SeriesCollection.NewSeries()
        'myChart.SeriesCollection(myCountField + 1).Name = String.Format("='{0}'!DistPCTName", osheet.Name)
        'myChart.SeriesCollection(myCountField + 1).Values = String.Format("='{0}'!DistPCTValue", osheet.Name)
        'myChart.SeriesCollection(myCountField + 1).XValues = String.Format("='{0}'!SeriesXValue", osheet.Name)

        'myDataLabel = myChart.SeriesCollection(myCountField + 1).DataLabels
        'myDataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionInsideBase

        'Worksheet2 shipment (by continent)
        oWB.Worksheets(2).select()
        Dim osheet As Excel.Worksheet = oWB.Worksheets(2)

        osheet.PivotTables("PivotTable2").PivotFields("YEAR_INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable2").PivotFields("YEAR_INQCONF").PivotItems

            If pi.Value = myFirstDate.Year Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable3").PivotFields("YEAR_INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable3").PivotFields("YEAR_INQCONF").PivotItems
            If pi.Value = myFirstDate.Year Or pi.Value = myFirstDate.Year + 1 Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable3").PivotFields("MONTH INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable3").PivotFields("MONTH INQCONF").PivotItems

            If pi.Value <= myLastDate.Month Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.Cells(6, 17).value = String.Format("{0} YTD", osheet.Cells(6, 17).value)
        osheet.Cells(6, 19).value = String.Format("{0} YTD", osheet.Cells(6, 19).value)

        'worksheet2 all continent
        oWB.Worksheets(3).select()
        osheet = oWB.Worksheets(3)
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable1").PivotFields("continent_group_emea").PivotItems()
            If pi.Value = "(blank)" Then
                pi.Visible = False            
            End If
        Next

        osheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").PivotItems

            If pi.Value <= myLastDate.Month Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next


        osheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").PivotItems()
            pi.Visible = False
            If pi.Value = myFirstDate.Year Or pi.Value = myLastDate.Year Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable2").PivotFields("continent_group_emea").PivotItems()
            If pi.Value = "(blank)" Then
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable2").PivotFields("MONTH INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable2").PivotFields("MONTH INQCONF").PivotItems

            If pi.Value <= myLastDate.Month Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        'worksheet2 distribution continent
        oWB.Worksheets(4).select()
        osheet = oWB.Worksheets(4)

        osheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").PivotItems()

            If pi.Value = myLastDate.Year Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        'osheet.PivotTables("PivotTable1").pivotfields("MONTH INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        'osheet.PivotTables("PivotTable1").pivotfields("MONTH INQCONF").Position = 1

        osheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").PivotItems

            If pi.Value <= myLastDate.Month Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next



        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable1").PivotFields("continent_group_emea").PivotItems()
            If pi.Value = "(blank)" Then
                pi.Visible = False
            End If
        Next

        'worksheet2 distribution by %
        oWB.Worksheets(5).select()
        osheet = oWB.Worksheets(5)
        osheet.PivotTables("PivotTable2").PivotFields("YEAR_INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable2").PivotFields("YEAR_INQCONF").PivotItems()

            If pi.Value = myLastDate.Year Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        'osheet.PivotTables("PivotTable2").pivotfields("MONTH INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        'osheet.PivotTables("PivotTable2").pivotfields("MONTH INQCONF").Position = 1

        osheet.PivotTables("PivotTable2").PivotFields("MONTH INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable2").PivotFields("MONTH INQCONF").PivotItems

            If pi.Value <= myLastDate.Month Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next


        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable2").PivotFields("continent_group_emea").PivotItems()
            If pi.Value = "(blank)" Then
                pi.Visible = False
            End If
        Next







        ''Worksheet6 Chart (order & shipment)
        'oWB.Worksheets(6).select()
        'osheet = oWB.Worksheets(6)
        'myChart = osheet.ChartObjects("Chart 1").Chart
        'myChart.ChartTitle.Text = String.Format("{0} YTD vs {1} YTD (by TXN) {2} ALL", myFirstDate.Year, myLastDate.Year, Chr(13))
        'myChart = osheet.ChartObjects("Chart 2").Chart
        'myChart.ChartTitle.Text = String.Format("{0} YTD vs {1} YTD (by TXN) {2} Shipment", myFirstDate.Year, myLastDate.Year, Chr(13))
        'myChart = osheet.ChartObjects("Chart 8").Chart
        'myChart.ChartTitle.Text = String.Format("{0} YTD vs {1} YTD (by Value) {2} ALL", myFirstDate.Year, myLastDate.Year, Chr(13))
        'myChart = osheet.ChartObjects("Chart 15").Chart
        'myChart.ChartTitle.Text = String.Format("{0} YTD vs {1} YTD (by Value) {2} Shipment", myFirstDate.Year, myLastDate.Year, Chr(13))


        'Worksheet7 order fulfillment & shipment
        oWB.Worksheets(7).select()
        osheet = oWB.Worksheets(7)
        osheet.Cells(3, 3).value = Year(myFirstDateShipment)
        osheet.Cells(3, 5).value = Year(myFirstDate)
        osheet.Cells(3, 7).value = String.Format("{0:yyyy} YTD", myLastDate)

        osheet.Cells(3, 18).value = Year(myFirstDateShipment)
        osheet.Cells(3, 20).value = Year(myFirstDate)
        osheet.Cells(3, 22).value = String.Format("{0:yyyy} YTD", myLastDate)

        'worksheet8 order & shipment chart data
        oWB.Worksheets(8).select()
        osheet = oWB.Worksheets(8)

        osheet.PivotTables("PivotTable13").PivotFields("MONTH INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable13").PivotFields("MONTH INQCONF").PivotItems

            If pi.Value <= myLastDate.Month Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable13").PivotFields("YEAR_INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable13").PivotFields("YEAR_INQCONF").PivotItems()

            If pi.Value = myLastDate.Year Or pi.Value = myFirstDate.Year Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").PivotItems

            If pi.Value <= myLastDate.Month Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable14").PivotFields("MONTH INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable14").PivotFields("MONTH INQCONF").PivotItems

            If pi.Value <= myLastDate.Month Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable14").PivotFields("YEAR_INQCONF").clearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable14").PivotFields("YEAR_INQCONF").PivotItems()

            If pi.Value = myLastDate.Year Or pi.Value = myFirstDate.Year Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next
        osheet.PivotTables("PivotTable2").PivotFields("MONTH INQCONF").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable2").PivotFields("MONTH INQCONF").PivotItems

            If pi.Value <= myLastDate.Month Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        'worksheet8 order & shipment table
        oWB.Worksheets(9).select()
        osheet = oWB.Worksheets(9)
        osheet.PivotTables("PivotTable1").PivotFields("details").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable1").PivotFields("details").PivotItems()

            If pi.Value = "#of market" Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable1").PivotFields("Shipment Type").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable1").PivotFields("Shipment Type").PivotItems()

            If pi.Value = "FG" Or pi.Value = "CP" Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable3").PivotFields("Shipment Type").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable3").PivotFields("Shipment Type").PivotItems()

            If pi.Value = "FG" Or pi.Value = "CP" Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable4").PivotFields("details").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable4").PivotFields("details").PivotItems()

            If pi.Value = "#of supplier" Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable4").PivotFields("vendortype").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable4").PivotFields("vendortype").PivotItems()
            If pi.Value = "non-group" Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable6").PivotFields("details").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable6").PivotFields("details").PivotItems()
            If pi.Value = "#of supplier" Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable6").PivotFields("Shipment Type").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable6").PivotFields("Shipment Type").PivotItems()

            If pi.Value = "FG" Or pi.Value = "CP" Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable6").PivotFields("vendortype").ClearAllFilters()
        For Each pi As Excel.PivotItem In osheet.PivotTables("PivotTable6").PivotFields("vendortype").PivotItems()

            If pi.Value = "non-group" Then
                pi.Visible = True
            Else
                pi.Visible = False
            End If
        Next

        'Worksheet 1 chart (by continent)
        oWB.Worksheets(1).select()
        'Dim osheet As Excel.Worksheet = oWB.Worksheets(1)
        osheet = oWB.Worksheets(1)

        Dim myChart As Excel.Chart = osheet.ChartObjects("Chart 1").Chart
        myChart.ChartTitle.Text = String.Format("{0} YTD vs {1} YTD (by TXN)", myFirstDate.Year, myLastDate.Year)

        myChart = osheet.ChartObjects("Chart 2").Chart
        myChart.ChartTitle.Text = String.Format("{0} YTD vs {1} YTD (by Value)", myFirstDate.Year, myLastDate.Year)

        myChart = osheet.ChartObjects("Chart 5").Chart
        For i = 0 To myChart.SeriesCollection.count - 1
            myChart.SeriesCollection(1).delete()
        Next
        'Get Count of Market


        osheet = oWB.Worksheets(5)
        oWB.Names.Add("DistPCTName", RefersToR1C1:=String.Format("='{0}'!R6C2", osheet.Name))
        oWB.Names.Add("DistPCTValue", RefersTo:=String.Format("=offset('{0}'!R7C2,0,0,counta('{0}'!C1)-5,1)", osheet.Name))

        osheet = oWB.Worksheets(4)
        oWB.Names.Add("ColumnLabel", RefersTo:=String.Format("=counta('{0}'!R7)-1", osheet.Name))
        Dim obj As Excel.Range
        obj = osheet.Range("Z1")
        obj.FormulaR1C1 = "=counta(R7)-1"

        Dim myCountField = obj.Value

        Dim myCountRow = osheet.PivotTables("PivotTable1").PivotFields("continent_group_emea").PivotItems.count
        If osheet.PivotTables("PivotTable1").PivotFields("continent_group_emea").PivotItems(myCountRow).value = "(blank)" Then
            myCountRow = myCountRow - 1
        End If
        oWB.Names.Add("SeriesXValue", RefersTo:=String.Format("=offset('{0}'!R8C1,0,0,counta('{0}'!C1)-6,1)", osheet.Name))
        For i = 1 To myCountField
            oWB.Names.Add("SeriesName" & i, RefersToR1C1:=String.Format("='{0}'!R7C{1}", osheet.Name, i + 1))
            oWB.Names.Add("SeriesValue" & i, RefersTo:=String.Format("=offset('{0}'!R8C{1},0,0,counta('{0}'!C1)-6,1)", osheet.Name, i + 1))
        Next



        osheet = oWB.Worksheets(1)
        Dim myDataLabel As Excel.DataLabels
        For i = 1 To myCountField
            myChart.SeriesCollection.NewSeries()
            myChart.SeriesCollection(i).Name = String.Format("='{0}'!SeriesName{1}", osheet.Name, i)
            myChart.SeriesCollection(i).Values = String.Format("='{0}'!SeriesValue{1}", osheet.Name, i)
            myChart.SeriesCollection(i).XValues = String.Format("='{0}'!SeriesXValue", osheet.Name, i)
            myDataLabel = myChart.SeriesCollection(i).DataLabels
            myDataLabel.ShowSeriesName = True
            myDataLabel.ShowValue = False
        Next

        osheet = oWB.Worksheets(5)

        myChart.SeriesCollection.NewSeries()
        myChart.SeriesCollection(myCountField + 1).Name = String.Format("='{0}'!DistPCTName", osheet.Name)
        myChart.SeriesCollection(myCountField + 1).Values = String.Format("='{0}'!DistPCTValue", osheet.Name)
        myChart.SeriesCollection(myCountField + 1).XValues = String.Format("='{0}'!SeriesXValue", osheet.Name)

        myDataLabel = myChart.SeriesCollection(myCountField + 1).DataLabels
        myDataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionInsideBase

        'Worksheet6 Chart (order & shipment)
        oWB.Worksheets(6).select()
        osheet = oWB.Worksheets(6)
        myChart = osheet.ChartObjects("Chart 1").Chart
        myChart.ChartTitle.Text = String.Format("{0} YTD vs {1} YTD (by TXN) {2} ALL", myFirstDate.Year, myLastDate.Year, Chr(13))
        myChart = osheet.ChartObjects("Chart 2").Chart
        myChart.ChartTitle.Text = String.Format("{0} YTD vs {1} YTD (by TXN) {2} Shipment", myFirstDate.Year, myLastDate.Year, Chr(13))
        myChart = osheet.ChartObjects("Chart 8").Chart
        myChart.ChartTitle.Text = String.Format("{0} YTD vs {1} YTD (by Value) {2} ALL", myFirstDate.Year, myLastDate.Year, Chr(13))
        myChart = osheet.ChartObjects("Chart 15").Chart
        myChart.ChartTitle.Text = String.Format("{0} YTD vs {1} YTD (by Value) {2} Shipment", myFirstDate.Year, myLastDate.Year, Chr(13))

    End Sub

    Public Function getLastdate(ByVal month As Integer, ByVal year As Integer) As Date
        'If month = 12 Then
        '    year = year + 1
        '    month = 2
        'ElseIf month + 2 > 12 Then
        '    year = year + 1
        '    month = month + 2 - 12
        'Else
        '    month = month + 2
        'End If

        'getLastdate = CDate(String.Format("{0}-{1}-1", year, month)).AddDays(-1)
        getLastdate = CDate(String.Format("{0}-{1}-{2}", year, month, DateTime.DaysInMonth(year, month)))
    End Function
End Class