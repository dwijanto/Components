Imports Components.PublicClass
Imports DJLib.Dbtools
Imports System.Threading
Public Class FormWORAllLines
    Dim mythread As New Thread(AddressOf doQuery)
    Dim AccessTableName As String = "tbl_WOR_FG"
    Dim AccessDbFullPath As String
    'Dim myreport As ExportToExcelFile
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        AccessDbFullPath = TextBox1.Text
        If CheckBox1.Checked Then
            If TextBox1.Text = "" Then
                MessageBox.Show("Please select Access MDB fullpath location.")
                Exit Sub
            End If

            If RadioButton3.Checked Then
                AccessTableName = "tbl_WOR_Comp"
            End If
        End If

        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""


        Dim mymessage As String = String.Empty

        ' Dim myuser As String = String.Empty
        ' Dim dateupload As Date
        Dim sqlstr As String
        Dim Dept As String = "FG"

        Dim alldata = "cxallworfg"

        If RadioButton3.Checked Then
            alldata = "cxallworcomp"
            Dept = "COMP"
        ElseIf RadioButton5.Checked Then
            Dept = "SIS"
        End If

        Dim openorder = " where osqty > 0"
        If RadioButton2.Checked Then
            openorder = " where receptiondate >= " & DateFormatyyyyMMdd(DateTimePicker1.Value.Date) & " and receptiondate <= " & DateFormatyyyyMMdd(DateTimePicker2.Value.Date)
        End If
        'ToolStripProgressBar1.

        Select Case Dept
            Case "SIS"
                sqlstr = " Select ordertype as ""Header / Item"", shiptoparty as ""Customer code"", shiptopartyname::character varying as ""Customer name"",soldtoparty as ""Sold To Party SAP Code"",soldtopartyname::character varying as ""Sold To Party Name"", sao as ""Sales Administration Office"", customerorderno as ""Customer Order No."", sebasiasalesorder as ""Seb Asia Sales Order"", solineno as ""Item No"", vendorcode as ""Vendor Code"",vendorname::character varying as ""Vendor"", ssm::character varying as ""SSM"", rir::character varying as ""Research and Industry Responsible"", cmmf as ""CMMF"", comfam as ""Commercial family CG2 n"", brand as ""Brand CG2"", brandname::character varying as ""Brand"",itemid::character varying as ""Item ID"", materialdesc::character varying as ""Description"", orderstatus as ""Order Status"", latestupdate as ""Latest Update"",updatesince as ""Updated since 7""" & _
                 ", curinq as ""Cur./Inq."", receptiondate as ""Reception Date"", fob::numeric as ""FOB"", unittp::numeric as ""Unit T.P."", inquiryeta as ""Inquiry ETA"", inquiryetd as ""Inquiry ETD"", inquiryqty as ""Inquiry Qty"", currentinquiryeta as ""Current Inquiry ETA"", currentinquiryetd as ""Current Inquiry ETD"", currentinquiryqty as ""Current Inquiry Qty"", confirmationstatus as ""Confirmation status"", stconfirmedetd as ""1st Confirmed ETD"", stconfirmedqty as ""1st Confirmed Qty"",currentconfirmedeta as ""Current Confirmed ETA"", currentconfirmedetd as ""Current Confirmed ETD"",currentconfirmedqty as ""Current Confirmed Qty"", deliveredqty as ""Delivered Qty"", shipdate as ""Ship.Date"",shipdateeta as ""Ship.Date ETA"", osqty as ""O/S Qty"", sebasiapono as ""Seb Asia P.O. No."", polineno as ""Line No"", ctrno as ""CTR No."", boatid as ""Boat ID"", packinglist as ""Packing List"", shipfrom as ""Ship From "",finalcustomerorder as ""Final Customer Order"" , comments as ""Comments"" " & _
                 ", sbu::character varying as ""SBU"",bu::character varying as ""BU"",familysbu as ""Family SBU"", purchasinggroup as ""PGr"",currency as ""Crcy"",sp as ""SP"",spm as ""SPM"",dicustomerorder as ""DI customer order""" &
                 " from cxallworfg " & openorder & " and purchasinggroup = 'FOC' " & " union all " &
                 " Select ordertype as ""Header / Item"", shiptoparty as ""Customer code"", shiptopartyname::character varying as ""Customer name"",soldtoparty as ""Sold To Party SAP Code"",soldtopartyname::character varying as ""Sold To Party Name"", sao as ""Sales Administration Office"", customerorderno as ""Customer Order No."", sebasiasalesorder as ""Seb Asia Sales Order"", solineno as ""Item No"", vendorcode as ""Vendor Code"",vendorname::character varying as ""Vendor"", ssm::character varying as ""SSM"", rir::character varying as ""Research and Industry Responsible"", cmmf as ""CMMF"", comfam as ""Commercial family CG2 n"", brand as ""Brand CG2"", brandname::character varying as ""Brand"",itemid::character varying as ""Item ID"", materialdesc::character varying as ""Description"", orderstatus as ""Order Status"", latestupdate as ""Latest Update"",updatesince as ""Updated since 7""" & _
                 ", curinq as ""Cur./Inq."", receptiondate as ""Reception Date"", fob::numeric as ""FOB"", unittp::numeric as ""Unit T.P."", inquiryeta as ""Inquiry ETA"", inquiryetd as ""Inquiry ETD"", inquiryqty as ""Inquiry Qty"", currentinquiryeta as ""Current Inquiry ETA"", currentinquiryetd as ""Current Inquiry ETD"", currentinquiryqty as ""Current Inquiry Qty"", confirmationstatus as ""Confirmation status"", stconfirmedetd as ""1st Confirmed ETD"", stconfirmedqty as ""1st Confirmed Qty"",currentconfirmedeta as ""Current Confirmed ETA"", currentconfirmedetd as ""Current Confirmed ETD"",currentconfirmedqty as ""Current Confirmed Qty"", deliveredqty as ""Delivered Qty"", shipdate as ""Ship.Date"",shipdateeta as ""Ship.Date ETA"", osqty as ""O/S Qty"", sebasiapono as ""Seb Asia P.O. No."", polineno as ""Line No"", ctrno as ""CTR No."", boatid as ""Boat ID"", packinglist as ""Packing List"", shipfrom as ""Ship From "",finalcustomerorder as ""Final Customer Order"" , comments as ""Comments"" " & _
                 ", sbu::character varying as ""SBU"",bu::character varying as ""BU"",familysbu as ""Family SBU"", purchasinggroup as ""PGr"",currency as ""Crcy"",sp as ""SP"",spm as ""SPM"",dicustomerorder as ""DI customer order""" &
                 " from cxallworfg " & openorder & " and (purchasinggroup in('FOA','FO9') and soldtoparty = 99009500) " & " union all " &
                 " Select ordertype as ""Header / Item"", shiptoparty as ""Customer code"", shiptopartyname::character varying as ""Customer name"",soldtoparty as ""Sold To Party SAP Code"",soldtopartyname::character varying as ""Sold To Party Name"", sao as ""Sales Administration Office"", customerorderno as ""Customer Order No."", sebasiasalesorder as ""Seb Asia Sales Order"", solineno as ""Item No"", vendorcode as ""Vendor Code"",vendorname::character varying as ""Vendor"", ssm::character varying as ""SSM"", rir::character varying as ""Research and Industry Responsible"", cmmf as ""CMMF"", comfam as ""Commercial family CG2 n"", brand as ""Brand CG2"", brandname::character varying as ""Brand"",itemid::character varying as ""Item ID"", materialdesc::character varying as ""Description"", orderstatus as ""Order Status"", latestupdate as ""Latest Update"",updatesince as ""Updated since 7""" & _
                 ", curinq as ""Cur./Inq."", receptiondate as ""Reception Date"", fob::numeric as ""FOB"", unittp::numeric as ""Unit T.P."", inquiryeta as ""Inquiry ETA"", inquiryetd as ""Inquiry ETD"", inquiryqty as ""Inquiry Qty"", currentinquiryeta as ""Current Inquiry ETA"", currentinquiryetd as ""Current Inquiry ETD"", currentinquiryqty as ""Current Inquiry Qty"", confirmationstatus as ""Confirmation status"", stconfirmedetd as ""1st Confirmed ETD"", stconfirmedqty as ""1st Confirmed Qty"",currentconfirmedeta as ""Current Confirmed ETA"", currentconfirmedetd as ""Current Confirmed ETD"",currentconfirmedqty as ""Current Confirmed Qty"", deliveredqty as ""Delivered Qty"", shipdate as ""Ship.Date"",shipdateeta as ""Ship.Date ETA"", osqty as ""O/S Qty"", sebasiapono as ""Seb Asia P.O. No."", polineno as ""Line No"", ctrno as ""CTR No."", boatid as ""Boat ID"", packinglist as ""Packing List"", shipfrom as ""Ship From "",finalcustomerorder as ""Final Customer Order"" , comments as ""Comments"" " & _
                 ", sbu::character varying as ""SBU"",bu::character varying as ""BU"",familysbu as ""Family SBU"", purchasinggroup as ""PGr"",currency as ""Crcy"",sp as ""SP"",spm as ""SPM"",dicustomerorder as ""DI customer order""" &
                 " from cxallworcomp " & openorder & " and soldtoparty = 99009500 "
            Case Else
                sqlstr = " Select ordertype as ""Header / Item"", shiptoparty as ""Customer code"", shiptopartyname::character varying as ""Customer name"",soldtoparty as ""Sold To Party SAP Code"",soldtopartyname::character varying as ""Sold To Party Name"", sao as ""Sales Administration Office"", customerorderno as ""Customer Order No."", sebasiasalesorder as ""Seb Asia Sales Order"", solineno as ""Item No"", vendorcode as ""Vendor Code"",vendorname::character varying as ""Vendor"", ssm::character varying as ""SSM"", rir::character varying as ""Research and Industry Responsible"", cmmf as ""CMMF"", comfam as ""Commercial family CG2 n"", brand as ""Brand CG2"", brandname::character varying as ""Brand"",itemid::character varying as ""Item ID"", materialdesc::character varying as ""Description"", orderstatus as ""Order Status"", latestupdate as ""Latest Update"",updatesince as ""Updated since 7""" & _
                 ", curinq as ""Cur./Inq."", receptiondate as ""Reception Date"", fob::numeric as ""FOB"", unittp::numeric as ""Unit T.P."", inquiryeta as ""Inquiry ETA"", inquiryetd as ""Inquiry ETD"", inquiryqty as ""Inquiry Qty"", currentinquiryeta as ""Current Inquiry ETA"", currentinquiryetd as ""Current Inquiry ETD"", currentinquiryqty as ""Current Inquiry Qty"", confirmationstatus as ""Confirmation status"", stconfirmedetd as ""1st Confirmed ETD"", stconfirmedqty as ""1st Confirmed Qty"",currentconfirmedeta as ""Current Confirmed ETA"", currentconfirmedetd as ""Current Confirmed ETD"",currentconfirmedqty as ""Current Confirmed Qty"", deliveredqty as ""Delivered Qty"", shipdate as ""Ship.Date"",shipdateeta as ""Ship.Date ETA"", osqty as ""O/S Qty"", sebasiapono as ""Seb Asia P.O. No."", polineno as ""Line No"", ctrno as ""CTR No."", boatid as ""Boat ID"", packinglist as ""Packing List"", shipfrom as ""Ship From "",finalcustomerorder as ""Final Customer Order"" , comments as ""Comments"" " & _
                 ", sbu::character varying as ""SBU"",bu::character varying as ""BU"",familysbu as ""Family SBU"", purchasinggroup as ""PGr"",currency as ""Crcy"",sp as ""SP"",spm as ""SPM"",dicustomerorder as ""DI customer order""" &
                 " from " & alldata & openorder
        End Select

        'sqlstr = " Select ordertype as ""Header / Item"", shiptoparty as ""Customer code"", shiptopartyname::character varying as ""Customer name"",soldtoparty as ""Sold To Party SAP Code"",soldtopartyname::character varying as ""Sold To Party Name"", sao as ""Sales Administration Office"", customerorderno as ""Customer Order No."", sebasiasalesorder as ""Seb Asia Sales Order"", solineno as ""Item No"", vendorcode as ""Vendor Code"",vendorname::character varying as ""Vendor"", ssm::character varying as ""SSM"", rir::character varying as ""Research and Industry Responsible"", cmmf as ""CMMF"", comfam as ""Commercial family CG2 n"", brand as ""Brand CG2"", brandname::character varying as ""Brand"",itemid::character varying as ""Item ID"", materialdesc::character varying as ""Description"", orderstatus as ""Order Status"", latestupdate as ""Latest Update"",updatesince as ""Updated since 7""" & _
        '         ", curinq as ""Cur./Inq."", receptiondate as ""Reception Date"", fob::numeric as ""FOB"", unittp::numeric as ""Unit T.P."", inquiryeta as ""Inquiry ETA"", inquiryetd as ""Inquiry ETD"", inquiryqty as ""Inquiry Qty"", currentinquiryeta as ""Current Inquiry ETA"", currentinquiryetd as ""Current Inquiry ETD"", currentinquiryqty as ""Current Inquiry Qty"", confirmationstatus as ""Confirmation status"", stconfirmedetd as ""1st Confirmed ETD"", stconfirmedqty as ""1st Confirmed Qty"",currentconfirmedeta as ""Current Confirmed ETA"", currentconfirmedetd as ""Current Confirmed ETD"",currentconfirmedqty as ""Current Confirmed Qty"", deliveredqty as ""Delivered Qty"", shipdate as ""Ship.Date"",shipdateeta as ""Ship.Date ETA"", osqty as ""O/S Qty"", sebasiapono as ""Seb Asia P.O. No."", polineno as ""Line No"", ctrno as ""CTR No."", boatid as ""Boat ID"", packinglist as ""Packing List"", shipfrom as ""Ship From "",finalcustomerorder as ""Final Customer Order"" , comments as ""Comments"" " & _
        '         ", sbu::character varying as ""SBU"",bu::character varying as ""BU"",familysbu as ""Family SBU"", purchasinggroup as ""PGr"",currency as ""Crcy"",sp as ""SP"",spm as ""SPM"",dicustomerorder as ""DI customer order""" &
        '         " from " & alldata & openorder

        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = "WOR-" & Dept
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
            Dim SpecificationName As String = "FGImport"

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, AccessDbFullPath, AccessTableName, SpecificationName)
            'myreport = New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, AccessDbFullPath, AccessTableName, SpecificationName)
            myreport.Run(Me, e)
        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged, RadioButton1.CheckedChanged
        DateTimePicker1.Enabled = RadioButton2.Checked
        DateTimePicker2.Enabled = RadioButton2.Checked
    End Sub

    Private Sub FormWORAllLines_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        showEkkoLastCreationDate()
    End Sub

    Private Sub showEkkoLastCreationDate()
        If Not mythread.IsAlive Then
            mythread = New Thread(AddressOf doQuery)
            mythread.Start()
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                
                Case 8
                    Label3.Text = message
            End Select

        End If

    End Sub
    Private Sub doQuery()
        Dim myresult As Date
        If DbAdapter1.ExecuteScalar("select getekkolastcreatedon();", myresult) Then
            ProgressReport(8, String.Format("Latest Ekko Creation Date : {0:dd-MMM-yyyy} ", myresult))
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        TextBox1.Visible = CheckBox1.Checked
        Button2.Visible = CheckBox1.Checked
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName
        End If
    End Sub
End Class