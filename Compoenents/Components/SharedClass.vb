Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports DJLib


Public Class SharedClass
    <DllImport("user32.dll")> _
    Public Shared Function EndTask(ByVal hWnd As IntPtr, ByVal fShutDown As Boolean, ByVal fForce As Boolean) As Boolean
    End Function
    Public Shared Function ValidateFileName(ByVal foldername As String, ByVal source As String) As String
        ValidateFileName = source
        Dim FileName = Path.GetFileName(source)
        Dim i As Integer = 0
        While File.Exists(foldername & "\" & FileName)
            i += 1
            FileName = "Copy(" & i & ") of " & Path.GetFileName(source)
            ValidateFileName = foldername & "\" & FileName
        End While
        Return ValidateFileName
    End Function

    Public Shared Sub releaseComObject(ByRef o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
        Finally
            o = Nothing
        End Try
    End Sub
    Public Shared Sub FillDataSource(ByRef owb As Excel.Workbook, ByVal SheetNum As Integer, ByVal sqlstr As String, ByVal dbAdapter As DbAdapter, Optional ByVal Location As String = "A1")

        owb.Worksheets(SheetNum).select()
        Dim osheet As Excel.Worksheet = owb.Worksheets(SheetNum)
        Dim oRange As Excel.Range
        Dim oExCon As String = My.Settings.oExCon.ToString '"ODBC;DSN=PostgreSQL30;"
        oExCon = oExCon.Insert(oExCon.Length, "UID=" & dbAdapter.ConnectionStringDict("USER ID") & ";PWD=" & dbAdapter.ConnectionStringDict("PASSWORD"))
        With osheet.QueryTables.Add(oExCon.Replace("Host=", "Server="), osheet.Range(Location))
            'With osheet.QueryTables.Add(oExCon, osheet.Range("A1"))
            .CommandText = sqlstr
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
            .SavePassword = True
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .Refresh(BackgroundQuery:=False)
            Application.DoEvents()
        End With
        If owb.Connections.Count > 0 Then
            owb.Connections(owb.Connections.Count).Delete()
        End If

        'oRange = osheet.Range("1:1")
        oRange = osheet.Range(Location)
        oRange.Select()
        osheet.Application.Selection.autofilter()
        osheet.Cells.EntireColumn.AutoFit()

    End Sub
    Public Shared Sub FillWorksheet(ByVal osheet As Excel.Worksheet, ByVal sqlstr As String, ByVal dbAdapter As DbAdapter, Optional ByVal Location As String = "A1")
        'Dim oRange As Excel.Range
        Dim oExCon As String = My.Settings.oExCon.ToString '"ODBC;DSN=PostgreSQL30;"
        oExCon = oExCon.Insert(oExCon.Length, "UID=" & dbAdapter.ConnectionStringDict("USER ID") & ";PWD=" & dbAdapter.ConnectionStringDict("PASSWORD"))
        Dim oRange As Excel.Range
        oRange = osheet.Range(Location)
        With osheet.QueryTables.Add(oExCon.Replace("Host=", "Server="), oRange)
            'With osheet.QueryTables.Add(oExCon, osheet.Range("A1"))
            .CommandText = sqlstr
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
            .SavePassword = True
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .Refresh(BackgroundQuery:=False)
            Application.DoEvents()
        End With

        oRange = Nothing
    End Sub

    Public Shared Sub ExportToExcel(ByRef FileName As String, ByVal sqlstr As String, ByVal dbadapter As DbAdapter)
        Dim Result As Boolean = False
        Dim handle As IntPtr
        Dim StringBuilder1 As New System.Text.StringBuilder

        Dim source As String = ""
        'ask export location
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Application.DoEvents()
            Cursor.Current = Cursors.WaitCursor
            source = DirectoryBrowser.SelectedPath & "\" & FileName

            'Excel Variable
            Dim oXl As Excel.Application = Nothing
            Dim oWb As Excel.Workbook = Nothing
            Dim oSheet As Excel.Worksheet = Nothing
            Dim SheetName As String = vbEmpty
            Dim oRange As Excel.Range = Nothing

            Try
                'Create Object Excel 
                oXl = CType(CreateObject("Excel.Application"), Excel.Application)
                Application.DoEvents()
                'oXl.Visible = True
                'get process pid
                handle = oXl.Hwnd
                oXl.Visible = False
                oXl.DisplayAlerts = False
                oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ExcelTemplate.xltx")
                'Loop for chart
                oSheet = oWb.Worksheets(1)

                FillDataSource(oWb, 1, sqlstr, dbadapter)

                oSheet.Cells.EntireColumn.AutoFit()
                FileName = ValidateFileName(DirectoryBrowser.SelectedPath, source)
                oWb.SaveAs(FileName)
                Result = True

            Catch ex As Exception
                'MsgBox(ex.Message)

            Finally
                'clear excel from memory
                oXl.Quit()
                releaseComObject(oRange)
                releaseComObject(oSheet)
                releaseComObject(oWb)
                releaseComObject(oXl)
                GC.Collect()
                GC.WaitForPendingFinalizers()
                Try
                    EndTask(handle, True, True)
                Catch ex As Exception
                End Try
                If Result Then
                    If MsgBox("File name: " & FileName & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                        Process.Start(FileName)
                    End If
                End If
            End Try
        End If
    End Sub
    Public Shared Sub CheckedListBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim CheckedListBox1 = CType(sender, CheckedListBox)
        Select Case sender.selectedindex
            Case 0
                Dim chkstate As CheckState
                chkstate = CheckedListBox1.GetItemCheckState(0)
                For i = 0 To sender.items.count - 1
                    CheckedListBox1.SetItemChecked(i, chkstate)

                Next
            Case Else
                CheckedListBox1.SetItemChecked(0, 0)
                Dim mycountlist As Integer = Countlist(CheckedListBox1)
                If CheckedListBox1.Items.Count = mycountlist + 1 Then
                    CheckedListBox1.SetItemChecked(0, True)
                End If

                Dim mycheck As CheckState
                If mycountlist = 0 Then
                    For i = 1 To sender.items.count - 1

                    Next
                Else
                    For i = 1 To sender.items.count - 1

                        mycheck = CheckedListBox1.GetItemCheckState(i)

                    Next
                End If


        End Select
        'End If
    End Sub

    Public Shared Function Countlist(ByVal myCheckedListbox As CheckedListBox) As Integer
        Dim count As Integer = 0
        For i = 0 To myCheckedListbox.Items.Count - 1
            If myCheckedListbox.GetItemCheckState(i) Then
                count += 1
            End If
        Next
        Return count
    End Function
    Public Shared Function DateFormatyyyyMMdd(ByRef DateInput As Object) As String
        Dim myRet As String
        myRet = "NULL"
        If IsDBNull(DateInput) Then
            Return "Null"
        ElseIf Not DateInput Is Nothing Then
            Dim arrDate(2) As String
            arrDate(0) = DateInput.Year
            arrDate(1) = DateInput.Month
            arrDate(2) = DateInput.Day
            myRet = "'" & String.Join("-", arrDate) & "'"
        End If
        Return myRet
    End Function

    Public Shared Function DateFormatyyyyMMddString(ByRef DateInput As Object) As String
        Dim myRet As String
        myRet = "Null"
        If DateInput = "" Then
            Return myRet
        End If
        If Not DateInput Is Nothing Then
            Dim arrDate(2) As String

            arrDate(0) = CDate(DateInput).Year
            arrDate(1) = CDate(DateInput).Month
            arrDate(2) = CDate(DateInput).Day
            myRet = "'" & String.Join("-", arrDate) & "'"
        End If
        Return myRet
    End Function
    Public Shared Function getPrevMonth(ByVal mydate As Date) As Date
        Dim newdate As Date
        Dim mymonth As Integer
        Dim myyear As Integer


        If mydate.Month = 1 Then
            mymonth = 12
            myyear = mydate.Year - 1
        Else
            mymonth = mydate.Month - 1
            myyear = mydate.Year
        End If
        newdate = CDate(myyear & "-" & mymonth & "-1")
        Return newdate
    End Function
    Public Shared Function validint(ByVal p1 As Object) As Object
        If p1 = "" Then
            Return "Null"
        Else
            Return CInt((Replace(p1, ",", "")))
        End If
    End Function
    Public Shared Function validlong(ByVal p1 As Object) As Object
        If p1 = "" Then
            Return "Null"
        Else
            Return CLng((Replace(p1, ",", "")))
        End If
    End Function

    Public Shared Function validstr(ByVal data As Object) As Object
        If data = "" Then
            Return "Null"
        End If
        'Return Replace(Replace(data, Chr(9), " "), "'", "''")
        Return Replace(Replace(data, Chr(9), " "), "'", "''")
    End Function
    Public Shared Function escapestr(ByVal data As Object) As String
        If data = "" Then
            Return "Null"
        End If
        Return "'" & data & "'"
    End Function

    Public Shared Function dateformatdotyyyymmdd(ByVal data As Object) As Object
        Dim myret As String = "Null"
        If data = "" Then
            Return myret
        End If
        If data = "00.00.0000" Then
            Return myret
        End If
        Dim mydate = data.ToString.Split(".")
        myret = "'" & mydate(2) & "-" & mydate(1) & "-" & mydate(0) & "'"
        Return myret
    End Function

    Public Shared Function dateformatdotyyyymmddstring(ByVal data As Object) As Object
        Dim myret As String = "Null"
        If data = "" Then
            Return myret
        End If
        If data = "00.00.0000" Then
            Return myret
        End If
        Dim mydate = data.ToString.Split(".")
        myret = mydate(2) & "-" & mydate(1) & "-" & mydate(0)
        Return myret
    End Function

    Public Shared Function validreal(ByVal data As Object) As Object
        Dim myret As String = "Null"
        If data = "" Then
            Return myret
        End If
        Return CDec(Replace(data, ",", ""))
    End Function

    Public Shared Function validzerotonull(ByVal podtlid As Long) As String
        Dim myret = podtlid
        If podtlid = 0 Then
            Return "Null"
        End If
        Return myret
    End Function
    Public Shared Function getfirstdate(ByVal p1 As Date) As Date
        Return CDate(p1.Year & "-" & p1.Month & "-1")
    End Function

    Public Shared Function getlastdate(ByVal p1 As Date) As Date
        If p1.Month = 12 Then
            Return CDate(p1.Year & "-12-31")
        Else
            Return CDate(p1.Year & "-" & p1.Month + 1 & "-1").AddDays(-1)
        End If
    End Function
End Class

Class ThreadPoolObj
    Public ObjectID As Integer
    Public signal As System.Threading.ManualResetEvent
    Public osheet As Excel.Worksheet
    Public ds As DataSet
    Public sb As System.Text.StringBuilder
    Public strsql As String
    Public Name As String
End Class

Class ThreadPoolManualResetEvent
    Public ObjectID As Integer
    Public signal As System.Threading.ManualResetEvent
End Class
Public Class ContentBaseEventArgs
    Inherits EventArgs
    Public Property dataset As DataSet
    Public Property message As String
    Public Property hasChanges As Boolean
    Public Property ra As Integer
    Public Property continueonerror As Boolean

    Public Sub New(ByVal dataset As DataSet, ByRef haschanges As Boolean, ByRef message As String, ByRef recordaffected As Integer, ByVal continueonerror As Boolean)
        Me.dataset = dataset
        Me.message = message
        Me.ra = ra
        Me.continueonerror = continueonerror
    End Sub
End Class