Imports Microsoft.Office.Interop
Public Class ImportOPLTComments
    Dim FileName As String
    Dim myform As Object
    Public Property ErrorMsg As String
    Dim OPLTCommentTxController As New OPLTCOMMENTTXController
    Dim OPLTCommentController As New OPLTCOMMENTController
    Public Sub New(ByVal myform As Object, ByVal Filename As String)
        Me.FileName = Filename
        Me.myform = myform
    End Sub
    Public Function ValidateFile() As Boolean
        Dim myret As Boolean = False
        Try
            If openExcelFile() Then
                myret = True
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return myret
    End Function

    Private Function openExcelFile() As Boolean
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Dim myret As Boolean = False

        Try
            'Create Object Excel 
            myform.ProgressReport(1, "Create object...")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)

            hwnd = oXl.Hwnd
            oXl.Visible = False
            oXl.DisplayAlerts = False
            myform.ProgressReport(1, String.Format("Open file {0}...", FileName))

            oWb = oXl.Workbooks.Open(FileName)
            myform.ProgressReport(1, String.Format("Open file Done..."))
            'Check FileType
            MessageBox.Show("Before owb.worksheets.count")
            Application.DoEvents()
            myform.ProgressReport(1, String.Format("Worksheet Count {0}...", oWb.Worksheets.Count))
            Application.DoEvents()
            MessageBox.Show("After owb.worksheets.count")
            For i = 1 To oWb.Worksheets.Count
                Application.DoEvents()
                oWb.Worksheets(i).select()
                MessageBox.Show("inside for i")
                oSheet = oWb.Worksheets(i)
                myform.ProgressReport(1, String.Format("Worksheet name {0}...", oSheet.Name))
                If oSheet.Name = "DATA" Then
                    myform.ProgressReport(1, "Save TXT File...")
                    oWb.SaveAs(Filename:=FileName.Replace("xlsx", "TXT"), FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
                    myret = True
                    Exit For
                End If
            Next
            If Not myret Then
                Throw New Exception("File is not valid.")
            End If
            myform.ProgressReport(1, "Save TXT File Done...")
        Catch ex As Exception
            ErrorMsg = ex.Message
        Finally
            oXl.Quit()
            ExcelStuff.releaseComObject(oSheet)
            ExcelStuff.releaseComObject(oWb)
            ExcelStuff.releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                ExcelStuff.EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
        End Try
        Return myret
    End Function

    Public Function DoImportFile() As Boolean
        Dim myret As Boolean
        Dim mylist As New List(Of String())
        Try
            Dim myrecord() As String
            Using objTFParser = New FileIO.TextFieldParser(FileName.Replace("xlsx", "TXT"))
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    myform.ProgressReport(1, "Read Data")
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count > 0 Then
                            mylist.Add(myrecord)
                        End If
                        count += 1
                    Loop
                End With
            End Using
            'Get OPLTCOMMENTTX Data
            If Not OPLTCommentTxController.loaddata() Then
                ErrorMsg = "Failed to load OPLTCommentTX Data."
                Return myret
            End If

            If Not OPLTCommentController.loaddata() Then
                ErrorMsg = "Failed to load OPLTComment Data."
                Return myret
            End If
            '
            For i = 0 To mylist.Count - 1
                'Check Data
                If mylist(i)(35) <> "NOT FAILED" AndAlso mylist(i)(35) <> "" Then
                    'Find opltcommenttx, if not avail create one, if found update
                    Dim result As DataRow
                    Dim mykey(1) As Object

                    Dim Commentid As Integer
                    Commentid = OPLTCommentController.getId(mylist(i)(35))
                    If IsDBNull(Commentid) Then
                        Err.Raise(500, Description:=String.Format("Comment {0} is not registered.", mylist(i)(35)))
                    End If
                    mykey(0) = mylist(i)(0)
                    mykey(1) = mylist(i)(1)
                    result = OPLTCommentTxController.DS.Tables("opltcommenttx").Rows.Find(mykey)

                    If IsNothing(result) Then
                        Dim dr As DataRow = OPLTCommentTxController.DS.Tables("opltcommenttx").NewRow
                        dr.Item("salesdoc") = mykey(0)
                        dr.Item("itemno") = mykey(1)
                        dr.Item("commentid") = Commentid
                        dr.Item("lt") = mylist(i)(21)
                        If mylist(i)(36) = "" Then
                            dr.Item("remarks") = DBNull.Value
                        Else
                            dr.Item("remarks") = mylist(i)(36)
                        End If

                        OPLTCommentTxController.DS.Tables("opltcommenttx").Rows.Add(dr)
                    Else
                        result.Item("commentid") = Commentid
                        result.Item("lt") = mylist(i)(21)
                        If mylist(i)(36) = "" Then
                            result.Item("remarks") = DBNull.Value
                        Else
                            result.Item("remarks") = mylist(i)(36)
                        End If

                    End If
                End If
            Next
            myform.ProgressReport(1, "Saving Data...")
            OPLTCommentTxController.save()
            myform.ProgressReport(1, "Saving Data Done...")
            myret = True
        Catch ex As Exception
            ErrorMsg = ex.Message
        End Try

        Return myret
    End Function
End Class
