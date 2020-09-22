Imports System.Text
Imports System.IO
Public Class CMMFSPSPMController
    Implements IController
    Implements IToolbarAction
    Public Model As New CMMFSPSPMModel
    Public BS As BindingSource
    Public DS As DataSet
    Dim readfilestatus As Boolean
    Dim errMsgSB As StringBuilder
    Public Property ImportFileName As String
    Dim Parent As Object
    Private CMMFSPSPMSB As StringBuilder
    Private UpdCMMFSB As StringBuilder
    Public BlankSPSPM As Boolean
    Public Sub New(ByVal Parent As Object)
        Me.Parent = Parent
    End Sub

    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.TableName).Copy()
        End Get
    End Property

    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim BS As New BindingSource
            BS.DataSource = GetTable
            BS.Sort = Model.SortField
            Return BS
        End Get
    End Property

    'Public Function GetSBUBS(ByVal criteria As String) As BindingSource
    '    Return Model.GetBindingSource(criteria)
    'End Function


    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New CMMFSPSPMModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("cmmf")          
            DS.Tables(0).PrimaryKey = pk
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret

    End Function
    Public Function Validate() As Boolean
        Dim myret As Boolean

        Return myret
    End Function

    Public Function save() As Boolean Implements IController.save
        Dim myret As Boolean = False

        BS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If

        Return myret
    End Function

    Public Function Save(ByVal mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Dim myret As Boolean = False
        If Model.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function

    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format(Model.FilterField, value)
        End Set
    End Property

    Public Function GetCurrentRecord() As DataRowView Implements IToolbarAction.GetCurrentRecord
        Return BS.Current
    End Function

    Public Function GetNewRecord() As DataRowView Implements IToolbarAction.GetNewRecord
        Return BS.AddNew
    End Function

    Public Sub RemoveAt(ByVal value As Integer) Implements IToolbarAction.RemoveAt
        BS.RemoveAt(value)
    End Sub

    Public Function ImportData() As Boolean
        'Dim myret As Boolean

        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        Parent.ProgressReport(1, "Read Folder..")
        BlankSPSPM = False
        readfilestatus = ImportTextFile(ImportFileName)
        If readfilestatus Then
            sw.Stop()
            Parent.ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            Parent.ProgressReport(5, "Set to continuous mode again")
        Else
            If Not errMsgSB.ToString.Contains(vbCrLf) Then
                Parent.ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done with error.{3}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString, errMsgSB.ToString))
            Else
                Using mystream As New StreamWriter(Application.StartupPath & "\error.txt")
                    mystream.WriteLine(errMsgSB.ToString)
                End Using
                Process.Start(Application.StartupPath & "\error.txt")
                Parent.ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done with Error.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))

            End If


        End If
        sw.Stop()




        Return readfilestatus
    End Function

    Private Function ImportTextFile(ByVal p1 As String) As Boolean
        Dim sb As New StringBuilder
        Dim myret As Boolean = False
        Dim list As New List(Of String())
        errMsgSB = New StringBuilder
        Dim i As Long
        Try
            Dim myrecord() As String
            Dim tcount As Long = 0
            Parent.ProgressReport(6, "Set To Marque")
            Parent.ProgressReport(1, String.Format("Read Text File...{0}", ImportFileName))
            Using objTFParser = New FileIO.TextFieldParser(ImportFileName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(vbTab)
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count > 0 Then
                            list.Add(myrecord)
                        End If
                        count += 1
                    Loop
                End With
            End Using
            If list.Count = 0 Then
                errMsgSB.Append("Text File Wrong Format")
                'Parent.ProgressReport(6, "Set To Marque")
                Return myret
            End If

            Parent.ProgressReport(1, String.Format("Build Data row..........."))
            Parent.ProgressReport(5, "Set To Continuous")
            CMMFSPSPMSB = New StringBuilder
            UpdCMMFSB = New StringBuilder
            For i = 0 To list.Count - 1
                Parent.ProgressReport(7, i + 1 & "," & list.Count)
                'Find existing if avail -> update else create
                Dim mycol(0) As Object
                mycol(0) = list(i)(0)
                Dim result As DataRow = DS.Tables(0).Rows.Find(mycol)
                If IsNothing(result) Then
                    buildSB(list(i))
                Else
                    'Update
                    Dim flagUpdate As Boolean = False

                    If list(i)(1) <> result.Item(1) Then
                        flagUpdate = True
                    End If
                    If list(i)(2) <> result.Item(2) Then
                        flagUpdate = True
                    End If

                    If flagUpdate Then
                        If UpdCMMFSB.Length > 0 Then
                            UpdCMMFSB.Append(",")
                        End If
                        UpdCMMFSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying]", list(i)(0), list(i)(1), list(i)(2)))
                    End If

                End If

            Next
            myret = True
            If errMsgSB.Length > 0 Then
                myret = False
            End If

            Parent.ProgressReport(1, String.Format("Copy To Db"))
            Parent.ProgressReport(6, "Set To Marque")
            If Not copyToDb() Then
                Return False
            End If
            Parent.ProgressReport(5, "Set To Continuous")
        Catch ex As Exception
            errMsgSB.Append(String.Format("Row : {0} ", i) & ex.Message)
        Finally
            Parent.ProgressReport(5, "Set To Continuous")
        End Try
        Return myret
    End Function

    Private Function copyToDb() As Boolean
        Dim myret As Boolean = False
        Dim mystr As New StringBuilder
        Dim errmessage As String
        If UpdCMMFSB.Length > 0 Then
            Parent.ProgressReport(2, "Update CMMF")
            'cmmf.activitycode, cmmf.brandid, cmmf.cmmftype, cmmf.comfam, cmmf.commercialref, cmmf.materialdesc, cmmf.modelcode, cmmf.plnt, cmmf.rir, cmmf.sbu, cmmf.sorg
            Dim sqlstr = "update cmmfspspm set sp= foo.sp,spm = foo.spm" &
                        " from (select * from array_to_set3(Array[" & UpdCMMFSB.ToString &
                     "]) as tb (cmmf character varying,sp character varying, spm character varying))foo where cmmfspspm.cmmf = foo.cmmf::bigint;"
            Dim ra As Long
            Dim errmsg As String = String.Empty
            If Not Model.myadapter.ExecuteNonQuery(sqlstr, ra, errmsg) Then
                errMsgSB.Append(errmsg & vbCrLf)
                Parent.ProgressReport(2, "Update CMMF" & "::" & errmsg)
                Return False
            End If
        End If

        If CMMFSPSPMSB.Length > 0 Then
            Parent.ProgressReport(1, "Start Add New Records")
            'mystr.Append("delete from cmmfspspm;")        
            Dim sqlstr As String = String.Format("begin;set statement_timeout to 0;end;{0};copy cmmfspspm(cmmf,sp,spm) from stdin with null as 'Null';", mystr.ToString)
            Dim ra As Long = 0
            Try

                errmessage = Model.copyToDb(sqlstr, CMMFSPSPMSB, myret)
                If myret Then
                    Parent.ProgressReport(1, "Add Records Done.")
                Else
                    errMsgSB.Append("Copy Error " & Model.ErrorMessage & vbCrLf)
                End If
            Catch ex As Exception
                errMsgSB.Append(ex.Message & vbCrLf)
            End Try
        Else
            myret = True
        End If

        Return myret
    End Function


    Private Sub buildSB(ByVal myrecord As String())
        Dim mymodel = New CMMFSPSPMModel With {.cmmf = myrecord(0),
                                                       .sp = myrecord(1),
                                                       .spm = myrecord(2)}
        If Not (mymodel.sp = "" Or mymodel.spm = "") Then
            CMMFSPSPMSB.Append(mymodel.cmmf & vbTab & mymodel.sp & vbTab & mymodel.spm & vbCrLf)
        Else
            BlankSPSPM = True
        End If

    End Sub
End Class
