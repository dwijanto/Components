Imports System.Text
Imports System.IO

Public Class VendorFamilySPSPMController
    Implements IController
    Implements IToolbarAction
    Public Model As New VendorFamilySPSPMModel
    Public BS As BindingSource
    Public DS As DataSet
    Dim readfilestatus As Boolean
    Dim errMsg As StringBuilder
    Public Property ImportFileName As String
    Dim Parent As Object
    Private VendorFamilySPSPMSB As StringBuilder

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
        Model = New VendorFamilySPSPMModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).Columns("id").AutoIncrement = True
            DS.Tables(0).Columns("id").AutoIncrementSeed = 0
            DS.Tables(0).Columns("id").AutoIncrementStep = -1
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
        Dim myret As Boolean

        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        Parent.ProgressReport(1, "Read Folder..")

        readfilestatus = ImportTextFile(ImportFileName)
        If readfilestatus Then
            sw.Stop()
            Parent.ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            Parent.ProgressReport(5, "Set to continuous mode again")
        Else
            If Not errmsg.ToString.Contains(vbCrLf) Then
                Parent.ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done with error.{3}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString, errMsg.ToString))
            Else
                Using mystream As New StreamWriter(Application.StartupPath & "\error.txt")
                    mystream.WriteLine(errmsg.ToString)
                End Using
                Process.Start(Application.StartupPath & "\error.txt")
                Parent.ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done with Error.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))

            End If


        End If
        sw.Stop()




        Return myret
    End Function

    Private Function ImportTextFile(ByVal p1 As String) As Boolean
        Dim sb As New StringBuilder
        Dim myret As Boolean = False
        Dim list As New List(Of String())
        errMsg = New StringBuilder
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
                errmsg.Append("Text File Wrong Format")
                Parent.ProgressReport(5, "Set To Marque")
                Return myret
            End If

            Parent.ProgressReport(1, String.Format("Build Data row..........."))
            Parent.ProgressReport(5, "Set To Continuous")
            VendorFamilySPSPMSB = New StringBuilder
            For i = 0 To list.Count - 1
                Parent.ProgressReport(7, i + 1 & "," & list.Count)
                buildSB(list(i))              
            Next
            myret = True
            If errmsg.Length > 0 Then
                myret = False
            End If
            Parent.ProgressReport(6, "Set To Marque")
            Parent.ProgressReport(1, String.Format("Copy To Db"))
            If Not copyToDb() Then
                Return False
            End If

        Catch ex As Exception
            errMsg.Append(String.Format("Row : {0} ", i) & ex.Message)
        Finally
            Parent.ProgressReport(5, "Set To Continuous")
        End Try
        Return myret
    End Function

    Private Function copyToDb() As Boolean
        Dim myret As Boolean = False
        Dim mystr As New StringBuilder
        Dim errmessage As String

        Parent.ProgressReport(1, "Start Add New Records")
        mystr.Append("delete from familyspspm;")
        mystr.Append("select setval('familyspspm_id_seq',1,false);")

        Dim sqlstr As String = String.Format("begin;set statement_timeout to 0;end;{0};copy familyspspm(familyid,vendorcode,sp,spm) from stdin with null as 'Null';", mystr.ToString)
        Dim ra As Long = 0
        Try

            errmessage = Model.copyToDb(sqlstr, VendorFamilySPSPMSB, myret)
            If myret Then
                Parent.ProgressReport(1, "Add Records Done.")
            Else
                errMsg.Append("Copy Error " & Model.ErrorMessage & vbCrLf)
            End If
        Catch ex As Exception
            errMsg.Append(ex.Message & vbCrLf)        
        End Try
        Return myret
    End Function


    Private Sub buildSB(ByVal myrecord As String())
        Dim mymodel = New VendorFamilySPSPMModel With {.familyid = myrecord(0),
                                                       .vendorcode = myrecord(1),
                                                       .sp = myrecord(3),
                                                       .spm = myrecord(4)}
        VendorFamilySPSPMSB.Append(mymodel.familyid & vbTab & mymodel.vendorcode & vbTab & mymodel.sp & vbTab & mymodel.spm & vbCrLf)
    End Sub
End Class
