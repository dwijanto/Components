Imports System.Text
Public Class CMMFAdapter
    Public AddCMMFSB As StringBuilder
    Public UpdCMMFSB As StringBuilder
    Public AddRangeSB As StringBuilder

    Public DS As DataSet
    Public dbadapter1 As DbAdapter = DbAdapter.getInstance
    Public errmsg As String = String.Empty


    Public Sub New()
        DS = New DataSet
        AddCMMFSB = New StringBuilder
        UpdCMMFSB = New StringBuilder
        AddRangeSB = New StringBuilder
    End Sub

    Public Function LoadCMMF() As Boolean
        Dim myret As Boolean = False
        'Dim sqlstr = "Select cmmf.activitycode, cmmf.brandid::character varying, cmmf.cmmf, cmmf.cmmftype, cmmf.comfam::character varying, cmmf.commercialref, cmmf.createon, cmmf.materialdesc, cmmf.modelcode, cmmf.plnt, cmmf.rangeid, cmmf.rir, cmmf.sbu, cmmf.sorg from cmmf;select rangeid,range,rangedesc from range;"
        Dim sqlstr = "Select * from cmmf;select rangeid,range,rangedesc from range;"
        If dbadapter1.TbgetDataSet(sqlstr, DS, errmsg) Then
            'Table Name
            DS.Tables(0).TableName = "CMMF"
            DS.Tables(1).TableName = "Range"

            'Primary Key
            Dim pk0(0) As DataColumn
            pk0(0) = DS.Tables(0).Columns("cmmf")
            DS.Tables(0).PrimaryKey = pk0

            Dim pk1(0) As DataColumn
            pk1(0) = DS.Tables(1).Columns("range")
            DS.Tables(1).PrimaryKey = pk1
            myret = True


        End If
        Return myret
    End Function

    Public Sub ValidateCMMF(ByVal cmmf As CMMFModel)
        'Check Range
        Dim updflag As Boolean = False
        If cmmf.cmmf = 1830007001 Then
            Debug.Print(String.Format("cmmf :{0}", cmmf.cmmf))
        End If
        If cmmf.range <> "" Then
            ValidateRange(New RangeModel With {.range = cmmf.range}, cmmf)
        End If
        'Check CMMF
        Dim pk0(0) As Object
        pk0(0) = cmmf.cmmf
        Dim result As DataRow = DS.Tables("CMMF").Rows.Find(pk0)
        If IsNothing(result) Then
            'Create CMMf
            Dim dr As DataRow = DS.Tables("CMMF").NewRow
            dr.Item("activitycode") = cmmf.activitycode '2
            dr.Item("brandid") = cmmf.brandid '3
            dr.Item("cmmf") = cmmf.cmmf '4
            dr.Item("cmmftype") = cmmf.cmmftype '5
            dr.Item("comfam") = cmmf.comfam '6
            dr.Item("commercialref") = cmmf.commercialref '7
            dr.Item("createon") = cmmf.createon '8
            dr.Item("materialdesc") = cmmf.materialdesc '9
            dr.Item("modelcode") = cmmf.modelcode '10
            dr.Item("plnt") = cmmf.plnt '11
            dr.Item("rangeid") = cmmf.rangeid '12
            dr.Item("rir") = cmmf.rir '13
            dr.Item("sbu") = cmmf.sbu '14
            dr.Item("sorg") = cmmf.sorg '15

            DS.Tables("CMMF").Rows.Add(dr)
            'AddCMMFSB.Append(String.Format("{2}{0}{3}{0}{4}{0}{5}{0}{6}{0}{7}{0}{8:yyyy-MM-dd}{0}{9}{0}{10}{0}{11}{0}{12}{0}{13}{0}{14}{0}{15}{1}", vbTab, vbCrLf, cmmf.activitycode, cmmf.brandid, cmmf.cmmf, cmmf.cmmftype, cmmf.comfam, cmmf.commercialref, cmmf.createon, cmmf.materialdesc, cmmf.modelcode, cmmf.plnt, cmmf.rangeid, cmmf.rir, cmmf.sbu, cmmf.sorg))
            AddCMMFSB.Append(validstr(cmmf.activitycode) & vbTab &
                             validint(cmmf.brandid) & vbTab &
                             cmmf.cmmf & vbTab &
                             validstr(cmmf.cmmftype) & vbTab &
                             validint(cmmf.comfam) & vbTab &
                             validstr(cmmf.commercialref) & vbTab &
                             cmmf.createon & vbTab &
                             validstr(cmmf.materialdesc) & vbTab &
                             validstr(cmmf.modelcode) & vbTab &
                             cmmf.plnt & vbTab &
                             validint(cmmf.rangeid) & vbTab &
                             validstr(cmmf.rir) & vbTab &
                             validstr(cmmf.sbu) & vbTab &
                             cmmf.sorg & vbCrLf)
        Else
            'Update CMMF
            If Not IsDBNull(result.Item("activitycode")) Then  '0
                If result.Item("activitycode") <> cmmf.activitycode Then
                    result.Item("activitycode") = cmmf.activitycode
                    updflag = True
                End If
            Else
                result.Item("activitycode") = cmmf.activitycode
                updflag = True
            End If
            If Not IsDBNull(result.Item("brandid")) Then '1
                If result.Item("brandid") <> cmmf.brandid Then
                    result.Item("brandid") = cmmf.brandid
                    updflag = True
                End If
            Else
                result.Item("brandid") = cmmf.brandid
                updflag = True
            End If
            If Not IsDBNull(result.Item("cmmftype")) Then '2
                If result.Item("cmmftype") <> cmmf.cmmftype Then
                    result.Item("cmmftype") = cmmf.cmmftype
                    updflag = True
                End If
            Else
                result.Item("cmmftype") = cmmf.cmmftype
                updflag = True
            End If
            If Not IsDBNull(result.Item("comfam")) Then '3
                If result.Item("comfam") <> cmmf.comfam Then
                    result.Item("comfam") = cmmf.comfam
                    updflag = True
                End If
            Else

                result.Item("comfam") = cmmf.comfam
                updflag = True

            End If
            Dim commercial = cmmf.commercialref
            If Not IsDBNull(result.Item("commercialref")) Then '4

                If commercial.Length > 15 Then
                    commercial = commercial.Substring(0, 15)
                    cmmf.commercialref = commercial
                End If

                If result.Item("commercialref").ToString.Trim <> cmmf.commercialref Then
                    result.Item("commercialref") = cmmf.commercialref
                    updflag = True
                End If
            Else
                If commercial.Length > 15 Then
                    commercial = commercial.Substring(0, 15)
                    cmmf.commercialref = commercial
                End If
                result.Item("commercialref") = cmmf.commercialref
                updflag = True
            End If
            If Not IsDBNull(result.Item("materialdesc")) Then '5
                If result.Item("materialdesc").ToString.Trim <> cmmf.materialdesc Then
                    result.Item("materialdesc") = cmmf.materialdesc
                    updflag = True
                End If
            Else
                result.Item("materialdesc") = cmmf.materialdesc
                updflag = True
            End If
            If Not IsDBNull(result.Item("modelcode")) Then '6
                If result.Item("modelcode") <> cmmf.modelcode Then
                    result.Item("modelcode") = cmmf.modelcode
                    updflag = True
                End If
            Else
                result.Item("modelcode") = cmmf.modelcode
                updflag = True
            End If
            If Not IsDBNull(result.Item("plnt")) Then '7
                If result.Item("plnt") <> cmmf.plnt Then
                    result.Item("plnt") = cmmf.plnt
                    updflag = True
                End If
            End If
            If Not IsDBNull(result.Item("rir")) Then '8
                If result.Item("rir") <> cmmf.rir Then
                    result.Item("rir") = cmmf.rir
                    updflag = True
                End If
            Else
                result.Item("rir") = cmmf.rir
                updflag = True
            End If

            If Not IsDBNull(result.Item("sbu")) Then '9
                If result.Item("sbu") <> cmmf.sbu Then
                    result.Item("sbu") = cmmf.sbu
                    updflag = True
                End If
            Else
                result.Item("sbu") = cmmf.sbu
                updflag = True
            End If

            If updflag Then
                If UpdCMMFSB.Length > 0 Then
                    UpdCMMFSB.Append(",")
                End If
                UpdCMMFSB.Append(String.Format("[{0}::character varying,{1}::character varying,{2}::character varying,{3}::character varying,{4}::character varying,{5}::character varying,{6}::character varying,{7}::character varying,{8}::character varying,{9}::character varying,{10}::character varying,{11}::character varying]", escstr(cmmf.cmmf), escstr(cmmf.activitycode), escstr(cmmf.brandid), escstr(cmmf.cmmftype), escstr(cmmf.comfam), escstr(cmmf.commercialref), escstr(cmmf.materialdesc), escstr(cmmf.modelcode), escstr(cmmf.plnt), escstr(cmmf.rir), escstr(cmmf.sbu), escstr(cmmf.sorg)))
            End If
        End If
    End Sub

    Public Sub ValidateRange(ByVal range As RangeModel, ByVal cmmf As CMMFModel)
        Dim pk0(0) As Object
        pk0(0) = range.range
        Dim result As DataRow = DS.Tables("Range").Rows.Find(pk0)
        If IsNothing(result) Then
            'Create Range
            Dim dr As DataRow = DS.Tables("Range").NewRow
            dr.Item(1) = range.range

            DS.Tables("Range").Rows.Add(dr)
            range.rangeid = range.getId
            cmmf.rangeid = range.rangeid
            AddRangeSB.Append(range.rangeid & vbTab &
                              range.range & vbCrLf)
        Else
            If Not IsDBNull(result.Item("rangeid")) Then
                cmmf.rangeid = result.Item("rangeid")
            End If


        End If

    End Sub


    Public Sub ClearList()
        AddCMMFSB.Clear()
        UpdCMMFSB.Clear()
    End Sub

    Private Function validstr(ByVal p1 As String) As String
        Return (IIf(p1 = "", "Null", p1)).ToString.Replace("'", "''")
    End Function

    Private Function validint(ByVal p1 As Integer) As String
        Return IIf(p1 = 0, "Null", p1)
    End Function

    Private Function escstr(ByVal p1 As Object) As Object
        If p1.ToString = "" Or p1.ToString = "0" Then
            p1 = "Null"
        Else
            p1 = String.Format("'{0}'", p1.ToString.Replace("'", "''"))
        End If
        Return p1
    End Function

End Class
