Public Class CMMFModel
    Public Property cmmf As Long
    Public Property range As String
    Public Property rangedesc As String
    Public Property sorg As Integer
    Public Property plnt As Integer
    Public Property materialdesc As String
    Public Property commercialref As String
    Public Property modelcode As String
    Public Property cmmftype As String
    Public Property sbu As String
    Public Property brandid As Integer
    Public Property rir As String
    Public Property activitycode As String
    Public Property comfam As Integer
    Public Property createon As Date
    Public Property rangeid As Integer    
End Class

Public Class RangeModel
    Public Property range As String  
    Public Property rangeid As Long
    Private dbAdapter1 As DbAdapter = DbAdapter.getInstance

    Public Function getId() As Long
        Dim sqlstr = "select nextval('range_rangeid_seq')"
        Dim ra As Long
        dbAdapter1.ExecuteScalar(sqlstr, ra) 
        Return ra
    End Function
End Class
