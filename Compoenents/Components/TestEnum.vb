Imports System.Threading
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Imports System.Text
Imports Components.PublicClass

Public Class TestEnum
    Dim ds As DataSet
    Dim bs As New BindingSource
    Dim bscb As New BindingSource
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        loaddata()
    End Sub

    Private Sub loaddata()
        ds = New DataSet
        DataGridView1.AutoGenerateColumns = False


        If DbAdapter1.TbgetDataSet("select teamtitlename,tt.teamid from teamtitle tt left join team t on t.teamid = tt.teamid order by teamtitleid;select teamid,teamname from team;", ds) Then

            'bind list

            bs.DataSource = ds.Tables(0)
            bscb.DataSource = ds.Tables(1)

            DataGridView1.DataSource = bs

            Dim mycol As New DataGridViewComboBoxColumn




            'mycol.DataSource = mylist
            'mycol.ValueMember = "Key"
            'mycol.DisplayMember = "Value"
            'DataGridView1.Columns(1).DataPropertyName = "MyTeam"
            Dim column As DataGridViewColumn = New DataGridViewTextBoxColumn()
            column.DataPropertyName = "teamtitlename"
            column.Name = "Knight"
            DataGridView1.Columns.Add(column)

            Dim column1 As DataGridViewColumn = New DataGridViewTextBoxColumn()
            column1.DataPropertyName = "teamid"
            column1.Name = "Knight"
            DataGridView1.Columns.Add(column1)

            'DataGridView1.Columns.Add(CreateComboBoxWithEnums())

            ' Add any initialization after the InitializeComponent() call.
            Dim mylist As New List(Of CTeamTitle)
            Dim mydict As New Dictionary(Of Integer, String)
            For Each myteam In [Enum].GetValues(GetType(Components.TeamTitle))
                Debug.Print(myteam)
                Debug.Print(myteam.ToString)
                'mydict.Add(myteam, myteam.ToString)
                mylist.Add(New CTeamTitle With {.myid = myteam, .myname = myteam.ToString})
            Next
            DirectCast(DataGridView1.Columns(0), DataGridViewComboBoxColumn).DataSource = bscb
            DirectCast(DataGridView1.Columns(0), DataGridViewComboBoxColumn).ValueMember = "teamid"
            DirectCast(DataGridView1.Columns(0), DataGridViewComboBoxColumn).DisplayMember = "teamname"
            DirectCast(DataGridView1.Columns(0), DataGridViewComboBoxColumn).DataPropertyName = "teamid"

        End If
    End Sub
    Private Function CreateComboBoxWithEnums() As DataGridViewComboBoxColumn
        Dim combo As New DataGridViewComboBoxColumn()
        combo.DataSource = [Enum].GetValues(GetType(TeamTitle))
        combo.DataPropertyName = "teamid"
        Return combo
    End Function


    Private Sub DataGridView1_CellPainting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting
        Debug.Print(e.FormattedValue)
    End Sub


    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        Debug.Print("hello")
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()


    End Sub
End Class
Public Class CTeamTitle
    Public Property myid As Integer
    Public Property myname As String
End Class
Public Enum TeamTitle
    ProductDevelopment = 1
    PurchasingTeam = 2
    QualityTeam = 3
    SupplyChain = 4
    SBU = 5
End Enum



'Imports System.Windows.Forms
'Imports System.Collections.Generic
'Public Enum Title
'    King = 10
'    Sir = 20
'End Enum
'Public Class TestEnum 'EnumsAndComboBox
'    Inherits Form

'    Private flow As New FlowLayoutPanel()
'    Private WithEvents checkForChange As Button = New Button()
'    Private knights As List(Of Knight)
'    'Private dataGridView1 As New DataGridView()

'    Public Sub New()
'        InitializeComponent()
'        'MyBase.New()
'        SetupForm()
'        SetupGrid()
'    End Sub

'    Private Sub SetupForm()
'        AutoSize = True
'    End Sub

'    Private Sub SetupGrid()
'        knights = New List(Of Knight)
'        knights.Add(New Knight(10, "Uther", True))
'        knights.Add(New Knight(10, "Arthur", True))
'        knights.Add(New Knight(20, "Mordred", False))
'        knights.Add(New Knight(20, "Gawain", True))
'        knights.Add(New Knight(20, "Galahad", True))

'        ' Initialize the DataGridView.
'        dataGridView1.AutoGenerateColumns = False
'        dataGridView1.AutoSize = True
'        dataGridView1.DataSource = knights

'        dataGridView1.Columns.Add(CreateComboBoxWithEnums())

'        ' Initialize and add a text box column.
'        Dim column As DataGridViewColumn = _
'            New DataGridViewTextBoxColumn()
'        column.DataPropertyName = "Name"
'        column.Name = "Knight"
'        dataGridView1.Columns.Add(column)

'        ' Initialize and add a check box column.
'        column = New DataGridViewCheckBoxColumn()
'        column.DataPropertyName = "GoodGuy"
'        column.Name = "Good"
'        dataGridView1.Columns.Add(column)

'        ' Initialize the form.
'        Controls.Add(dataGridView1)
'        Me.AutoSize = True
'        Me.Text = "DataGridView object binding demo"
'    End Sub

'    Private Function CreateComboBoxWithEnums() As DataGridViewComboBoxColumn
'        Dim combo As New DataGridViewComboBoxColumn()
'        combo.DataSource = [Enum].GetValues(GetType(Title))
'        combo.DataPropertyName = "Title"
'        combo.Name = "Title"
'        Return combo
'    End Function

'#Region "business object"
'    Private Class Knight
'        Private hisName As String
'        Private good As Boolean
'        Private hisTitle As Title

'        Public Sub New(ByVal title As Title, ByVal name As String, _
'            ByVal good As Boolean)

'            hisTitle = title
'            hisName = name
'            Me.good = good
'        End Sub

'        Public Property Name() As String
'            Get
'                Return hisName
'            End Get

'            Set(ByVal Value As String)
'                hisName = Value
'            End Set
'        End Property

'        Public Property GoodGuy() As Boolean
'            Get
'                Return good
'            End Get
'            Set(ByVal Value As Boolean)
'                good = Value
'            End Set
'        End Property

'        Public Property Title() As Title
'            Get
'                Return hisTitle
'            End Get
'            Set(ByVal Value As Title)
'                hisTitle = Value
'            End Set
'        End Property
'    End Class
'#End Region

'    'Public Shared Sub Main()
'    '    Application.Run(New EnumsAndComboBox())
'    'End Sub

'    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

'    End Sub
'End Class

