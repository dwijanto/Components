Public Class BuildTreeView
    Private orig As List(Of OrigObj)
    Public Sub New()
        orig = New List(Of OrigObj)() From
           {New OrigObj() With {.id = 1, .parentid = 0, .category = "About"},
            New OrigObj() With {.id = 2, .parentid = 0, .category = "About Content"},
            New OrigObj() With {.id = 3, .parentid = 1, .category = "About", .mypage = "/About/About1"},
            New OrigObj() With {.id = 4, .parentid = 2, .category = "About Content Folder"},
            New OrigObj() With {.id = 5, .parentid = 4, .category = "About Content", .mypage = "/About/AboutContent"},
            New OrigObj() With {.id = 6, .parentid = 0, .category = "Product"},
            New OrigObj() With {.id = 7, .parentid = 6, .category = "Product List", .mypage = "/Product/ProductList"},
            New OrigObj() With {.id = 8, .parentid = 6, .category = "Customer List", .mypage = "/Product/CustomerList"},
            New OrigObj() With {.id = 9, .parentid = 6, .category = "Product Category", .mypage = "/Product/ProductCategoryList"},
            New OrigObj() With {.id = 10, .parentid = 6, .category = "Product Parameter", .mypage = "/Product/ProductParameter"},
            New OrigObj() With {.id = 11, .parentid = 0, .category = "MVVM"},
            New OrigObj() With {.id = 12, .parentid = 11, .category = "MVVM 1", .mypage = "/MVVM/MVVM1"}
           }
    End Sub
    Public Sub New(ByVal datatable As DataTable)
        orig = New List(Of OrigObj)
        For Each dt As Object In datatable.Rows

            orig.Add(New OrigObj() With {.id = dt.item(0), .parentid = dt.item(1), .category = dt.item(3), .mypage = dt.item(8).ToString})
        Next
    End Sub
    Public Function BuildTreeView() As List(Of NewObj)
        Dim newones As List(Of NewObj) = BuildTree(orig)
        Return newones
        'mytree.ItemsSource = newones
    End Function

    Private Function BuildTree(ByRef source As List(Of OrigObj)) As List(Of NewObj)
        Dim tree As New List(Of NewObj)()
        addLevel(source, tree, 0)
        Return tree
    End Function

    Private Sub addLevel(ByVal source As List(Of OrigObj), ByRef tree As List(Of NewObj), ByVal p3 As Integer)
        Dim child As NewObj
        Dim childcollection As New List(Of NewObj)
        Dim o As List(Of OrigObj) = (From achild In source
                                    Where achild.parentid = p3
                                    Select achild).ToList
        For Each origi In o
            child = New NewObj() With {.id = origi.id, .category = origi.category, .mypage = origi.mypage, .subcategorys = New List(Of NewObj)}
            tree.Add(child)
            childcollection = child.subcategorys
            addLevel(source, childcollection, origi.id)
        Next
    End Sub

    'Private Sub mytree_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.Windows.RoutedPropertyChangedEventArgs(Of System.Object))
    '    Dim view As TreeView = TryCast(sender, TreeView)
    '    If Not (view.SelectedValue Is Nothing) Then
    '        Dim myframe As Frame = CType(CType(Application.Current.RootVisual, ContentControl).Content, FrameworkElement).FindName("ContentFrame")
    '        myframe.Source = New Uri(view.SelectedValue.ToString, UriKind.RelativeOrAbsolute)

    '    End If
    'End Sub
End Class
Public Class OrigObj
    Public Property id As Integer
    Public Property parentid As Integer
    Public Property mypage As String
    Public Property category As String
End Class

Public Class NewObj
    Public Property id As Integer
    Public Property category As String
    Public Property mypage As String
    Public Property subcategorys As List(Of NewObj)
End Class

