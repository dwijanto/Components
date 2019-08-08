Imports System.Xml
Imports System.IO
Imports System.Text
Public Class FormReadXML

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.FileName = ""
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            Dim widthdict As New Dictionary(Of Integer, Integer)

            Dim xmldoc As New XmlDocument
            Dim xmldoc2 As New XmlDocument
            Dim xmlnodelist As XmlNodeList
            Dim i As Long
            Dim str As New StringBuilder
            Try
                Using fs As New FileStream(OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read)
                    xmldoc.Load(fs)
                End Using

                xmlnodelist = xmldoc.GetElementsByTagName("ss:Column")
                For Each node As XmlNode In xmlnodelist
                    Dim attribute = node.Attributes("ss:Width")
                    If attribute IsNot Nothing Then
                        Debug.Print(attribute.Value)
                    End If
                Next

                Dim mykey As Integer = 0
                Using reader As XmlReader = XmlReader.Create(OpenFileDialog1.FileName)
                    While reader.Read()
                        If reader.IsStartElement() Then
                            If reader.Name = "ss:Column" Then
                                Dim attribute As String = reader("ss:Width")
                                If attribute IsNot Nothing Then
                                    widthdict.Add(mykey, CInt(attribute))
                                    mykey = mykey + 1
                                End If

                            End If
                        End If
                    End While
                End Using


                ListView1.View = View.Details
                ListView1.FullRowSelect = True
                ListView1.Items.Clear()
                ListView1.Columns.Clear()
                'create column

                xmlnodelist = xmldoc.GetElementsByTagName("ss:Row")
                For i = 0 To xmlnodelist(0).ChildNodes.Count - 1
                    ListView1.Columns.Add(xmlnodelist(0).ChildNodes.Item(i).InnerText.Trim, widthdict(i))
                Next
                Dim myarray As Integer = xmlnodelist(0).ChildNodes.Count - 1
                For i = 1 To xmlnodelist.Count - 1
                    xmlnodelist(i).ChildNodes.Item(0).InnerText.Trim()

                    'Dim mytext As String() = {xmlnodelist(i).ChildNodes.Item(0).InnerText.Trim().ToString,
                    '                          xmlnodelist(i).ChildNodes.Item(1).InnerText.Trim().ToString,
                    '                          xmlnodelist(i).ChildNodes.Item(2).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(3).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(4).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(5).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(6).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(7).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(8).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(9).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(10).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(11).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(12).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(13).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(14).InnerText.Trim(),
                    '                          xmlnodelist(i).ChildNodes.Item(15).InnerText.Trim()}
                    Dim mytext(myarray) As String
                    For j = 0 To myarray
                        mytext(j) = xmlnodelist(i).ChildNodes.Item(j).InnerText.Trim().ToString
                    Next

                    Dim item As New ListViewItem(mytext)
                    ListView1.Items.Add(item)



                Next
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
            

        End If
    End Sub

    Private Sub FormReadXML_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class