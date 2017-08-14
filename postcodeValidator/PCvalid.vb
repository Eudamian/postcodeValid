Imports System.Xml
' PCVal.xml has been compiled using data from OS OpenData and is free to use under the OGL - http://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/ 
'Contains OS data © Crown copyright And database right 2017
'Contains Royal Mail data © Royal Mail copyright And Database right 2017
'Contains National Statistics data © Crown copyright And database right 2017
Public Class PCvalid

    Private PClist As tree

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' load in xml of valid UK postcodes (Aug 2017)
            Dim doc As New XmlDocument()
            doc.Load("PCval.xml")
            'Create a tree class with the xml 
            PClist = New tree(doc)
            Label1.Text = PClist.getNextChildren("")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If Not IsNothing(PClist) Then
            Label1.Text = PClist.getNextChildren(TextBox1.Text.Replace(" ", ""))
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim TB1 As String = UCase(TextBox1.Text.Replace(" ", ""))
        ' Format the postcode with space - no need to do this for the validator but it looks better 
        TextBox1.Text = Strings.Left(TB1, TB1.Length - 3) & " " & Strings.Right(TB1, 3)

        If PClist.isValid(TB1) Then
            MsgBox("Postcode is Valid")
        Else
            MsgBox("Postcode is NOT Valid")
        End If

    End Sub

    Private Class tree
        Private c_tree As XmlDocument

        Public Sub New(newXMLdoc As XmlDocument)
            c_tree = newXMLdoc
        End Sub

        Public Function isValid(postcode As String) As Boolean
            postcode = UCase(postcode.Replace(" ", ""))
            isValid = False
            Dim node1 As XmlNode
            Dim res As String = ""
            ' get a single node for the postcode less 1 letter by converting to an xpath. The last letter is stored in an element rather than a node
            node1 = c_tree.SelectSingleNode(StringToXpath(Strings.Left(postcode, postcode.Length - 1)))
            ' if the xmlnode is not nothing then it is a valid part xpath and hence a valid part code 
            If Not IsNothing(node1) Then
                ' check if the last letter is in the element inner text
                If Not node1.InnerText.IndexOf(Strings.Right(postcode, 1)) Then
                    Return True
                End If
            End If
        End Function
        Public Function getNextChildren(partCode As String) As String
            partCode = UCase(partCode.Replace(" ", ""))
            Dim node1 As XmlNode
            Dim res As String = ""
            ' get a single node for the partial postcode by converting to an xpath
            node1 = c_tree.SelectSingleNode(StringToXpath(partCode))
            ' if the xmlnode is not nothing then it is a valid part xpath and hence a valid part code 
            If Not IsNothing(node1) Then
                res = getChildNames(node1)
            End If
            Return res ' invalid codes return "" valid codes return a unsorted string of possible next characters
        End Function
        Private Function StringToXpath(partPC As String) As String
            ' this function turns a part or full postcode into an xpath
            Dim xp As String = "/tree"
            Dim nd As Char() = partPC.ToArray
            For Each nn As Char In nd
                xp = xp & "/" & toXMLok(nn)
            Next
            Return xp
        End Function
        Private Function getChildNames(node As XmlNode) As String
            Dim res As String = ""
            ' In the xml, "zz" denotes that last character posibilities for the postcode (reduces xml file size)
            If node.FirstChild.Name = "zz" Then
                Return node.FirstChild.InnerText
            End If
            ' if not the last character then iterate the nodes and return each node name (viable characters)
            For Each n1 As XmlNode In node
                res = res & fromXMLok(n1.Name)
            Next
            Return res
        End Function
        Private Function toXMLok(a As Char) As Char
            ' 0-9 are not legal xml names so use "a-j" in place 
            toXMLok = UCase(a)
            If Asc(a) < 65 Then
                ' if 0-9 then return a-j
                Return Chr(Asc(a) + 49)
            End If
        End Function
        Private Function fromXMLok(a As Char) As Char
            ' 0-9 are not legal xml names so use "a-j" in place 
            fromXMLok = UCase(a)
            If Asc(a) > 96 Then
                ' if a-j return 0-9 
                Return (Asc(a) - 97).ToString
            End If
        End Function
    End Class
End Class