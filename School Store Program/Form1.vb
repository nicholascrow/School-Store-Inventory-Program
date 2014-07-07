Imports System.Data.OleDb

Public Class Form1
    Sub Reload()
        Dim db As String = "SELECT ItemName FROM Items"

        Using cn As New OleDbConnection(Mycn)
            Using da As New OleDbDataAdapter(db, cn)
                da.Fill(ds, "ItemName")
            End Using
        End Using

        With ListBox1
            .DisplayMember = "ItemName"
            .DataSource = ds.Tables("ItemName")
            .SelectedIndex = 0
        End With
        With ComboBox2
            .DisplayMember = "ItemName"
            .DataSource = ds.Tables("ItemName")
            .SelectedIndex = 0
        End With
        With ComboBox1
            .DisplayMember = "ItemName"
            .DataSource = ds.Tables("ItemName")
            .SelectedIndex = 0
        End With
    End Sub
    Dim Mycn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & My.Computer.FileSystem.CurrentDirectory & "\Store Database.mdb;"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim x As New OleDbConnection(Mycn)
        x.Open()
        Dim SQLstr = "INSERT INTO Items VALUES('" & TextBox1.Text & "','" & TextBox2.Text.Replace("$", "").Replace(" ", "") & "','" & TextBox5.Text & "')"
        Dim Command As OleDbCommand = New OleDbCommand(SQLstr, x)
        Command.ExecuteNonQuery()
        x.Close()
    End Sub
    Private ds As New DataSet()

    Protected Overrides Sub OnLoad(e As EventArgs)
        Dim db As String = "SELECT ItemName FROM Items"

        Using cn As New OleDbConnection(Mycn)
            Using da As New OleDbDataAdapter(db, cn)
                da.Fill(ds, "ItemName")
            End Using
        End Using

        With ListBox1
            .DisplayMember = "ItemName"
            .DataSource = ds.Tables("ItemName")
            .SelectedIndex = 0
        End With
        With ComboBox2
            .DisplayMember = "ItemName"
            .DataSource = ds.Tables("ItemName")
            .SelectedIndex = 0
        End With
        With ComboBox1
            .DisplayMember = "ItemName"
            .DataSource = ds.Tables("ItemName")
            .SelectedIndex = 0
        End With
        MyBase.OnLoad(e)
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick

        For Each selectedItem As Object In ListBox1.SelectedItems
            Dim dr As DataRowView = DirectCast(selectedItem, DataRowView)
            Dim result As [String] = dr("ItemName").ToString
            ListBox2.Items.Add(result)
        Next
        Total()
    End Sub
    Sub Total()
        Dim ttl As Integer = 0
        For Each item In ListBox2.Items
            Dim x As New OleDbConnection(Mycn)
            x.Open()
            Dim db As String = "SELECT Cost FROM Items WHERE ItemName='" & item & "';"

            Using cn As New OleDbConnection(Mycn)
                Using da As New OleDbDataAdapter(db, cn)
                    da.Fill(ds, "Cost")
                End Using
            End Using
            ttl += ds.Tables("Cost").Rows(0).Item(0)
            Label14.Text = "Total: $" & ttl
            x.Close()
            ds.Clear()
            Reload()
        Next
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ListBox2.Items.Clear()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ListBox2.Items.Remove(ListBox2.SelectedItem)
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim msgboxstring As String
        For i = 0 To ListBox2.Items.Count - 1
            If i = ListBox2.Items.Count - 1 Then
                msgboxstring = msgboxstring & ListBox2.Items(i) & vbNewLine
            Else
                msgboxstring = msgboxstring & ListBox2.Items(i) & vbNewLine
            End If

        Next
        If MsgBox("Are You Sure You Want To Sell:" & vbNewLine & vbNewLine & msgboxstring, MsgBoxStyle.YesNo) = vbYes Then

            For Each item In ListBox2.Items
                Dim x As New OleDbConnection(Mycn)
                x.Open()
                Dim db As String = "SELECT Cost FROM Items WHERE ItemName='" & item & "';"

                Using cn As New OleDbConnection(Mycn)
                    Using da As New OleDbDataAdapter(db, cn)
                        da.Fill(ds, "Cost")
                    End Using
                End Using
                Dim SQLstr = "INSERT INTO Sales VALUES('" & item & "','" & ds.Tables("Cost").Rows(0).Item(0) & "','" & Date.Now & "')"
                Dim Command As OleDbCommand = New OleDbCommand(SQLstr, x)
                Command.ExecuteNonQuery()
                x.Close()
                ds.Clear()
                Reload()
            Next
            ListBox2.Items.Clear()
            Label14.Text = "Total: $0"
            MsgBox("Sale Complete!")
        Else
        End If
    End Sub
#Region "Update Database"
    Sub UpdateDatabase()
        ' For Each line In TextBox1.Lines
        For i = 0 To 1
            If TextBox3.Text = "" Then GoTo a
            Dim con As New OleDb.OleDbConnection
            Dim ds1 As New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim sql As String
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & My.Computer.FileSystem.SpecialDirectories.Desktop & "\Store Database.mdb;"
            con.Open()
            sql = "UPDATE Items SET ItemCount='" & TextBox3.Text & "' WHERE ItemName='" & AddToCombobox(ComboBox2) & "';"

            da = New OleDb.OleDbDataAdapter(sql, con)

            da.Fill(ds1, "Items")


a:
            con.Close()
            'filling the corresponding textbox named txtName which I have on my windows form
            '.Text = ds.Tables("ClientList").Rows(0).Item("Name")
        Next
        ' Main()

    End Sub
    Shared OleDbConnection As System.Data.OleDb.OleDbConnection
    Shared AddressBookDataAdapter As System.Data.OleDb.OleDbDataAdapter
#End Region
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        UpdateDatabase()
    End Sub
    Function AddToCombobox(cb As ComboBox)
        'For Each selectedItem As Object In cb.SelectedItem
        Dim dr As DataRowView = DirectCast(cb.SelectedItem, DataRowView)
        Dim result As [String] = dr("ItemName").ToString
        Return result
        ' Next
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Reload()
    End Sub
End Class
