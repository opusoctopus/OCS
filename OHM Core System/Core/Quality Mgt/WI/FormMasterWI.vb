Imports MySql.Data.MySqlClient
Imports System.IO
Public Class FormMasterWI
    Dim Myconnection As String = "server=192.168.12.203;user id=ohmuser;password=;database=stockflow_system" 'con2
    Dim result As Integer
    Dim imgpath As String
    Dim arrImage() As Byte
    Dim sql As String

    Sub bersih()
        TextBox1.Text = ""
        Label2.Text = ""
        Label3.Text = ""
        PictureBox1.Image = Nothing
        TextBox1.Focus()
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Try
            Dim OFD As FileDialog = New OpenFileDialog()
            OFD.Filter = "Image File (*.jpg;*.bmp;*.gif)|*.jpg;*.bmp;*.gif"
            If OFD.ShowDialog() = DialogResult.OK Then
                imgpath = OFD.FileName
                Label3.Text = OFD.FileName
                PictureBox1.ImageLocation = imgpath
            End If
            OFD = Nothing
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("CAN'T EMPTY", vbInformation)
        Else
            Call connection()
            str = "SELECT * FROM prodsys2_wi_tbl WHERE ebt='" & TextBox1.Text & "'"
            cmd = New MySql.Data.MySqlClient.MySqlCommand(str, conn)
            rd = cmd.ExecuteReader
            rd.Read()
            Try
                If rd.HasRows Then
                    Label2.Text = rd("id")
                    Dim tanya
                    tanya = MessageBox.Show("Duplicate WI, are you sure want to replace?", "Replace", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If tanya = vbYes Then
                        conn.Close()
                        Dim mstream As New System.IO.MemoryStream()
                        PictureBox1.Image.Save(mstream, System.Drawing.Imaging.ImageFormat.Jpeg)
                        arrImage = mstream.GetBuffer()
                        Dim FileSize As UInt32
                        FileSize = mstream.Length
                        mstream.Close()
                        conn.ConnectionString = Myconnection
                        conn.Open()

                        sql = "UPDATE prodsys2_wi_tbl SET wi = @wi WHERE id = @id"
                        cmd.Connection = conn
                        cmd.CommandText = sql
                        cmd.Parameters.AddWithValue("@id", Label2.Text)
                        cmd.Parameters.AddWithValue("@wi", arrImage)
                        Dim r As Integer
                        r = cmd.ExecuteNonQuery()
                        cmd.Parameters.Clear()
                        conn.Close()
                        MsgBox("Update successfully!", MsgBoxStyle.Information)
                        Call bersih()
                    End If
                Else
                    conn.Close()
                    Dim mstream As New System.IO.MemoryStream()
                    PictureBox1.Image.Save(mstream, System.Drawing.Imaging.ImageFormat.Jpeg)
                    arrImage = mstream.GetBuffer()
                    Dim FileSize As UInt32
                    FileSize = mstream.Length
                    mstream.Close()
                    conn.ConnectionString = Myconnection
                    conn.Open()

                    sql = "INSERT INTO `prodsys2_wi_tbl`(`ebt`, `wi`) VALUES (@ebt, @wi)"
                    cmd.Connection = conn
                    cmd.CommandText = sql
                    cmd.Parameters.AddWithValue("@ebt", TextBox1.Text)
                    cmd.Parameters.AddWithValue("@wi", arrImage)
                    Dim r As Integer
                    r = cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                    conn.Close()
                    MsgBox("Saved successfully!", MsgBoxStyle.Information)
                    Call bersih()
                End If
            Catch ex As Exception
                MessageBox.Show("ERROR: " & ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub FormMasterWI_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call bersih()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim img() As Byte
        Dim Query As String
        Call connection()
        Query = "SELECT * FROM prodsys2_wi_tbl WHERE ebt = '" & TextBox1.Text & "'"
        cmd = New MySqlCommand(Query, conn)
        rd = cmd.ExecuteReader
        While rd.Read()
            Label2.Text = rd("id")
            img = rd.Item("wi")
            Dim mstream As New MemoryStream(img)
            PictureBox1.Image = Image.FromStream(mstream)
        End While
        rd.Close()
        conn.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call bersih()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            Dim img() As Byte
            Dim Query As String
            Call connection()
            Query = "SELECT * FROM prodsys2_wi_tbl WHERE ebt = '" & TextBox1.Text & "'"
            cmd = New MySqlCommand(Query, conn)
            rd = cmd.ExecuteReader
            While rd.Read()
                Label2.Text = rd("id")
                img = rd.Item("wi")
                Dim mstream As New MemoryStream(img)
                PictureBox1.Image = Image.FromStream(mstream)
            End While
            rd.Close()
            conn.Close()
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        FormListBOM.Show()
    End Sub
End Class