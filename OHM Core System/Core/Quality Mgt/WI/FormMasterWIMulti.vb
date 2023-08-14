Imports MySql.Data.MySqlClient
Imports System.IO
Public Class FormMasterWIMulti
    Dim arrImage() As Byte
    Dim arrImage2() As Byte
    Dim arrImage3() As Byte


    Dim Myconnection As String = "server=192.168.12.203;user id=ohmuser;password=;database=stockflow_system" 'con2
    Dim result As Integer
    Dim imgpath As String
    Dim sql As String

    Sub bersih()
        TextBox1.Clear()
        Label2.Text = ""
        PictureBox1.Image = Nothing
        TextBox1.Select()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Try
            Dim OFD As FileDialog = New OpenFileDialog()
            OFD.Filter = "All Files|*.*|Bitmaps|*.bmp|GIFs|*.gif|JPEGs|*.jpg|PNGs|*.png"
            If OFD.ShowDialog() = DialogResult.OK Then
                imgpath = OFD.FileName
                PictureBox1.ImageLocation = imgpath
            End If
            OFD = Nothing
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Dim OFD As FileDialog = New OpenFileDialog()
            OFD.Filter = "All Files|*.*|Bitmaps|*.bmp|GIFs|*.gif|JPEGs|*.jpg|PNGs|*.png"
            If OFD.ShowDialog() = DialogResult.OK Then
                imgpath = OFD.FileName
                PictureBox2.ImageLocation = imgpath
            End If
            OFD = Nothing
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim OFD As FileDialog = New OpenFileDialog()
            OFD.Filter = "All Files|*.*|Bitmaps|*.bmp|GIFs|*.gif|JPEGs|*.jpg|PNGs|*.png"
            If OFD.ShowDialog() = DialogResult.OK Then
                imgpath = OFD.FileName
                PictureBox3.ImageLocation = imgpath
            End If
            OFD = Nothing
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("Input ID!")
        Else
            Try
                Dim mstream As New System.IO.MemoryStream()
                PictureBox1.Image.Save(mstream, PictureBox1.Image.RawFormat)
                arrImage = mstream.GetBuffer()
                Dim FileSize As UInt32
                FileSize = mstream.Length
                mstream.Close()
                conn = New MySqlConnection
                conn.ConnectionString = "server=10.217.4.115; user id=ohmuser; password=; database=stockflow_system;"
                conn.Open()
                sql = "INSERT INTO prodsys2_wi_multi_tbl (`pn`, `wi1`) VALUES (@pn, @wi1);"
                cmd = New MySqlCommand
                With cmd
                    .Connection = conn
                    .CommandText = sql
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@pn", TextBox1.Text)
                    .Parameters.AddWithValue("@wi1", arrImage)
                    .ExecuteNonQuery()
                End With
            Catch ex As MySqlException
                MsgBox(ex.Message)
            Finally
                conn.Close()
                'Try
                'Dim mstream As New System.IO.MemoryStream()
                'PictureBox2.Image.Save(mstream, PictureBox2.Image.RawFormat)
                'arrImage2 = mstream.GetBuffer()
                'Dim FileSize As UInt32
                'FileSize = mstream.Length
                'mstream.Close()
                'conn = New MySqlConnection
                'conn.ConnectionString = "server=10.217.4.115; user id=ohmuser; password=; database=stockflow_system;"
                'conn.Open()
                'sql = "UPDATE prodsys2_wi_multi_tbl SET wi2 = @wi2 WHERE pn = @pn;"
                'cmd = New MySqlCommand
                'With cmd
                '.Connection = conn
                '.CommandText = sql
                '.Parameters.Clear()
                '.Parameters.AddWithValue("@pn", TextBox1.Text)
                '.Parameters.AddWithValue("@wi2", arrImage2)
                '.ExecuteNonQuery()
                'End With
                'Catch ex As MySqlException
                'MsgBox(ex.Message)
                'Finally
                'conn.Close()
                'Try
                'Dim mstream As New System.IO.MemoryStream()
                'PictureBox3.Image.Save(mstream, PictureBox1.Image.RawFormat)
                'arrImage3 = mstream.GetBuffer()
                'Dim FileSize As UInt32
                'FileSize = mstream.Length
                'mstream.Close()
                'conn = New MySqlConnection
                'conn.ConnectionString = "server=10.217.4.115; user id=ohmuser; password=; database=stockflow_system;"
                'conn.Open()
                'sql = "UPDATE prodsys2_wi_multi_tbl SET wi3 = @wi3 WHERE pn = @pn;"
                'With cmd
                '.Connection = conn
                '.CommandText = sql
                '.Parameters.Clear()
                '.Parameters.AddWithValue("@pn", TextBox1.Text)
                '.Parameters.AddWithValue("@wi3", arrImage3)
                'result = .ExecuteNonQuery()
                'End With
                'Catch ex As MySqlException
                'MsgBox(ex.Message)
                'Finally
                'conn.Close()
                If result = 0 Then
                    MsgBox("Error in registering!")
                Else
                    MsgBox("Successfully added images!")
                End If
                TextBox1.Text = ""
                PictureBox1.Image = Nothing
                PictureBox2.Image = Nothing
                PictureBox3.Image = Nothing
                'End Try
                'End Try
            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Input Partnumber for retrieval!")
        Else
            Try
                conn = New MySqlConnection
                conn.ConnectionString = "server=10.217.4.115; user id=ohmuser; password=; database=stockflow_system;"
                conn.Open()
                sql = "SELECT pn, wi1 FROM prodsys2_wi_multi_tbl WHERE pn = @pn;"
                cmd = New MySqlCommand
                With cmd
                    .Connection = conn
                    .CommandText = sql
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@pn", TextBox1.Text)
                    .ExecuteNonQuery()
                End With
                Dim arrImage(), arrImage2(), arrImage3() As Byte
                Dim dt As New DataTable
                da = New MySqlDataAdapter
                da.SelectCommand = cmd
                da.Fill(dt)
                If dt.Rows.Count > 0 Then
                    arrImage = dt.Rows(0).Item(1)
                    arrImage2 = dt.Rows(0).Item(1)
                    arrImage3 = dt.Rows(0).Item(1)

                    Dim mstream As New MemoryStream(arrImage)
                    Dim mstream2 As New MemoryStream(arrImage2)
                    Dim mstream3 As New MemoryStream(arrImage3)

                    PictureBox1.Image = Image.FromStream(mstream)
                    PictureBox2.Image = Image.FromStream(mstream2)
                    PictureBox3.Image = Image.FromStream(mstream3)

                    FormWIDMSInlineMNT1.PictureBox1.Image = Image.FromStream(mstream)
                    FormWIDMSInlineMNT2.PictureBox1.Image = Image.FromStream(mstream2)
                    FormWIDMSInlineMNT3.PictureBox1.Image = Image.FromStream(mstream3)

                    FormWIDMSInlineMNT1.PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                    FormWIDMSInlineMNT2.PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                    FormWIDMSInlineMNT3.PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage

                    Dim screen As Screen
                    screen = Screen.AllScreens(0)
                    FormWIDMSInlineMNT1.StartPosition = FormStartPosition.Manual
                    FormWIDMSInlineMNT1.Location = screen.Bounds.Location + New Point(100, 100)
                    FormWIDMSInlineMNT1.Show()

                    Dim screen2 As Screen
                    screen2 = Screen.AllScreens(1)
                    FormWIDMSInlineMNT2.StartPosition = FormStartPosition.Manual
                    FormWIDMSInlineMNT2.Location = screen2.Bounds.Location + New Point(100, 100)
                    FormWIDMSInlineMNT2.Show()

                    Dim screen3 As Screen
                    screen3 = Screen.AllScreens(2)
                    FormWIDMSInlineMNT3.StartPosition = FormStartPosition.Manual
                    FormWIDMSInlineMNT3.Location = screen3.Bounds.Location + New Point(100, 100)
                    FormWIDMSInlineMNT3.Show()

                Else
                    MsgBox("WI Belum diupload!", vbExclamation)
                End If
            Catch ex As MySqlException
                MsgBox(ex.Message)
            Finally
                conn.Close()
                da.Dispose()
            End Try
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox1.Text = "" Then
            MsgBox("CAN'T EMPTY", vbInformation)
        Else
            Call connection()
            str = "SELECT id, wi1 FROM prodsys2_wi_multi_tbl WHERE pn='" & TextBox1.Text & "'"
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
                        PictureBox1.Image.Save(mstream, PictureBox1.Image.RawFormat)
                        arrImage = mstream.GetBuffer()
                        Dim FileSize As UInt32
                        FileSize = mstream.Length
                        mstream.Close()
                        conn.ConnectionString = Myconnection
                        conn.Open()

                        sql = "UPDATE prodsys2_wi_multi_tbl SET wi1 = @wi1 WHERE id = @id"
                        cmd.Connection = conn
                        cmd.CommandText = sql
                        cmd.Parameters.AddWithValue("@id", Label2.Text)
                        cmd.Parameters.AddWithValue("@wi1", arrImage)
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
                    PictureBox1.Image.Save(mstream, PictureBox1.Image.RawFormat)
                    arrImage = mstream.GetBuffer()
                    Dim FileSize As UInt32
                    FileSize = mstream.Length
                    mstream.Close()
                    conn.ConnectionString = Myconnection
                    conn.Open()

                    sql = "INSERT INTO `prodsys2_wi_multi_tbl`(`pn`, `wi1`) VALUES (@pn, @wi1)"
                    cmd.Connection = conn
                    cmd.CommandText = sql
                    cmd.Parameters.AddWithValue("@pn", TextBox1.Text)
                    cmd.Parameters.AddWithValue("@wi1", arrImage)
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

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If TextBox1.Text = "" Then
            MsgBox("Input Partnumber for retrieval!")
        Else
            Try
                conn = New MySqlConnection
                conn.ConnectionString = "server=10.217.4.115; user id=ohmuser; password=; database=stockflow_system;"
                conn.Open()
                sql = "SELECT pn, wi1 FROM prodsys2_wi_multi_tbl WHERE pn = @pn;"
                cmd = New MySqlCommand
                With cmd
                    .Connection = conn
                    .CommandText = sql
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@pn", TextBox1.Text)
                    .ExecuteNonQuery()
                End With
                'Dim arrImage(), arrImage2(), arrImage3() As Byte

                Dim dt As New DataTable
                da = New MySqlDataAdapter
                da.SelectCommand = cmd
                da.Fill(dt)
                If dt.Rows.Count > 0 Then
                    arrImage = dt.Rows(0).Item(1)
                    arrImage2 = dt.Rows(0).Item(1)
                    arrImage3 = dt.Rows(0).Item(1)

                    Dim mstream As New MemoryStream(arrImage)
                    Dim mstream2 As New MemoryStream(arrImage2)
                    Dim mstream3 As New MemoryStream(arrImage3)

                    PictureBox1.Image = Image.FromStream(mstream)
                    PictureBox2.Image = Image.FromStream(mstream2)
                    PictureBox3.Image = Image.FromStream(mstream3)

                    FormWIDMSInlineMNT1.PictureBox1.Image = Image.FromStream(mstream)
                    FormWIDMSInlineMNT2.PictureBox1.Image = Image.FromStream(mstream2)
                    FormWIDMSInlineMNT3.PictureBox1.Image = Image.FromStream(mstream3)

                    FormWIDMSInlineMNT1.PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                    FormWIDMSInlineMNT2.PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                    FormWIDMSInlineMNT3.PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage

                    Dim screen As Screen
                    screen = Screen.AllScreens(0)
                    FormWIDMSInlineMNT1.StartPosition = FormStartPosition.Manual
                    FormWIDMSInlineMNT1.Location = screen.Bounds.Location + New Point(100, 100)
                    FormWIDMSInlineMNT1.Show()

                    Dim screen2 As Screen
                    screen2 = Screen.AllScreens(1)
                    FormWIDMSInlineMNT2.StartPosition = FormStartPosition.Manual
                    FormWIDMSInlineMNT2.Location = screen2.Bounds.Location + New Point(100, 100)
                    FormWIDMSInlineMNT2.Show()

                    Dim screen3 As Screen
                    screen3 = Screen.AllScreens(2)
                    FormWIDMSInlineMNT3.StartPosition = FormStartPosition.Manual
                    FormWIDMSInlineMNT3.Location = screen3.Bounds.Location + New Point(100, 100)
                    FormWIDMSInlineMNT3.Show()

                Else
                    MsgBox("WI Belum diupload!", vbExclamation)
                End If
            Catch ex As MySqlException
                MsgBox(ex.Message)
            Finally
                conn.Close()
                da.Dispose()
            End Try
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) = 13 Then
            If TextBox1.Text = "" Then
                MsgBox("Input Partnumber for retrieval!")
            Else
                Try
                    conn = New MySqlConnection
                    conn.ConnectionString = "server=10.217.4.115; user id=ohmuser; password=; database=stockflow_system;"
                    conn.Open()
                    sql = "SELECT pn, wi1 FROM prodsys2_wi_multi_tbl WHERE pn = @pn;"
                    cmd = New MySqlCommand
                    With cmd
                        .Connection = conn
                        .CommandText = sql
                        .Parameters.Clear()
                        .Parameters.AddWithValue("@pn", TextBox1.Text)
                        .ExecuteNonQuery()
                    End With
                    Dim arrImage(), arrImage2(), arrImage3() As Byte
                    Dim dt As New DataTable
                    da = New MySqlDataAdapter
                    da.SelectCommand = cmd
                    da.Fill(dt)
                    If dt.Rows.Count > 0 Then
                        arrImage = dt.Rows(0).Item(1)
                        arrImage2 = dt.Rows(0).Item(1)
                        arrImage3 = dt.Rows(0).Item(1)

                        Dim mstream As New MemoryStream(arrImage)
                        Dim mstream2 As New MemoryStream(arrImage2)
                        Dim mstream3 As New MemoryStream(arrImage3)

                        PictureBox1.Image = Image.FromStream(mstream)
                        PictureBox2.Image = Image.FromStream(mstream2)
                        PictureBox3.Image = Image.FromStream(mstream3)

                        FormWIDMSInlineMNT1.PictureBox1.Image = Image.FromStream(mstream)
                        FormWIDMSInlineMNT2.PictureBox1.Image = Image.FromStream(mstream2)
                        FormWIDMSInlineMNT3.PictureBox1.Image = Image.FromStream(mstream3)

                        FormWIDMSInlineMNT1.PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                        FormWIDMSInlineMNT2.PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
                        FormWIDMSInlineMNT3.PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage

                        Dim screen As Screen
                        screen = Screen.AllScreens(0)
                        FormWIDMSInlineMNT1.StartPosition = FormStartPosition.Manual
                        FormWIDMSInlineMNT1.Location = screen.Bounds.Location + New Point(100, 100)
                        FormWIDMSInlineMNT1.Show()

                        Dim screen2 As Screen
                        screen2 = Screen.AllScreens(1)
                        FormWIDMSInlineMNT2.StartPosition = FormStartPosition.Manual
                        FormWIDMSInlineMNT2.Location = screen2.Bounds.Location + New Point(100, 100)
                        FormWIDMSInlineMNT2.Show()

                        Dim screen3 As Screen
                        screen3 = Screen.AllScreens(2)
                        FormWIDMSInlineMNT3.StartPosition = FormStartPosition.Manual
                        FormWIDMSInlineMNT3.Location = screen3.Bounds.Location + New Point(100, 100)
                        FormWIDMSInlineMNT3.Show()

                    Else
                        MsgBox("WI Belum diupload!", vbExclamation)
                    End If
                Catch ex As MySqlException
                    MsgBox(ex.Message)
                Finally
                    conn.Close()
                    da.Dispose()
                End Try
            End If
        End If
    End Sub
End Class