Imports System
Imports System.Data
Imports System.Data.OleDb
Public Class Form1
    Dim _koneksiString As String
    Dim _koneksi As New OleDbConnection
    Dim komandambil As New OleDbCommand
    Dim datatabelku As New DataTable
    Dim dataadapterku As New OleDbDataAdapter
    Dim x As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        _koneksiString = "Provider=Microsoft.Jet.OleDb.4.0;" + "Data Source=D:\Campus\Semester V\Kecerdasan Komputasi\Aplikasi-Apotek\database\Aplikasi Apotek.mdb;"
        _koneksi.ConnectionString = _koneksiString
        _koneksi.Open()

        komandambil.Connection = _koneksi
        komandambil.CommandType = CommandType.Text

        komandambil.CommandText = "SELECT * FROM Pegawai"
        dataadapterku.SelectCommand = komandambil
        dataadapterku.Fill(datatabelku)
        Bs_coba.DataSource = datatabelku
        dgv_coba.DataSource = Bs_coba
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim cmdTambah As New OleDbCommand
        Dim tanya As String
        Dim x As DataRow
        cmdTambah.Connection = _koneksi
        cmdTambah.CommandText = "INSERT INTO " + "Pegawai ([ID Pegawai], [Nama Pegawai], Alamat, [Jenis Kelamin], Jabatan)" + "VALUES ('" + TextBox1.Text + "','" + TextBox2.Text + "','" + TextBox3.Text + "','" + ComboBox1.Text + "','" + TextBox4.Text + " ')"
        tanya = MessageBox.Show("Data Ini di Simpan ?", "info", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If tanya = vbYes Then
            cmdTambah.ExecuteNonQuery()
            x = datatabelku.NewRow
            x("ID Pegawai") = TextBox1.Text
            x("Nama Pegawai") = TextBox2.Text
            x("Alamat") = TextBox3.Text
            x("Jenis Kelamin") = ComboBox1.Text
            x("Jabatan") = TextBox4.Text
            datatabelku.Rows.Add(x)
            Bs_coba.DataSource = Nothing
            Bs_coba.DataSource = datatabelku

            dgv_coba.Refresh()
            Bs_coba.MoveFirst()
        End If
    End Sub

    Private Sub dgv_coba_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_coba.CellContentClick
        TextBox1.Text = dgv_coba.CurrentRow.Cells(0).Value.ToString
        TextBox2.Text = dgv_coba.CurrentRow.Cells(1).Value.ToString
        TextBox3.Text = dgv_coba.CurrentRow.Cells(2).Value.ToString 
        ComboBox1.Text = dgv_coba.CurrentRow.Cells(3).Value.ToString
        TextBox4.Text = dgv_coba.CurrentRow.Cells(4).Value.ToString
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim cmdHapus As New OleDbCommand
        cmdHapus.Connection = _koneksi
        cmdHapus.CommandType = CommandType.Text
        x = MessageBox.Show("Yakin Data Akan di Hapus ?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If x = vbYes Then
            cmdHapus.CommandText = "DELETE FROM " + "Pegawai WHERE [ID Pegawai]=" + TextBox1.Text
            cmdHapus.ExecuteNonQuery()
        End If
        Bs_coba.RemoveCurrent()
        dgv_coba.Refresh()

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        ComboBox1.Items.Clear()
        TextBox4.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim cmdUpdate As New OleDbCommand
        cmdUpdate.Connection = _koneksi
        cmdUpdate.CommandType = CommandType.Text
        x = MessageBox.Show("Yakin Data Ingin di Perbarui?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If x = vbYes Then
            cmdUpdate.CommandText = "UPDATE Pegawai SET " +
                "[Nama Pegawai] = '" + TextBox2.Text + "', " +
                "Alamat = '" + TextBox3.Text + "', " +
                "[Jenis Kelamin] = '" + ComboBox1.Text + "', " +
                "Jabatan = '" + TextBox4.Text + "' " +
                "WHERE [ID Pegawai] = " + TextBox1.Text  '
            cmdUpdate.ExecuteNonQuery()
            Dim rowToUpdate As DataRow = datatabelku.Select("[ID Pegawai] = " + TextBox1.Text).FirstOrDefault()
            If rowToUpdate IsNot Nothing Then
                rowToUpdate("ID Pegawai") = TextBox1.Text
                rowToUpdate("Nama Pegawai") = TextBox2.Text
                rowToUpdate("Alamat") = TextBox3.Text
                rowToUpdate("Jenis Kelamin") = ComboBox1.Text
                rowToUpdate("Jabatan") = TextBox4.Text
            End If
            Bs_coba.DataSource = Nothing
            Bs_coba.DataSource = datatabelku
            dgv_coba.Refresh()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        ComboBox1.Items.Clear()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        datatabelku.Clear()

        komandambil.Connection = _koneksi
        komandambil.CommandType = CommandType.Text
        komandambil.CommandText = "SELECT * FROM Pegawai WHERE [ID Pegawai] LIKE '%" + TextBox5.Text + "%' " +
                                       "OR [Nama Pegawai] LIKE '%" + TextBox5.Text + "%' " +
                                       "OR [Alamat] LIKE '%" + TextBox5.Text + "%' " +
                                       "OR [Jenis Kelamin] LIKE '%" + TextBox5.Text + "%' " +
                                       "OR [Jabatan] LIKE '%" + TextBox5.Text + "%'"

        dataadapterku.SelectCommand = komandambil
        dataadapterku.Fill(datatabelku)

        Bs_coba.DataSource = datatabelku
        dgv_coba.DataSource = Bs_coba
        dgv_coba.Refresh()
    End Sub
End Class
