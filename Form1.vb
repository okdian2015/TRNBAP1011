Imports SettingLib
Imports SD3Fungsi
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.OracleClient
Imports System.IO
Imports System.String
Imports System.Net.Mail
Imports System.Security.Cryptography.X509Certificates
Public Class frmtrnbap
    Dim DBNull As String
    Dim strupd As String = ""
    Dim con As New OleDb.OleDbConnection(ora_conn)
    Dim cmd As New OleDb.OleDbCommand(strupd, con)
    Dim tblheader, tblgudang, tblppbr, tblcek, tblalamat, tblapprove, tbljab, tblnobap, tblporosesbap As DataTable
    Public nomor, rekakhir, recbap, nmrrecbap As Integer
    Public mplu, tbl_dcid, tbl_gudangid, tbl_lokasiid, tbl_dc_kode, tbl_dc_nama, lokasi_nama, gudang_nama, bap_hdr_id, foldernya As String
    Public arow As Integer
    Dim nrow, ncel As Integer


   

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        AddHandler CType(e.Control, TextBox).KeyPress, AddressOf TextBoxNumeric_keyPress
        Dim tb As TextBox = CType(e.Control, TextBox)
        AddHandler tb.PreviewKeyDown, AddressOf TextBox_PreviewKeyDown
    End Sub
    Private Sub TextBoxNumeric_keyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) 'numeric only

        If Char.IsDigit(CChar(CStr(e.KeyChar))) = False And e.KeyChar <> ControlChars.Back Then e.Handled = True
        'e.Handled = NumericOnly(e)


    End Sub
    Private Sub TextBox_PreviewKeyDown(ByVal sender As Object, ByVal e As PreviewKeyDownEventArgs)
        If e.KeyCode = Keys.Enter Then
            Dim clmName As String = DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).Name
            If clmName = "PLUID" Then
                'SendKeys.Send("{Tab}")
                'SendKeys.Send("{Tab}")
                'SendKeys.Send("{Tab}")
                DataGridView1.CurrentCell = DataGridView1.Rows(arow).Cells("QTY")
            ElseIf clmName = "QTY" Then
                'arow = arow + 1
                DataGridView1.CurrentCell = DataGridView1.Rows(arow).Cells("PLUID")
                'SendKeys.Send("{Left}")
                'SendKeys.Send("{Left}")
                'SendKeys.Send("{Left}")
            End If
            
            'SendKeys.Send("{Up}")
            'MessageBox.Show("Success")    '''''WILL WORK
        End If
        'If e.KeyCode = Keys.Space Then

        '    MessageBox.Show("Success")    '''''WORKS
        'End If
    End Sub
    Private Sub cekfolder()
        Dim strSystemDir As String = Environment.GetEnvironmentVariable("SystemDrive")
        foldernya = strSystemDir & "\backuptempbap"
        If Not IO.Directory.Exists(foldernya) Then
            IO.Directory.CreateDirectory(foldernya)
        End If
    End Sub
    Private Sub frmtrnbap_load(sender As Object, e As EventArgs) Handles MyBase.Load
        cekfolder()
        strKoneksiORCL = getStrKoneksiORCL()
        serverDB = getServerDB()
        userDB = getUserDB()
        passDB = getPassDB()
        Me.Text &= String.Concat(New String() {" V.", Application.ProductVersion, " | TNS : ", serverDB, ""})
        ambilgudang()
        nomor = 0
        nmrrecbap = 0
        ambilalamat()
        dgalamatalert()
        ambildata()
        ambildatabap()
        If recbap = 0 Then
            MessageBox.Show("Belum ada data PPBR yang siap di BAP", "Data kosong", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            isidatabap()
        End If
        If rekakhir = 0 Then
            MessageBox.Show("Belum ada data PPBR", "Data kosong", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            'Me.Close()
        Else
            isidata()
            isigrid()
            ttlnilai()
        End If
    End Sub
    Private Sub ambilgudang()
        query = dtgudang()
        tblgudang = getDS(query, strKoneksiORCL).Tables(0)
        If tblgudang.Rows.Count > 0 Then
            tbl_dcid = tblgudang.Rows(0)("tbl_dcid")
            tbl_gudangid = tblgudang.Rows(0)("tbl_gudangid")
            tbl_lokasiid = tblgudang.Rows(0)("tbl_lokasiid")
            tbl_dc_kode = tblgudang.Rows(0)("tbl_dc_kode")
            tbl_dc_nama = tblgudang.Rows(0)("tbl_dc_nama")
            lokasi_nama = tblgudang.Rows(0)("tbl_lokasi_nama")
            gudang_nama = tblgudang.Rows(0)("tbl_gudang_nama")
        Else
            MessageBox.Show("Data gudang belum ada", "Kode gudang", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Me.Close()
        End If
    End Sub
    Private Sub ambildata()

        query = dt_header()
        tblheader = getDS(query, strKoneksiORCL).Tables(0)
        rekakhir = tblheader.Rows.Count

    End Sub

    Private Sub ambilalamat()
        Dim njab As String
        query = dtalamat()
        tblalamat = getDS(query, strKoneksiORCL).Tables(0)
        TextBoxX7.Text = UCase(tblalamat.Rows(0)("DB_ALAMAT").ToString)

        query = alamatapproval()
        tblapprove = getDS(query, strKoneksiORCL).Tables(0)
        If tblapprove.Rows.Count = 0 Then
            'tambahalamat(strKoneksiORCL)
            'tbljab = getDS(query, strKoneksiORCL).Tables(1)
            dgvapproval.AllowUserToAddRows = True
            dgvapproval.Rows.Insert(0, New String() {"SOMEONE", "SOMEONE@ANYWHERE.CO.ID", "LOGISTIK MGR"})
            'dgvapproval.Rows(0).Cells("NAMA").Value = "SOMEONE"
            'dgvapproval.Rows(0).Cells("ALAMAT").Value = "SOMEONE@ANYWHERE.CO.ID"
            'dgvapproval.Rows(0).Cells("JABATAN").Value = "LOGISTIK MGR"
            dgvapproval.Refresh()
            'njab = "LOGISTIK MGR"
            dgvapproval.AllowUserToAddRows = False
        Else
            Dim con = New Oracle.ManagedDataAccess.Client.OracleConnection(strKoneksiORCL)
            Dim ds As New DataSet
            daAddFunctionDGVSumUpl = New Oracle.ManagedDataAccess.Client.OracleDataAdapter(query, con)
            daAddFunctionDGVSumUpl.Fill(ds, "SumUpl")
            dgvapproval.DataSource = ds.Tables("SumUpl")    'njab = tblapprove.Rows(0)("JABATAN")
        End If
        
        'Dim cols As New DataGridViewComboBoxColumn
        'cols.DataSource = ds.Tables("SumUpl")
        'cols.DisplayMember = "jabatan"
        'cols.DefaultCellStyle.NullValue = njab
        'cols.ValueMember = "JABATAN"
        'dgvapproval.Columns.Add(cols)
        'cols.HeaderText = "JABATAN"
    End Sub
    Private Sub dgalamatalert()
      
        query = alamatalert()
        Dim con As New Oracle.ManagedDataAccess.Client.OracleConnection(strKoneksiORCL)
        Dim dn As New DataSet
        daAddFunctionDGVSumUpl = New Oracle.ManagedDataAccess.Client.OracleDataAdapter(query, con)
        daAddFunctionDGVSumUpl.Fill(dn, "tbalert")
        dgvnotif.DataSource = dn.Tables("tbalert")
        dgvnotif.Refresh()
    End Sub
    Private Sub ambildatabap()
        query = dt_bap()
        tblnobap = getDS(query, strKoneksiORCL).Tables(0)
        recbap = tblnobap.Rows.Count
    End Sub

    Private Sub isidatabap()
        If tblnobap.Rows.Count > 0 Then
            TextBoxX5.Text = tblnobap.Rows(nmrrecbap)("no_ppbr")
            TextBoxX4.Text = tblnobap.Rows(nmrrecbap)("tgl_ppbr")
            TextBoxX6.Text = tblnobap.Rows(nmrrecbap)("jenis_ppbr")
        End If
        
    End Sub
    Private Sub dgvnotif_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvnotif.CellClick
        Dim A As New System.Drawing.Rectangle
        Dim clmName As String = dgvnotif.Columns(dgvnotif.CurrentCell.ColumnIndex).Name
        nrow = dgvnotif.CurrentCell.RowIndex
        ncel = dgvnotif.CurrentCell.ColumnIndex
        If clmName = "Jabatan1" Then
            A = dgvnotif.CurrentCell.DataGridView.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True)
            cmbjab.Height = A.Height
            cmbjab.Width = A.Width - 1

            cmbjab.Left = A.Left + 22
            cmbjab.Top = A.Top + 220

            cmbjab.Items.Clear()
            cmbjab.Items.Add("DCM/DDCM")
            cmbjab.Items.Add("Receiving Retur Supv./ Jr. Mgr")
            cmbjab.Items.Add("Adm. Support Sr. Clerk")
            cmbjab.Items.Add("Development Cab. Mgr. / Jr. Mgr")
            cmbjab.Items.Add("IC Mgr. /Jr. Mgr")
            cmbjab.Items.Add("GA Cab. Mgr. /Jr. Mgr")

            cmbjab.Visible = True
        Else
            cmbjab.Visible = False
        End If
    End Sub



    Private Sub isidata()
        Dim htgotp As String
        Dim table4 As DataTable
        query = cekcountotp(tblheader.Rows(nomor)("no_ppbr"))
        table4 = getDS(query, strKoneksiORCL).Tables(0)
        If table4.Rows.Count > 0 Then
            htgotp = table4.Rows(0)("COUNT_OTP")
        Else
            htgotp = "0"
        End If
        If nomor > tblheader.Rows.Count - 1 Then
            nomor = nomor - 1
        End If
        txtnoppbr1.Text = tblheader.Rows(nomor)("no_ppbr")
        txttglppbr1.Text = tblheader.Rows(nomor)("tgl_ppbr")
        If IsDBNull(tblheader.Rows(nomor)("status_ppbr")) Then
            ButtonX1.Text = "Re-Send Email"
            ButtonX1.Enabled = False
        ElseIf tblheader.Rows(nomor)("status_ppbr") = "PENDING" Then
            ButtonX1.Text = "Re-Send Email (" & htgotp & ")"
            ButtonX1.Enabled = True
        Else
            ButtonX1.Text = "Re-Send Email"
            ButtonX1.Enabled = False
        End If
        If IsDBNull(tblheader.Rows(nomor)("jenis_ppbr")) Then
            ddlppbr.Text = "KOSONG"
        Else
            ddlppbr.Text = tblheader.Rows(nomor)("jenis_ppbr")
        End If
        If IsDBNull(tblheader.Rows(nomor)("status_ppbr")) Then
            lblstatus.Text = "NULL"
        Else
            lblstatus.Text = tblheader.Rows(nomor)("status_ppbr")
        End If
    End Sub

    Private Sub isigrid()
        Dim strppbr As String = txtnoppbr1.Text
        query = dtgridppbr(strppbr)
        tblppbr = getTB(query, strKoneksiORCL).Tables(0)
        Dim con = New Oracle.ManagedDataAccess.Client.OracleConnection(strKoneksiORCL)
        Dim ds As New DataSet
        daAddFunctionDGVSumUpl = New Oracle.ManagedDataAccess.Client.OracleDataAdapter(query, con)
        daAddFunctionDGVSumUpl.Fill(ds, "SumUpl")
        DataGridView1.DataSource = ds.Tables("SumUpl")
    End Sub

    Private Sub trnbapClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Application.Exit()
    End Sub


    Private Sub tblnext_Click(sender As Object, e As EventArgs) Handles tblnext.Click
        If rekakhir > 0 Then
            nomor = nomor + 1
            If nomor = rekakhir Then
                MessageBox.Show("Akhir Data")
                nomor = nomor - 1
            Else
                isidata()
                isigrid()
                ttlnilai()
                DataGridView1.Enabled = False
                tblsave.Enabled = False
            End If
        Else
            MessageBox.Show("Akhir Data")
        End If
        
    End Sub

    Private Sub tblprev_Click(sender As Object, e As EventArgs) Handles tblprev.Click
        If rekakhir > 0 Then
            nomor = nomor - 1
            If nomor < 0 Then
                MessageBox.Show("Awal Data")
                nomor = nomor + 1
            Else
                isidata()
                isigrid()
                ttlnilai()
                DataGridView1.Enabled = False
                tblsave.Enabled = False
            End If
        Else
            MessageBox.Show("Awal Data")
        End If
        
    End Sub

    Private Sub tbladd_Click(sender As Object, e As EventArgs) Handles tbladd.Click
        Dim rcek, rtambah As String
        query = cekdtkosong()
        tblcek = getDS(query, strKoneksiORCL).Tables(0)
        rcek = tblcek.Rows.Count
        If rcek > 0 Then
            MessageBox.Show("Masih ada PPBR yang masih kosong", "Cek data", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            rtambah = tambahdtpbbr(strKoneksiORCL, varKodeDC)
            If rtambah = "Y" Then
                ambildata()
                nomor = rekakhir - 1
                isidata()
                isigrid()
                ddlppbr.Enabled = True
                ddlppbr.SelectedIndex = 1
                DataGridView1.Enabled = True
                formatgrid()
                DataGridView1.Refresh()
                tblsave.Enabled = True

            Else
                MessageBox.Show(rtambah, "error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub tbldel_Click(sender As Object, e As EventArgs) Handles tbldel.Click
        Dim nonya As String = txtnoppbr1.Text
        Dim tglnya As String = txttglppbr1.Text
        If lblstatus.Text = "REJECTED" Or lblstatus.Text = "APPROVED" Then
            'MessageBox.Show("Maaf no PPBR " & nonya & " tidak bisa di hapus", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            MessageBox.Show("PPBR dengan status REJECT/APPROVE tidak bisa dihapus", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        ElseIf lblstatus.Text = "PENDING" Then
            Dim tglskg As String = Today
            'Dim tglskg As String = Format(Date.Now(), "MM/dd/yyyy")
            If txttglppbr1.Text = tglskg Then
                MessageBox.Show("Permohonan Approval sudah terkirim, tidak bisa di hapus!", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                Dim result As Integer = MsgBox("Yakin Hapus Data PPBR no " & nonya & " ?", MsgBoxStyle.YesNo, "Konfirmasi")
                If result = MsgBoxResult.Yes Then
                    Dim shps As String
                    shps = hapusdtpbbr(strKoneksiORCL, nonya)
                    If shps = "Y" Then
                        If nomor > 0 Then
                            nomor = nomor - 1
                        End If
                        MessageBox.Show("No PPBR " & nonya & " berhasil di hapus", "Hapus data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ambildata()
                        If rekakhir > 0 Then
                            isidata()
                            isigrid()
                            ttlnilai()
                        Else
                            DataGridView1.DataSource = Nothing
                            DataGridView1.Refresh()
                            txtnoppbr1.Text = ""
                            txttglppbr1.Text = ""
                        End If

                        DataGridView1.Enabled = False
                        tblsave.Enabled = False
                    Else
                        MessageBox.Show(shps, "Gagal Hapus", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            End If

        Else
                Dim result As Integer = MsgBox("Yakin Hapus Data PPBR no " & nonya & " ?", MsgBoxStyle.YesNo, "Konfirmasi")
                If result = MsgBoxResult.Yes Then
                    Dim shps As String
                    shps = hapusdtpbbr(strKoneksiORCL, nonya)
                    If shps = "Y" Then
                        If nomor > 0 Then
                            nomor = nomor - 1
                        End If
                        MessageBox.Show("No PPBR " & nonya & " berhasil di hapus", "Hapus data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ambildata()
                        If rekakhir > 0 Then
                            isidata()
                            isigrid()
                            ttlnilai()
                        Else
                            DataGridView1.DataSource = Nothing
                            DataGridView1.Refresh()
                            txtnoppbr1.Text = ""
                            txttglppbr1.Text = ""
                        End If

                        DataGridView1.Enabled = False
                        tblsave.Enabled = False
                    Else
                        MessageBox.Show(shps, "Gagal Hapus", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
        End If
    End Sub

    Private Sub tbledit_Click(sender As Object, e As EventArgs) Handles tbledit.Click
        Dim stts As String = lblstatus.Text
        Dim nonya As String = txtnoppbr1.Text
        If stts = "PENDING" Or stts = "APPROVED" Then
            MessageBox.Show("Maaf no PPBR " & nonya & " Tidak bisa di Edit", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            tbledit.Enabled = False
            DataGridView1.Enabled = True
            tblsave.Enabled = True
            ddlppbr.Enabled = True
            tbladd.Enabled = False
            tbldel.Enabled = False
            tblprev.Enabled = False
            tblnext.Enabled = False
            tblhelp.Enabled = False
            ButtonX1.Enabled = False
        End If
    End Sub
  
    

   


    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        Dim clmName As String = DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).Name
        arow = e.RowIndex
        If clmName = "PLUID" Then
            If IsDBNull(DataGridView1.CurrentCell.Value) Then

            Else
                mplu = DataGridView1.CurrentCell.Value
                Dim nmrow As Integer = e.RowIndex
                Dim tnd As String = "N"
                Dim xplu As String
                For i = 0 To nmrow - 1
                    xplu = DataGridView1.Rows(i).Cells("PLUID").Value
                    If xplu = mplu Then
                        tnd = "Y"
                    End If
                Next
                If tnd = "N" Then
                    ambildtplu()
                    If table2.Rows.Count < 1 Then
                        MessageBox.Show("PLU tidak terdaftar", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        DataGridView1.CurrentCell.Value = "0"
                    ElseIf table2.Rows(0)("stok") <= 0 Then
                        MessageBox.Show("Stok untuk PLU " & mplu & " Tidak mencukupi, Stok LPP : " & table2.Rows(0)("stok"), "Portal", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        DataGridView1.CurrentCell.Value = "0"
                    ElseIf IsDBNull(table2.Rows(0)("harga")) Then
                        MessageBox.Show("Acost null tidak bisa di proses", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        DataGridView1.CurrentCell.Value = "0"
                    ElseIf table2.Rows(0)("harga") <= 0 Then
                        MessageBox.Show("Acost 0 tidak bisa di proses", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        DataGridView1.CurrentCell.Value = "0"
                    Else
                        Dim rs As String = DataGridView1.Rows.Count
                        If rs = 2 Then
                            DataGridView1.Rows(e.RowIndex).Cells("NO").Value = 1
                        Else
                            DataGridView1.Rows(e.RowIndex).Cells("NO").Value = DataGridView1.Rows(e.RowIndex - 1).Cells("NO").Value + 1
                        End If
                        DataGridView1.Rows(e.RowIndex).Cells("DESKRIPSI").Value = table2.Rows(0)("deskripsi")
                        DataGridView1.Rows(e.RowIndex).Cells("SATUAN").Value = table2.Rows(0)("satuan")
                        DataGridView1.Rows(e.RowIndex).Cells("HARGA").Value = table2.Rows(0)("harga")

                        'If e.RowIndex = 0 Then
                        '    DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells("QTY")
                        'Else
                        '    DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex - 1).Cells("QTY")
                        'End If

                    End If
                Else
                    MessageBox.Show("PLU " & mplu & " sudah pernah di input", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    DataGridView1.CurrentCell.Value = "0"
                End If

            End If
        End If
        If clmName = "QTY" Then
            mplu = DataGridView1.Rows(e.RowIndex).Cells("PLUID").Value
            cekstok()
            If table3.Rows.Count <= 0 Then
                MessageBox.Show("PLU " & mplu & " Tidak terdaftar di lokasi 03", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim qtyin As String = DataGridView1.CurrentCell.Value
                Dim qtystok As String = table3.Rows(0)("stok3")
                If CInt(qtyin) > CInt(qtystok) Then
                    MessageBox.Show("Qty input lebih besar dari stok ( stok : " & qtystok & " )", "Qty lebih", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    DataGridView1.CurrentCell.Value = 0
                    DataGridView1.Rows(e.RowIndex).Cells("NILAI").Value = 0
                    ttlnilai()
                Else
                    DataGridView1.Rows(e.RowIndex).Cells("NILAI").Value = DataGridView1.Rows(e.RowIndex).Cells("HARGA").Value * qtyin
                    ttlnilai()
                End If
            End If


        End If

    End Sub
    Private Sub ambildtplu()
        'table2.Clear()
        query = cekplu(mplu)
        table2 = getDS(query, strKoneksiORCL).Tables(0)
    End Sub
    Private Sub cekstok()
        'table3.Clear()
        query = cekstokrusak(mplu)
        table3 = getDS(query, strKoneksiORCL).Tables(0)
    End Sub
    Private Sub ttlnilai()
        Dim ttl As Double = 0
        For s = 0 To DataGridView1.Rows.Count - 1
            ttl = ttl + Val(DataGridView1.Rows(s).Cells("NILAI").Value)
        Next
        TextBoxX3.Text = FormatNumber(ttl, 0)

    End Sub

    Private Sub tblsave_Click(sender As Object, e As EventArgs) Handles tblsave.Click
        Dim result As Integer = MsgBox("Simpan data sekarang ?", MsgBoxStyle.YesNo, "Konfirmasi")
        If result = MsgBoxResult.Yes Then
            Dim noppbr, jnsppbr, totitm, totqty, totnilai As String
            Dim plu, sat, qty, harga, nilai, hslupd, hsltbh, updhdr As String
            Dim tblcek1, tblcek2 As DataTable
            noppbr = txtnoppbr1.Text
            jnsppbr = ddlppbr.Text
            'query = ceknoppbr(noppbr)
            'tblcek1 = getDS(query, strKoneksiORCL).Tables(0)
            'If tblcek1.Rows.Count > 0 Then
            If Len(DataGridView1.Rows(0).Cells("PLUID").Value) < 1 Then
                MessageBox.Show("Tidak ada data yang akan disimpan", "Data kosong", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                tblsave.Enabled = False
                tbledit.Enabled = True
                DataGridView1.Enabled = False
                ddlppbr.Enabled = False
                tbladd.Enabled = True
                tbldel.Enabled = True
                tblprev.Enabled = True
                tblnext.Enabled = True
                tblhelp.Enabled = True
                ButtonX1.Enabled = True
                Exit Sub
            End If
            For s = 0 To DataGridView1.Rows.Count - 1

                plu = DataGridView1.Rows(s).Cells("PLUID").Value
                If Len(plu) >= 8 Then
                    query = cekplu(plu)
                    Dim tblaa As DataTable = getDS(query, getStrKoneksiORCL).Tables(0)
                    If tblaa.Rows.Count > 0 Then
                        harga = tblaa.Rows(0)("HARGA")
                    Else
                        harga = DataGridView1.Rows(s).Cells("HARGA").Value
                    End If
                    sat = DataGridView1.Rows(s).Cells("SATUAN").Value
                    qty = DataGridView1.Rows(s).Cells("QTY").Value
                    nilai = harga * qty
                    'nilai = DataGridView1.Rows(s).Cells("NILAI").Value
                    query = cekpluppbr(plu, noppbr)
                    tblcek2 = getDS(query, strKoneksiORCL).Tables(0)
                    If tblcek2.Rows.Count > 0 Then
                        hslupd = updatedtlppbr(strKoneksiORCL, plu, qty, harga, nilai, noppbr)
                        If hslupd = "Y" Then
                        Else
                            MessageBox.Show(hslupd, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    Else
                        hsltbh = tambahdtlppbr(strKoneksiORCL, plu, qty, harga, nilai, noppbr, sat)
                        If hsltbh = "Y" Then
                        Else
                            MessageBox.Show(hsltbh, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                    End If
                End If

            Next

            query = ttlppbr(noppbr)
            tblcek1 = getDS(query, strKoneksiORCL).Tables(0)
            If tblcek1.Rows.Count > 0 Then
                totitm = tblcek1.Rows(0)("ITEM")
                totqty = tblcek1.Rows(0)("QTY")
                totnilai = tblcek1.Rows(0)("NILAI")
                updhdr = updthdrppbr(strKoneksiORCL, noppbr, totitm, totqty, totnilai, jnsppbr, username)
                If updhdr = "Y" Then
                Else
                    MessageBox.Show(updhdr, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
            'Else

            'End If
            tblsave.Enabled = False
            tbledit.Enabled = True
            DataGridView1.Enabled = False
            ddlppbr.Enabled = False
            tbladd.Enabled = True
            tbldel.Enabled = True
            tblprev.Enabled = True
            tblnext.Enabled = True
            tblhelp.Enabled = True
            If lblstatus.Text = "REJECTED" Then
                ButtonX1.Enabled = False
            ElseIf lblstatus.Text = "PENDING" Then
                ButtonX1.Enabled = True
            Else
                ButtonX1.Enabled = True
            End If


            ButtonX10.Enabled = True
            MessageBox.Show("Proses simpan selesai", "Simpan data", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub tblhelp_Click(sender As Object, e As EventArgs) Handles tblhelp.Click
        Me.Enabled = False 'tombol cari toko
        hstatus = "PPBR"
        nstatus = "WHERE STATUS_PPBR = 'PENDING' OR STATUS_PPBR IS NULL OR STATUS_PPBR = 'REJECTED'"
        If frmhelp.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim rowhslcari() As DataRow
            Me.Enabled = True
            rowhslcari = tblheader.Select("NO_PPBR = '" & crnoppbr & "'")
            nomor = rowhslcari(0).Item("noid") - 1
            isidata()
            isigrid()
            ttlnilai()
        End If
    End Sub

    Private Sub ButtonX2_Click(sender As Object, e As EventArgs) Handles ButtonX2.Click
        Dim tblcekal, tblcekjns As DataTable
        query = alamatapproval()
        tblcekal = getDS(query, strKoneksiORCL).Tables(0)
        If tblcekal.Rows.Count > 0 Then
            Dim hsl As String
            Dim noppbrnya As String = txtnoppbr1.Text
            Dim stts As String = lblstatus.Text
            Dim jns As String = ddlppbr.Text
            query = cekjns(noppbrnya)
            tblcekjns = getDS(query, strKoneksiORCL).Tables(0)
            If tblcekjns.Rows.Count > 0 Then
                'Dim jnsnya As String = tblcekjns.Rows(0)("JENIS_PPBR")
                If IsDBNull(tblcekjns.Rows(0)("JENIS_PPBR")) Then
                    MessageBox.Show("Jenis PPBR belum ada", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                    Exit Sub
                Else

                    If stts = "PENDING" Then
                        MessageBox.Show("No PPBR " & noppbrnya & " Sudah pernah di Posting", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                    ElseIf jns = "KOSONG" Then
                        MessageBox.Show("Jenis PPBR belum ada", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                    Else
                        Dim result As Integer = MsgBox("Posting Data Sekarang ? ", MsgBoxStyle.YesNo, "Konfirmasi")
                        If result = MsgBoxResult.Yes Then
                            hsl = postingppbr(strKoneksiORCL, noppbrnya)
                            updlog(strKoneksiORCL, noppbrnya)
                            Dim tblemail As DataTable
                            Dim alamats As String
                            Dim daftaralm As String = ""
                            query = ambildtalamatemail()
                            tblemail = getDS(query, strKoneksiORCL).Tables(0)
                            If tblemail.Rows.Count > 0 Then

                                For s = 0 To tblemail.Rows.Count - 1
                                    alamats = tblemail.Rows(s)("EMAIL_ALAMAT")
                                    ' Email.To.Add(alamats)
                                    If s = 0 Then
                                        daftaralm = alamats
                                    Else
                                        daftaralm = "," & alamats
                                    End If
                                Next
                            End If
                            updatelogmail(daftaralm, strKoneksiORCL, noppbrnya)
                            hsl = postingppbr(strKoneksiORCL, noppbrnya)
                            If hsl = "Y" Or hsl = "APPROVE berhasil" Then
                                query = cekapprovedtgl()
                                Dim tblcekapp As DataTable = getDS(query, strKoneksiORCL).Tables(0)
                                If tblcekapp.Rows.Count > 0 Then
                                    ambildatabap()
                                    isidatabap()
                                    MessageBox.Show("Posting Berhasil", "Posting", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                    nomor = nomor - 1
                                Else
                                    updlog(strKoneksiORCL, noppbrnya)
                                    kirimemail()
                                    ambildatabap()
                                    isidatabap()
                                    MessageBox.Show("Posting Berhasil", "Posting", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                End If
                                ambildata()
                                isidata()
                                isigrid()
                                ttlnilai()
                                DataGridView1.Enabled = False
                                tblsave.Enabled = False
                            Else
                                MessageBox.Show(hsl, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End If
                        End If
                    End If
                End If
            Else
                MessageBox.Show("No PPBR " & noppbrnya & " tidak ditemukan", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End If
        Else
            MessageBox.Show("Tidak dapat posting, alamat email LM belum ada", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub dgvnotif_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvnotif.CellEndEdit
        Dim clmName As String = dgvnotif.Columns(dgvnotif.CurrentCell.ColumnIndex).Name
        'arow = e.RowIndex
        If clmName.ToUpper = "NAMA1" Then



            If IsDBNull(dgvnotif.CurrentCell.Value) Then

            Else
                Dim nm As String = dgvnotif.CurrentCell.Value
                Dim nma As String
                Dim nocek As Integer = 0
                For i = 0 To dgvnotif.Rows.Count - 2
                    nma = dgvnotif.Rows(i).Cells("NAMA1").Value
                    If nm.Trim = nma.Trim Then
                        nocek = nocek + 1
                        If nocek > 1 Then
                            MessageBox.Show("Nama " & nm & " sudah ada", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            dgvnotif.CurrentCell.Value = " "
                            Exit Sub
                        End If
                    End If
                Next
                'Dim ceknm As String = ceknama(nm)
                'If ceknm = "Y" Then

                'End If
            End If

        ElseIf clmName.ToUpper = "ALAMAT1" Then
            If IsDBNull(dgvnotif.CurrentCell.Value) Then

            Else
                Dim alm As String = dgvnotif.CurrentCell.Value
                Dim alma As String
                Dim almcek As Integer = 0
                For s = 0 To dgvnotif.Rows.Count - 2
                    alma = dgvnotif.Rows(s).Cells("ALAMAT1").Value
                    If alm.Trim = alma.Trim Then
                        almcek = almcek + 1
                        If almcek > 1 Then
                            MessageBox.Show("Email " & alm & " sudah ada", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            dgvnotif.CurrentCell.Value = " "
                            Exit Sub
                        End If
                    End If
                Next
                'Dim cekalm As String = cekalamatemail(alm)
                'If cekalm = "Y" Then


            End If
        End If

    End Sub

  

    Private Sub dgvnotif_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgvnotif.CellMouseEnter
        'Dim clmName As String = DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).Name
        'If clmName = "jabatan1" Then
        '    'dgvnotif.Rows(e.RowIndex).Cells(e.ColumnIndex) = New DataGridViewComboBoxCell()
        '    Dim jb As New DataGridViewComboBoxCell
        '    With dgvnotif
        '        'Dim i = dgvnotif.CurrentRow.Index

        '        jb.Items.Add("DCM/DDCM")
        '        jb.Items.Add("Receiving Retur Supv./ Jr. Mgr")
        '        jb.Items.Add("Adm. Support Sr. Clerk")
        '        jb.Items.Add("Development Cab. Mgr. / Jr. Mgr")
        '        jb.Items.Add("IC Mgr. /Jr. Mgr")
        '        jb.Items.Add("GA Cab. Mgr. /Jr. Mgr")

        '        .Item(e.ColumnIndex, e.RowIndex) = jb
        '    End With

        'End If
    End Sub

    Private Sub cmbjab_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbjab.SelectedIndexChanged
        Dim aa As String = cmbjab.SelectedItem
        dgvnotif.Rows(nrow).Cells(ncel).Value = aa
        dgvnotif.Refresh()
        cmbjab.Visible = False
    End Sub

    Private Sub ButtonX8_Click(sender As Object, e As EventArgs) Handles ButtonX8.Click
        Dim cek, nm, almt, jab As String
        Dim result As Integer = MsgBox("Simpan Data Sekarang ? ", MsgBoxStyle.YesNo, "Konfirmasi")
        If result = MsgBoxResult.Yes Then
            'For i = 0 To dgvapproval.Rows.Count - 1
            '    If IsDBNull(dgvapproval.Rows(i).Cells("NAMA").Value) Then
            '    Else

            '        nm = dgvapproval.Rows(i).Cells("NAMA").Value
            '        almt = dgvapproval.Rows(i).Cells("ALAMAT").Value
            '        jab = "LOGISTIK MGR"
            '        cek = ceknama(nm)
            '        If cek = "Y" Then
            '            updatedtemail(strKoneksiORCL, nm, almt, jab)
            '        Else
            '            tambahdtemail(strKoneksiORCL, nm, almt, jab)
            '        End If
            '    End If
            'Next
            'dgvapproval.AllowUserToAddRows = False

            For s = 0 To dgvnotif.Rows.Count - 1
                nm = dgvnotif.Rows(s).Cells("NAMA1").Value
                almt = dgvnotif.Rows(s).Cells("ALAMAT1").Value
                jab = dgvnotif.Rows(s).Cells("JABATAN1").Value
                'jab = dgvnotif.Rows(s).Cells("JABATAN1").Value
                If nm <> "" Then
                    nm = nm.Trim

                    cek = ceknama(nm)
                    If cek = "Y" Then
                        updatedtemail(strKoneksiORCL, nm, almt, jab)
                    Else
                        tambahdtemail(strKoneksiORCL, nm, almt, jab)
                    End If
                End If
            Next
            MessageBox.Show("Proses Simpan Selesai", "Simpan Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub kirimemail()
        'Try
        Dim mailIP, mailHost, mailPort, mailUser, mailPassword, mailSender, alamats, alamatweb, daftaralm As String
        Dim tblotp, tblemail As DataTable
        Dim noppbr As String = txtnoppbr1.Text
        daftaralm = ""
        alamatweb = ambilweb()
        query = ambildtotp(noppbr)
        tblotp = getDS(query, strKoneksiORCL).Tables(0)
        If tblotp.Rows.Count > 0 Then
            Dim tgltrans, jnstrans, notrans, brttrans, kodeotp, linkapp, judul, kodehal As String
            tgltrans = tblotp.Rows(0)("TGL_DOC")
            kodehal = tblotp.Rows(0)("KODE_HAL")
            Dim abc As Date = Convert.ToDateTime(tgltrans).ToString("MM/dd/yyyy")
            Dim tglpros As String = IIf(abc.Day.ToString.Length = 1, "0" & abc.Day, abc.Day) & " - " & IIf(abc.Month.ToString.Length = 1, "0" & abc.Month, abc.Month) & " - " & abc.Year
            Dim jamtrans As Date = Convert.ToDateTime(tgltrans).ToString("HH:mm:ss")
            judul = tblotp.Rows(0)("SUBJECT")
            jnstrans = tblotp.Rows(0)("TYPE_TRANS")
            notrans = tblotp.Rows(0)("NO_DOC")
            brttrans = tblotp.Rows(0)("KETER")
            kodeotp = tblotp.Rows(0)("KODE_OTP")
            linkapp = alamatweb & kodehal
            query = "select * from dc_listmailserver_t where mail_dckode = '" & varKodeDC & "'"
            Dim tblsetmail As New DataTable
            tblsetmail = getDS(query, strKoneksiORCL).Tables(0)
            If tblsetmail.Rows.Count > 0 Then
                mailIP = tblsetmail.Rows(0)("mail_ip")
                mailHost = tblsetmail.Rows(0)("mail_hostname")
                mailPort = tblsetmail.Rows(0)("mail_port")
                mailUser = tblsetmail.Rows(0)("mail_username")
                mailPassword = tblsetmail.Rows(0)("mail_password")
                mailSender = tblsetmail.Rows(0)("mail_sender")
            Else
                mailIP = "mail.indomaret.lan" 'SMTP Indomaret : 192.168.2.240
                mailHost = "virtual.indomaret.co.id"
                mailPort = 587
                mailUser = "oracle_mail"
                mailPassword = "oracle_mail"
                mailSender = "oracle_mail@virtual.indomaret.co.id"
            End If
            
            Dim host As String = mailIP
            Dim mailFrom As String = mailSender
            Dim mailTo As String = ""
            Dim mailSubject As String = judul
            Dim mailTgl As String = Date.Now.ToString("dd-MMM-yyyy HH:mm:ss")
            Dim mailBody As String = "Yth. Logistik Manager <br/><br>Silahkan melakukan approval Permohonan Pemusnahan Barang Rusak (BAP) dengan<br/>klik link " & _
                                     "Approval dan masukkan kode OTP dibawah,<br/><br/>" & _
                                     "Tgl Transaksi &nbsp; &nbsp; : &nbsp;" & tglpros & "<br/>" & _
                                     "Jam Transaksi &nbsp; &nbsp; : &nbsp;" & jamtrans & "<br/>" & _
                                     "Jenis Transaksi &nbsp; : &nbsp;" & jnstrans & "<br/>" & _
                                     "No Transaksi &nbsp; &nbsp; &nbsp; : &nbsp;" & notrans & "<br/>" & _
                                     "Berita &emsp; &emsp; &emsp; : &nbsp;" & brttrans & "<br/>" & _
                                     "Kode OTP &emsp; &emsp; : &nbsp;" & kodeotp & "<br/>" & _
                                     "Link Approval &nbsp; &ensp; : &nbsp;" & linkapp & "<br/><br/>" & _
                                     "Harap jangan membalas e-mail ini, karena e-mail ini dikirim otomatis oleh system.<br/><br/>" & _
                                     "Auto Message Program<br/>" & mailTgl

            Dim Email As New System.Net.Mail.MailMessage()
            Email.From = New MailAddress(mailFrom)
            'Email.To.Add(mailTo)
            query = ambildtalamatemail()
            tblemail = getDS(query, strKoneksiORCL).Tables(0)
            If tblemail.Rows.Count > 0 Then

                For s = 0 To tblemail.Rows.Count - 1
                    alamats = tblemail.Rows(s)("EMAIL_ALAMAT")
                    Email.To.Add(alamats)
                    If s = 0 Then
                        daftaralm = alamats
                    Else
                        daftaralm = "," & alamats
                    End If
                Next
            End If
            'Email.To.Add("rizky.purnomo@indomaret.co.id")
            Email.Subject = mailSubject
            Email.IsBodyHtml = True
            Email.Body = mailBody
            'Email.Attachments.Add(Attachment)
            Email.Priority = MailPriority.High
            Dim mailClient As New System.Net.Mail.SmtpClient()
            mailClient.Host = host

            'Set the port number, if it was provided
            Dim port As Int32 = 0
            Int32.TryParse(mailPort, port) 'port : 557 / 25 / 110

            If port <> 0 Then _
                mailClient.Port = port
            Dim emailCredential As String = mailUser '& "@" & mailHost
            Dim passwordCredential As String = mailPassword
            Dim authenticationInfo As New System.Net.NetworkCredential(emailCredential, passwordCredential)
            mailClient.UseDefaultCredentials = True
            mailClient.Credentials = authenticationInfo
            Try
                mailClient.Send(Email)
                query = updatecountotp(noppbr)
                execQuery(query, strKoneksiORCL)
                updatelogmail(daftaralm, strKoneksiORCL, noppbr)
                MessageBox.Show("Email untuk Approval sudah terkirim, silahkan menunggu approval LM", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MsgBox(ex.ToString)
                'Finally
                '    mailClient.Dispose()
            End Try
        Else
            MessageBox.Show("Kode OTP untuk no PPBR " & noppbr & " tidak ditemukan", "Data salah", MessageBoxButtons.OK, MessageBoxIcon.Hand)
        End If

    End Sub

    Private Sub ButtonX1_Click(sender As Object, e As EventArgs) Handles ButtonX1.Click
        Dim tglppbr As String = txttglppbr1.Text
        resendemail()
        'Dim noppbrnya As String = txtnoppbr1.Text
        'Dim hsl As String = postingppbr(strKoneksiORCL, noppbrnya)
        'If hsl = "Y" Then
        '    kirimemail()
        'Else
        '    MessageBox.Show(hsl, "Gagal kirim OTP", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End If

    End Sub
    Private Sub resendemail()
        'Try
        Dim mailIP, mailHost, mailPort, mailUser, mailPassword, mailSender, alamats, alamatweb As String
        Dim tblotp, tblemail As DataTable
        'Dim noppbr As String = txtnoppbr1.Text
        Dim tglppbr As String = txttglppbr1.Text
        alamatweb = ambilweb()
        query = kirimulangotp(tglppbr)
        tblotp = getDS(query, strKoneksiORCL).Tables(0)
        If tblotp.Rows.Count > 0 Then
            Dim tgltrans, jnstrans, notrans, brttrans, kodeotp, linkapp, judul, kodehal, noppbr, cotp As String
            tgltrans = tblotp.Rows(0)("TGL_DOC")
            kodehal = tblotp.Rows(0)("KODE_HAL")
            Dim abc As Date = Convert.ToDateTime(tgltrans).ToString("MM/dd/yyyy")
            Dim tglpros As String = IIf(abc.Day.ToString.Length = 1, "0" & abc.Day, abc.Day) & " - " & IIf(abc.Month.ToString.Length = 1, "0" & abc.Month, abc.Month) & " - " & abc.Year
            Dim jamtrans As Date = Convert.ToDateTime(tgltrans).ToString("HH:mm:ss")
            judul = tblotp.Rows(0)("SUBJECT")
            jnstrans = tblotp.Rows(0)("TYPE_TRANS")
            notrans = tblotp.Rows(0)("NO_DOC")
            brttrans = tblotp.Rows(0)("KETER")
            kodeotp = tblotp.Rows(0)("KODE_OTP")
            noppbr = tblotp.Rows(0)("NO_DOC")
            cotp = tblotp.Rows(0)("COUNT_OTP") + 1
            linkapp = alamatweb & kodehal
            'mailIP = "192.168.2.34" 'SMTP Indomaret : 192.168.2.240
            'mailIP = "mail.indomaret.lan"
            'mailHost = "virtual.indomaret.co.id"
            'mailPort = 587
            'mailUser = "oracle_mail"
            'mailPassword = "oracle_mail"
            'mailSender = "oracle_mail@virtual.indomaret.co.id"
            query = "select * from dc_listmailserver_t where mail_dckode = '" & varKodeDC & "'"
            Dim tblsetmail As New DataTable
            tblsetmail = getDS(query, strKoneksiORCL).Tables(0)
            If tblsetmail.Rows.Count > 0 Then
                mailIP = tblsetmail.Rows(0)("mail_ip")
                mailHost = tblsetmail.Rows(0)("mail_hostname")
                mailPort = tblsetmail.Rows(0)("mail_port")
                mailUser = tblsetmail.Rows(0)("mail_username")
                mailPassword = tblsetmail.Rows(0)("mail_password")
                mailSender = tblsetmail.Rows(0)("mail_sender")
            Else
                mailIP = "mail.indomaret.lan" 'SMTP Indomaret : 192.168.2.240
                mailHost = "virtual.indomaret.co.id"
                mailPort = 587
                mailUser = "oracle_mail"
                mailPassword = "oracle_mail"
                mailSender = "oracle_mail@virtual.indomaret.co.id"
            End If
            Dim host As String = mailIP
            Dim mailFrom As String = mailSender
            Dim mailTo As String = ""
            Dim mailSubject As String = "[Kirim Ulang Email ke " & cotp & "] " & judul
            Dim mailTgl As String = Date.Now.ToString("dd-MMM-yyyy HH:mm:ss")
            Dim mailBody As String = "[Kirim Ulang Email ke " & cotp & "] " & _
                                     "Yth. Logistik Manager <br/><br>Silahkan melakukan approval Permohonan Pemusnahan Barang Rusak (BAP) dengan<br/>klik link " & _
                                     "Approval dan masukkan kode OTP dibawah,<br/><br/>" & _
                                     "Tgl Transaksi &nbsp; &nbsp; : &nbsp;" & tglpros & "<br/>" & _
                                     "Jam Transaksi &nbsp; &nbsp; : &nbsp;" & jamtrans & "<br/>" & _
                                     "Jenis Transaksi &nbsp; : &nbsp;" & jnstrans & "<br/>" & _
                                     "No Transaksi &nbsp; &nbsp; &nbsp; : &nbsp;" & notrans & "<br/>" & _
                                     "Berita &emsp; &emsp; &emsp; : &nbsp;" & brttrans & "<br/>" & _
                                     "Kode OTP &emsp; &emsp; : &nbsp;" & kodeotp & "<br/>" & _
                                     "Link Approval &nbsp; &ensp; : &nbsp;" & linkapp & "<br/><br/>" & _
                                     "Harap jangan membalas e-mail ini, karena e-mail ini dikirim otomatis oleh system.<br/><br/>" & _
                                     "Auto Message Program<br/>" & mailTgl

            Dim Email As New System.Net.Mail.MailMessage()
            Email.From = New MailAddress(mailFrom)
            'Email.To.Add(mailTo)
            query = ambildtalamatemail()
            tblemail = getDS(query, strKoneksiORCL).Tables(0)
            If tblemail.Rows.Count > 0 Then
                For s = 0 To tblemail.Rows.Count - 1
                    alamats = tblemail.Rows(s)("EMAIL_ALAMAT")
                    Email.To.Add(alamats)
                Next
            End If
            'Email.To.Add("rizky.purnomo@indomaret.co.id")
            Email.Subject = mailSubject
            Email.IsBodyHtml = True
            Email.Body = mailBody
            'Email.Attachments.Add(Attachment)
            Email.Priority = MailPriority.High
            Dim mailClient As New System.Net.Mail.SmtpClient()
            mailClient.Host = host

            'Set the port number, if it was provided
            Dim port As Int32 = 0
            Int32.TryParse(mailPort, port) 'port : 557 / 25 / 110

            If port <> 0 Then _
                mailClient.Port = port
            Dim emailCredential As String = mailUser '& "@" & mailHost
            Dim passwordCredential As String = mailPassword
            Dim authenticationInfo As New System.Net.NetworkCredential(emailCredential, passwordCredential)
            mailClient.UseDefaultCredentials = True
            mailClient.Credentials = authenticationInfo
            Try
                mailClient.Send(Email)
                query = updatecountotp(noppbr)
                execQuery(query, strKoneksiORCL)
                MessageBox.Show("Email untuk Approval sudah terkirim, silahkan menunggu approval LM", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MsgBox(ex.ToString)
                'Finally
                '    mailClient.Dispose()
            End Try
        Else
            'MessageBox.Show("Kode OTP untuk no PPBR " & noppbr & " tidak ditemukan", "Data salah", MessageBoxButtons.OK, MessageBoxIcon.Hand)
            MessageBox.Show("Gagal kirim ulang Email", "Gagal", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

    End Sub

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        System.Net.ServicePointManager.ServerCertificateValidationCallback = Function(s, cert, chain, sslPolicyErrors)
                                                                                 Return True
                                                                             End Function
    End Sub

    Private Sub tblload_Click(sender As Object, e As EventArgs) Handles tblload.Click
        tblload.Enabled = False
        tblproses.Enabled = True
        Dim noppbr As String = TextBoxX5.Text
        ' Dim tblgbap As DataTable
        query = dtgridbap(noppbr)
        'tblgbap = getDS(query, strKoneksiORCL).Tables(0)
        Dim con = New Oracle.ManagedDataAccess.Client.OracleConnection(strKoneksiORCL)
        Dim ds As New DataSet
        daAddFunctionDGVSumUpl = New Oracle.ManagedDataAccess.Client.OracleDataAdapter(query, con)
        daAddFunctionDGVSumUpl.Fill(ds, "SumUpl")
        DataGridView2.DataSource = ds.Tables("SumUpl")
    End Sub

    Private Sub ButtonX3_Click(sender As Object, e As EventArgs) Handles ButtonX3.Click
        If recbap > 0 Then
            nmrrecbap = nmrrecbap + 1
            If nmrrecbap = recbap Then
                MessageBox.Show("Akhir Data")
                nmrrecbap = nmrrecbap - 1
            Else
                isidatabap()
                Dim tblksg As New DataTable
                tblksg.Clear()
                DataGridView2.DataSource = tblksg
                DataGridView2.Refresh()
                'Call CType(DataGridView2.DataSource, DataTable).Rows.Clear()
                'DataGridView2.Rows.Clear()
                tblload.Enabled = True
                tblproses.Enabled = False
            End If
        Else
            MessageBox.Show("Akhir Data")
        End If

    End Sub

    Private Sub ButtonX4_Click(sender As Object, e As EventArgs) Handles ButtonX4.Click
        If recbap > 0 Then
            nmrrecbap = nmrrecbap - 1
            If nmrrecbap < 0 Then
                MessageBox.Show("Awal Data")
                nmrrecbap = nmrrecbap + 1
            Else
                isidatabap()
                Dim tblksg As New DataTable
                tblksg.Clear()
                DataGridView2.DataSource = tblksg
                DataGridView2.Refresh()
                'Call CType(DataGridView2.DataSource, DataTable).Rows.Clear()
                'DataGridView2.Rows.Clear()
                tblload.Enabled = True
                tblproses.Enabled = False
            End If
        Else
            MessageBox.Show("Awal Data")
        End If

    End Sub

    Private Sub ButtonX5_Click(sender As Object, e As EventArgs) Handles ButtonX5.Click
        Me.Enabled = False 'tombol cari toko
        nstatus = "WHERE STATUS_PPBR = 'APPROVED' AND NO_BAP IS NULL "
        hstatus = "PPBR"
        If frmhelp.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim rowhslcari() As DataRow
            Me.Enabled = True
            rowhslcari = tblnobap.Select("NO_PPBR = '" & crnoppbr & "'")
            nmrrecbap = rowhslcari(0).Item("noid") - 1
            isidatabap()
        End If
    End Sub
    Public nohdrid As String
    Private Sub tblproses_Click(sender As Object, e As EventArgs) Handles tblproses.Click
        Dim noppbr As String = TextBoxX5.Text
        Dim tbltmp As DataTable
        Dim result As Integer = MsgBox("Proses BAP untuk PPBR " & noppbr & "  ? ", MsgBoxStyle.YesNo, "Konfirmasi")
        If result = MsgBoxResult.Yes Then

            Dim hsl As String
            hsl = loaddtbap(strKoneksiORCL, noppbr)
            If hsl = "Y" Then
                query = ambldtTRNbap(noppbr)
                tbltmp = getDS(query, strKoneksiORCL).Tables(0)
                If tbltmp.Rows.Count > 0 Then
                    Dim noref As String = tbltmp.Rows(0)("BAP_HDR_ID")
                    nohdrid = noref
                    Dim rslt As String = prosesdtbap(strKoneksiORCL, noref)
                    If rslt = "Y" Then
                        query = AMBILDTBAPPROS(noppbr)
                        Dim tblss As DataTable = getDS(query, strKoneksiORCL).Tables(0)
                        If tblss.Rows.Count > 0 Then
                            Dim nonya As Integer = CInt(tblss.Rows(0)("NO_BAP"))
                            Dim tglnya As Date = tblss.Rows(0)("TGL_BAP")
                            If createbapdev(strKoneksiORCL, nonya, tglnya) Then
                                kirimftp()
                                kirimulangftp()
                            Else
                                MsgBox("Gagal kirim data ke development")
                                'Exit Sub
                            End If
                            MessageBox.Show("Proses BAP Berhasil", "Proses BAP", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            sttslap = "BAP"
                            lapnobap = tblss.Rows(0)("NO_BAP")
                            laptglbapitm1 = tblss.Rows(0)("TGL_BAP")
                            ambillapbap(lapnobap, laptglbapitm1)
                            ambiljnsppbr(lapnoppbr)
                            Me.Enabled = False
                            frmlaporan.ShowDialog()
                            Me.Enabled = True
                            Me.Focus()
                            tblload.Enabled = True
                            tblproses.Enabled = False
                            ambildatabap()
                            If recbap = 0 Then
                                MessageBox.Show("Belum ada data PPBR yang siap di BAP", "Data kosong", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                TextBoxX4.Text = ""
                                TextBoxX5.Text = ""
                                TextBoxX6.Text = ""
                                DataGridView2.DataSource = Nothing
                                DataGridView2.Refresh()
                            Else
                                If nmrrecbap <= 0 Then
                                    nmrrecbap = 0
                                Else
                                    nmrrecbap = nmrrecbap - 1
                                End If
                                isidatabap()
                            End If
                            sttslap = "BAP"
                        End If
                    Else
                        MessageBox.Show(rslt, "Gagal Proses BAP", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If

                End If
            Else
                MessageBox.Show(hsl, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub

    Private Sub rboppbr_CheckedChanged(sender As Object, e As EventArgs) Handles rboppbr.CheckedChanged
        If rboppbr.Checked Then
            ButtonX6.Enabled = True
        Else
            ButtonX6.Enabled = False
        End If
    End Sub

    Private Sub rbobap_CheckedChanged(sender As Object, e As EventArgs) Handles rbobap.CheckedChanged
        If rbobap.Checked Then
            ButtonX7.Enabled = True
        Else
            ButtonX7.Enabled = False
        End If
    End Sub

    Private Sub rborkpbap1_CheckedChanged(sender As Object, e As EventArgs) Handles rborkpbap1.CheckedChanged
        If rborkpbap1.Checked Then
            tglrbapn1.Enabled = True
            tglrbapn2.Enabled = True
        Else
            tglrbapn1.Enabled = False
            tglrbapn2.Enabled = False
        End If
    End Sub

    Private Sub rborkpbap2_CheckedChanged(sender As Object, e As EventArgs) Handles rborkpbap2.CheckedChanged
        If rborkpbap2.Checked Then
            tglrbapi1.Enabled = True
            tglrbapi2.Enabled = True
            ddlbapi.Enabled = True
        Else
            tglrbapi1.Enabled = False
            tglrbapi2.Enabled = False
            ddlbapi.Enabled = False
        End If
    End Sub

    Private Sub btncetak_Click(sender As Object, e As EventArgs) Handles btncetak.Click
        If rboppbr.Checked Then
            lapppbr()
        ElseIf rbobap.Checked Then
            lapbap()
        ElseIf rborkpbap1.Checked Then
            rkpbapno()
        ElseIf rborkpbap2.Checked Then
            rkpbapitem()
        End If
    End Sub
    Private Sub lapppbr()
        If Len(txttglppbr.Text) < 1 Then
            MessageBox.Show("Silahkan pilih no PPBR dahulu", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Else
            sttslap = "PPBR"
            lapnoppbr = txtnoppbr.Text
            Dim tbltemp As DataTable
            query = lapjudul(lapnoppbr)
            tbltemp = getDS(query, strKoneksiORCL).Tables(0)
            If tbltemp.Rows.Count > 0 Then
                lapjnsppbr = tbltemp.Rows(0)("JENIS_PPBR")
                laptglppbr = tbltemp.Rows(0)("TGL_PPBR")
                Me.Enabled = False
                frmlaporan.ShowDialog()
                Me.Enabled = True
                Me.Focus()
            Else
                MessageBox.Show("Data PPBR tidak ada", "Gagal Lap", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End If

        End If
    End Sub
    Private Sub lapbap()
        If Len(txtnobap.Text) < 1 Then
            MessageBox.Show("Silahkan pilih no BAP dahulu", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Else
            sttslap = "BAP"
            lapnobap = txtnobap.Text
            laptglbapitm1 = txttglbap.Text
            ambillapbap(lapnobap, laptglbapitm1)
            ambiljnsppbr(lapnoppbr)
            Me.Enabled = False
            frmlaporan.ShowDialog()
            Me.Enabled = True
            Me.Focus()
        End If
    End Sub
    Private Sub rkpbapno()
        If Len(tglrbapn1.Text) < 1 Or Len(tglrbapn2.Text) < 1 Then
            MessageBox.Show("Silahkan pilih tanggal periode BAP dahulu", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Else
            sttslap = "REPRKPNOMOR"
            laptglbapnmr1 = tglrbapn1.Text
            laptglbapnmr2 = tglrbapn2.Text
            Me.Enabled = False
            frmlaporan.ShowDialog()
            Me.Enabled = True
            Me.Focus()
        End If
    End Sub
    Private Sub rkpbapitem()
        If Len(tglrbapi1.Text) < 1 Or Len(tglrbapi2.Text) < 1 Then
            MessageBox.Show("Silahkan pilih tanggal periode BAP dahulu", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Else
            sttslap = "REPRKPITEM"
            laptglbapitm1 = tglrbapi1.Text
            laptglbapitm2 = tglrbapi2.Text
            If ddlbapi.SelectedItem = "ALL" Then
                sitmbap = "DIV,DEPT,KAT"
            Else
                sitmbap = ddlbapi.SelectedItem
            End If

            Me.Enabled = False
            frmlaporan.ShowDialog()
            Me.Enabled = True
            Me.Focus()
        End If
    End Sub

    Private Sub ButtonX6_Click(sender As Object, e As EventArgs) Handles ButtonX6.Click
        hstatus = "LAPPPBR"
        nstatus = "WHERE STATUS_PPBR IS NOT NULL "
        Me.Enabled = False
        If frmhelp.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.Enabled = True
            txtnoppbr.Text = nocr
            txttglppbr.Text = tglcr
        Else
            Me.Enabled = True
        End If
    End Sub

    Private Sub ButtonX7_Click(sender As Object, e As EventArgs) Handles ButtonX7.Click
        hstatus = "LAPBAP"
        Me.Enabled = False
        If frmhelp.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.Enabled = True
            txtnobap.Text = nocr
            txttglbap.Text = tglcr
        Else
            Me.Enabled = True
        End If
    End Sub

    Private Sub ButtonX10_Click(sender As Object, e As EventArgs) Handles ButtonX10.Click
        Dim fileDlg As New FolderBrowserDialog
        Dim dir_plu As String
        'fileDlg.Filter = "CSV Files | *.csv"
        'fileDlg.FilterIndex = 1
        'fileDlg.RestoreDirectory = True
        If fileDlg.ShowDialog() = DialogResult.OK Then
            dir_plu = fileDlg.SelectedPath
            Dim filecsv As String = dir_plu & "\PPBR.CSV"
            If System.IO.File.Exists(filecsv) Then
                Dim csvConn As New System.Data.Odbc.OdbcConnection(KoneksiToTXT(dir_plu))
                Dim csvCom As New Odbc.OdbcCommand("", csvConn)
                Dim csvAdapt As New Odbc.OdbcDataAdapter("", csvConn)
                Dim Datacsv As New DataTable
                Dim Datacsv1 As New DataTable
                Dim strupd As String = ""
                Dim con As New OleDb.OleDbConnection(ora_conn)
                Dim cmd As New OleDb.OleDbCommand(strupd, con)
                con.Open()
                csvConn.Open()
                DeklarasiCreateSchemaBAP()
                CreateSchemaBAP("PPBR", dir_plu, 3, True)
                csvAdapt.SelectCommand.CommandText = "SELECT PLU,QTY FROM PPBR.CSV"
                Datacsv.Clear()
                csvAdapt.Fill(Datacsv)
                If Datacsv.Rows.Count > 0 Then

                    Try
                        Me.Cursor = Cursors.WaitCursor
                        'Application.DoEvents()
                        'Dim hdrid As String = DC_TRNBAP_HDR_TDataGridView.CurrentCell.DataGridView(2, DC_TRNBAP_HDR_TDataGridView.CurrentCell.RowIndex).Value.ToString
                        cmd.CommandText = "delete from dc_ppbr_csv"
                        cmd.ExecuteNonQuery()
                        Dim plu, qty As String
                        For i = 0 To Datacsv.Rows.Count - 1
                            plu = Datacsv.Rows(i)(0)
                            qty = Datacsv.Rows(i)(1)
                            tambahitemppbr(strKoneksiORCL, plu, qty)
                        Next
                        query = "SELECT * FROM DC_PPBR_CSV"
                        Dim cekupl As DataTable = getDS(query, strKoneksiORCL).Tables(0)
                        If cekupl.Rows.Count > 0 Then
                            MessageBox.Show("Upload file berhasil", "oke", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            filterppbr(strKoneksiORCL)

                            Me.Cursor = Cursors.Default
                            MessageBox.Show("Proses filter file berhasil", "oke", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            sttslap = "UPLOADCSV"
                            Me.Enabled = False
                            frmlaporan.ShowDialog()
                            Me.Enabled = True
                            Me.Focus()
                            Dim tblcekcsv As DataTable = cekuploadcsv(strKoneksiORCL)
                            If tblcekcsv.Rows.Count > 0 Then
                                Dim rcek, noppbr As String
                                query = cekdtkosong()
                                tblcek = getDS(query, strKoneksiORCL).Tables(0)
                                rcek = tblcek.Rows.Count
                                If rcek > 0 Then
                                    noppbr = tblcek.Rows(0)("NO_PPBR")
                                Else
                                    tambahdtpbbr(strKoneksiORCL, varKodeDC)
                                    Dim tbltbh As DataTable = ambilnoppbr(strKoneksiORCL)
                                    noppbr = tbltbh.Rows(0)("NO_PPBR")
                                End If
                                Dim tblitem As DataTable = ambildtuploadcsv(strKoneksiORCL)
                                If tblitem.Rows.Count > 0 Then

                                    Dim uplu, uqty, sat, harga, nilai As String
                                    For i = 0 To tblitem.Rows.Count - 1
                                        uplu = tblitem.Rows(i)("PLU")
                                        uqty = tblitem.Rows(i)("QTY")
                                        sat = tblitem.Rows(i)("SATUAN")
                                        harga = tblitem.Rows(i)("HARGA")
                                        nilai = tblitem.Rows(i)("NILAI")
                                        tambahdtlppbr(strKoneksiORCL, uplu, uqty, harga, nilai, noppbr, sat)
                                    Next
                                    Dim totitm, totqty, totnilai, jnsppbr, updhdr As String
                                    query = ttlppbr(noppbr)
                                    Dim tblcek1 As DataTable = getDS(query, strKoneksiORCL).Tables(0)
                                    If tblcek1.Rows.Count > 0 Then
                                        totitm = tblcek1.Rows(0)("ITEM")
                                        totqty = tblcek1.Rows(0)("QTY")
                                        totnilai = tblcek1.Rows(0)("NILAI")
                                        jnsppbr = ""
                                        updhdr = updthdrppbr(strKoneksiORCL, noppbr, totitm, totqty, totnilai, jnsppbr, username)
                                        If updhdr = "Y" Then
                                            ambildata()
                                            nomor = rekakhir - 1
                                            isidata()
                                            isigrid()
                                            ddlppbr.Enabled = True
                                            ddlppbr.SelectedIndex = 1
                                            DataGridView1.Enabled = True
                                            tblsave.Enabled = True
                                            tbladd.Enabled = False
                                            tbledit.Enabled = False
                                            ButtonX1.Enabled = False
                                            ButtonX10.Enabled = False
                                        Else
                                            MessageBox.Show(updhdr, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        End If
                                    End If
                                End If
                            Else
                                MessageBox.Show("Tidak ada item untuk di proses PPBR", "Data kosong", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            End If
                        End If

                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                End If
            Else
                MessageBox.Show("Tidak ada file LISTPPBR di direktori " & dir_plu, "File tidak ada", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End If
        End If
    End Sub
    Public fieldlist() As String
    Public Sub DeklarasiCreateSchemaBAP()
        '  1    2       3   4     5       6   7   8     9     10      11    12    13      14    15     16  17   18    19   20      21    22  23    24  25   26   27     28 
        'RECID|RTYPE|DOCNO|PRDCD|SINGKAT|DIV|QTY|TOKO|TGL_PB|PKM_AKH|LENGTH|WIDTH|HEIGHT|MINOR|GUDANG|
        'MAX|PKM_G|QTYM1|SPD|PRICE|GROSS|PTAG|QTY_MAN|PPN|KM|JAMIN|JAMOUT|NO_SJ
        ReDim fieldlist(2)
        fieldlist(0) = "PLU CHAR"
        fieldlist(1) = "QTY CHAR"
        'fieldlist(2) = "KETERANGAN CHAR"
    End Sub
    Sub CreateSchemaBAP(ByVal Namatabel As String, ByVal path As String, ByVal JmlCol As Integer, ByVal OVR As Boolean)
        If OVR = True Then
            If IO.File.Exists(path & "\Schema.Ini") Then
                IO.File.Delete(path & "\Schema.Ini")
            End If
        End If


        Dim Sw As New IO.StreamWriter(path & "\Schema.Ini", True)
        Dim StrLine As String
        Dim WrLine As String

        Sw.WriteLine("[" & Namatabel & ".CSV]")
        Sw.WriteLine("ColNameHeader=True")
        Sw.WriteLine("CharacterSet=ANSI")
        Sw.WriteLine("Format=Delimited(,)")
        Sw.WriteLine("TextDelimiter=none")

        For X As Integer = 0 To JmlCol - 1
            StrLine = "COL" & X + 1 & "="
            'StrLine = "COL" & Trim(Str(X + 1)) & "="
            WrLine = StrLine & fieldlist(X)
            Sw.WriteLine(WrLine)
        Next
        Sw.Flush() : Sw.Close()
    End Sub



    Private Sub ButtonX9_Click(sender As Object, e As EventArgs) Handles ButtonX9.Click
        Dim nama As String = dgvnotif.CurrentRow.Cells("NAMA1").Value
        nama = nama.Trim
        Dim result As Integer = MsgBox("Yakin Hapus Data mail " & nama & " ?", MsgBoxStyle.YesNo, "Konfirmasi")
        If result = MsgBoxResult.Yes Then
            Dim hasil As String
            hasil = hapusemail(strKoneksiORCL, nama)
            If hasil = "Y" Then
                MessageBox.Show("Data berhasil dihapus", "OKE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                dgalamatalert()
                dgvnotif.Refresh()
            Else
                MessageBox.Show(hasil, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub
    Private Sub formatgrid()

        DataGridView1.Columns("NO").ReadOnly = True
        DataGridView1.Columns("DESKRIPSI").ReadOnly = True
        DataGridView1.Columns("SATUAN").ReadOnly = True
        DataGridView1.Columns("HARGA").ReadOnly = True
        DataGridView1.Columns("NILAI").ReadOnly = True
        DataGridView1.Columns("NO").DefaultCellStyle.BackColor = Color.FromArgb(192, 192, 255)
        DataGridView1.Columns("DESKRIPSI").DefaultCellStyle.BackColor = Color.FromArgb(192, 192, 255)
        DataGridView1.Columns("SATUAN").DefaultCellStyle.BackColor = Color.FromArgb(192, 192, 255)
        DataGridView1.Columns("HARGA").DefaultCellStyle.BackColor = Color.FromArgb(192, 192, 255)
        DataGridView1.Columns("NILAI").DefaultCellStyle.BackColor = Color.FromArgb(192, 192, 255)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ambildata()
        isidata()
        isigrid()
        ttlnilai()
        ambildatabap()
        If recbap = 0 Then
            'MessageBox.Show("Belum ada data PPBR yang siap di BAP", "Data kosong", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TextBoxX4.Text = ""
            TextBoxX5.Text = ""
            TextBoxX6.Text = ""
            DataGridView2.DataSource = Nothing
            DataGridView2.Refresh()
        Else
            If nmrrecbap <= 0 Then
                nmrrecbap = 0
            Else
                nmrrecbap = nmrrecbap - 1
            End If
            isidatabap()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        dgvapproval.Enabled = True
        Button3.Enabled = True
        Button2.Enabled = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim nm, almt, jab, cek, cek2 As String
        Dim result As Integer = MsgBox("Simpan Data Sekarang ? ", MsgBoxStyle.YesNo, "Konfirmasi")
        If result = MsgBoxResult.Yes Then
            For i = 0 To dgvapproval.Rows.Count - 1
                If IsDBNull(dgvapproval.Rows(i).Cells("NAMA").Value) Then
                Else

                    nm = dgvapproval.Rows(i).Cells("NAMA").Value
                    almt = dgvapproval.Rows(i).Cells("ALAMAT").Value
                    jab = "LOGISTIK MGR"
                    nm = nm.Trim
                    almt = almt.Trim
                    cek = ceknama(nm)
                    If cek = "Y" Then
                        MessageBox.Show("Nama " & nm & " sudah digunakan", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        'updatedtemail(strKoneksiORCL, nm, almt, jab)
                    ElseIf cek = "X" Then
                        cek2 = cekalamatemail(almt)
                        If cek2 = "Y" Then
                            MessageBox.Show("Alamat " & almt & " sudah digunakan", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                        ElseIf cek2 = "X" Then
                            Dim cek3 As String = ceklm()
                            If cek3 = "Y" Then
                                Dim ceksimpan As String = updatedtemaillm(strKoneksiORCL, nm, almt, jab)
                                If ceksimpan = "Y" Then
                                    MessageBox.Show("Data berhasil di simpan", "OKE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Else
                                    MessageBox.Show(ceksimpan, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                End If
                            Else
                                Dim ceksimpan As String = tambahdtemail(strKoneksiORCL, nm, almt, jab)
                                If ceksimpan = "Y" Then
                                    MessageBox.Show("Data berhasil di simpan", "OKE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Else
                                    MessageBox.Show(ceksimpan, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                End If
                            End If
                        End If

                    Else
                        MessageBox.Show(cek, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            Next
        End If
        Button3.Enabled = False
        Button2.Enabled = True
        dgvapproval.Enabled = False
    End Sub
    Sub DataTable2CSV(ByVal table As DataTable, ByVal filename As String, _
       ByVal sepChar As String, ByVal lokasiFileBackup As String)
        Dim writer As System.IO.StreamWriter = Nothing
        Try
            writer = New System.IO.StreamWriter(lokasiFileBackup & "/" & filename)

            ' first write a line with the columns name
            Dim sep As String = ""
            Dim builder As New System.Text.StringBuilder
            For Each col As DataColumn In table.Columns
                builder.Append(sep).Append(col.ColumnName)
                sep = sepChar
            Next
            writer.WriteLine(builder.ToString())

            ' then write all the rows
            For Each row As DataRow In table.Rows
                sep = ""
                builder = New System.Text.StringBuilder

                For Each col As DataColumn In table.Columns
                    builder.Append(sep).Append(row(col.ColumnName))
                    sep = sepChar
                Next
                writer.WriteLine(builder.ToString())
            Next
        Catch ex As Exception
            'LogUtility.CatatStartTransfer(Application.ProductName & " - " & Application.ProductVersion, ex.Message)
            'WriteErrorGeneral(ex.Message)
        Finally
            If Not writer Is Nothing Then writer.Close()
        End Try
    End Sub
    Private Sub kirimftp()
        query = "SELECT * FROM Q_TRF_CSV WHERE q_filename ='BAP_NODOC'"
        Dim tblftp As DataTable = getDS(query, strKoneksiORCL).Tables(0)
        If tblftp.Rows.Count > 0 Then
            Dim cquery, cseparator, cfilename As String
            cquery = tblftp.Rows(0)("Q_QUERY")
            cseparator = tblftp.Rows(0)("Q_SEPERATOR")
            cfilename = tblftp.Rows(0)("Q_NAMAFILE")
            Dim tblisiftp As DataTable = getDS(cquery, strKoneksiORCL).Tables(0)
            If tblisiftp.Rows.Count > 0 Then
                Dim namafilenya As String = foldernya & "\" & cfilename
                If IO.File.Exists(namafilenya) Then
                    IO.File.Delete(namafilenya)
                End If
                DataTable2CSV(tblisiftp, cfilename, cseparator, foldernya)
                If IO.File.Exists(namafilenya) Then
                    Dim alamatftp, usrftp, passftp, folderftp As String
                    Dim ukuran As Long = FileIO.FileSystem.GetFileInfo(namafilenya).Length()
                    query = "select * from dc_ftp_t where pga_type = 'NPK_DEV'"
                    Dim tblalamat As DataTable = getDS(query, strKoneksiORCL).Tables(0)
                    If tblalamat.Rows.Count > 0 Then
                        alamatftp = tblalamat.Rows(0)("PGA_IPADDRESS")
                        usrftp = tblalamat.Rows(0)("PGA_USERNAME")
                        passftp = tblalamat.Rows(0)("PGA_PASSWORD")
                        folderftp = tblalamat.Rows(0)("PGA_FOLDER")
                        If UploadFTPFiles("ftp://" & alamatftp & "/" & folderftp, usrftp, passftp, foldernya, cfilename, cfilename, False) Then
                            updatelogftp(cfilename, alamatftp, folderftp, Now().ToString, "SUKSES", ukuran, nohdrid)
                            MsgBox("File " & cfilename & "berhasil dikirim")
                        Else
                            updatelogftp(cfilename, alamatftp, folderftp, Now().ToString, "Error :" & ftperrmsg, ukuran, nohdrid)
                            MsgBox("Gagal kirim file via FTP")
                        End If
                    Else
                        MsgBox("Setting FTP belum ada")
                        Exit Sub
                    End If
                Else
                    MsgBox("Gagal create file")
                End If
            Else
                MsgBox("Gagal kirim data ke development")
                Exit Sub
            End If
        Else
            MsgBox("Gagal kirim data ke development")
            Exit Sub
        End If
    End Sub
    Private Sub updatelogftp(ByVal fname As String, ftpaddrs As String, ftpfolder As String, ByVal timelog As String, ByVal status As String, ByVal filesize As String, ByVal nhdrid As String)
        query = "Insert into FTP_TRANS_TRFLOG(FTP_FILENAME,FTP_APPNAME,FTP_IP,FTP_LOCATION,FTP_TIMELOG_TRF,FTP_FILESIZE_TRF,FTP_STATUS_TRF,FTP_TYPE_TRANS,FTP_HDRID)" & _
                " VALUES('" & fname & "','DCTRNBAP.EXE','" & ftpaddrs & "','" & ftpfolder & "',SYSDATE,'" & filesize & "','" & status & "','BAP_DC','" & nhdrid & "')"
        execQuery(query, strKoneksiORCL)
    End Sub
    Private Sub kirimulangftp()
        query = "select * from ftp_trans_trflog where ftp_status_trf <> 'SUKSES' and trunc(ftp_timelog_trf) > trunc(sysdate -7)"
        Dim tblcek As DataTable = getDS(query, strKoneksiORCL).Tables(0)
        If tblcek.Rows.Count > 0 Then
            query = "select * from dc_ftp_t where pga_type = 'NPK_DEV'"
            Dim tblalamat As DataTable = getDS(query, strKoneksiORCL).Tables(0)
            If tblalamat.Rows.Count > 0 Then
                Dim alamatftp, usrftp, passftp, folderftp As String
                alamatftp = tblalamat.Rows(0)("PGA_IPADDRESS")
                usrftp = tblalamat.Rows(0)("PGA_USERNAME")
                passftp = tblalamat.Rows(0)("PGA_PASSWORD")
                folderftp = tblalamat.Rows(0)("PGA_FOLDER")
                For d = 0 To tblcek.Rows.Count - 1
                    Dim namafile As String = tblcek.Rows(d)("ftp_filename")
                    Dim cekfile As String = foldernya & "\" & namafile
                    If IO.File.Exists(cekfile) Then
                        If UploadFTPFiles("ftp://" & alamatftp & "/" & folderftp, usrftp, passftp, foldernya, namafile, namafile, False) Then
                            query = "update ftp_trans_trflog set ftp_status_trf = 'SUKSES', ftp_timelog_trf = sysdate where ftp_filename = '" & namafile & "'"
                            execQuery(query, strKoneksiORCL)
                            execQuery("commit", strKoneksiORCL)
                        End If
                    End If
                Next
            End If
        End If
    End Sub
End Class
