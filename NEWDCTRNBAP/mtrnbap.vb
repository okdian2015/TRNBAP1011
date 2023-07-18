Imports SettingLib
Imports SD3Fungsi
'Imports System.Data
Imports System.Data.OleDb
'Imports System.Data.OracleClient
Imports System.IO
Imports System.Windows
Imports System.Net
Imports Oracle.ManagedDataAccess.Client
Module mtrnbap
    Public newencrypt As New SettingLib.Class1
    Public ora_conn As String = "Provider=MSDAORA.1;User ID=" & newencrypt.GetVariabel("UserOrcl") & ";Password = " & newencrypt.GetVariabel("PasswordOrcl") & " ;Data Source=" & newencrypt.GetVariabel("ServerOrcl")
    Public dc As String
    Public kodedc As String
    Public kodedc1 As String
    Public pluid As String
    Public pluid1 As String
    Public curlok As String
    Public curalok As String
    Public curseq As String
    Public curbrg As String
    Public alokid As String
    Public curtoc As String
    Public supid As String
    Public tipe As String
    Public kodeid As String
    Public DCID As String
    Public gudangid As String
    Public lokasiid As String
    Public hdrid As String
    Public plukode As String
    Public nodoc As String
    Public trans_id As String
    Public login As Integer = 0
    Public crnoppbr As String

    Public strKoneksiORCL As String
    Public strKoneksi_OleDB As String
    Public userDB, passDB, serverDB As String
    Public daAddFunctionDGVSumUpl As Oracle.ManagedDataAccess.Client.OracleDataAdapter
    Public daAddFunctionDGVDetUpl As Oracle.ManagedDataAccess.Client.OracleDataAdapter
    Public mDataSet As DataSet

    Public query, nmproses As String
    Public tblusr, table1, table2, table3, carino, tblcekrcd, tblppbr As DataTable
    Public varKodeDC, varDCID, varLokID, varNamaDC, varNpwp, varAlamatdc As String
    Public daAddFunctionDtl1 As OleDbDataAdapter
    Public daAddFunctionDtl2 As OleDbDataAdapter
    Public ip_address, m_dc_id, username, password, nik, nstatus, sttslap, hstatus, nocr, tglcr As String
    Public lapnoppbr, lapnobap, laptglbapnmr1, laptglbapnmr2, laptglbapitm1, laptglbapitm2, sitmbap, lapjnsppbr, laptglppbr As String
    Public Function ValidateLogin()
        Return "select a.USER_NAME , a.USER_PASSWORD," & _
            " a.USER_NIK, a.USER_FK_TBL_DCID, b.TBL_LOKASI_KODE" & _
            " from DC_User_T a, dc_dc_full_v b where a.USER_FK_TBL_LOKASIID = b.TBL_LOKASIID" & _
            " and USER_NAME = '" & username & "' and USER_PASSWORD = '" & password & "' And USER_FLAG_HANDHELD = 'Y'" & _
            " and TBL_LOKASI_KODE = '01'" ' harus user baik yg bisa login
    End Function
    Public Function getUserDB() As String
        Return newencrypt.GetVariabel("UserOrcl")
    End Function

    Public Function getPassDB() As String
        Return newencrypt.GetVariabel("PasswordOrcl")
    End Function

    Public Function getServerDB() As String
        Return newencrypt.GetVariabel("ServerOrcl")
    End Function

    Public Function getStrKoneksiORCL() As String

        strKoneksiORCL = ""
        Try
            userDB = getUserDB()
            passDB = getPassDB()
            serverDB = getServerDB()

            If serverDB <> "" Then
                Dim odp As String = getODP()
                strKoneksiORCL = "Data Source=" & odp & ";;User Id=" & userDB & ";;Password=" & passDB & ""
                'strKoneksiORCL = "Provider=MSDAORA.1;User ID=" & userDB & ";Password = " & passDB & ";Data Source = " & serverDB
                'strKoneksiORCL = "User ID=" & userDB & ";Password = " & passDB & ";Data Source = " & serverDB
            Else
                strKoneksiORCL = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return strKoneksiORCL
    End Function

    Public Function getODP() As String
        Return newencrypt.GetVariabel("ODPOrcl")
    End Function

    Public Function getStrKoneksi_OleDB() As String
        strKoneksi_OleDB = ""
        Try
            userDB = getUserDB()
            passDB = getPassDB()
            serverDB = getServerDB()
            If serverDB <> "" Then
                'strKoneksiORCL = "Provider=MSDAORA.1;User ID=" & userDB & ";Password = " & passDB & ";Data Source = " & serverDB
                strKoneksi_OleDB = "Provider=MSDAORA.1;User ID=" & userDB & ";Password = " & passDB & ";Data Source = " & serverDB
            Else
                strKoneksi_OleDB = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return strKoneksi_OleDB
    End Function

    Public Function getDS(ByVal sql As String, ByVal conStr As String) As DataSet
        Dim connection As New Oracle.ManagedDataAccess.Client.OracleConnection(conStr)
        Dim command As New Oracle.ManagedDataAccess.Client.OracleCommand("", connection)
        Dim adapter As New Oracle.ManagedDataAccess.Client.OracleDataAdapter
        Dim ds As New DataSet
        If (connection.State = ConnectionState.Open) Then
            connection.Close()
        End If
        connection.Open()
        'MessageBox.Show(sql)
        command.CommandText = sql
        adapter.SelectCommand = command
        adapter.Fill(ds, "data")
        connection.Close()
        Return ds
    End Function
    Dim strHostName As String = System.Net.Dns.GetHostName()
    Dim strIPAddress As String = System.Net.Dns.GetHostByName(strHostName).AddressList(0).ToString()
    Private IpLocal As String = strIPAddress
    Private VAR_PROD_NAME As String = "RETURTOKOPC" '(Application.ProductName).ToUpper
    Private VAR_PROD_VER As String = Application.ProductVersion
    Public Sub Cek_Program_2(ByVal KodeDc)
        Try
            Dim cls As New SettingLib.Class1
            Dim Ret
            strKoneksi_OleDB = getStrKoneksi_OleDB()
            Ret = cls.GetVersiODP(strKoneksiORCL, KodeDc, "DCTRNBAP.EXE", VAR_PROD_VER, IpLocal)
            If Ret.ToString.Contains("OKE") = False Then
                MsgBox(Ret)
                End
            End If
        Catch ex As Exception
            MsgBox(ex.StackTrace.ToString())
            'MsgBox(strKoneksi_OleDB & KodeDc & "RETURTOKOPC.EXE" & VAR_PROD_VER & IpLocal)
        End Try
    End Sub
    Public Function getTB(ByVal sql As String, ByVal conStr As String) As DataSet
        Dim connection As New OracleConnection(conStr)
        Dim command As New OracleCommand("", connection)
        Dim adapter As New OracleDataAdapter
        Dim TB As New DataSet
        If (connection.State = ConnectionState.Open) Then
            connection.Close()
        End If
        connection.Open()
        'MessageBox.Show(sql)
        command.CommandText = sql
        adapter.SelectCommand = command
        adapter.Fill(TB, "data")
        connection.Close()
        Return TB
    End Function

    Public Function execScalar(ByVal sql As String, ByVal conStr As String) As Integer
        Dim connection As New OracleConnection(conStr)
        Dim cmd As New OracleCommand("", connection)
        Dim adapter As New OracleDataAdapter
        Dim ds As New DataSet
        If (connection.State = ConnectionState.Open) Then
            connection.Close()
        End If
        connection.Open()
        cmd.CommandText = sql
        Return cmd.ExecuteScalar()
        connection.Close()
    End Function

    Public Function execQuery(ByVal sql As String, ByVal conStr As String) As Boolean
        Dim connection As New OracleConnection(conStr)
        Dim cmd As New OracleCommand("", connection)
        Dim adapter As New OracleDataAdapter
        Dim ds As New DataSet
        If (connection.State = ConnectionState.Open) Then
            connection.Close()
        End If
        connection.Open()
        cmd.CommandText = sql
        cmd.ExecuteNonQuery()
        connection.Close()
        Return True
    End Function
    Public Function getKodeDC() As String
        query = " SELECT a.tbl_dc_kode, a.tbl_dc_nama, a.tbl_dcid, a.tbl_npwp_dc, b.db_alamat" & _
                " FROM dc_tabel_dc_t a, DC_SETUP_DB b" & _
                " WHERE a.tbl_dc_kode=NVL(b.db_kode_Dc,a.tbl_dc_kode)" & _
                " ORDER BY a.tbl_dc_kode ASC"
        table1 = getDS(query, strKoneksiORCL).Tables(0)
        varKodeDC = table1.Rows(0)(0).ToString
        varNamaDC = table1.Rows(0)("tbl_dc_nama").ToString
        varNpwp = table1.Rows(0)("tbl_npwp_dc").ToString
        varAlamatdc = table1.Rows(0)("db_alamat").ToString
        Return varKodeDC
        Return varNamaDC
        Return varNpwp
        Return varAlamatdc
    End Function
    Public Function cekUser(ByVal ConstrOra As String, ByVal txtUsername As String, ByVal txtPassword As String) As String
        Dim Scon As New OracleConnection(ConstrOra)
        Dim Scom As New OracleCommand("", Scon)

        Try
            'cUser = txtUsername.Replace("'", "").ToUpper
            'cPass = txtPassword.Replace("'", "").ToUpper
            Scon.Open()
            Scom.CommandText = "SELECT COUNT(*) AS JUM" & _
                                " FROM DC_USER_T" & _
                                " WHERE USER_PRIVS = 'Y' " & _
                                    " AND USER_NAME = '" & username & "'" & _
                                    " AND UPPER(USER_PASSWORD)='" & password & "'"
            If Scom.ExecuteScalar = 0 Then
                MessageBox.Show("NIK Atau Password Salah !!", "Perhatian", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            Else

                'MessageBox.Show("Berhasil masuk !!", "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return True
            End If
        Catch ex As Exception
            MessageBox.Show("Error saat ambil data user - " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
    Public Function dtgudang()
        Return "SELECT TBL_DCID, TBL_GUDANGID, TBL_LOKASIID, TBL_DC_KODE, TBL_DC_NAMA, TBL_LOKASI_NAMA, TBL_GUDANG_NAMA FROM DC_DC_FULL_V WHERE TBL_LOKASI_NAMA = 'BARANG RUSAK'"
    End Function
    Public Function dtalamat()
        Return "Select * from DC_SETUP_DB"
    End Function
    Public Function alamatapproval()
        Return "SELECT EMAIL_NAMA AS NAMA, EMAIL_ALAMAT AS ALAMAT, EMAIL_JABATAN AS JABATAN FROM DC_BAP_EMAIL_T WHERE TRIM(EMAIL_JABATAN) = 'LOGISTIK MGR'"
    End Function
    Public Function alamatalert()
        Return "SELECT EMAIL_NAMA AS NAMA1, EMAIL_ALAMAT AS ALAMAT1, EMAIL_JABATAN AS JABATAN1 FROM DC_BAP_EMAIL_T WHERE TRIM(EMAIL_JABATAN) <> 'LOGISTIK MGR'"
    End Function
    Public Function cekjns(ByVal noppbr As String) As String
        Return "SELECT TBL_DC_KODE, JENIS_PPBR, NO_PPBR, TGL_PPBR," & _
               " STATUS_PPBR, USER_PPBR, NO_BAP, TGL_BAP, TOT_ITEM, TOT_QTY, TOT_NILAI FROM dc_ppbr_hdr_t" & _
               " WHERE NO_PPBR = '" & noppbr & "'"
    End Function

    Public Function dt_header() As String
        Return "SELECT row_number() OVER(ORDER BY no_ppbr) AS noid, TBL_DC_KODE, JENIS_PPBR, NO_PPBR, TGL_PPBR," & _
               " STATUS_PPBR, USER_PPBR, NO_BAP, TGL_BAP, TOT_ITEM, TOT_QTY, TOT_NILAI FROM dc_ppbr_hdr_t" & _
               " WHERE STATUS_PPBR = 'PENDING' OR STATUS_PPBR IS NULL OR STATUS_PPBR = 'REJECTED' ORDER BY NO_PPBR"
    End Function
    Public Function dt_bap() As String
        Return "SELECT row_number() OVER(ORDER BY no_ppbr) AS noid, TBL_DC_KODE, JENIS_PPBR, NO_PPBR, TGL_PPBR," & _
               " STATUS_PPBR, USER_PPBR, NO_BAP, TGL_BAP, TOT_ITEM, TOT_QTY, TOT_NILAI FROM dc_ppbr_hdr_t" & _
               " WHERE STATUS_PPBR = 'APPROVED' AND NO_BAP IS NULL ORDER BY NO_PPBR"
    End Function
    Public Function cekcountotp(ByVal noppbr As String) As String
        Return "select * from dc_otp_log where no_doc = '" & noppbr & "' and type_trans = 'PPBR'"
    End Function
    Public Function dtgridppbr(ByVal noppbr As String) As String
        Return "SELECT row_number() over(order by no_fk_ppbr) as no, a.pluid, b.mbr_singkatan as deskripsi, a.satuan, a.qty, TRUNC(a.harga,3) AS HARGA, TRUNC(a.nilai,3) AS NILAI" & _
               " FROM dc_ppbr_dtl_t a, dc_barang_t b where b.mbr_pluid = a.pluid and a.no_fk_ppbr = '" & noppbr & "'"
    End Function
    Public Function dtgridbap(ByVal noppbr As String) As String
        Return "SELECT row_number() OVER(ORDER BY no_fk_ppbr) AS no, a.pluid as plu, b.mbr_singkatan as deskripsi, a.satuan as sat, a.qty, trunc(a.harga,3) as harga, trunc(a.nilai,3) as nilai" & _
               " FROM dc_ppbr_dtl_t a, dc_barang_t b where b.mbr_pluid = a.pluid and a.no_fk_ppbr = '" & noppbr & "' order by a.pluid"
    End Function
    Public Function cekdtkosong() As String
        Return "select * from dc_ppbr_hdr_t where tot_item is null"
    End Function

    Public Function updatecountotp(ByVal noppbr As String) As String
        Return "update dc_otp_log set count_otp = count_otp + 1 where no_doc = '" & noppbr & "' and type_trans = 'PPBR'"
    End Function
    Public Function updlog(ByVal constr As String, noppbr As String) As String
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "UPDATE DC_OTP_LOG SET status = 'PENDING', Tgl_stat = SYSDATE WHERE no_doc = '" & noppbr & "' AND Type_trans = 'PPBR'"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()

            conection.Open()
            query = "UPDATE DC_PPBR_HDR_T SET STATUS_PPBR = 'PENDING' WHERE NO_PPBR = '" & noppbr & "' AND TGL_PPBR = TRUNC(SYSDATE)"
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()
            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function updatelogmail(ByVal alamat As String, ByVal constr As String, ByVal noppbr As String) As String
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "UPDATE DC_OTP_LOG SET email = '" & alamat & "', tgl_sendmail = sysdate, count_otp = nvl(count_otp,0) + 1 where trunc(tgl_stat) = trunc(SYSDATE) and no_doc = '" & noppbr & "' AND Type_trans = 'PPBR'"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()
            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function tambahalamat(ByVal constr As String) As String
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "INSERT INTO DC_BAP_EMAIL_T (EMAIL_NAMA, EMAIL_ALAMAT, EMAIL_JABATAN) VALUES ('SOMEONE', 'SOMEONE@EXAMPLE.CO.ID', 'LOGISTIK MGR')"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()
            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function tambahdtpbbr(ByVal constr As String, ByVal kodedc As String) As String
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "INSERT INTO DC_PPBR_HDR_T (TBL_DC_KODE) VALUES ('" & kodedc & "')"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()
            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function hapusdtpbbr(ByVal constr As String, ByVal noppbr As String) As String
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "DELETE FROM DC_PPBR_DTL_T WHERE NO_FK_PPBR = '" & noppbr & "'"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            'cmd.Dispose()
            conection.Close()


            conection.Open()
            query = "DELETE FROM DC_PPBR_HDR_T WHERE NO_PPBR = '" & noppbr & "'"
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            ' cmd.Dispose()
            conection.Close()

            conection.Open()
            query = "DELETE FROM DC_OTP_LOG WHERE NO_DOC = '" & noppbr & "' AND TYPE_TRANS = 'PPBR'"
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            'cmd.Dispose()
            conection.Close()



            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function cekplu(ByVal PLU As String) As String
        Return "SELECT a.mbr_pluid as prdcd, a.mbr_full_nama as deskripsi, a.mbr_kemasan as satuan, CAST(b.mbr_acost as VARCHAR(100)) as harga, (c.stk_qty_sld_akhir+c.stk_out_qty_pick) as stok, b.mbr_tgl_plumati" &
               " FROM dc_barang_t a, dc_barang_dc_t b, dc_allstok_akhir_rusak_v c where" &
               " a.mbr_pluid = '" & PLU & "' and a.mbr_pluid = b.mbr_fk_pluid and a.mbr_pluid = c.stk_fk_pluid and b.mbr_tgl_plumati is null"
    End Function

    Public Function cekstokrusak(ByVal plu As String) As String
        Return "SELECT STK_QTY_SLD_AKHIR AS STOK3 FROM DC_ALLSTOK_AKHIR_RUSAK_V WHERE STK_FK_PLUID = '" & plu & "'"
    End Function

    Public Function ttlppbr(ByVal noppbr As String) As String
        Return "SELECT COUNT(PLUID) AS ITEM, SUM(QTY) AS QTY, TRUNC(SUM(NILAI),3) AS NILAI FROM DC_PPBR_DTL_T WHERE NO_FK_PPBR = '" & noppbr & "'"
    End Function

    Public Function cekpluppbr(ByVal mplu As String, ByVal noppbr As String) As String
        Return "SELECT NO_FK_PPBR, PLUID, SATUAN, QTY, CAST(HARGA AS VARCHAR(100)) AS HARGA, CAST(NILAI AS VARCHAR(100)) AS NILAI FROM DC_PPBR_DTL_T WHERE PLUID = '" & mplu & "' AND NO_FK_PPBR = '" & noppbr & "'"
    End Function

    Public Function updatedtlppbr(ByVal constr As String, ByVal plu As String, ByVal qty As String, harga As String, nilai As String, noppbr As String) As String
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "UPDATE DC_PPBR_DTL_T SET QTY = '" & qty & "', NILAI = (HARGA*QTY) WHERE NO_FK_PPBR = '" & noppbr & "' AND PLUID = '" & plu & "'"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()
            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function tambahdtlppbr(ByVal constr As String, ByVal plu As String, ByVal qty As String, ByVal harga As String, ByVal nilai As String, ByVal noppbr As String, ByVal satuan As String) As String
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "INSERT INTO DC_PPBR_DTL_T (NO_FK_PPBR, PLUID, SATUAN, QTY, HARGA, NILAI) VALUES ('" & noppbr & "', '" & plu & "', '" & satuan & "', '" & qty & "', '" & harga & "', '" & nilai & "')"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()
            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function updthdrppbr(ByVal constr As String, ByVal noppbr As String, ByVal totitem As String, ByVal totqty As String, ByVal totnilai As String, ByVal jenis As String, ByVal user As String) As String
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "UPDATE DC_PPBR_HDR_T SET JENIS_PPBR = '" & jenis & "', USER_PPBR = '" & user & "', TOT_ITEM = '" & totitem & "', TOT_QTY = '" & totqty & "', TOT_NILAI = '" & totnilai & "' WHERE NO_PPBR = '" & noppbr & "'"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()
            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function pbbrhelp(ByVal status As String) As String
        Return "SELECT NO_PPBR, TGL_PPBR, STATUS_PPBR, TOT_ITEM, TOT_QTY, TOT_NILAI FROM DC_PPBR_HDR_T " & status & "ORDER BY NO_PPBR"
    End Function
    Public Function baphelp()
        Return "SELECT a.hdr_no_doc AS no_ppbr, TRUNC(a.hdr_tgl_doc) AS tgl_ppbr, a.hdr_no_inv AS status_ppbr," &
               " b.tot_item, b.tot_qty, trunc(a.hdr_nilai) as tot_nilai FROM DC_HEADER_TRANSAKSI_T a," &
               " (SELECT his_hdr_fk_id, SUM(his_qty) AS tot_qty, COUNT(his_fk_plukode) AS tot_item FROM dc_history_transaksi_t WHERE his_type_trans = 'BAP DC' GROUP BY his_hdr_fk_id) b" &
               " WHERE a.hdr_type_trans = 'BAP DC' AND a.hdr_hdr_id = b.his_hdr_fk_id ORDER BY hdr_tgl_doc DESC"
    End Function

    Public Function pbbrfilter(ByVal noppbr As String)
        Return "SELECT NO_PPBR, TGL_PPBR, STATUS_PPBR FROM DC_PPBR_HDR_T WHERE NO_PPBR = '" & noppbr & "' ORDER BY NO_PPBR"
    End Function

    Public Function ambildtalamatemail()
        Return "Select * from dc_bap_email_t where TRIM(email_jabatan) = 'LOGISTIK MGR'"
    End Function
    Public Function ceknama(ByVal nama As String) As String
        Dim tblhtg As DataTable
        Try
            query = "SELECT * FROM DC_BAP_EMAIL_T WHERE TRIM(EMAIL_NAMA) = '" & nama & "'"
            tblhtg = getDS(query, strKoneksiORCL).Tables(0)
            If tblhtg.Rows.Count > 0 Then
                Return "Y"
            Else
                Return "X"
            End If
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function cekalamatemail(ByVal alamat As String) As String
        Dim tblalmt As DataTable
        Try
            query = "SELECT * FROM DC_BAP_EMAIL_T WHERE TRIM(EMAIL_ALAMAT) = '" & alamat & "'"
            tblalmt = getDS(query, strKoneksiORCL).Tables(0)
            If tblalmt.Rows.Count > 0 Then
                Return "Y"
            Else
                Return "X"
            End If
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function ceklm() As String
        Dim tbllm As DataTable
        Try
            query = "Select * from dc_bap_email_t where TRIM(email_jabatan) = 'LOGISTIK MGR'"
            tbllm = getDS(query, strKoneksiORCL).Tables(0)
            If tbllm.Rows.Count > 0 Then
                Return "Y"
            Else
                Return "X"
            End If
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function updatedtemaillm(ByVal constr As String, ByVal nama As String, ByVal alamat As String, ByVal jabatan As String) As String
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "UPDATE DC_BAP_EMAIL_T SET EMAIL_ALAMAT = '" & alamat & "', EMAIL_NAMA = '" & nama & "' WHERE  TRIM(EMAIL_JABATAN) = 'LOGISTIK MGR'"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()
            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function updatedtemail(ByVal constr As String, ByVal nama As String, ByVal alamat As String, ByVal jabatan As String) As String
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "UPDATE DC_BAP_EMAIL_T SET EMAIL_ALAMAT = '" & alamat & "', EMAIL_JABATAN = '" & jabatan & "' WHERE EMAIL_NAMA = '" & nama & "'"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()
            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function tambahdtemail(ByVal constr As String, ByVal nama As String, ByVal alamat As String, ByVal jabatan As String) As String
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "INSERT INTO DC_BAP_EMAIL_T (EMAIL_NAMA, EMAIL_ALAMAT, EMAIL_JABATAN) VALUES (' " & nama & "', '" & alamat & "', ' " & jabatan & "')"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()
            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function postingppbr(ByVal constr As String, noppbr As String) As String
        Dim ret As String = ""
        Dim connection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
        connection.Open()
        Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand("", connection)
        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "BAP.SEND_OTP"
            cmd.Parameters.Add(New Oracle.ManagedDataAccess.Client.OracleParameter("p_noppbr", Oracle.ManagedDataAccess.Client.OracleDbType.Varchar2, 20))
            cmd.Parameters.Add(New Oracle.ManagedDataAccess.Client.OracleParameter("p_nootp", Oracle.ManagedDataAccess.Client.OracleDbType.Varchar2, 200, ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
            cmd.Parameters.Add(New Oracle.ManagedDataAccess.Client.OracleParameter("p_msg", Oracle.ManagedDataAccess.Client.OracleDbType.Varchar2, 200, ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
            cmd.Parameters.Item(0).Value = noppbr
            cmd.ExecuteNonQuery()
            If cmd.Parameters.Item(2).Value.ToString = "null" Then
                ret = ""
            Else
                ret = cmd.Parameters.Item(2).Value.ToString
            End If

        Catch ex As Exception
            ret = ex.ToString
        Finally
            cmd.Dispose()
            connection.Close()
        End Try

        If ret = "" Then
            Return "Y"
        Else
            Return ret
        End If
    End Function

    Public Function ambildtotp(ByVal noppbr As String) As String
        Return "select * from dc_otp_log where no_doc = '" & noppbr & "' and Type_trans = 'PPBR'"
    End Function

    Public Function kirimulangotp(ByVal tgl As String) As String
        Return "SELECT * FROM DC_OTP_LOG" &
               " WHERE no_doc = (SELECT MIN(no_doc) FROM DC_OTP_LOG WHERE Type_trans = 'PPBR' AND TRUNC(Tgl_otp) = TO_DATE('" & tgl & "','mm/dd/yyyy'))" &
               " and type_trans = 'PPBR'"
    End Function

    Public Function loaddtbap(ByVal constr As String, noppbr As String) As String
        Dim ret As String = ""
        Dim connection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
        connection.Open()
        Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand("", connection)
        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "BAP.LOAD"
            cmd.Parameters.Add(New Oracle.ManagedDataAccess.Client.OracleParameter("p_noppbr", Oracle.ManagedDataAccess.Client.OracleDbType.Varchar2, 20))
            cmd.Parameters.Add(New Oracle.ManagedDataAccess.Client.OracleParameter("p_msg", Oracle.ManagedDataAccess.Client.OracleDbType.Varchar2, 200, ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
            cmd.Parameters.Item(0).Value = noppbr
            cmd.ExecuteNonQuery()
            If cmd.Parameters.Item(1).Value.ToString = "null" Then
                ret = ""
            Else
                ret = cmd.Parameters.Item(1).Value.ToString
            End If
            'ret = cmd.Parameters.Item(1).Value.ToString

        Catch ex As Exception
            ret = ex.ToString
        Finally
            cmd.Dispose()
            connection.Close()
        End Try

        If ret = "" Then
            Return "Y"
        Else
            Return ret
        End If
    End Function

    Public Function ambldtTRNbap(ByVal noppbr As String)
        Return "SELECT * FROM DC_TRNBAP_HDR_T WHERE BAP_NO_REF = '" & noppbr & "'"
    End Function
    Public Function createbapdev(ByVal constr As String, ByVal nobap As Integer, ByVal tglbap As Date) As Boolean
        Dim connection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
        connection.Open()
        Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand("", connection)
        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "CREATE_BAPDEVNOMOR_EVO"
            cmd.Parameters.Add(New Oracle.ManagedDataAccess.Client.OracleParameter("p_tgl", Oracle.ManagedDataAccess.Client.OracleDbType.Date))
            cmd.Parameters.Add(New Oracle.ManagedDataAccess.Client.OracleParameter("p_nomor", Oracle.ManagedDataAccess.Client.OracleDbType.Double))
            'cmd.Parameters.Add(New OracleClient.OracleParameter("p_msg", OracleClient.OracleType.VarChar, 200, ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
            cmd.Parameters.Item(0).Value = tglbap
            cmd.Parameters.Item(1).Value = nobap
            cmd.ExecuteNonQuery()
            'ret = cmd.Parameters.Item(1).Value.ToString
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function prosesdtbap(ByVal constr As String, bapid As String) As String
        Dim ret As String = ""
        Dim connection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
        connection.Open()
        Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand("", connection)
        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "BAP.TRANSAKSI"
            cmd.Parameters.Add(New Oracle.ManagedDataAccess.Client.OracleParameter("p_bap_hdrid", Oracle.ManagedDataAccess.Client.OracleDbType.Varchar2, 20))
            cmd.Parameters.Add(New Oracle.ManagedDataAccess.Client.OracleParameter("p_msg", Oracle.ManagedDataAccess.Client.OracleDbType.Varchar2, 200, ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
            cmd.Parameters.Item(0).Value = bapid
            cmd.ExecuteNonQuery()
            If cmd.Parameters.Item(1).Value.ToString = "null" Then
                ret = ""
            Else
                ret = cmd.Parameters.Item(1).Value.ToString
            End If
            ret = cmd.Parameters.Item(1).Value.ToString
        Catch ex As Exception
            ret = ex.ToString
        Finally
            cmd.Dispose()
            connection.Close()
        End Try

        If ret = "" Then
            Return "Y"
        Else
            Return ret
        End If
    End Function
    Public Function ambillapppbr(ByVal noppbr As String)
        Return "SELECT a.pluid AS plu, b.mbr_full_nama AS deskripsi, a.satuan, a.qty, trunc(a.harga,3) as harga, trunc(a.nilai,3) as nilai" &
               " FROM dc_ppbr_dtl_t a, dc_barang_t b WHERE a.no_fk_ppbr = '" & noppbr & "'" &
               " AND a.pluid = b.mbr_pluid ORDER BY a.pluid"
    End Function

    Public Function lapjudul(ByVal noppbr As String)
        Return "SELECT * FROM DC_PPBR_HDR_T WHERE NO_PPBR = '" & noppbr & "'"
    End Function

    Public Function ambiljnsppbr(ByVal noppbr As String
                                 )
        Dim tblaja As DataTable
        query = "SELECT * FROM DC_PPBR_HDR_T WHERE NO_PPBR = '" & noppbr & "'"
        tblaja = getDS(query, getStrKoneksiORCL).Tables(0)
        If tblaja.Rows.Count > 0 Then
            lapjnsppbr = tblaja.Rows(0)("JENIS_PPBR")
        Else
            lapjnsppbr = "KOSONG"
        End If

        Return lapjnsppbr
    End Function

    Public Function ambillapbap(ByVal nobap As String, ByVal tglbap As String)
        Dim tbleks As DataTable
        query = "SELECT HDR_NO_DOC, HDR_TGL_DOC, HDR_NO_INV FROM DC_HEADER_TRANSAKSI_T WHERE HDR_NO_INV = '" & nobap & "' AND HDR_TYPE_TRANS = 'BAP DC' AND TRUNC(HDR_TGL_DOC) = TO_DATE('" & tglbap & "','DD/MM/YYYY')"
        tbleks = getDS(query, strKoneksiORCL).Tables(0)
        If tbleks.Rows.Count > 0 Then
            lapnoppbr = tbleks.Rows(0)("HDR_NO_INV")
        Else
            lapnoppbr = "0"
        End If
        Return lapnoppbr
    End Function

    Public Function querylapbap(ByVal nobap As String, tglbap As String)
        Return "SELECT a.HIS_FK_PLUKODE AS PLU, b.mbr_full_nama AS deskripsi, b.mbr_kemasan AS satuan, a.his_qty AS qty, TRUNC(a.his_price,3) AS harga, TRUNC(a.his_nilai,3) AS nilai, c.hdr_hdr_id AS idtrans" &
               " FROM DC_HISTORY_TRANSAKSI_T a, DC_BARANG_T b, DC_HEADER_TRANSAKSI_T c WHERE A.HIS_TYPE_TRANS = 'BAP DC' AND c.hdr_no_doc = '" & nobap & "' and trunc(c.hdr_tgl_doc) = to_date('" & tglbap & "','DD/MM/YYYY') AND a.his_hdr_fk_id = c.hdr_hdr_id AND a.his_fk_pluid =  b.mbr_pluid"
    End Function

    Public Function queryrkpbapno(ByVal tglawl As String, ByVal tglakh As String)
        Return "SELECT HDR_NO_DOC AS NOBAP, TRUNC(TO_DATE(HDR_TGL_DOC,'DD/MON/YY')) AS TANGGAL, TRUNC(HDR_NILAI,3) AS NILAI FROM DC_HEADER_TRANSAKSI_T WHERE TRUNC(HDR_TGL_DOC) >= TO_DATE('" & tglawl & "', 'dd/mm/YYYY') AND TRUNC(HDR_TGL_DOC) <= TO_DATE('" & tglakh & "', 'dd/mm/YYYY') AND HDR_TYPE_TRANS = 'BAP DC' ORDER BY HDR_TGL_DOC"
    End Function

    Public Function queryrkpbapitem(ByVal tglawl As String, ByVal tglakh As String, ByVal urut As String)
        Return "SELECT b.his_fk_pluid AS PLU, c.mbr_full_nama AS DESKRIPSI, a.hdr_no_doc AS NOBAP, TRUNC(a.hdr_tgl_doc) AS TANGGAL," &
               " Get_Hrgbeli_Unit (Get_Dckode(a.hdr_fk_dcid),b.his_fk_pluid) AS SATUAN, b.his_qty AS QTY, TRUNC(b.his_nilai,3) AS NILAI, c.mbr_fk_div AS DIV," &
               " c.mbr_fk_dep AS DEPT, c.mbr_fk_KAT AS kat FROM DC_HEADER_TRANSAKSI_T a," &
               " DC_HISTORY_TRANSAKSI_T b, DC_BARANG_T c WHERE a.hdr_hdr_id = b.his_hdr_fk_id AND b.his_fk_pluid=c.mbr_pluid(+) AND a.hdr_type_trans = 'BAP DC'" &
               " AND a.hdr_tgl_doc >= TO_DATE('" & tglawl & "','dd/mm/yyyy') AND a.hdr_tgl_doc <= TO_DATE('" & tglakh & " 23:59:59','dd/mm/yyyy hh24:mi:ss')" &
               " ORDER BY " & urut & " ASC"
    End Function

    Public Function AMBILDTBAPPROS(ByVal NOPPBR As String)
        Return "SELECT NO_PPBR, NO_BAP, TRUNC(TGL_BAP) AS TGL_BAP FROM DC_PPBR_HDR_T WHERE NO_PPBR = '" & NOPPBR & "'"
    End Function
    Function KoneksiToTXT(ByVal sPath As String) As String
        Return "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" + sPath + ";"
        ' "Driver={Microsoft Text Driver (*.txt; *.csv)};DriverId=27;FIL=text;DefaultDir=" & _
        '                       sPath & ";DBQ=" & sPath
    End Function
    Public Function tambahitemppbr(ByVal constr As String, ByVal plu As String, ByVal qty As String) As String
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "INSERT INTO DC_PPBR_CSV(PLU, QTY) VALUES (' " & plu & "', '" & qty & "')"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()
            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function filterppbr(ByVal constr As String) As String
        Dim ret As String = ""
        Dim connection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
        connection.Open()
        Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand("", connection)
        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "BAP.FILTERPPBR"
            cmd.Parameters.Add(New Oracle.ManagedDataAccess.Client.OracleParameter("p_msg", Oracle.ManagedDataAccess.Client.OracleDbType.Varchar2, 200, ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            ret = ex.ToString
        Finally
            cmd.Dispose()
            connection.Close()
        End Try

        If ret = "" Then
            Return "Y"
        Else
            Return ret
        End If
    End Function
    Public Function cekuploadcsv(ByVal constr As String) As DataTable
        Dim ss As DataTable
        query = "SELECT * FROM DC_PPBR_CSV WHERE ERROR_MSG IS NULL"
        ss = getDS(query, constr).Tables(0)
        Return ss
    End Function
    Public Function ambildtuploadcsv(ByVal constr As String) As DataTable
        Dim ss As DataTable
        'query = "SELECT a.plu, a.qty, b.mbr_kemasan AS satuan, TRUNC(c.mbr_acost,3) AS harga, trunc((a.qty * c.mbr_acost),3) AS nilai  FROM DC_PPBR_CSV a, dc_barang_t b, dc_barang_dc_t c WHERE a.error_msg IS NULL AND a.plu = b.mbr_pluid AND a.plu = c.mbr_fk_pluid"
        query = "SELECT a.plu, a.qty, b.mbr_kemasan AS satuan, CAST(c.mbr_acost AS VARCHAR(100)) AS harga, TRUNC((a.qty * c.mbr_acost),3) AS nilai  FROM DC_PPBR_CSV a, dc_barang_t b, dc_barang_dc_t c WHERE a.error_msg IS NULL AND a.plu = b.mbr_pluid AND a.plu = c.mbr_fk_pluid"
        ss = getDS(query, constr).Tables(0)
        Return ss
    End Function
    Public Function ambillapupload(ByVal constr As String) As DataTable
        Dim ss As DataTable
        query = "SELECT a.PLU, b.MBR_FULL_NAMA AS DESKRIPSI, a.QTY, a.ERROR_MSG as KETERANGAN FROM DC_PPBR_CSV a, DC_BARANG_T b WHERE a.ERROR_MSG IS NOT NULL AND a.plu = b.MBR_PLUID ORDER BY PLU"
        ss = getDS(query, constr).Tables(0)
        Return ss
    End Function
    Public Function ambilnoppbr(ByVal constr As String) As DataTable
        Dim ss As DataTable
        query = "SELECT * from dc_ppbr_hdr_t where TGL_PPBR = TRUNC(SYSDATE) AND TOT_ITEM IS NULL"
        ss = getDS(query, constr).Tables(0)
        Return ss
    End Function
    Public Function ambilweb() As String
        Dim sa As DataTable
        Dim almt As String

        query = "SELECT * FROM DC_WEBSITES_T WHERE WEB_TYPE = 'PPBR'"
        sa = getDS(query, strKoneksiORCL).Tables(0)
        If sa.Rows.Count > 0 Then
            almt = sa.Rows(0)("WEB_URL")
        Else
            almt = "NULL"
        End If
        Return almt
    End Function

    Public Function hapusemail(ByVal constr As String, ByVal nama As String)
        Try
            Dim conection = New Oracle.ManagedDataAccess.Client.OracleConnection(constr)
            conection.Open()
            query = "DELETE FROM DC_BAP_EMAIL_T WHERE TRIM(EMAIL_NAMA) = '" & nama & "'"
            Dim cmd = New Oracle.ManagedDataAccess.Client.OracleCommand
            cmd.Connection = conection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            conection.Close()
            Return "Y"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function cekapprovedtgl() As String
        Return "SELECT * FROM DC_OTP_LOG" & _
               " WHERE Type_trans = 'PPBR' AND status = 'APPROVED'" & _
               " AND TRUNC(tgl_stat) = TRUNC(SYSDATE)"
    End Function
    Public ftperrmsg As String
    Public Function UploadFTPFiles(ftpAddress As String, ftpUser As String, ftpPassword As String, folderfile As String,
                               fileToUpload As String, targetFileName As String,
                               deleteAfterUpload As Boolean
                               ) As Boolean
        Dim str As String = ""
        Dim credential As NetworkCredential
        Try
            credential = New NetworkCredential(ftpUser, ftpPassword)
            If ftpAddress.EndsWith("/") = False Then ftpAddress = ftpAddress & "/"
            If folderfile.EndsWith("\") = False Then folderfile = folderfile & "\"
            Dim sFtpFile As String = ftpAddress & fileToUpload
            Dim request As FtpWebRequest = DirectCast(WebRequest.Create(sFtpFile), FtpWebRequest)
            request.KeepAlive = False
            request.Method = WebRequestMethods.Ftp.UploadFile
            request.Credentials = credential
            request.UsePassive = False
            request.Timeout = (60 * 1000) * 3 '3 mins
            Using reader As New FileStream(folderfile & fileToUpload, FileMode.Open)
                Dim buffer(Convert.ToInt32(reader.Length - 1)) As Byte
                reader.Read(buffer, 0, buffer.Length)
                reader.Close()
                request.ContentLength = buffer.Length
                Dim stream As Stream = request.GetRequestStream
                stream.Write(buffer, 0, buffer.Length)
                stream.Close()
                Using response As FtpWebResponse = DirectCast(request.GetResponse, FtpWebResponse)
                    If deleteAfterUpload Then
                        My.Computer.FileSystem.DeleteFile(fileToUpload)
                    End If
                    response.Close()
                End Using
            End Using
            Return True
        Catch ex As WebException
            str = ex.Message
            ftperrmsg = ex.Message
            'ExceptionInfo = ex
            Return False
        Finally
        End Try
    End Function
End Module
