Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.DataSet
Imports System.Data.OleDb
Imports System.IO
Imports System.Threading
Imports System.Windows.Forms
Imports System.Reflection
Imports SD3Fungsi
Imports System.Net
Imports Oracle.ManagedDataAccess.Client
Public Class frmlogin
    Private mHostEntry As System.Net.IPHostEntry
    Private mDate As Date
    Private mIpEntry As System.Net.IPAddress
    Dim m_ds_login As DataSet
    Dim versionNumber As Version
    Dim progName As String
    Dim ceklos As String
    Private Sub frmlogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim strIPAddress As String
        Dim strHostName As String
        Try

            versionNumber = Assembly.GetExecutingAssembly().GetName().Version
            progName = Assembly.GetExecutingAssembly().GetName().Name
            Try
                mDate = Date.Now

                'Dim kodeDC, Constr As String

                'Constr = StringKoneksiOracle()
                'kodeDC = getKodeDC()
                strKoneksiORCL = getStrKoneksiORCL()
                serverDB = getServerDB()
                varKodeDC = getKodeDC()
                Cek_Program_2(varKodeDC)
                'mHostEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName)
                'ip_address = mHostEntry.AddressList.GetValue(0).ToString
                strHostName = System.Net.Dns.GetHostName()
                strIPAddress = System.Net.Dns.GetHostEntry(strHostName).AddressList(1).ToString()
                Label1.Text = " IP : " & strIPAddress
                Me.txtuser.Focus()
                Me.Text = "[ DCTRNBAP v " & versionNumber.ToString & "] [Tns: " & serverDB & "]"
                'Me.Text &= String.Concat(New String() {" Versi ", Application.ProductVersion, " | TNS : ", serverDB, ""})
                mHostEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName)
                'If Not Cek_Program(varKodeDC) Then
                '    Application.Exit()
                'End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Me.txtuser.Focus()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Me.txtuser.Focus()
    End Sub

    Private Sub ambluser()
        query = ValidateLogin()
        Try
            tblusr = getDS(query, strKoneksiORCL).Tables(0)
            username = tblusr.Rows(0)("user_name")
            nik = tblusr.Rows(0)("user_nik")
            m_dc_id = tblusr.Rows(0)("user_fk_tbl_dcid")
            ceklos = "Y"
        Catch aa As Exception
            ceklos = "N"
        End Try


    End Sub
    Private Sub ButtonX1_Click(sender As Object, e As EventArgs) Handles ButtonX1.Click
        username = txtuser.Text
        password = txtpassword.Text
        If username = "" And password = "" Then
            MessageBox.Show("Data login belum lengkap !!", "Erorr")
            txtuser.Focus()
        Else
            If cekUser(strKoneksiORCL, username, password) Then
                ambluser()
                If ceklos = "Y" Then
                    Me.Hide()
                    Try
                        frmtrnbap.Show()
                    Catch eX As Exception
                        MessageBox.Show(eX.Message)
                    End Try

                Else
                    MessageBox.Show("User tidak terdaftar di lokasi 01 !!", "Erorr")
                    txtuser.Focus()
                End If
            End If
        End If



    End Sub

    Private Sub ButtonX2_Click(sender As Object, e As EventArgs) Handles ButtonX2.Click
        Me.Close()
    End Sub
    Private Sub LOGIN_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Application.Exit()
    End Sub

    Private Sub usernamepress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtuser.KeyPress
        e.KeyChar = UCase(e.KeyChar)
    End Sub
    Private Sub usernamedown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtuser.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.txtpassword.Focus()
        End If
    End Sub
    Private Sub Passwordtextbox_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtpassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call ButtonX1_Click(sender, e)
        End If
    End Sub
End Class