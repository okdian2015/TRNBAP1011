Imports System.Data.OracleClient
Public Class frmhelp

    Private Sub frmhelp_load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        strKoneksiORCL = getStrKoneksiORCL()
        serverDB = getServerDB()
        userDB = getUserDB()
        passDB = getPassDB()
        If hstatus = "PPBR" Then
            Label1.Text = "Masukkan no PPBR"
            DGPPBR.Columns(0).HeaderText = "NO PPBR"
            DGPPBR.Columns(1).HeaderText = "TGL_PPBR"
            DGPPBR.Columns(2).HeaderText = "STATUS"
            query = pbbrhelp(nstatus)
            tblppbr = getTB(query, strKoneksiORCL).Tables(0)
            Dim con = New Oracle.ManagedDataAccess.Client.OracleConnection(strKoneksiORCL)
            Dim ds As New DataSet
            daAddFunctionDGVSumUpl = New Oracle.ManagedDataAccess.Client.OracleDataAdapter(query, con)
            daAddFunctionDGVSumUpl.Fill(ds, "SumUpl")
            DGPPBR.DataSource = ds.Tables("SumUpl")
            TextBoxX1.Text = ""
            TextBoxX2.Text = tblppbr.Rows(0)("TOT_ITEM")
            TextBoxX3.Text = tblppbr.Rows(0)("TOT_QTY")
            TextBoxX4.Text = tblppbr.Rows(0)("TOT_NILAI")
        ElseIf hstatus = "LAPPPBR" Then
            Label1.Text = "Masukkan no PPBR"
            DGPPBR.Columns(0).HeaderText = "NO PPBR"
            DGPPBR.Columns(1).HeaderText = "TGL_PPBR"
            DGPPBR.Columns(2).HeaderText = "STATUS"
            query = pbbrhelp(nstatus)
            tblppbr = getTB(query, strKoneksiORCL).Tables(0)
            Dim con = New Oracle.ManagedDataAccess.Client.OracleConnection(strKoneksiORCL)
            Dim ds As New DataSet
            daAddFunctionDGVSumUpl = New Oracle.ManagedDataAccess.Client.OracleDataAdapter(query, con)
            daAddFunctionDGVSumUpl.Fill(ds, "SumUpl")
            DGPPBR.DataSource = ds.Tables("SumUpl")
            TextBoxX1.Text = ""
            TextBoxX2.Text = tblppbr.Rows(0)("TOT_ITEM")
            TextBoxX3.Text = tblppbr.Rows(0)("TOT_QTY")
            TextBoxX4.Text = tblppbr.Rows(0)("TOT_NILAI")
        ElseIf hstatus = "LAPBAP" Then
            Label1.Text = "Masukkan no BAP"
            DGPPBR.Columns(0).HeaderText = "NO BAP"
            DGPPBR.Columns(1).HeaderText = "TGL BAP"
            DGPPBR.Columns(2).HeaderText = "NO REFF"
            query = baphelp()
            tblppbr = getTB(query, strKoneksiORCL).Tables(0)
            Dim con = New Oracle.ManagedDataAccess.Client.OracleConnection(strKoneksiORCL)
            Dim ds As New DataSet
            daAddFunctionDGVSumUpl = New Oracle.ManagedDataAccess.Client.OracleDataAdapter(query, con)
            daAddFunctionDGVSumUpl.Fill(ds, "SumUpl")
            DGPPBR.DataSource = ds.Tables("SumUpl")
            TextBoxX1.Text = ""
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        crnoppbr = TextBoxX1.Text
        If crnoppbr = "" Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
            frmtrnbap.Enabled = True
            frmtrnbap.Focus()
        Else
            nocr = crnoppbr
            ' tglcr =
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
        frmtrnbap.Enabled = True
        frmtrnbap.Focus()
    End Sub

    Private Sub ButtonX3_Click(sender As Object, e As EventArgs) Handles ButtonX3.Click
        If Len(TextBoxX1.Text) > 0 Then
            Dim no_ppbr As String = TextBoxX1.Text
            query = pbbrfilter(no_ppbr)
            tblppbr = getTB(query, strKoneksiORCL).Tables(0)
            If tblppbr.Rows.Count > 0 Then
                Dim con = New Oracle.ManagedDataAccess.Client.OracleConnection(strKoneksiORCL)
                Dim ds As New DataSet
                daAddFunctionDGVSumUpl = New Oracle.ManagedDataAccess.Client.OracleDataAdapter(query, con)
                daAddFunctionDGVSumUpl.Fill(ds, "SumUpl")
                DGPPBR.DataSource = ds.Tables("SumUpl")
                TextBoxX2.Text = tblppbr.Rows(0)("TOT_ITEM")
                TextBoxX3.Text = tblppbr.Rows(0)("TOT_QTY")
                TextBoxX4.Text = tblppbr.Rows(0)("TOT_NILAI")
                tglcr = tblppbr.Rows(0)("TGLPPBR")
            Else
                MessageBox.Show("No Transaksi yang anda masukkan salah", "Portal", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBoxX1.Text = ""
                TextBoxX1.Focus()
            End If
        End If
    End Sub

    Private Sub DGPPBR_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DGPPBR.CellMouseClick
        Dim clmName As String = DGPPBR.Columns(DGPPBR.CurrentCell.ColumnIndex).Name
        TextBoxX1.Text = DGPPBR.Rows(e.RowIndex).Cells("NOPPBR").Value
        tglcr = DGPPBR.Rows(e.RowIndex).Cells("TGLPPBR").Value
        TextBoxX2.Text = IIf(IsDBNull(DGPPBR.Rows(e.RowIndex).Cells("ITEM").Value), 0, DGPPBR.Rows(e.RowIndex).Cells("ITEM").Value)
        TextBoxX3.Text = IIf(IsDBNull(DGPPBR.Rows(e.RowIndex).Cells("QTY").Value), 0, DGPPBR.Rows(e.RowIndex).Cells("QTY").Value)
        TextBoxX4.Text = FormatNumber(IIf(IsDBNull(DGPPBR.Rows(e.RowIndex).Cells("NILAI").Value), 0, DGPPBR.Rows(e.RowIndex).Cells("NILAI").Value), 0)

        'DGPPBR.Columns(0).HeaderText = "Test ajah"
    End Sub
End Class