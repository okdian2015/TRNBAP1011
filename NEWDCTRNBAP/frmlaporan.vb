Imports System.Data
Imports System.IO
Imports System.Net
Imports Microsoft.Reporting.WinForms
Public Class frmlaporan

    Private Sub frmlaporan_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

    End Sub

    Private Sub frmlaporan_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If sttslap = "BAP" Then
            Dim dll As DataTable
            query = querylapbap(lapnobap, laptglbapitm1)
            dll = getDS(query, strKoneksiORCL).Tables(0)
            Dim tglajah As Date = Convert.ToDateTime(laptglbapitm1).ToString("dd/MM/yyyy")
            Dim tgl As String = tglajah.Day
            Dim bln As Integer = tglajah.Month
            Dim thn As String = tglajah.Year
            Dim nn As String = tglajah.DayOfWeek
            Dim hr As String = ""
            Dim bl As String = ""
            Dim rbl As String = ""
            If nn = "0" Then
                hr = "Minggu"
            ElseIf nn = "1" Then
                hr = "Senin"
            ElseIf nn = "2" Then
                hr = "Selasa"
            ElseIf nn = "3" Then
                hr = "Rabu"
            ElseIf nn = "4" Then
                hr = "Kamis"
            ElseIf nn = "5" Then
                hr = "Jum'at"
            ElseIf nn = "6" Then
                hr = "Sabtu"
            End If

            If bln = 1 Then
                bl = "Januari"
                rbl = "I/" & thn
            ElseIf bln = 2 Then
                bl = "Februari"
                rbl = "II/" & thn
            ElseIf bln = 3 Then
                bl = "Maret"
                rbl = "III/" & thn
            ElseIf bln = 4 Then
                bl = "April"
                rbl = "IV/" & thn
            ElseIf bln = 5 Then
                bl = "Mei"
                rbl = "V/" & thn
            ElseIf bln = 6 Then
                bl = "Juni"
                rbl = "VI/" & thn
            ElseIf bln = 7 Then
                bl = "Juli"
                rbl = "VII/" & thn
            ElseIf bln = 8 Then
                bl = "Agustus"
                rbl = "VIII/" & thn
            ElseIf bln = 9 Then
                bl = "September"
                rbl = "IX/" & thn
            ElseIf bln = 10 Then
                bl = "Oktober"
                rbl = "X/" & thn
            ElseIf bln = 11 Then
                bl = "November"
                rbl = "XI/" & thn
            ElseIf bln = 12 Then
                bl = "Desember"
                rbl = "XII/" & thn
            End If
            Dim jenis As String = ""
            If lapjnsppbr = "BARANG DAGANGAN RUSAK YANG TIDAK DAPAT DIMANFAATKAN" Then
                jenis = "Penyerahan fisik ke GA Cabang Idm"
            ElseIf lapjnsppbr = "BARANG DAGANGAN RUSAK YANG MASIH DAPAT DIMANFAATKAN" Then
                jenis = "Penyerahan fisik ke Development Cabang Idm"
            Else
                jenis = "Tidak Spesifik"
            End If

            If dll.Rows.Count > 0 Then
                ReportViewer1.Clear()
                Dim kdc As String = varKodeDC & "-" & varNamaDC
                Dim rpp1 As New ReportParameter("KODEDC", kdc)
                Dim rpp2 As New ReportParameter("NOBAP", lapnobap)
                Dim rpp3 As New ReportParameter("TGLBAP", rbl)
                Dim rpp4 As New ReportParameter("ALAMAT", varAlamatdc)
                Dim rpp5 As New ReportParameter("USER", username)
                Dim rpp6 As New ReportParameter("NPWP", varNpwp)
                Dim rpp7 As New ReportParameter("NOPPBR", lapnoppbr)
                Dim rpp8 As New ReportParameter("HARI", hr)
                Dim rpp9 As New ReportParameter("BULAN", bl)
                Dim rpp10 As New ReportParameter("TANGGAL", tgl)
                Dim rpp11 As New ReportParameter("TAHUN", thn)
                Dim rpp12 As New ReportParameter("JENIS", jenis)
                ReportViewer1.Reset()
                ReportViewer1.LocalReport.ReportEmbeddedResource = "NEWDCTRNBAP.RPTBAP.rdlc"
                ReportViewer1.ProcessingMode = ProcessingMode.Local
                ReportViewer1.LocalReport.DataSources.Clear()
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp1})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp2})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp3})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp4})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp5})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp6})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp7})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp8})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp9})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp10})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp11})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp12})
                Dim rds As New ReportDataSource
                rds.Name = "dtbap"
                rds.Value = dll
                ReportViewer1.LocalReport.DataSources.Add(rds)
                ReportViewer1.LocalReport.Refresh()
                ReportViewer1.Refresh()
            End If
        ElseIf sttslap = "PPBR" Then
            Dim dll As DataTable
            query = ambillapppbr(lapnoppbr)
            dll = getDS(query, strKoneksiORCL).Tables(0)
            If dll.Rows.Count > 0 Then
                ReportViewer1.Clear()
                Dim kdc As String = varKodeDC & "-" & varNamaDC
                Dim rpp1 As New ReportParameter("KODEDC", kdc)
                Dim rpp2 As New ReportParameter("JENISPPBR", lapjnsppbr)
                Dim rpp3 As New ReportParameter("NOPPBR", lapnoppbr)
                Dim rpp4 As New ReportParameter("TGLPPBR", laptglppbr)
                Dim rpp5 As New ReportParameter("USER", username)
                ReportViewer1.Reset()
                ReportViewer1.LocalReport.ReportEmbeddedResource = "NEWDCTRNBAP.RPTPPBR.rdlc"
                ReportViewer1.ProcessingMode = ProcessingMode.Local
                ReportViewer1.LocalReport.DataSources.Clear()
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp1})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp2})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp3})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp4})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp5})

                Dim rds As New ReportDataSource
                rds.Name = "dtlapppbr"
                rds.Value = dll
                ReportViewer1.LocalReport.DataSources.Add(rds)
                ReportViewer1.LocalReport.Refresh()
                ReportViewer1.Refresh()
            Else
                MessageBox.Show("Tidak ada data untuk ditampilkan", "Kosong", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        ElseIf sttslap = "REPRKPNOMOR" Then
            Dim dll As DataTable
            query = queryrkpbapno(laptglbapnmr1, laptglbapnmr2)
            dll = getDS(query, strKoneksiORCL).Tables(0)
            If dll.Rows.Count > 0 Then
                ReportViewer1.Clear()
                Dim kdc As String = varKodeDC & "-" & varNamaDC
                Dim periode As String = laptglbapnmr1 & " s/d " & laptglbapnmr2
                Dim rpp1 As New ReportParameter("KODEDC", kdc)
                Dim rpp2 As New ReportParameter("PERIODE", periode)
                Dim rpp3 As New ReportParameter("USER", username)
                ReportViewer1.Reset()
                ReportViewer1.LocalReport.ReportEmbeddedResource = "NEWDCTRNBAP.Rptrkpbap.rdlc"
                ReportViewer1.ProcessingMode = ProcessingMode.Local
                ReportViewer1.LocalReport.DataSources.Clear()
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp1})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp2})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp3})
                Dim rds As New ReportDataSource
                rds.Name = "dtrkpbapnomor"
                rds.Value = dll
                ReportViewer1.LocalReport.DataSources.Add(rds)
                ReportViewer1.LocalReport.Refresh()
                ReportViewer1.Refresh()
            Else
                MessageBox.Show("Tidak ada data untuk ditampilkan", "Kosong", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        ElseIf sttslap = "REPRKPITEM" Then
            Dim dll As DataTable
            query = queryrkpbapitem(laptglbapitm1, laptglbapitm2, sitmbap)
            dll = getDS(query, strKoneksiORCL).Tables(0)
            If dll.Rows.Count > 0 Then
                ReportViewer1.Clear()
                Dim kdc As String = varKodeDC & "-" & varNamaDC
                Dim periode As String = laptglbapitm1 & " s/d " & laptglbapitm2
                Dim rpp1 As New ReportParameter("KODEDC", kdc)
                Dim rpp2 As New ReportParameter("PERIODE", periode)
                Dim rpp3 As New ReportParameter("USER", username)
                Dim rpp4 As New ReportParameter("URUT", sitmbap)
                ReportViewer1.Reset()
                ReportViewer1.LocalReport.ReportEmbeddedResource = "NEWDCTRNBAP.Rptrkpbapitem.rdlc"
                ReportViewer1.ProcessingMode = ProcessingMode.Local
                ReportViewer1.LocalReport.DataSources.Clear()
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp1})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp2})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp3})
                ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp4})
                Dim rds As New ReportDataSource
                rds.Name = "DataSet1"
                rds.Value = dll
                ReportViewer1.LocalReport.DataSources.Add(rds)
                ReportViewer1.LocalReport.Refresh()
                ReportViewer1.Refresh()
            Else
                MessageBox.Show("Tidak ada data untuk ditampilkan", "Kosong", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        ElseIf sttslap = "UPLOADCSV" Then
            ReportViewer1.Clear()
            Dim tbllap As DataTable = ambillapupload(strKoneksiORCL)
            Dim kdc As String = varKodeDC & "-" & varNamaDC
            Dim rpp1 As New ReportParameter("KODEDC", kdc)
            'Dim rpp2 As New ReportParameter("USER", username)
            ReportViewer1.Reset()
            ReportViewer1.LocalReport.ReportEmbeddedResource = "NEWDCTRNBAP.lapcsv.rdlc"
            ReportViewer1.ProcessingMode = ProcessingMode.Local
            ReportViewer1.LocalReport.DataSources.Clear()
            ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp1})
            'ReportViewer1.LocalReport.SetParameters(New ReportParameter() {rpp2})
            Dim rds As New ReportDataSource
            rds.Name = "lapcsv"
            rds.Value = tbllap
            ReportViewer1.LocalReport.DataSources.Add(rds)
            ReportViewer1.LocalReport.Refresh()
            ReportViewer1.Refresh()
        End If
        Me.ReportViewer1.RefreshReport()
    End Sub
End Class