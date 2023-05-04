<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmlaporan
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim ReportDataSource1 As Microsoft.Reporting.WinForms.ReportDataSource = New Microsoft.Reporting.WinForms.ReportDataSource()
        Me.ReportViewer1 = New Microsoft.Reporting.WinForms.ReportViewer()
        Me.Dsbap = New NEWDCTRNBAP.Dsbap()
        Me.dtlapppbrBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.dtrkpbapnomorBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.dtrkpbapitemBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.Dsbap, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtlapppbrBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtrkpbapnomorBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtrkpbapitemBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ReportViewer1
        '
        ReportDataSource1.Name = "DataSet1"
        ReportDataSource1.Value = Me.dtlapppbrBindingSource
        Me.ReportViewer1.LocalReport.DataSources.Add(ReportDataSource1)
        Me.ReportViewer1.LocalReport.ReportEmbeddedResource = "NEWDCTRNBAP.RPTBAP.rdlc"
        Me.ReportViewer1.Location = New System.Drawing.Point(2, 0)
        Me.ReportViewer1.Name = "ReportViewer1"
        Me.ReportViewer1.Size = New System.Drawing.Size(780, 479)
        Me.ReportViewer1.TabIndex = 0
        '
        'Dsbap
        '
        Me.Dsbap.DataSetName = "Dsbap"
        Me.Dsbap.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'dtlapppbrBindingSource
        '
        Me.dtlapppbrBindingSource.DataMember = "dtlapppbr"
        Me.dtlapppbrBindingSource.DataSource = Me.Dsbap
        '
        'dtrkpbapnomorBindingSource
        '
        Me.dtrkpbapnomorBindingSource.DataMember = "dtrkpbapnomor"
        Me.dtrkpbapnomorBindingSource.DataSource = Me.Dsbap
        '
        'dtrkpbapitemBindingSource
        '
        Me.dtrkpbapitemBindingSource.DataMember = "dtrkpbapitem"
        Me.dtrkpbapitemBindingSource.DataSource = Me.Dsbap
        '
        'frmlaporan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(782, 481)
        Me.Controls.Add(Me.ReportViewer1)
        Me.Name = "frmlaporan"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Laporan Transaksi BAP"
        CType(Me.Dsbap, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtlapppbrBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtrkpbapnomorBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtrkpbapitemBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ReportViewer1 As Microsoft.Reporting.WinForms.ReportViewer
    Friend WithEvents dtlapppbrBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Dsbap As NEWDCTRNBAP.Dsbap
    Friend WithEvents dtrkpbapnomorBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents dtrkpbapitemBindingSource As System.Windows.Forms.BindingSource
End Class
