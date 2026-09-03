<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMath
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
        components = New System.ComponentModel.Container
        Me.wbMath = New Microsoft.Web.WebView2.WinForms.WebView2()
        CType(Me.wbMath, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1440, 900)
        Me.Controls.Add(Me.wbMath)
        Me.MinimumSize = New System.Drawing.Size(1100, 700)
        Me.Name = "frmMath"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Math Canvas"
        Me.wbMath.AllowExternalDrop = False
        Me.wbMath.CreationProperties = Nothing
        Me.wbMath.DefaultBackgroundColor = System.Drawing.Color.FromArgb(CType(CType(30, Byte), Integer), CType(CType(30, Byte), Integer), CType(CType(46, Byte), Integer))
        Me.wbMath.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbMath.Location = New System.Drawing.Point(0, 0)
        Me.wbMath.Name = "wbMath"
        Me.wbMath.Size = New System.Drawing.Size(1440, 900)
        Me.wbMath.TabIndex = 0
        Me.wbMath.ZoomFactor = 1.0R
        CType(Me.wbMath, System.ComponentModel.ISupportInitialize).EndInit()
    End Sub

    Friend WithEvents wbMath As Microsoft.Web.WebView2.WinForms.WebView2
End Class
