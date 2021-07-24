<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrenomMixte
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
        Me.cmdAnalyser = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmdAnalyser
        '
        Me.cmdAnalyser.Location = New System.Drawing.Point(31, 25)
        Me.cmdAnalyser.Name = "cmdAnalyser"
        Me.cmdAnalyser.Size = New System.Drawing.Size(76, 39)
        Me.cmdAnalyser.TabIndex = 0
        Me.cmdAnalyser.Text = "Analyser"
        Me.cmdAnalyser.UseVisualStyleBackColor = True
        '
        'frmPrenomMixte
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(583, 332)
        Me.Controls.Add(Me.cmdAnalyser)
        Me.Name = "frmPrenomMixte"
        Me.Text = "Prénom mixte"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cmdAnalyser As Button
End Class
