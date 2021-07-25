<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrenomMixte
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cmdAnalyser = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmdAnalyser
        '
        Me.cmdAnalyser.Location = New System.Drawing.Point(24, 21)
        Me.cmdAnalyser.Name = "cmdAnalyser"
        Me.cmdAnalyser.Size = New System.Drawing.Size(93, 36)
        Me.cmdAnalyser.TabIndex = 0
        Me.cmdAnalyser.Text = "Analyser"
        Me.cmdAnalyser.UseVisualStyleBackColor = True
        '
        'frmPrenomMixte
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(287, 178)
        Me.Controls.Add(Me.cmdAnalyser)
        Me.Name = "frmPrenomMixte"
        Me.Text = "Prénom mixte"
        Me.ResumeLayout(False)

End Sub
    Friend WithEvents cmdAnalyser As System.Windows.Forms.Button

End Class
