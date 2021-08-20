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
        Me.components = New System.ComponentModel.Container()
        Me.cmdAnalyser = New System.Windows.Forms.Button()
        Me.cmdExporter = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'cmdAnalyser
        '
        Me.cmdAnalyser.Location = New System.Drawing.Point(24, 21)
        Me.cmdAnalyser.Name = "cmdAnalyser"
        Me.cmdAnalyser.Size = New System.Drawing.Size(93, 36)
        Me.cmdAnalyser.TabIndex = 0
        Me.cmdAnalyser.Text = "Analyser"
        Me.ToolTip1.SetToolTip(Me.cmdAnalyser, "Analyser le fichier des prénoms de l'INSEE et produire les rapports sur les préno" & _
        "ms mixtes et prénoms fréquents")
        Me.cmdAnalyser.UseVisualStyleBackColor = True
        '
        'cmdExporter
        '
        Me.cmdExporter.Location = New System.Drawing.Point(24, 76)
        Me.cmdExporter.Name = "cmdExporter"
        Me.cmdExporter.Size = New System.Drawing.Size(93, 36)
        Me.cmdExporter.TabIndex = 1
        Me.cmdExporter.Text = "Exporter"
        Me.ToolTip1.SetToolTip(Me.cmdExporter, "Exporter le fichier des prénoms de l'INSEE en y ajoutant les colonnes corresponda" & _
        "ntes à l'identification des prénoms mixtes")
        Me.cmdExporter.UseVisualStyleBackColor = True
        '
        'frmPrenomMixte
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(287, 178)
        Me.Controls.Add(Me.cmdExporter)
        Me.Controls.Add(Me.cmdAnalyser)
        Me.Name = "frmPrenomMixte"
        Me.Text = "Prénom mixte"
        Me.ResumeLayout(False)

End Sub
    Friend WithEvents cmdAnalyser As System.Windows.Forms.Button
    Friend WithEvents cmdExporter As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip

End Class
