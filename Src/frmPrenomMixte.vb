
' frmPrenomMixte.vb
' -----------------

Public Class frmPrenomMixte

    Private Sub frmPrenomMixte_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Dim sVersion$ = " - V" & sVersionAppli & " (" & sDateVersionAppli & ")"
        Dim sDebug$ = " - Debug"
        Dim sTxt$ = Me.Text & sVersion
        If bDebug Then sTxt &= sDebug
        Me.Text = sTxt

    End Sub

    Private Sub cmdAnalyser_Click(sender As Object, e As EventArgs) Handles cmdAnalyser.Click

        Me.cmdAnalyser.Enabled = False
        Me.cmdExporter.Enabled = False
        AnalyserPrenoms(Application.StartupPath)
        Me.cmdAnalyser.Enabled = True
        Me.cmdExporter.Enabled = True

    End Sub

    Private Sub cmdExporter_Click(sender As Object, e As EventArgs) Handles cmdExporter.Click

        Me.cmdAnalyser.Enabled = False
        Me.cmdExporter.Enabled = False
        AnalyserPrenoms(Application.StartupPath, bExporter:=True)
        Me.cmdAnalyser.Enabled = True
        Me.cmdExporter.Enabled = True

    End Sub

End Class