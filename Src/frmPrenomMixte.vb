
Public Class frmPrenomMixte

    Private Sub cmdAnalyser_Click(sender As Object, e As EventArgs) Handles cmdAnalyser.Click

        Me.cmdAnalyser.Enabled = False
        Me.cmdExporter.Enabled = False
        AnalyserPrenoms()
        Me.cmdAnalyser.Enabled = True
        Me.cmdExporter.Enabled = True

    End Sub

    Private Sub cmdExporter_Click(sender As Object, e As EventArgs) Handles cmdExporter.Click

        Me.cmdAnalyser.Enabled = False
        Me.cmdExporter.Enabled = False
        AnalyserPrenoms(bExporter:=True)
        Me.cmdAnalyser.Enabled = True
        Me.cmdExporter.Enabled = True

    End Sub

End Class