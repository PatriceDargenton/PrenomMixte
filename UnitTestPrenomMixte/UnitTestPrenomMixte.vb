
Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UnitTestPrenomMixte

    <TestMethod()>
    Public Sub VerificationAccents()

        ' Vérifier si les traitements (correction des accents, puis suppression des accents)
        '  redonne bien exactement le fichier original
        ' S'il y a la moindre erreur lors de la correction des accents, ce test échouera

        Dim sStartupPath$ = AppDomain.CurrentDomain.BaseDirectory
        Dim sChemin1$ = IO.Path.GetDirectoryName(sStartupPath)
        Dim sChemin2$ = IO.Path.GetDirectoryName(sChemin1)
        Dim sChemin3$ = IO.Path.GetDirectoryName(sChemin2)
        Dim sDossierAppli$ = sChemin3 & "\bin"
        PrenomMixte.modPrenom.AnalyserPrenoms(sDossierAppli, bExporter:=True, bTest:=True)
        Dim sOriginal$ = PrenomMixte.modUtil.sLireFichier(sDossierAppli & "\" &
            PrenomMixte.sFichierPrenomsInsee)
        Dim sCopie$ = PrenomMixte.modUtil.sLireFichier(sDossierAppli & "\" &
            PrenomMixte.sFichierPrenomsInseeCorrige)

        ' Trop gros pour le mettre directement dans l'Assert : ça fait planter VS en cas d'échec !
        Dim bIdentique = (sOriginal = sCopie)
        Assert.AreEqual(bIdentique, True)

    End Sub

End Class