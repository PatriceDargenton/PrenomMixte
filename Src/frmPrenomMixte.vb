
Imports System.Text

Public Class frmPrenomMixte

    Private Sub cmdAnalyser_Click(sender As Object, e As EventArgs) Handles cmdAnalyser.Click

        TestPrenoms()

    End Sub

    Public Class clsPrenom
        Public sPrenom$
        Public bMasc As Boolean
        Public bFem As Boolean
        Public bMixte As Boolean
        Public iNbOccMasc%, iNbOccFem%, iNbOcc%
        Public rFreqRelative#, rFreqRelativeMasc#, rFreqRelativeFem#
        Public rFreqTotale#, rFreqTotaleMasc#, rFreqTotaleFem#
        Public iAnnee%, rAnneeMoy#, rAnneeMoyMasc#, rAnneeMoyFem#
    End Class

    Private Sub TestPrenoms()

        Dim sChemin = Application.StartupPath & "\nat2019.csv"
        If Not IO.File.Exists(sChemin) Then
            MsgBox("Veuillez télécharger nat2019_csv.zip !" & vbLf & sChemin,
                MsgBoxStyle.Exclamation, "Prénom mixte")
            Exit Sub
        End If
        Dim asLignes$() = IO.File.ReadAllLines(sChemin, Encoding.UTF8)
        If IsNothing(asLignes) Then Exit Sub

        Dim dico As New DicoTri(Of String, clsPrenom)
        Dim iNbLignes% = 0
        Dim iNbLignesOk% = 0
        Dim iNbPrenomsTot% = 0
        Dim iNbPrenomsTotOk% = 0
        Dim iNbPrenomsIgnores% = 0
        Dim iNbPrenomsIgnoresDate% = 0

        For Each sLigne As String In asLignes

            iNbLignes += 1
            If iNbLignes = 1 Then Continue For ' Entête
            Dim asChamps() As String
            asChamps = Split(sLigne, ";"c)
            Dim iNumChampMax% = asChamps.GetUpperBound(0)
            Dim iNumChamp% = 0
            Dim bOk As Boolean = True
            Dim sCodeSexe$ = ""
            Dim sPrenomOrig$ = ""
            Dim sAnnee$ = ""
            Dim sOcc$ = ""
            For Each sChamp As String In asChamps
                iNumChamp += 1
                If IsNothing(sChamp) Then sChamp = ""
                If sChamp.Length = 0 Then bOk = False : Exit For
                Select Case iNumChamp
                    Case 1 : sCodeSexe = sChamp
                    Case 2 : sPrenomOrig = sChamp
                    Case 3 : sAnnee = sChamp
                    Case 4 : sOcc = sChamp
                End Select
            Next
            If Not bOk Then Continue For

            Dim sPrenom = sPrenomOrig.ToLower

            If sPrenom = "adelaide" Then sPrenom = "adélaide"
            If sPrenom = "aloise" Then sPrenom = "aloïse"
            If sPrenom = "amedee" Then sPrenom = "amedée"
            If sPrenom = "anael" Then sPrenom = "anaël"
            If sPrenom = "andrea" Then sPrenom = "andréa"
            If sPrenom = "arsene" Then sPrenom = "arsène"
            If sPrenom = "barthelemy" Then sPrenom = "barthélemy"
            If sPrenom = "celeste" Then sPrenom = "céleste"
            If sPrenom = "cleo" Then sPrenom = "cléo"
            If sPrenom = "come" Then sPrenom = "côme"
            If sPrenom = "dorothee" Then sPrenom = "dorothée"
            If sPrenom = "eden" Then sPrenom = "éden"
            If sPrenom = "elia" Then sPrenom = "élia"
            If sPrenom = "elie" Then sPrenom = "élie"
            If sPrenom = "elisee" Then sPrenom = "elisée"
            If sPrenom = "esperance" Then sPrenom = "espérance"
            If sPrenom = "evariste" Then sPrenom = "évariste"
            If sPrenom = "felicite" Then sPrenom = "félicité"
            If sPrenom = "gael" Then sPrenom = "gaël"
            If sPrenom = "gwenael" Then sPrenom = "gwenaël"
            If sPrenom = "heidi" Then sPrenom = "heïdi"
            If sPrenom = "irenee" Then sPrenom = "irenée"
            If sPrenom = "judicael" Then sPrenom = "judicaël"
            If sPrenom = "leocadie" Then sPrenom = "léocadie"
            If sPrenom = "leonard" Then sPrenom = "léonard"
            If sPrenom = "leonce" Then sPrenom = "léonce"
            If sPrenom = "mae" Then sPrenom = "maé"
            If sPrenom = "mael" Then sPrenom = "maël"
            If sPrenom = "mederic" Then sPrenom = "médéric"
            If sPrenom = "medine" Then sPrenom = "médine"
            If sPrenom = "meryl" Then sPrenom = "méryl"
            If sPrenom = "nael" Then sPrenom = "naël"
            If sPrenom = "sylvere" Then sPrenom = "sylvère"
            If sPrenom = "thais" Then sPrenom = "thaïs"
            If sPrenom = "theodore" Then sPrenom = "théodore"
            If sPrenom = "valere" Then sPrenom = "valère"
            If sPrenom = "valery" Then sPrenom = "valéry"
            If sPrenom = "yael" Then sPrenom = "yaël"

            Dim prenom As New clsPrenom
            prenom.sPrenom = FirstCharToUpper(sPrenom)
            If sCodeSexe = "1" Then prenom.bMasc = True
            If sCodeSexe = "2" Then prenom.bFem = True
            Dim iNbOccN%? = iConvN(sOcc)
            If IsNothing(iNbOccN) Then Continue For ' Tous ces nombres sont bien formés
            Dim iNbOcc% = CInt(iNbOccN)

            If sPrenomOrig = "_PRENOMS_RARES" Then
                iNbPrenomsIgnores += iNbOcc
                iNbPrenomsTot += iNbOcc
                Continue For
            End If
            If sAnnee = "XXXX" Then
                iNbPrenomsIgnoresDate += iNbOcc
                iNbPrenomsTot += iNbOcc
                Continue For
            End If
            iNbLignesOk += 1

            If prenom.bMasc Then prenom.iNbOccMasc = iNbOcc
            If prenom.bFem Then prenom.iNbOccFem = iNbOcc
            prenom.iNbOcc = iNbOcc
            iNbPrenomsTotOk += iNbOcc
            iNbPrenomsTot += iNbOcc

            Dim iAnneeN%? = iConvN(sAnnee)
            If IsNothing(iAnneeN) Then Continue For ' Il ne reste plus de date invalide
            prenom.iAnnee = CInt(iAnneeN)
            prenom.rAnneeMoy = prenom.iAnnee * iNbOcc
            If prenom.bMasc Then prenom.rAnneeMoyMasc = prenom.iAnnee * iNbOcc
            If prenom.bFem Then prenom.rAnneeMoyFem = prenom.iAnnee * iNbOcc

            Dim sCle$ = prenom.sPrenom
            If dico.ContainsKey(sCle) Then
                Dim prenom0 = dico(sCle)
                If prenom0.bMasc AndAlso prenom.bFem Then prenom0.bFem = True
                If prenom0.bFem AndAlso prenom.bMasc Then prenom0.bMasc = True
                prenom0.iNbOccFem += prenom.iNbOccFem
                prenom0.iNbOccMasc += prenom.iNbOccMasc
                prenom0.iNbOcc += prenom.iNbOcc
                prenom0.rAnneeMoy += prenom.rAnneeMoy
                prenom0.rAnneeMoyMasc += prenom.rAnneeMoyMasc
                prenom0.rAnneeMoyFem += prenom.rAnneeMoyFem
            Else
                dico.Add(sCle, prenom)
            End If

        Next

        Const rSeuilFreqRel# = 0.01 ' 1% (par exemple 1% de masc. et 99% de fém.)
        'Const rSeuilFreqRel# = 0.02 ' 2% (par exemple 2% de masc. et 98% de fém.)
        Const iSeuilMin% = 2000 '10000 ' Nombre minimal d'occurrences du prénom sur plus d'un siècle
        Dim iNbPrenomsVerif% = 0
        Dim iNbPrenomsVerifMF% = 0
        For Each prenom In dico.Trier("")
            prenom.rFreqTotale = prenom.iNbOcc / iNbPrenomsTot
            prenom.rFreqTotaleMasc = prenom.iNbOccMasc / iNbPrenomsTot
            prenom.rFreqTotaleFem = prenom.iNbOccFem / iNbPrenomsTot
            prenom.rAnneeMoy = prenom.rAnneeMoy / prenom.iNbOcc
            prenom.rAnneeMoyMasc = prenom.rAnneeMoyMasc / prenom.iNbOccMasc
            prenom.rAnneeMoyFem = prenom.rAnneeMoyFem / prenom.iNbOccFem
            prenom.rFreqRelativeMasc = prenom.iNbOccMasc / prenom.iNbOcc
            prenom.rFreqRelativeFem = prenom.iNbOccFem / prenom.iNbOcc
            iNbPrenomsVerif += prenom.iNbOcc
            iNbPrenomsVerifMF += prenom.iNbOccMasc + prenom.iNbOccFem

            If prenom.bMasc AndAlso prenom.bFem AndAlso
                prenom.iNbOcc >= iSeuilMin Then
                If prenom.iNbOccMasc > prenom.iNbOccFem Then
                    prenom.rFreqRelative = prenom.rFreqRelativeFem
                Else
                    prenom.rFreqRelative = prenom.rFreqRelativeMasc
                End If
                If prenom.rFreqRelative >= rSeuilFreqRel Then prenom.bMixte = True
            End If

        Next

        'Debug.WriteLine("Tot. Ok : " & sFormaterNum(iNbPrenomsVerif) & "=" & sFormaterNum(iNbPrenomsTotOk))
        'Debug.WriteLine("Tot. Ok : " & sFormaterNum(iNbPrenomsVerifMF) & "=" &
        '    sFormaterNum(iNbPrenomsTotOk))
        'Debug.WriteLine("Tot.: " &
        '    sFormaterNum(iNbPrenomsVerif + iNbPrenomsIgnores + iNbPrenomsIgnoresDate) & "=" &
        '    sFormaterNum(iNbPrenomsTot))

        Dim iNbPrenomsMixtes% = 0
        Dim sb As New StringBuilder
        AfficherInfo(sb, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel)
        Dim sbMD As New StringBuilder ' Syntaxe MarkDown
        AfficherInfo(sbMD, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, bDoublerRAL:=True)

        sbMD.AppendLine("|n° |Occurrences|Prénom|Année moyenne|Année moyenne masc.|Année moyenne fém.|Occurrences masc.|Occurrences fém.|Fréq.|Fréq. rel. masc.|Fréq. rel. fém.|")
        sbMD.AppendLine("|--:|--:|:--|:-:|:-:|:-:|--:|--:|--:|--:|--:|")

        Dim iNbLignesFin = 0
        For Each prenom In dico.Trier("bMixte desc, rFreqTotale desc")
            iNbLignesFin += 1
            If Not prenom.bMixte Then Continue For
            iNbPrenomsMixtes += 1
            If iNbLignesFin > 1000 Then Exit For
            Dim sGenre$ = "(f) ="
            If prenom.iNbOccMasc < prenom.iNbOccFem Then sGenre = "(m) ="
            Const sFormatFreq$ = "0.000%"
            sb.AppendLine(
                iNbPrenomsMixtes &
                " : " & sFormaterNum(prenom.iNbOcc) & " : " & prenom.sPrenom &
                ", " & prenom.rAnneeMoy.ToString("0") &
                ", " & prenom.rAnneeMoyMasc.ToString("0") & " (m)" &
                ", " & prenom.rAnneeMoyFem.ToString("0") & " (f)" &
                ", " & sFormaterNum(prenom.iNbOccMasc) & " (m)" &
                ", " & sFormaterNum(prenom.iNbOccFem) & " (f)" &
                ", freq. tot.=" & prenom.rFreqTotale.ToString(sFormatFreq) &
                ", freq. rel. m. " & sGenre & prenom.rFreqRelativeMasc.ToString("0%") &
                ", freq. rel. f. " & sGenre & prenom.rFreqRelativeFem.ToString("0%") &
                ", mixte=" & prenom.bMixte)

            sbMD.AppendLine(
                "|" & iNbPrenomsMixtes &
                "|" & sFormaterNum(prenom.iNbOcc) &
                "|" & prenom.sPrenom &
                "|" & prenom.rAnneeMoy.ToString("0") &
                "|" & prenom.rAnneeMoyMasc.ToString("0") &
                "|" & prenom.rAnneeMoyFem.ToString("0") &
                "|" & sFormaterNum(prenom.iNbOccMasc) &
                "|" & sFormaterNum(prenom.iNbOccFem) &
                "|" & prenom.rFreqTotale.ToString(sFormatFreq) &
                "|" & prenom.rFreqRelativeMasc.ToString("0%") &
                "|" & prenom.rFreqRelativeFem.ToString("0%"))

        Next

        Dim s$
        's = sb.ToString
        'Debug.WriteLine(s)

        s = sbMD.ToString
        Debug.WriteLine(s)

        MsgBox("Terminé !", MsgBoxStyle.Information, "Prénom mixte")

    End Sub

    Private Sub AfficherInfo(sb As StringBuilder,
            iNbPrenomsTotOk%, iNbPrenomsTot%, iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%, iSeuilMin%, rSeuilFreqRel!,
            Optional bDoublerRAL As Boolean = False)

        sb.AppendLine("Date début = 1900")
        If bDoublerRAL Then sb.AppendLine("")
        sb.AppendLine("Date fin   = 2019")
        If bDoublerRAL Then sb.AppendLine("")
        sb.AppendLine("Nb. total de prénoms identifiés et datés = " & sFormaterNum(iNbPrenomsTotOk))
        If bDoublerRAL Then sb.AppendLine("")
        sb.AppendLine("Nb. total de prénoms = " & sFormaterNum(iNbPrenomsTot))
        If bDoublerRAL Then sb.AppendLine("")
        sb.AppendLine("Nb. prénoms ignorés ('_PRENOMS_RARES') = " &
            sFormaterNum(iNbPrenomsIgnores) & " : " &
            (iNbPrenomsIgnores / iNbPrenomsTot).ToString("0.0%"))
        If bDoublerRAL Then sb.AppendLine("")
        sb.AppendLine("Nb prénoms ignorés (date 'XXXX') = " &
            sFormaterNum(iNbPrenomsIgnoresDate) & " : " &
            (iNbPrenomsIgnoresDate / iNbPrenomsTot).ToString("0.0%"))
        If bDoublerRAL Then sb.AppendLine("")
        sb.AppendLine("Seuil min. = " & sFormaterNum(iSeuilMin))
        If bDoublerRAL Then sb.AppendLine("")
        sb.AppendLine("Fréquence relative min. = " & rSeuilFreqRel.ToString("0.0%"))
        If bDoublerRAL Then sb.AppendLine("")

    End Sub

    Private Function iConvN(ByVal sVal$) As Nullable(Of Integer)

        If String.IsNullOrEmpty(sVal) Then Return Nothing

        Dim iVal%
        If Integer.TryParse(sVal, iVal) Then
            Return iVal
        Else
            Return Nothing
        End If

    End Function

    Private Function FirstCharToUpper$(sTexte$)

        If String.IsNullOrEmpty(sTexte) Then Throw New ArgumentException("Chaîne vide !")
        Return sTexte.First().ToString().ToUpper() + sTexte.Substring(1)

    End Function

    Private Function sFormaterNum$(iNum%)
        Return iNum.ToString("n").Replace(".00", "")
    End Function

End Class