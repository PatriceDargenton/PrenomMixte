
Imports System.Text

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

    Public Class clsPrenom : Implements ICloneable

        Public Const sPrenomRare$ = "_PRENOMS_RARES"
        Public Const sDateXXXX$ = "XXXX"

        Private Function IClone() As Object Implements ICloneable.Clone
            Return MemberwiseClone()
        End Function
        Public Function Clone() As clsPrenom
            Return DirectCast(Me.IClone(), clsPrenom)
        End Function

        Public sPrenom$, sPrenomOrig$, sPrenomHomophone$, sAnnee$, sCodeSexe$, sNbOcc$
        Public bMasc As Boolean
        Public bFem As Boolean
        Public bMixteEpicene As Boolean
        Public bMixteHomophone As Boolean
        Public iNbOccMasc%, iNbOccFem%, iNbOcc%
        Public rFreqRelative#, rFreqRelativeMasc#, rFreqRelativeFem#
        Public rFreqTotale#, rFreqTotaleMasc#, rFreqTotaleFem#
        Public iAnnee%, rAnneeMoy#, rAnneeMoyMasc#, rAnneeMoyFem#
        Public bSelect As Boolean
        Public dicoVariantes As New DicoTri(Of String, clsPrenom)

        Public Sub Calculer(iNbPrenomsTot%)
            If iNbPrenomsTot > 0 Then
                Me.rFreqTotale = Me.iNbOcc / iNbPrenomsTot
                Me.rFreqTotaleMasc = Me.iNbOccMasc / iNbPrenomsTot
                Me.rFreqTotaleFem = Me.iNbOccFem / iNbPrenomsTot
            End If
            If Me.iNbOccMasc > 0 Then Me.rAnneeMoyMasc = Me.rAnneeMoyMasc / Me.iNbOccMasc
            If Me.iNbOccFem > 0 Then Me.rAnneeMoyFem = Me.rAnneeMoyFem / Me.iNbOccFem
            If Me.iNbOcc > 0 Then
                Me.rAnneeMoy = Me.rAnneeMoy / Me.iNbOcc
                Me.rFreqRelativeMasc = Me.iNbOccMasc / Me.iNbOcc
                Me.rFreqRelativeFem = Me.iNbOccFem / Me.iNbOcc
            End If
        End Sub

    End Class

    Private Sub AnalyserPrenoms(Optional bExporter As Boolean = False)

        Dim sChemin = Application.StartupPath & "\nat2019.csv"
        If Not IO.File.Exists(sChemin) Then
            MsgBox("Veuillez télécharger nat2019_csv.zip !" & vbLf & sChemin,
                MsgBoxStyle.Exclamation, "Prénom mixte")
            Exit Sub
        End If
        Dim asLignes$() = IO.File.ReadAllLines(sChemin, Encoding.UTF8)
        If IsNothing(asLignes) Then Exit Sub

        Dim sCheminPrenomsMixtesEpicenes$ = Application.StartupPath &
            "\CorrectionsPrenomsMixtesEpicenes.csv"
        Dim dicoCorrectionsPrenomMixteEpicene = LireFichier(sCheminPrenomsMixtesEpicenes)
        Dim dicoCorrectionsPrenomMixteEpiceneUtil As New DicoTri(Of String, String)

        ' Vérifier si le fichier de correction des prénoms mixtes épicènes corrige bien uniquement les accents
        For Each kvp In dicoCorrectionsPrenomMixteEpicene
            Dim sPrenomOrig$ = kvp.Key
            Dim sPrenomCorrige$ = kvp.Value
            Dim sVerif$ = sEnleverAccents(sPrenomCorrige)
            If sVerif <> sPrenomOrig Then
                Debug.WriteLine(sVerif & "<>" & sPrenomOrig)
            End If
        Next

        Dim sCheminPrenomsMixtesHomophones$ = Application.StartupPath &
            "\CorrectionsPrenomsMixtesHomophones.csv"
        Dim dicoCorrectionsPrenomMixteHomophone = LireFichier(sCheminPrenomsMixtesHomophones)
        Dim dicoCorrectionsPrenomMixteHomophoneUtil As New DicoTri(Of String, String)

        Dim dicoE As New DicoTri(Of String, clsPrenom) ' épicène
        Dim dicoH As New DicoTri(Of String, clsPrenom) ' homophone
        'Dim dicoT As New DicoTri(Of String, String) ' Détection des prénoms homophones 
        Dim iNbLignes% = 0
        Dim iNbLignesOk% = 0
        Dim iNbPrenomsTot% = 0
        Dim iNbPrenomsTotOk% = 0
        Dim iNbPrenomsIgnores% = 0
        Dim iNbPrenomsIgnoresDate% = 0

        For Each sLigne As String In asLignes

            iNbLignes += 1
            If iNbLignes = 1 Then Continue For ' Entête

            Dim prenom As New clsPrenom
            If Not bAnalyserPrenom(sLigne$, prenom,
                dicoCorrectionsPrenomMixteEpicene,
                dicoCorrectionsPrenomMixteHomophone,
                dicoCorrectionsPrenomMixteEpiceneUtil,
                dicoCorrectionsPrenomMixteHomophoneUtil) Then Continue For
            ConvertirPrenom(prenom)

            If prenom.sPrenomOrig = clsPrenom.sPrenomRare Then
                iNbPrenomsIgnores += prenom.iNbOcc
                iNbPrenomsTot += prenom.iNbOcc
                Continue For
            End If
            If prenom.sAnnee = clsPrenom.sDateXXXX Then
                iNbPrenomsIgnoresDate += prenom.iNbOcc
                iNbPrenomsTot += prenom.iNbOcc
                Continue For
            End If
            iNbLignesOk += 1
            iNbPrenomsTotOk += prenom.iNbOcc
            iNbPrenomsTot += prenom.iNbOcc

            prenom.rAnneeMoy = prenom.iAnnee * prenom.iNbOcc
            If prenom.bMasc Then prenom.rAnneeMoyMasc = prenom.iAnnee * prenom.iNbOcc
            If prenom.bFem Then prenom.rAnneeMoyFem = prenom.iAnnee * prenom.iNbOcc

            Dim sCle$ = prenom.sPrenom
            If dicoE.ContainsKey(sCle) Then
                Dim prenom0 = dicoE(sCle)
                If prenom0.bMasc AndAlso prenom.bFem Then prenom0.bFem = True
                If prenom0.bFem AndAlso prenom.bMasc Then prenom0.bMasc = True
                prenom0.iNbOccFem += prenom.iNbOccFem
                prenom0.iNbOccMasc += prenom.iNbOccMasc
                prenom0.iNbOcc += prenom.iNbOcc
                prenom0.rAnneeMoy += prenom.rAnneeMoy
                prenom0.rAnneeMoyMasc += prenom.rAnneeMoyMasc
                prenom0.rAnneeMoyFem += prenom.rAnneeMoyFem
            Else
                dicoE.Add(sCle, prenom)
            End If

            ' Dico des prénoms homophones
            Dim prenomH = prenom.Clone() ' Il faut faire une copie pour que l'objet soit distinct
            If Not prenomH.dicoVariantes.ContainsKey(prenomH.sPrenom) Then
                prenomH.dicoVariantes.Add(prenomH.sPrenom, prenom)
            End If
            Dim sCleH$ = prenomH.sPrenomHomophone
            If dicoH.ContainsKey(sCleH) Then
                Dim prenom0 = dicoH(sCleH)
                If prenom0.bMasc AndAlso prenom.bFem Then prenom0.bFem = True
                If prenom0.bFem AndAlso prenom.bMasc Then prenom0.bMasc = True
                prenom0.iNbOccFem += prenom.iNbOccFem
                prenom0.iNbOccMasc += prenom.iNbOccMasc
                prenom0.iNbOcc += prenom.iNbOcc
                prenom0.rAnneeMoy += prenom.rAnneeMoy
                prenom0.rAnneeMoyMasc += prenom.rAnneeMoyMasc
                prenom0.rAnneeMoyFem += prenom.rAnneeMoyFem

                For Each kvp In prenomH.dicoVariantes
                    If Not prenom0.dicoVariantes.ContainsKey(kvp.Key) Then
                        prenom0.dicoVariantes.Add(kvp.Key, prenom)
                    End If
                Next

            Else
                dicoH.Add(sCleH, prenomH)
            End If

            '' Détection des prénoms homophones 
            'Dim sPrenomF1$ = prenom.sPrenom & "e"
            'If dicoE.ContainsKey(sPrenomF1) AndAlso Not dicoT.ContainsKey(sPrenomF1) Then
            '    dicoT.Add(sPrenomF1, prenom.sPrenom)
            'End If
            'Dim sPrenomF2$ = prenom.sPrenom & "le"
            'If dicoE.ContainsKey(sPrenomF2) AndAlso Not dicoT.ContainsKey(sPrenomF2) Then
            '    dicoT.Add(sPrenomF2, prenom.sPrenom)
            'End If

        Next

        '' Liste des prénoms homophones trouvées
        'Dim iNbPrenomU% = 0
        'For Each kvp In dicoT
        '    iNbPrenomU += 1
        '    If iNbPrenomU > 200 Then Exit For
        '    Dim sPrenom$ = kvp.Key
        '    Debug.WriteLine(sPrenom & " : " & dicoT(sPrenom))
        'Next
        'GoTo Fin

        'Const rSeuilFreqRel# = 0.001 ' 0.1% (par exemple 0.1% de masc. et 99.9% de fém.)
        Const rSeuilFreqRel# = 0.01 ' 1% (par exemple 1% de masc. et 99% de fém.)
        'Const rSeuilFreqRel# = 0.02 ' 2% (par exemple 2% de masc. et 98% de fém.)
        Const iSeuilMinEpicene% = 2000 ' Nombre minimal d'occurrences du prénom sur plus d'un siècle
        FiltrerPrenomMixteEpicene(dicoE, iNbPrenomsTot, iSeuilMinEpicene, rSeuilFreqRel,
            iNbPrenomsTotOk, iNbPrenomsIgnores, iNbPrenomsIgnoresDate)

        Const iSeuilMinHomophone% = 1 ' Nombre minimal d'occurrences du prénom sur plus d'un siècle
        FiltrerPrenomMixteHomophone(dicoH, dicoE, iNbPrenomsTot,
            iNbPrenomsTotOk, iNbPrenomsIgnores, iNbPrenomsIgnoresDate)

        If bExporter Then
            EcrireFichierFiltre(asLignes, dicoE,
                dicoCorrectionsPrenomMixteEpicene,
                dicoCorrectionsPrenomMixteHomophone,
                dicoCorrectionsPrenomMixteEpiceneUtil,
                dicoCorrectionsPrenomMixteHomophoneUtil)
            GoTo Fin
        End If

        Const iSeuilMin% = 50000
        Const iNbLignesMaxPrenom% = 0 ' 32346 prénoms en tout (reste quelques accents à corriger)
        AfficherSynthesePrenomsFrequents(dicoE, dicoH, iNbPrenomsTotOk, iNbPrenomsTot,
            iNbPrenomsIgnores, iNbPrenomsIgnoresDate, iSeuilMin, 0, iNbLignesMaxPrenom)

        Const iNbLignesMax% = 10000
        AfficherSyntheseEpicene(dicoE, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores,
            iNbPrenomsIgnoresDate, iSeuilMinEpicene, rSeuilFreqRel, iNbLignesMax,
            dicoCorrectionsPrenomMixteEpicene,
            dicoCorrectionsPrenomMixteEpiceneUtil)

        AfficherSyntheseHomophone(dicoH, iNbPrenomsTotOk, iNbPrenomsTot,
            iNbPrenomsIgnores, iNbPrenomsIgnoresDate, iSeuilMinHomophone, 0, iNbLignesMax,
            dicoCorrectionsPrenomMixteHomophone,
            dicoCorrectionsPrenomMixteHomophoneUtil)

Fin:
        MsgBox("Terminé !", MsgBoxStyle.Information, "Prénom mixte")

    End Sub

    Private Function LireFichier(sChemin$) As DicoTri(Of String, String)

        ' Lire un fichier de corrections de prénoms et retourner un dictionnaire

        Dim dico As New DicoTri(Of String, String)
        If Not IO.File.Exists(sChemin) Then
            MsgBox("Impossible de trouver le fichier suivant :" & vbLf &
                sChemin, MsgBoxStyle.Exclamation, "Prénom mixte")
            Return dico
        End If
        Dim asLignes$() = IO.File.ReadAllLines(sChemin, Encoding.UTF8)
        If IsNothing(asLignes) Then Return dico

        Dim hsDoublons As New HashSet(Of String)
        Dim iNbLignes% = 0
        For Each sLigne As String In asLignes
            iNbLignes += 1
            If iNbLignes = 1 Then Continue For ' Entête
            If sLigne.StartsWith("'") Then Continue For ' Commentaire
            'If sLigne.StartsWith("michel;michelle") Then
            '    Debug.WriteLine("!")
            'End If
            Dim asChamps() As String
            asChamps = Split(sLigne, ";"c)
            Dim iNumChampMax% = asChamps.GetUpperBound(0)
            Dim iNumChamp% = 0
            Dim sValeurOrig$ = ""
            Dim sValeurCorrigee$ = ""
            For Each sChamp As String In asChamps
                iNumChamp += 1
                If IsNothing(sChamp) Then sChamp = ""
                If sChamp.Length = 0 Then Exit For
                Select Case iNumChamp
                    Case 1 : sValeurCorrigee = sChamp
                    Case 2 : sValeurOrig = sChamp
                End Select
                If sValeurOrig.Contains("'") Then ' Commentaire à la fin de la ligne
                    Dim iPosQuote% = sValeurOrig.IndexOf("'")
                    sValeurOrig = sValeurOrig.Substring(0, iPosQuote)
                    sValeurOrig = sValeurOrig.TrimEnd
                End If
            Next

            ' Vérifier les doublons
            Dim sCle$ = sValeurCorrigee & ";" & sValeurOrig
            If hsDoublons.Contains(sCle) Then
                MsgBox("Doublon : " & sCle & vbLf & IO.Path.GetFileName(sChemin),
                    MsgBoxStyle.Information, "Prénom Mixte")
            Else
                hsDoublons.Add(sCle)
            End If

            If String.IsNullOrEmpty(sValeurCorrigee) OrElse
               String.IsNullOrEmpty(sValeurOrig) Then Continue For
            If Not dico.ContainsKey(sValeurOrig) Then dico.Add(sValeurOrig, sValeurCorrigee)
        Next

        Return dico

    End Function

    Private Sub FiltrerPrenomMixteEpicene(dico As DicoTri(Of String, clsPrenom),
            iNbPrenomsTot%, iSeuilMin%, rSeuilFreqRel#,
            iNbPrenomsTotOk%, iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%)

        Dim iNbPrenomsVerif% = 0
        Dim iNbPrenomsVerifMF% = 0
        For Each prenom In dico.Trier("")

            prenom.Calculer(iNbPrenomsTot)

            iNbPrenomsVerif += prenom.iNbOcc
            iNbPrenomsVerifMF += prenom.iNbOccMasc + prenom.iNbOccFem

            If prenom.bMasc AndAlso prenom.bFem AndAlso
                prenom.iNbOcc >= iSeuilMin Then
                If prenom.iNbOccMasc > prenom.iNbOccFem Then
                    prenom.rFreqRelative = prenom.rFreqRelativeFem
                Else
                    prenom.rFreqRelative = prenom.rFreqRelativeMasc
                End If
                If prenom.rFreqRelative >= rSeuilFreqRel Then prenom.bMixteEpicene = True
            End If

        Next

        'Debug.WriteLine("Tot. Ok : " & sFormaterNum(iNbPrenomsVerif) & "=" &
        '    sFormaterNum(iNbPrenomsTotOk))
        'Debug.WriteLine("Tot. Ok : " & sFormaterNum(iNbPrenomsVerifMF) & "=" &
        '    sFormaterNum(iNbPrenomsTotOk))
        'Debug.WriteLine("Tot.: " &
        '    sFormaterNum(iNbPrenomsVerif + iNbPrenomsIgnores + iNbPrenomsIgnoresDate) & "=" &
        '    sFormaterNum(iNbPrenomsTot))

    End Sub

    Private Sub FiltrerPrenomMixteHomophone(
            dicoH As DicoTri(Of String, clsPrenom),
            dico As DicoTri(Of String, clsPrenom),
            iNbPrenomsTot%, iNbPrenomsTotOk%, iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%)

        Dim iNbPrenomsVerif% = 0
        Dim iNbPrenomsVerifMF% = 0
        For Each prenom In dicoH.Trier("")

            If dico.ContainsKey(prenom.sPrenom) Then
                Dim prenom0 = dico(prenom.sPrenom)
                prenom.bMixteEpicene = prenom0.bMixteEpicene
            End If

            prenom.Calculer(iNbPrenomsTot)
            iNbPrenomsVerif += prenom.iNbOcc
            iNbPrenomsVerifMF += prenom.iNbOccMasc + prenom.iNbOccFem

            If prenom.dicoVariantes.Count > 1 Then prenom.bMixteHomophone = True

        Next

        'Debug.WriteLine("Tot. Ok : " & sFormaterNum(iNbPrenomsVerif) & "=" &
        '    sFormaterNum(iNbPrenomsTotOk))
        'Debug.WriteLine("Tot. Ok : " & sFormaterNum(iNbPrenomsVerifMF) & "=" &
        '    sFormaterNum(iNbPrenomsTotOk))
        'Debug.WriteLine("Tot.: " &
        '    sFormaterNum(iNbPrenomsVerif + iNbPrenomsIgnores + iNbPrenomsIgnoresDate) & "=" &
        '    sFormaterNum(iNbPrenomsTot))

    End Sub

    Private Sub AfficherSynthesePrenomsFrequents(
            dicoE As DicoTri(Of String, clsPrenom),
            dicoH As DicoTri(Of String, clsPrenom),
            iNbPrenomsTotOk%, iNbPrenomsTot%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, iNbLignesMax%)

        ' Produire la synthèse statistique des prénoms fréquents

        Dim sb As New StringBuilder
        AfficherInfo(sb, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel)
        Dim sbMD As New StringBuilder ' Syntaxe MarkDown
        sbMD.AppendLine("Synthèse statistique des prénoms fréquents")
        AfficherInfo(sbMD, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, bDoublerRAL:=True)
        sbMD.AppendLine(sEnteteMarkDown())

        Dim sbWK As New StringBuilder ' Syntaxe Wiki
        AfficherInfo(sbWK, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, bDoublerRAL:=True)
        sbWK.AppendLine(sEnteteWiki("Synthèse statistique des prénoms fréquents"))

        Dim iNbPrenoms% = 0
        Dim iNbLignesFin = 0
        For Each prenom In dicoE.Trier("iNbOcc desc") '"sPrenom" : Ordre alphabétique
            iNbLignesFin += 1
            If iSeuilMin > 0 AndAlso prenom.iNbOcc < iSeuilMin Then Continue For

            iNbPrenoms += 1
            If iNbLignesMax > 0 AndAlso iNbLignesFin > iNbLignesMax Then Exit For

            prenom.bSelect = True

            Const sFormatFreq$ = "0.000%"
            sb.AppendLine(sLigneDebug(prenom, prenom.sPrenom, iNbPrenoms, sFormatFreq))

            Dim bItalique = False
            Dim iNumVariante% = -1
            If prenom.bMixteEpicene Then iNumVariante = 1 ' Gras
            If dicoH.ContainsKey(prenom.sPrenom) Then
                Dim prenomH = dicoH(prenom.sPrenom)
                If prenomH.bMixteHomophone Then bItalique = True
            End If
            sbMD.AppendLine(sLigneMarkDown(prenom, prenom.sPrenom, iNbPrenoms, sFormatFreq,
                iNumVariante, bItalique))

            sbWK.AppendLine(sLigneWiki(prenom, prenom.sPrenom, iNbPrenoms, sFormatFreq,
                iNumVariante, bItalique))

        Next
        sbWK.AppendLine("|}")

        Dim sChemin$ = Application.StartupPath & "\PrenomsFrequents.md"
        EcrireFichier(sChemin, sbMD)
        Dim sCheminWK$ = Application.StartupPath & "\PrenomsFrequents.wiki"
        EcrireFichier(sCheminWK, sbWK)

    End Sub

    Private Sub AfficherSyntheseEpicene(dico As DicoTri(Of String, clsPrenom),
            iNbPrenomsTotOk%, iNbPrenomsTot%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, iNbLignesMax%,
            dicoCorrectionsPrenomMixteEpicene As DicoTri(Of String, String),
            dicoCorrectionsPrenomMixteEpiceneUtil As DicoTri(Of String, String))

        ' Produire la synthèse statistique des prénoms mixtes épicènes

        Dim sb As New StringBuilder
        AfficherInfo(sb, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel)
        Dim sbMD As New StringBuilder ' Syntaxe MarkDown
        sbMD.AppendLine("Synthèse statistique des prénoms mixtes épicènes")
        AfficherInfo(sbMD, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, bDoublerRAL:=True)
        sbMD.AppendLine(sEnteteMarkDown())

        Dim sbWK As New StringBuilder ' Syntaxe Wiki
        AfficherInfo(sbWK, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, bDoublerRAL:=True)
        sbWK.AppendLine(sEnteteWiki("Synthèse statistique des prénoms mixtes épicènes"))

        Dim iNbPrenomsMixtes% = 0
        Dim iNbLignesFin = 0
        For Each prenom In dico.Trier("bMixteEpicene desc, rFreqTotale desc")
            iNbLignesFin += 1
            If Not prenom.bMixteEpicene Then Continue For
            iNbPrenomsMixtes += 1
            If iNbLignesMax > 0 AndAlso iNbLignesFin > iNbLignesMax Then Exit For

            prenom.bSelect = True

            Const sFormatFreq$ = "0.000%"
            sb.AppendLine(sLigneDebug(prenom, prenom.sPrenom, iNbPrenomsMixtes, sFormatFreq))

            sbMD.AppendLine(sLigneMarkDown(prenom, prenom.sPrenom, iNbPrenomsMixtes, sFormatFreq))

            sbWK.AppendLine(sLigneWiki(prenom, prenom.sPrenom, iNbPrenomsMixtes, sFormatFreq))

        Next
        sbWK.AppendLine("|}")

        For Each kvp In dicoCorrectionsPrenomMixteEpicene
            If Not dicoCorrectionsPrenomMixteEpiceneUtil.ContainsKey(kvp.Key) Then
                Dim sLigne$ = "Correction de prénom épicène non trouvée : " & kvp.Key
                sb.AppendLine(sLigne)
                sbMD.AppendLine(sLigne)
                sbWK.AppendLine(sLigne)
            End If
        Next

        'Debug.WriteLine(sb.ToString)

        Dim sCheminMD$ = Application.StartupPath & "\PrenomsMixtesEpicenes.md"
        EcrireFichier(sCheminMD, sbMD)
        Dim sCheminWK$ = Application.StartupPath & "\PrenomsMixtesEpicenes.wiki"
        EcrireFichier(sCheminWK, sbWK)

    End Sub

    Private Sub AfficherSyntheseHomophone(dico As DicoTri(Of String, clsPrenom),
            iNbPrenomsTotOk%, iNbPrenomsTot%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, iNbLignesMax%,
            dicoCorrectionsPrenomMixteHomophone As DicoTri(Of String, String),
            dicoCorrectionsPrenomMixteHomophoneUtil As DicoTri(Of String, String))

        ' Produire la synthèse statistique des prénoms mixtes homophones

        Dim sb As New StringBuilder
        AfficherInfo(sb, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel)
        Dim sbMD As New StringBuilder ' Syntaxe MarkDown
        sbMD.AppendLine("Synthèse statistique des prénoms mixtes homophones")
        AfficherInfo(sbMD, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, bDoublerRAL:=True)
        sbMD.AppendLine(sEnteteMarkDown())

        Dim sbWK As New StringBuilder ' Syntaxe Wiki
        AfficherInfo(sbWK, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, bDoublerRAL:=True)
        sbWK.AppendLine(sEnteteWiki("Synthèse statistique des prénoms mixtes homophones"))

        Dim iNbPrenomsMixtes% = 0
        Dim iNbLignesFin = 0
        For Each prenom In dico.Trier("bMixteHomophone desc, rFreqTotale desc")
            iNbLignesFin += 1
            If Not prenom.bMixteHomophone Then Continue For
            iNbPrenomsMixtes += 1
            If iNbLignesMax > 0 AndAlso iNbLignesFin > iNbLignesMax Then Exit For

            prenom.bSelect = True

            Const sFormatFreq$ = "0.000%"

            Dim sPrenom$ = prenom.sPrenomHomophone
            Dim sPrenomMD$ = sPrenom
            Dim sPrenomWiki$ = sPrenom
            Dim bVariantes = False
            If prenom.dicoVariantes.Count > 1 Then
                bVariantes = True
                Dim lst = prenom.dicoVariantes.ToList
                Dim sPrenomMajoritaire$ = ""
                For Each prenomV In prenom.dicoVariantes.Trier("iNbOcc desc")
                    sPrenomMajoritaire = prenomV.sPrenom
                    Exit For
                Next
                sPrenomMD = sListerCleTxt(lst, sPrenomMajoritaire, "**")
                sPrenomWiki = sListerCleTxt(lst, sPrenomMajoritaire, "'''")
            End If

            sb.AppendLine(sLigneDebug(prenom, sPrenom, iNbPrenomsMixtes, sFormatFreq))
            sbMD.AppendLine(sLigneMarkDown(prenom, sPrenomMD, iNbPrenomsMixtes, sFormatFreq,
                iNumVariante:=0, bSuffixeNumVariante:=True))
            sbWK.AppendLine(sLigneWiki(prenom, sPrenomWiki, iNbPrenomsMixtes, sFormatFreq,
                iNumVariante:=0, bSuffixeNumVariante:=True))

            If bVariantes Then
                Dim iNumVariante% = 0
                For Each prenomV In prenom.dicoVariantes.Trier("iNbOcc desc")
                    iNumVariante += 1
                    sb.AppendLine(sLigneDebug(prenomV, prenomV.sPrenom, iNbPrenomsMixtes, sFormatFreq))
                    sbMD.AppendLine(sLigneMarkDown(prenomV, prenomV.sPrenom, iNbPrenomsMixtes,
                        sFormatFreq, iNumVariante, bSuffixeNumVariante:=True))
                    sbWK.AppendLine(sLigneWiki(prenomV, prenomV.sPrenom, iNbPrenomsMixtes,
                        sFormatFreq, iNumVariante, bSuffixeNumVariante:=True))
                Next
            End If

        Next
        sbWK.AppendLine("|}")

        For Each kvp In dicoCorrectionsPrenomMixteHomophone
            If Not dicoCorrectionsPrenomMixteHomophoneUtil.ContainsKey(kvp.Key) Then
                Dim sLigne$ = "Correction de prénom homophone non trouvée : " & kvp.Key
                sb.AppendLine(sLigne)
                sbMD.AppendLine(sLigne)
                sbWK.AppendLine(sLigne)
            End If
        Next

        'Debug.WriteLine(sb.ToString)

        Dim sCheminMD$ = Application.StartupPath & "\PrenomsMixtesHomophones.md"
        EcrireFichier(sCheminMD, sbMD)
        Dim sCheminWK$ = Application.StartupPath & "\PrenomsMixtesHomophones.wiki"
        EcrireFichier(sCheminWK, sbWK)

    End Sub

    Private Function sEnteteMarkDown$()

        Dim s$ = ""
        s &= vbLf & "|n° |Occurrences|Prénom|Année moyenne|Année moyenne masc.|Année moyenne fém.|Occurrences masc.|Occurrences fém.|Fréq.|Fréq. rel. masc.|Fréq. rel. fém.|"
        s &= vbLf & "|--:|--:|:--|:-:|:-:|:-:|--:|--:|--:|--:|--:|"
        Return s

    End Function

    Private Function sEnteteWiki$(sTitre$)

        ' https://fr.wikipedia.org/wiki/Aide:Insérer_un_tableau_(wikicode,_avancé)

        Dim s$ = ""
        s &= vbLf & "{|class='wikitable sortable' style='text-align:center; width:80%;'"
        s &= vbLf & "|+ " & sTitre
        s &= vbLf & "! scope='col' | n°"
        s &= vbLf & "! scope='col' | Occurrences"
        s &= vbLf & "! scope='col' | Prénom"
        s &= vbLf & "! scope='col' | Année moyenne"
        s &= vbLf & "! scope='col' | Année moyenne masc."
        s &= vbLf & "! scope='col' | Année moyenne fém."
        s &= vbLf & "! scope='col' | Occurrences masc."
        s &= vbLf & "! scope='col' | Occurrences fém."
        s &= vbLf & "! scope='col' | Fréq."
        s &= vbLf & "! scope='col' | Fréq. rel. masc."
        s &= vbLf & "! scope='col' | Fréq. rel. fém."
        Return s

    End Function

    Private Function sLigneDebug$(prenom As clsPrenom, sPrenom$, iNumPrenom%, sFormatFreq$)

        Dim sGenre$ = "(f) ="
        If prenom.iNbOccMasc < prenom.iNbOccFem Then sGenre = "(m) ="
        Dim s$ =
            iNumPrenom &
            " : " & sFormaterNum(prenom.iNbOcc) & " : " & sPrenom &
            ", " & prenom.rAnneeMoy.ToString("0") &
            ", " & prenom.rAnneeMoyMasc.ToString("0") & " (m)" &
            ", " & prenom.rAnneeMoyFem.ToString("0") & " (f)" &
            ", " & sFormaterNum(prenom.iNbOccMasc) & " (m)" &
            ", " & sFormaterNum(prenom.iNbOccFem) & " (f)" &
            ", freq. tot.=" & prenom.rFreqTotale.ToString(sFormatFreq) &
            ", freq. rel. m. " & sGenre & prenom.rFreqRelativeMasc.ToString("0%") &
            ", freq. rel. f. " & sGenre & prenom.rFreqRelativeFem.ToString("0%") &
            ", mixte=" & prenom.bMixteEpicene
        Return s

    End Function

    Private Function sLigneMarkDown$(prenom As clsPrenom, sPrenom$, iNumPrenom%, sFormatFreq$,
            Optional iNumVariante% = -1, Optional bItalique As Boolean = False,
            Optional bSuffixeNumVariante As Boolean = False)

        Dim sMiseEnForme$ = ""
        Dim sNumVariante$ = ""
        If bSuffixeNumVariante AndAlso iNumVariante >= 0 Then sNumVariante = "." & iNumVariante
        If iNumVariante = 1 Then sMiseEnForme = "**" ' Gras
        If bItalique Then sMiseEnForme = "*" ' Italique
        If iNumVariante = 1 AndAlso bItalique Then sMiseEnForme = "***" ' Italique en gras

        Dim s$ =
            "|" & iNumPrenom & sNumVariante &
            "|" & sFormaterNum(prenom.iNbOcc) &
            "|" & sMiseEnForme & sPrenom & sMiseEnForme &
            "|" & prenom.rAnneeMoy.ToString("0") &
            "|" & prenom.rAnneeMoyMasc.ToString("0") &
            "|" & prenom.rAnneeMoyFem.ToString("0") &
            "|" & sFormaterNum(prenom.iNbOccMasc) &
            "|" & sFormaterNum(prenom.iNbOccFem) &
            "|" & prenom.rFreqTotale.ToString(sFormatFreq) &
            "|" & prenom.rFreqRelativeMasc.ToString("0%") &
            "|" & prenom.rFreqRelativeFem.ToString("0%")
        Return s

    End Function

    Private Function sLigneWiki$(prenom As clsPrenom, sPrenom$, iNumPrenom%, sFormatFreq$,
            Optional iNumVariante% = -1, Optional bItalique As Boolean = False,
            Optional bSuffixeNumVariante As Boolean = False)

        Dim sMiseEnForme$ = ""
        Dim sNumVariante$ = ""
        If bSuffixeNumVariante AndAlso iNumVariante >= 0 Then sNumVariante = "." & iNumVariante
        If iNumVariante = 1 Then sMiseEnForme = "'''" ' Gras
        If bItalique Then sMiseEnForme = "''" ' Italique
        If iNumVariante = 1 AndAlso bItalique Then sMiseEnForme = "'''''" ' Italique en gras

        Dim s$ = "|-" & vbLf &
                "|" & iNumPrenom & sNumVariante &
                "|| align='right' | " & sFormaterNumWiki(prenom.iNbOcc) &
                "|| " & sMiseEnForme & sPrenom & sMiseEnForme &
                "||" & prenom.rAnneeMoy.ToString("0") &
                "||" & prenom.rAnneeMoyMasc.ToString("0") &
                "||" & prenom.rAnneeMoyFem.ToString("0") &
                "|| align='right' | " & sFormaterNumWiki(prenom.iNbOccMasc) &
                "|| align='right' | " & sFormaterNumWiki(prenom.iNbOccFem) &
                "||" & prenom.rFreqTotale.ToString(sFormatFreq) &
                "||" & prenom.rFreqRelativeMasc.ToString("0%") &
                "||" & prenom.rFreqRelativeFem.ToString("0%")
        Return s

    End Function

    Private Function bAnalyserPrenom(sLigne$, prenom As clsPrenom,
            dicoCorrectionsPrenomMixteEpicene As DicoTri(Of String, String),
            dicoCorrectionsPrenomMixteHomophone As DicoTri(Of String, String),
            dicoCorrectionsPrenomMixteEpiceneUtil As DicoTri(Of String, String),
            dicoCorrectionsPrenomMixteHomophoneUtil As DicoTri(Of String, String)) As Boolean

        Dim asChamps() As String
        asChamps = Split(sLigne, ";"c)
        Dim iNumChampMax% = asChamps.GetUpperBound(0)
        Dim iNumChamp% = 0
        Dim bOk As Boolean = True
        Dim sCodeSexe$ = ""
        Dim sPrenomOrig$ = ""
        Dim sAnnee$ = ""
        Dim sNbOcc$ = ""
        For Each sChamp As String In asChamps
            iNumChamp += 1
            If IsNothing(sChamp) Then sChamp = ""
            If sChamp.Length = 0 Then bOk = False : Exit For
            Select Case iNumChamp
                Case 1 : sCodeSexe = sChamp
                Case 2 : sPrenomOrig = sChamp
                Case 3 : sAnnee = sChamp
                Case 4 : sNbOcc = sChamp
            End Select
        Next
        If Not bOk Then Return False

        Dim sPrenom = sPrenomOrig.ToLower

        For Each kvp In dicoCorrectionsPrenomMixteEpicene
            If sPrenom = kvp.Key Then
                If Not dicoCorrectionsPrenomMixteEpiceneUtil.ContainsKey(sPrenom) Then
                    dicoCorrectionsPrenomMixteEpiceneUtil.Add(sPrenom, kvp.Value)
                End If
                sPrenom = kvp.Value
            End If
        Next

        Dim sPrenomHomophone = sPrenom
        For Each kvp In dicoCorrectionsPrenomMixteHomophone
            If sPrenom = kvp.Key Then
                If Not dicoCorrectionsPrenomMixteHomophoneUtil.ContainsKey(sPrenom) Then
                    dicoCorrectionsPrenomMixteHomophoneUtil.Add(sPrenom, kvp.Value)
                End If
                sPrenomHomophone = kvp.Value
            End If
        Next

        ' 3ème rapport : prénoms féminisés
        'If sPrenom = "aloïse" Then sPrenomMasc = "aloïs"
        'If sPrenom = "adrienne" Then sPrenomMasc = "adrien"
        'If sPrenom = "antoinnette" Then sPrenomMasc = "antoinne"
        'If sPrenom = "charline" Then sPrenomMasc = "charles"
        'If sPrenom = "claudette" Then sPrenomMasc = "claude"
        'If sPrenom = "claudie" Then sPrenomMasc = "claude"
        'If sPrenom = "claudine" Then sPrenomMasc = "claude"
        'If sPrenom = "claudy" Then sPrenomMasc = "claude"
        'If sPrenom = "denise" Then sPrenomMasc = "denis"
        'If sPrenom = "edwige" Then sPrenomMasc = "edwig"
        'If sPrenom = "fernande" Then sPrenomMasc = "fernand"

        prenom.sPrenom = FirstCharToUpper(sPrenom)
        prenom.sPrenomHomophone = FirstCharToUpper(sPrenomHomophone)
        prenom.sPrenomOrig = sPrenomOrig

        prenom.sCodeSexe = sCodeSexe
        prenom.sAnnee = sAnnee
        prenom.sNbOcc = sNbOcc

        Return True

    End Function

    Private Sub ConvertirPrenom(prenom As clsPrenom)

        If prenom.sCodeSexe = "1" Then prenom.bMasc = True
        If prenom.sCodeSexe = "2" Then prenom.bFem = True

        Dim iNbOccN%? = iConvN(prenom.sNbOcc)
        If Not IsNothing(iNbOccN) Then prenom.iNbOcc = CInt(iNbOccN%)

        Dim iAnneeN%? = iConvN(prenom.sAnnee)
        If Not IsNothing(iAnneeN) Then prenom.iAnnee = CInt(iAnneeN)

        If prenom.bMasc Then prenom.iNbOccMasc = prenom.iNbOcc
        If prenom.bFem Then prenom.iNbOccFem = prenom.iNbOcc

    End Sub

    Private Sub EcrireFichierFiltre(asLignes$(), dico As DicoTri(Of String, clsPrenom),
        dicoCorrectionsPrenomMixteEpicene As DicoTri(Of String, String),
        dicoCorrectionsPrenomMixteHomophone As DicoTri(Of String, String),
        dicoCorrectionsPrenomMixteEpiceneUtil As DicoTri(Of String, String),
        dicoCorrectionsPrenomMixteHomophoneUtil As DicoTri(Of String, String))

        ' Génération d'un nouveau fichier csv filtré ou pas

        ' Vérifier si le traitement appliqué préserve entièrement le fichier d'origine
        Const bTestPrenomOrig = False
        Const bFiltrerPrenomEpicene = False

        Dim sb As New StringBuilder
        Dim iNbLignes = 0
        Dim sAjoutEntete$ = ""
        ' ToDo : Féminisé
        If Not bTestPrenomOrig Then sAjoutEntete = ";Prénom d'origine;Prénom épicène;Prénom homophone"
        For Each sLigne As String In asLignes

            iNbLignes += 1
            If iNbLignes = 1 Then sb.AppendLine(sLigne & sAjoutEntete) : Continue For

            Dim prenom As New clsPrenom
            If Not bAnalyserPrenom(sLigne$, prenom,
                dicoCorrectionsPrenomMixteEpicene,
                dicoCorrectionsPrenomMixteHomophone,
                dicoCorrectionsPrenomMixteEpiceneUtil,
                dicoCorrectionsPrenomMixteHomophoneUtil) Then Continue For

            ConvertirPrenom(prenom)

            Dim sAjout$ = ""
            Dim sPrenom$ = prenom.sPrenom
            If bTestPrenomOrig Then
                ' Si on remet en majuscule et qu'on rétablit les accents corrigés
                '  on doit retrouver exactement le fichier d'origine
                sPrenom = sPrenom.ToUpper
                If prenom.sPrenomOrig <> sPrenom Then
                    sPrenom = sEnleverAccents(sPrenom, bMinuscule:=False)
                    If prenom.sPrenomOrig <> sPrenom Then
                        ' Vérifier si tous les accents sont bien retirés
                        Debug.WriteLine(sPrenom)
                        Stop
                    End If
                End If
                If prenom.sPrenomOrig = clsPrenom.sPrenomRare Then GoTo Suite
                If prenom.sAnnee = clsPrenom.sDateXXXX Then GoTo Suite
            End If

            Dim sCle$ = prenom.sPrenom
            If dico.ContainsKey(sCle) Then
                Dim prenom0 = dico(sCle)
                If bFiltrerPrenomEpicene AndAlso Not prenom0.bSelect Then Continue For

                If Not bTestPrenomOrig Then
                    ' ToDo : Féminisé
                    sPrenom = prenom0.sPrenom
                    sAjout = ";" & prenom.sPrenomOrig & ";" & prenom.sPrenom & ";" & prenom.sPrenomHomophone
                End If

            Else
                Continue For
            End If

Suite:
            Dim sLigneC$ = prenom.sCodeSexe & ";" & sPrenom & ";" & prenom.sAnnee & ";" & prenom.iNbOcc & sAjout
            sb.AppendLine(sLigneC)

        Next

        Dim sCheminOut$ = Application.StartupPath & "\nat2019_corrige.csv"
        EcrireFichier(sCheminOut, sb, bConserverFormatOrigine:=bTestPrenomOrig)

    End Sub

    Private Sub EcrireFichier(sChemin$, sb As StringBuilder,
            Optional bConserverFormatOrigine As Boolean = False)

        ' Encodage classique : encoding:=Encoding.UTF8
        ' Pour comparer avec le format d'origine du fichier INSEE :
        '  encoding:=New UTF8Encoding(encoderShouldEmitUTF8Identifier:=False)
        Dim enc = Encoding.UTF8
        If bConserverFormatOrigine Then enc = New UTF8Encoding(encoderShouldEmitUTF8Identifier:=False)
        Using sw As New IO.StreamWriter(sChemin, append:=False, encoding:=enc)
            sw.Write(sb.ToString())
        End Using 'sw.Close()

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
        If iSeuilMin > 1 Then
            sb.AppendLine("Seuil min. = " & sFormaterNum(iSeuilMin))
            If bDoublerRAL Then sb.AppendLine("")
        End If
        If rSeuilFreqRel > 0 Then
            sb.AppendLine("Fréquence relative min. = " & rSeuilFreqRel.ToString("0.0%"))
            If bDoublerRAL Then sb.AppendLine("")
        End If

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

    Private Function sFormaterNumWiki$(iNum%)

        ' Pour la syntaxe wiki, éviter d'appliquer un format si le nombre est inférieur à 1000
        ' (sinon un bot corrigera cela)
        If Math.Abs(iNum) < 1000 Then Return iNum.ToString

        Return "{{formatnum:" & iNum & "}}"

    End Function

    Private Function sListerCleTxt$(
            lstTxt As List(Of KeyValuePair(Of String, clsPrenom)),
            sPrenomMajoritaire$, sGras$,
            Optional iNbMax% = 0)

        Dim sb As New StringBuilder("")
        Dim iNumOcc% = 0
        For Each kvp In lstTxt
            If sb.Length > 0 Then sb.Append(", ")
            Dim sPrenom$ = kvp.Key
            If sPrenom = sPrenomMajoritaire Then sPrenom = sGras & sPrenom & sGras
            sb.Append(sPrenom)
            iNumOcc += 1
            If iNbMax > 0 Then
                If iNumOcc >= iNbMax Then sb.Append("...") : Exit For
            End If
        Next

        Return sb.ToString

    End Function

    Private Function sEnleverAccents$(sChaine$, Optional bMinuscule As Boolean = True)

        ' Enlever les accents

        If sChaine.Length = 0 Then Return ""

        Dim sTexteSansAccents$ = sRemoveDiacritics(sChaine)
        If bMinuscule Then Return sTexteSansAccents.ToLower
        Return sTexteSansAccents

    End Function

    Private Function sRemoveDiacritics$(sTexte$)

        Dim sb As StringBuilder = sbRemoveDiacritics(sTexte)
        Dim sTexteDest$ = sb.ToString
        Return sTexteDest

    End Function

    Private Function sbRemoveDiacritics(sTexte$) As StringBuilder

        Dim sNormalizedString$ = sTexte.Normalize(NormalizationForm.FormD)
        Dim sb As New StringBuilder
        Const cChar_ae As Char = "æ"c
        Const cChar_oe As Char = "œ"c
        Const cChar_o As Char = "o"c
        Const cChar_e As Char = "e"c
        Const cChar_a As Char = "a"c
        Const cCharAE As Char = "Æ"c
        Const cCharOE As Char = "Œ"c
        Const cCharO As Char = "O"c
        Const cCharE As Char = "E"c
        Const cCharA As Char = "A"c
        Const cChar3P As Char = "…"c
        For Each c As Char In sNormalizedString
            Dim unicodeCategory As Globalization.UnicodeCategory = _
                Globalization.CharUnicodeInfo.GetUnicodeCategory(c)
            If (unicodeCategory <> Globalization.UnicodeCategory.NonSpacingMark) Then

                If c = cCharAE Then
                    sb.Append(cCharA)
                    sb.Append(cCharE)
                ElseIf c = cCharOE Then
                    sb.Append(cCharO)
                    sb.Append(cCharE)
                ElseIf c = cChar_ae Then
                    sb.Append(cChar_a)
                    sb.Append(cChar_e)
                ElseIf c = cChar_oe Then
                    sb.Append(cChar_o)
                    sb.Append(cChar_e)
                ElseIf c = cChar3P Then
                    sb.Append("...")
                Else
                    sb.Append(c)
                End If

            End If
        Next

        Return sb

    End Function

End Class