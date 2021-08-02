
Imports System.Text

Public Class frmPrenomMixte

    Private Sub cmdAnalyser_Click(sender As Object, e As EventArgs) Handles cmdAnalyser.Click

        Me.cmdAnalyser.Enabled = False
        AnalyserPrenoms()
        Me.cmdAnalyser.Enabled = True

    End Sub

    Public Class clsPrenom : Implements ICloneable

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
        Public hsVariantes As New HashSet(Of String)
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

    Private Sub AnalyserPrenoms()

        Dim sChemin = Application.StartupPath & "\nat2019.csv"
        If Not IO.File.Exists(sChemin) Then
            MsgBox("Veuillez télécharger nat2019_csv.zip !" & vbLf & sChemin,
                MsgBoxStyle.Exclamation, "Prénom mixte")
            Exit Sub
        End If
        Dim asLignes$() = IO.File.ReadAllLines(sChemin, Encoding.UTF8)
        If IsNothing(asLignes) Then Exit Sub

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
            If Not bAnalyserPrenom(sLigne$, prenom) Then Continue For
            ConvertirPrenom(prenom)

            If prenom.sPrenomOrig = "_PRENOMS_RARES" Then
                iNbPrenomsIgnores += prenom.iNbOcc
                iNbPrenomsTot += prenom.iNbOcc
                Continue For
            End If
            If prenom.sAnnee = "XXXX" Then
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
            'If prenomH.sPrenomHomophone <> prenomH.sPrenom Then
                If Not prenomH.hsVariantes.Contains(prenomH.sPrenom) Then
                    prenomH.hsVariantes.Add(prenomH.sPrenom)
                End If
                If Not prenomH.dicoVariantes.ContainsKey(prenomH.sPrenom) Then
                    prenomH.dicoVariantes.Add(prenomH.sPrenom, prenom)
                End If
            'End If
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

                For Each variante In prenomH.hsVariantes
                    If Not prenom0.hsVariantes.Contains(variante) Then
                        prenom0.hsVariantes.Add(variante)
                    End If
                    If Not prenom0.dicoVariantes.ContainsKey(variante) Then
                        prenom0.dicoVariantes.Add(variante, prenom)
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
        FiltrerPrenomMixteEpicene(dicoE, iNbPrenomsTot, iSeuilMinEpicene, rSeuilFreqRel)

        Const iSeuilMinHomophone% = 1 ' Nombre minimal d'occurrences du prénom sur plus d'un siècle
        FiltrerPrenomMixteHomophone(dicoH, dicoE, iNbPrenomsTot)

        EcrireFichierFiltre(asLignes, dicoE)

        'Const iSeuilMin% = 50000
        'Const iNbLignesMaxPrenom% = 0 ' 32346 prénoms en tout (reste quelques accents à corriger)
        'AfficherSynthesePrenoms(dicoE, iNbPrenomsTotOk, iNbPrenomsTot,
        '    iNbPrenomsIgnores, iNbPrenomsIgnoresDate, iSeuilMin, 0, iNbLignesMaxPrenom)
        'GoTo Fin

        Const iNbLignesMax% = 10000
        AfficherSyntheseEpicene(dicoE, iNbPrenomsTotOk, iNbPrenomsTot,
            iNbPrenomsIgnores, iNbPrenomsIgnoresDate, iSeuilMinEpicene, rSeuilFreqRel, iNbLignesMax)

        AfficherSyntheseHomophone(dicoH, iNbPrenomsTotOk, iNbPrenomsTot,
            iNbPrenomsIgnores, iNbPrenomsIgnoresDate, iSeuilMinHomophone, 0, iNbLignesMax)

Fin:
        MsgBox("Terminé !", MsgBoxStyle.Information, "Prénom mixte")

    End Sub

    Private Sub FiltrerPrenomMixteEpicene(dico As DicoTri(Of String, clsPrenom),
            iNbPrenomsTot%, iSeuilMin%, rSeuilFreqRel#)

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

        'Debug.WriteLine("Tot. Ok : " & sFormaterNum(iNbPrenomsVerif) & "=" & sFormaterNum(iNbPrenomsTotOk))
        'Debug.WriteLine("Tot. Ok : " & sFormaterNum(iNbPrenomsVerifMF) & "=" &
        '    sFormaterNum(iNbPrenomsTotOk))
        'Debug.WriteLine("Tot.: " &
        '    sFormaterNum(iNbPrenomsVerif + iNbPrenomsIgnores + iNbPrenomsIgnoresDate) & "=" &
        '    sFormaterNum(iNbPrenomsTot))

    End Sub

    Private Sub FiltrerPrenomMixteHomophone(
            dicoH As DicoTri(Of String, clsPrenom),
            dico As DicoTri(Of String, clsPrenom),
            iNbPrenomsTot%)

        For Each prenom In dicoH.Trier("")

            If dico.ContainsKey(prenom.sPrenom) Then
                Dim prenom0 = dico(prenom.sPrenom)
                prenom.bMixteEpicene = prenom0.bMixteEpicene
            End If

            prenom.Calculer(iNbPrenomsTot)

            If prenom.hsVariantes.Count > 1 Then prenom.bMixteHomophone = True

        Next

    End Sub

    Private Sub AfficherSynthesePrenoms(dico As DicoTri(Of String, clsPrenom),
            iNbPrenomsTotOk%, iNbPrenomsTot%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, iNbLignesMax%)

        ' Afficher la synthèse statistique des prénoms fréquents dans la fenêtre Debug de Visual Studio

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
        For Each prenom In dico.Trier("sPrenom")
            iNbLignesFin += 1
            If iSeuilMin > 0 AndAlso prenom.iNbOcc < iSeuilMin Then Continue For

            iNbPrenoms += 1
            If iNbLignesMax > 0 AndAlso iNbLignesFin > iNbLignesMax Then Exit For

            prenom.bSelect = True

            Const sFormatFreq$ = "0.000%"
            sb.AppendLine(sLigneDebug(prenom, prenom.sPrenom, iNbPrenoms, sFormatFreq))

            sbMD.AppendLine(sLigneMarkDown(prenom, prenom.sPrenom, iNbPrenoms, sFormatFreq))

            sbWK.AppendLine(sLigneWiki(prenom, prenom.sPrenom, iNbPrenoms, sFormatFreq))

        Next
        sbWK.AppendLine("|}")

        Dim sChemin$ = Application.StartupPath & "\Prenoms.md"
        EcrireFichier(sChemin, sbMD)
        Dim sCheminWK$ = Application.StartupPath & "\Prenoms.wiki"
        EcrireFichier(sCheminWK, sbWK)

    End Sub

    Private Sub AfficherSyntheseEpicene(dico As DicoTri(Of String, clsPrenom),
            iNbPrenomsTotOk%, iNbPrenomsTot%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, iNbLignesMax%)

        ' Afficher la synthèse statistique des prénoms mixtes épicènes dans la fenêtre Debug de Visual Studio


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

        Dim s$
        's = sb.ToString
        'Debug.WriteLine(s)

        's = sbMD.ToString
        'Debug.WriteLine(s)

        's = sbWK.ToString
        'Debug.WriteLine(s)

    End Sub

    Private Sub AfficherSyntheseHomophone(dico As DicoTri(Of String, clsPrenom),
            iNbPrenomsTotOk%, iNbPrenomsTot%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, iNbLignesMax%)

        ' Afficher la synthèse statistique des prénoms mixtes homophones dans la fenêtre Debug de Visual Studio

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
            Dim bVariantes = False
            If prenom.hsVariantes.Count > 1 Then
                bVariantes = True
                Dim lst = prenom.hsVariantes.ToList
                'If Not prenom.hsVariantes.Contains(sPrenom) Then lst.Add(sPrenom)
                sPrenom = sListerTxt(lst)
            End If

            sb.AppendLine(sLigneDebug(prenom, sPrenom, iNbPrenomsMixtes, sFormatFreq))
            sbMD.AppendLine(sLigneMarkDown(prenom, sPrenom, iNbPrenomsMixtes, sFormatFreq,
                iNumVariante:=0))
            sbWK.AppendLine(sLigneWiki(prenom, sPrenom, iNbPrenomsMixtes, sFormatFreq,
                iNumVariante:=0))

            If bVariantes Then
                Dim iNumVariante% = 0
                For Each prenomV In prenom.dicoVariantes.Trier("iNbOcc desc")
                    iNumVariante += 1
                    sb.AppendLine(sLigneDebug(prenomV, prenomV.sPrenom, iNbPrenomsMixtes, sFormatFreq))
                    sbMD.AppendLine(sLigneMarkDown(prenomV, prenomV.sPrenom, iNbPrenomsMixtes,
                        sFormatFreq, iNumVariante))
                    sbWK.AppendLine(sLigneWiki(prenomV, prenomV.sPrenom, iNbPrenomsMixtes,
                        sFormatFreq, iNumVariante))
                Next
            End If

        Next
        sbWK.AppendLine("|}")

        Dim s$
        's = sb.ToString
        'Debug.WriteLine(s)

        s = sbMD.ToString
        Debug.WriteLine(s)

        's = sbWK.ToString
        'Debug.WriteLine(s)

    End Sub

    Private Function sEnteteMarkDown$()

        Dim s$
        s &= vbLf & "|n° |Occurrences|Prénom|Année moyenne|Année moyenne masc.|Année moyenne fém.|Occurrences masc.|Occurrences fém.|Fréq.|Fréq. rel. masc.|Fréq. rel. fém.|"
        s &= vbLf & "|--:|--:|:--|:-:|:-:|:-:|--:|--:|--:|--:|--:|"
        Return s

    End Function

    Private Function sEnteteWiki$(sTitre$)

        ' https://fr.wikipedia.org/wiki/Aide:Insérer_un_tableau_(wikicode,_avancé)

        Dim s$
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
            Optional iNumVariante% = -1)

        Dim sNumVariante$ = ""
        If iNumVariante >= 0 Then sNumVariante = "." & iNumVariante

        Dim s$ =
            "|" & iNumPrenom & sNumVariante &
            "|" & sFormaterNum(prenom.iNbOcc) &
            "|" & sPrenom &
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
            Optional iNumVariante% = -1)

        Dim sNumVariante$ = ""
        If iNumVariante >= 0 Then sNumVariante = "." & iNumVariante

        Dim s$ = "|-" & vbLf &
                "|" & iNumPrenom & sNumVariante &
                "|| align='right' | {{formatnum:" & prenom.iNbOcc & "}}" &
                "|| " & sPrenom &
                "||" & prenom.rAnneeMoy.ToString("0") &
                "||" & prenom.rAnneeMoyMasc.ToString("0") &
                "||" & prenom.rAnneeMoyFem.ToString("0") &
                "|| align='right' | {{formatnum:" & prenom.iNbOccMasc & "}}" &
                "|| align='right' | {{formatnum:" & prenom.iNbOccFem & "}}" &
                "||" & prenom.rFreqTotale.ToString(sFormatFreq) &
                "||" & prenom.rFreqRelativeMasc.ToString("0%") &
                "||" & prenom.rFreqRelativeFem.ToString("0%")
        Return s

    End Function

    Private Function bAnalyserPrenom(sLigne$, prenom As clsPrenom) As Boolean

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

        If sPrenom = "adelaide" Then sPrenom = "adélaïde"
        If sPrenom = "adele" Then sPrenom = "adèle"
        If sPrenom = "aimee" Then sPrenom = "aimée"
        If sPrenom = "aissa" Then sPrenom = "aïssa"
        If sPrenom = "alfrede" Then sPrenom = "alfrède"
        If sPrenom = "alois" Then sPrenom = "aloïs"
        If sPrenom = "aloise" Then sPrenom = "aloïse"
        If sPrenom = "amedee" Then sPrenom = "amedée"
        If sPrenom = "anael" Then sPrenom = "anaël"
        If sPrenom = "anaelle" Then sPrenom = "anaëlle"
        If sPrenom = "andre" Then sPrenom = "andré"
        If sPrenom = "andrea" Then sPrenom = "andréa"
        If sPrenom = "andree" Then sPrenom = "andrée"
        If sPrenom = "arsene" Then sPrenom = "arsène"
        If sPrenom = "barthelemy" Then sPrenom = "barthélemy"
        If sPrenom = "benedict" Then sPrenom = "bénédict"
        If sPrenom = "benedicte" Then sPrenom = "bénédicte"
        If sPrenom = "celeste" Then sPrenom = "céleste"
        If sPrenom = "cleo" Then sPrenom = "cléo"
        If sPrenom = "come" Then sPrenom = "côme"
        If sPrenom = "danael" Then sPrenom = "danaël"
        If sPrenom = "danaelle" Then sPrenom = "danaëlle"
        If sPrenom = "daniele" Then sPrenom = "danièle"
        If sPrenom = "dorothee" Then sPrenom = "dorothée"
        If sPrenom = "eden" Then sPrenom = "éden"
        If sPrenom = "edme" Then sPrenom = "edmé"
        If sPrenom = "edmee" Then sPrenom = "edmée"
        If sPrenom = "eleonor" Then sPrenom = "éléonor"
        If sPrenom = "eleonore" Then sPrenom = "éléonore"
        If sPrenom = "eli" Then sPrenom = "éli"
        If sPrenom = "elia" Then sPrenom = "élia"
        If sPrenom = "elie" Then sPrenom = "élie"
        If sPrenom = "elisee" Then sPrenom = "élisée"
        If sPrenom = "emmanuel" Then sPrenom = "émmanuel"
        If sPrenom = "emmanuelle" Then sPrenom = "émmanuelle"
        If sPrenom = "esperance" Then sPrenom = "espérance"
        If sPrenom = "evariste" Then sPrenom = "évariste"
        If sPrenom = "felicite" Then sPrenom = "félicité"
        If sPrenom = "frederic" Then sPrenom = "frédéric"
        If sPrenom = "frederique" Then sPrenom = "frédérique"
        If sPrenom = "gael" Then sPrenom = "gaël"
        If sPrenom = "gaelle" Then sPrenom = "gaëlle"
        If sPrenom = "guenaelle" Then sPrenom = "guénaëlle"
        If sPrenom = "gwenael" Then sPrenom = "gwenaël"
        If sPrenom = "gwenaelle" Then sPrenom = "gwenaëlle"
        If sPrenom = "heidi" Then sPrenom = "heïdi"
        If sPrenom = "irenee" Then sPrenom = "irenée"
        If sPrenom = "joel" Then sPrenom = "joël"
        If sPrenom = "joelle" Then sPrenom = "joëlle"
        If sPrenom = "jose" Then sPrenom = "josé"
        If sPrenom = "josee" Then sPrenom = "josée"
        If sPrenom = "josephe" Then sPrenom = "josèphe"
        If sPrenom = "judicael" Then sPrenom = "judicaël"
        If sPrenom = "leocadie" Then sPrenom = "léocadie"
        If sPrenom = "leonard" Then sPrenom = "léonard"
        If sPrenom = "leonce" Then sPrenom = "léonce"
        If sPrenom = "mae" Then sPrenom = "maé"
        If sPrenom = "mael" Then sPrenom = "maël"
        If sPrenom = "maelle" Then sPrenom = "maëlle"
        If sPrenom = "mederic" Then sPrenom = "médéric"
        If sPrenom = "medine" Then sPrenom = "médine"
        If sPrenom = "meryl" Then sPrenom = "méryl"
        If sPrenom = "michael" Then sPrenom = "michaël"
        If sPrenom = "michaelle" Then sPrenom = "michaëlle"
        If sPrenom = "michele" Then sPrenom = "michèle"
        If sPrenom = "mickaelle" Then sPrenom = "mickaëlle"
        If sPrenom = "nael" Then sPrenom = "naël"
        If sPrenom = "nais" Then sPrenom = "naïs"
        If sPrenom = "noe" Then sPrenom = "noé"
        If sPrenom = "noee" Then sPrenom = "noée"
        If sPrenom = "noel" Then sPrenom = "noël"
        If sPrenom = "noelle" Then sPrenom = "noëlle"
        If sPrenom = "raphael" Then sPrenom = "raphaël"
        If sPrenom = "raphaelle" Then sPrenom = "raphaëlle"
        If sPrenom = "rene" Then sPrenom = "rené"
        If sPrenom = "renee" Then sPrenom = "renée"
        If sPrenom = "sylvere" Then sPrenom = "sylvère"
        If sPrenom = "thais" Then sPrenom = "thaïs"
        If sPrenom = "theodore" Then sPrenom = "théodore"
        If sPrenom = "valere" Then sPrenom = "valère"
        If sPrenom = "valery" Then sPrenom = "valéry"
        If sPrenom = "yael" Then sPrenom = "yaël"
        If sPrenom = "yaelle" Then sPrenom = "yaëlle"

        Dim sPrenomHomophone = sPrenom
        If sPrenom = "aarone" Then sPrenomHomophone = "aaron"
        If sPrenom = "achrafe" Then sPrenomHomophone = "achraf"
        If sPrenom = "adame" Then sPrenomHomophone = "adam"
        If sPrenom = "adèle" Then sPrenomHomophone = "adel"
        If sPrenom = "adrianne" Then sPrenomHomophone = "adrian"
        If sPrenom = "aimée" Then sPrenomHomophone = "aimé"
        If sPrenom = "alexie" Then sPrenomHomophone = "alexis"
        If sPrenom = "alfrède" Then sPrenomHomophone = "alfred"
        If sPrenom = "amane" Then sPrenomHomophone = "aman"
        If sPrenom = "amaan" Then sPrenomHomophone = "aman"
        If sPrenom = "anaëlle" Then sPrenomHomophone = "anaël"
        If sPrenom = "andrée" Then sPrenomHomophone = "andré"
        If sPrenom = "andie" Then sPrenomHomophone = "andy"
        If sPrenom = "anne" Then sPrenomHomophone = "ann"
        If sPrenom = "arielle" Then sPrenomHomophone = "ariel"
        If sPrenom = "armelle" Then sPrenomHomophone = "armel"
        If sPrenom = "axelle" Then sPrenomHomophone = "axel"
        If sPrenom = "ayane" Then sPrenomHomophone = "ayan"
        If sPrenom = "aydane" Then sPrenomHomophone = "aydan"
        If sPrenom = "bayane" Then sPrenomHomophone = "bayan"
        If sPrenom = "bénédicte" Then sPrenomHomophone = "bénédict"
        If sPrenom = "carole" Then sPrenomHomophone = "carol"
        If sPrenom = "camerone" Then sPrenomHomophone = "cameron"
        If sPrenom = "charly" Then sPrenomHomophone = "charlie"
        If sPrenom = "cyrille" Then sPrenomHomophone = "cyril"
        If sPrenom = "dane" Then sPrenomHomophone = "dan"
        If sPrenom = "danie" Then sPrenomHomophone = "dani"
        If sPrenom = "danaëlle" Then sPrenomHomophone = "danaël"
        If sPrenom = "danielle" Then sPrenomHomophone = "daniel"
        If sPrenom = "danièle" Then sPrenomHomophone = "daniel"
        If sPrenom = "davide" Then sPrenomHomophone = "david"
        If sPrenom = "dilane" Then sPrenomHomophone = "dilan"
        If sPrenom = "doctrovee" Then sPrenomHomophone = "doctrove"
        If sPrenom = "dominic" Then sPrenomHomophone = "dominique"
        If sPrenom = "doriane" Then sPrenomHomophone = "dorian"
        If sPrenom = "dorianne" Then sPrenomHomophone = "dorian"
        If sPrenom = "émmanuelle" Then sPrenomHomophone = "émmanuel"
        If sPrenom = "edmée" Then sPrenomHomophone = "edmé"
        If sPrenom = "élie" Then sPrenomHomophone = "éli"
        If sPrenom = "éléonore" Then sPrenomHomophone = "éléonor"
        If sPrenom = "frédérique" Then sPrenomHomophone = "frédéric"
        If sPrenom = "gabrielle" Then sPrenomHomophone = "gabriel"
        If sPrenom = "gaëlle" Then sPrenomHomophone = "gaël"
        If sPrenom = "george" Then sPrenomHomophone = "georges"
        If sPrenom = "guénaëlle" Then sPrenomHomophone = "guénaël"
        If sPrenom = "gwenaëlle" Then sPrenomHomophone = "gwenaël"
        If sPrenom = "jessie" Then sPrenomHomophone = "jessy"
        If sPrenom = "joëlle" Then sPrenomHomophone = "joël"
        If sPrenom = "josée" Then sPrenomHomophone = "josé"
        If sPrenom = "josèphe" Then sPrenomHomophone = "joseph"
        If sPrenom = "karime" Then sPrenomHomophone = "karim"
        If sPrenom = "kiliane" Then sPrenomHomophone = "kilian"
        If sPrenom = "lilianne" Then sPrenomHomophone = "lilian"
        If sPrenom = "mahé" Then sPrenomHomophone = "maé"
        If sPrenom = "maëlle" Then sPrenomHomophone = "maël"
        If sPrenom = "mallorie" Then sPrenomHomophone = "mallory"
        If sPrenom = "malorie" Then sPrenomHomophone = "mallory"
        If sPrenom = "mallaurie" Then sPrenomHomophone = "mallory"
        If sPrenom = "malaurie" Then sPrenomHomophone = "mallory"
        If sPrenom = "manuelle" Then sPrenomHomophone = "manuel"
        If sPrenom = "marcelle" Then sPrenomHomophone = "marcel"
        If sPrenom = "michaëlle" Then sPrenomHomophone = "michaël"
        If sPrenom = "mickaëlle" Then sPrenomHomophone = "michaël"
        If sPrenom = "michèle" Then sPrenomHomophone = "michel"
        If sPrenom = "michelle" Then sPrenomHomophone = "michel"
        If sPrenom = "morgann" Then sPrenomHomophone = "morgan"
        If sPrenom = "morgane" Then sPrenomHomophone = "morgan"
        If sPrenom = "murielle" Then sPrenomHomophone = "muriel"
        If sPrenom = "nathanielle" Then sPrenomHomophone = "nathaniel"
        If sPrenom = "noée" Then sPrenomHomophone = "noé"
        If sPrenom = "noëlle" Then sPrenomHomophone = "noël"
        If sPrenom = "pascale" Then sPrenomHomophone = "pascal"
        If sPrenom = "paule" Then sPrenomHomophone = "paul"
        If sPrenom = "raphaëlle" Then sPrenomHomophone = "raphaël"
        If sPrenom = "romane" Then sPrenomHomophone = "roman"
        If sPrenom = "renée" Then sPrenomHomophone = "rené"
        If sPrenom = "swan" Then sPrenomHomophone = "swann"
        If sPrenom = "vivianne" Then sPrenomHomophone = "vivian"
        If sPrenom = "valérie" Then sPrenomHomophone = "valéry"
        If sPrenom = "yaëlle" Then sPrenomHomophone = "yaël"

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

    Private Sub EcrireFichierFiltre(asLignes$(), dico As DicoTri(Of String, clsPrenom))

        ' Génération d'un nouveau fichier csv filtré

        Dim sb As New StringBuilder
        Dim iNbLignes = 0
        For Each sLigne As String In asLignes

            iNbLignes += 1
            If iNbLignes = 1 Then sb.AppendLine(sLigne) : Continue For

            Dim prenom As New clsPrenom
            If Not bAnalyserPrenom(sLigne$, prenom) Then Continue For

            ConvertirPrenom(prenom)

            Dim sCle$ = prenom.sPrenom
            If dico.ContainsKey(sCle) Then
                Dim prenom0 = dico(sCle)
                If Not prenom0.bSelect Then Continue For
            Else
                Continue For
            End If

            Dim sLigneC$ = prenom.sCodeSexe & ";" & prenom.sPrenom & ";" & prenom.sAnnee & ";" & prenom.iNbOcc
            sb.AppendLine(sLigneC)

        Next

        Dim sCheminOut$ = Application.StartupPath & "\nat2019_.csv"
        EcrireFichier(sCheminOut, sb)

    End Sub

    Private Sub EcrireFichier(sChemin$, sb As StringBuilder)

        Using sw As New IO.StreamWriter(sChemin, append:=False, encoding:=Encoding.UTF8)
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

    Public Function sListerTxt$(lstTxt As List(Of String), Optional iNbMax% = 0)
        Dim sb As New StringBuilder("")
        Dim iNumOcc% = 0
        For Each sDef0 In lstTxt
            If sb.Length > 0 Then sb.Append(", ")
            sb.Append(sDef0)
            iNumOcc += 1
            If iNbMax > 0 Then
                If iNumOcc >= iNbMax Then sb.Append("..") : Exit For
            End If
        Next
        Return sb.ToString
    End Function

End Class