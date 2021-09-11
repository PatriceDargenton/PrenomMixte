
' modPrenom.vb : module pour le traitement des prénoms mixtes et prénoms fréquents
' ------------

' Conventions de nommage des variables :
' ------------------------------------
' b pour Boolean (booléen vrai ou faux)
' i pour Integer : % (en VB .Net, l'entier a la capacité du VB6.Long)
' l pour Long : &
' r pour nombre Réel (Single!, Double# ou Decimal : D)
' s pour String : $
' c pour Char ou Byte
' d pour Date
' a pour Array (tableau) : ()
' o pour Object : objet instancié localement
' refX pour reference à un objet X préexistant qui n'est pas sensé être fermé
' m_ pour variable Membre de la classe ou de la feuille (Form)
'  (mais pas pour les constantes)
' frm pour Form
' cls pour Classe
' mod pour Module
' ...
' ------------------------------------

Imports System.Text

Public Module modPrenom

#If DEBUG Then
    Public Const bDebug As Boolean = True
    Public Const bRelease As Boolean = False
#Else
    Public Const bDebug As Boolean = False
    Public Const bRelease As Boolean = True
#End If

    Public Const sTitreAppli$ = "Prénom mixte"
    Public Const sDateVersionAppli$ = "11/09/2021"

    Public ReadOnly sVersionAppli$ =
        My.Application.Info.Version.Major & "." &
        My.Application.Info.Version.Minor &
        My.Application.Info.Version.Build

    Public Const sFichierPrenomsInsee$ = "nat2020.csv"
    Public Const sFichierPrenomsInseeCorrige$ = "nat2020_corrige.csv"
    Public Const sFichierPrenomsInseeZip$ = "nat2020_csv.zip"
    Const iDateFin% = 2020
    Const iDateMinExport% = 1900
    Const iDateMaxExport% = 2020

    ' Seuils de fréquence relative min.
    'Const rSeuilFreqRel# = 0.001 ' 0.1% (par exemple 0.1% de masc. et 99.9% de fém.)
    'Const sFormatFreqRel$ = "0.0%" ' 0.1%
    Const rSeuilFreqRel# = 0.01 ' 1% (par exemple 1% de masc. et 99% de fém.)
    'Const rSeuilFreqRel# = 0.02 ' 2% (par exemple 2% de masc. et 98% de fém.)
    Const sFormatFreqRel$ = "0%"    '   1%

    Const rSeuilFreqRelPrenomsEpicenes# = rSeuilFreqRel

    ' Le filtre n'est pas programmé, il faut rajouter la condition, le cas échéant :
    Const rSeuilFreqRelPrenomsFrequents# = 0
    Const rSeuilFreqRelPrenomsHomophones# = 0
    Const rSeuilFreqRelPrenomsSimilaires# = 0

    ' Fréquence relative minimale de la variante (homophone ou spécifiquement genrée)
    '  par rapport à la somme des variantes
    Const sFormatFreqRelVariante$ = "0%"
    Const rSeuilFreqRelVariante# = 0.01 ' 1%
    'Const sFormatFreqRelVariante$ = "0.0%"
    'Const rSeuilFreqRelVariante# = 0.001 ' 0.1%

    Const iSeuilMinPrenomsFrequents% = 50000 ' 4000 minimum pour une page wiki (sinon ça plante)
    ' Nombre minimal d'occurrences du prénom sur plus d'un siècle
    Const iSeuilMinPrenomsEpicenes% = 2000
    Const iSeuilMinPrenomsHomophones% = 20000
    Const iSeuilMinPrenomsSimilaires% = 20000
    Const iSeuilMinPrenomsUnigenre% = 10000

    ' Seuil min. pour la détection des prénoms homophones potentiels
    Const iSeuilMinPrenomsHomophonesPotentiels% = 10000
    Const iNbLignesMaxPrenoms% = 0 ' 32346 prénoms en tout (reste quelques accents à corriger)

    Const sGrasMD$ = "**"
    Const sGrasWiki$ = "**"
    'Const sGrasWiki$ = "'''" ' Ancienne version de MédiaWiki (1.26.3)
    Const sItaliqueMD$ = "*"
    Const sItaliqueWiki$ = "''"
    Const sItaliqueGrasMD$ = "***"
    Const sItaliqueGrasWiki$ = "'''''"
    Const sSautDeLigneMD$ = "<br>"

    Const sFormatFreq$ = "0.000%"

    Public Sub AnalyserPrenoms(sDossierAppli$,
            Optional bExporter As Boolean = False, Optional bTest As Boolean = False)

        Dim sChemin = sDossierAppli & "\" & sFichierPrenomsInsee
        If Not IO.File.Exists(sChemin) Then
            MsgBox("Veuillez télécharger " & sFichierPrenomsInseeZip & " !" & vbLf & sChemin,
                MsgBoxStyle.Exclamation, sTitreAppli)
            Exit Sub
        End If
        Dim asLignes$() = IO.File.ReadAllLines(sChemin, Encoding.UTF8)
        If IsNothing(asLignes) Then Exit Sub

        Dim sCheminCorrectionsPrenoms$ = sDossierAppli & "\CorrectionsPrenoms.csv"
        Dim dicoCorrectionsPrenoms = LireFichier(sCheminCorrectionsPrenoms)
        Dim dicoCorrectionsPrenomsUtil As New DicoTri(Of String, String)

        ' Vérifier si le fichier de correction des prénoms mixtes corrige bien uniquement les accents
        For Each kvp In dicoCorrectionsPrenoms
            Dim sPrenomOrig$ = kvp.Key
            sPrenomOrig = sEnleverAccents(sPrenomOrig) ' Pour pouvoir corriger aussi : jérôme en jérome
            Dim sPrenomCorrige$ = kvp.Value
            Dim sVerif$ = sEnleverAccents(sPrenomCorrige)
            If sVerif <> sPrenomOrig Then
                MsgBox("Erreur de correction d'accent : " & sVerif & " <> " & sPrenomOrig,
                    MsgBoxStyle.Exclamation, sTitreAppli)
            End If
        Next

        Dim sCheminDefPrenomsMixtesHomophones$ = sDossierAppli &
            "\DefinitionsPrenomsMixtesHomophones.csv"
        Dim dicoDefinitionsPrenomsMixtesHomophones = LireFichier(sCheminDefPrenomsMixtesHomophones)
        Dim dicoDefinitionsPrenomsMixtesHomophonesUtil As New DicoTri(Of String, String)

        Dim sCheminDefPrenomsSimilaires$ = sDossierAppli &
            "\DefinitionsPrenomsSimilaires.csv"
        Dim dicoDefinitionsPrenomsSimilaires = LireFichier(sCheminDefPrenomsSimilaires)
        Dim dicoDefinitionsPrenomsSimilairesUtil As New DicoTri(Of String, String)

        ' Ajouter les définitions de prénoms mixtes homophones aux
        '  définitions de prénoms similaires
        ' (le prénom pivot doit être le même, le cas échéant, sinon une alerte sera générée)
        For Each kvp In dicoDefinitionsPrenomsMixtesHomophones
            If Not dicoDefinitionsPrenomsSimilaires.ContainsKey(kvp.Key) Then
                dicoDefinitionsPrenomsSimilaires.Add(kvp.Key, kvp.Value)
            End If
        Next

        Dim dicoE As New DicoTri(Of String, clsPrenom) ' épicène
        Dim dicoH As New DicoTri(Of String, clsPrenom) ' homophone
        Dim dicoS As New DicoTri(Of String, clsPrenom) ' similaire

        Dim iNbLignes% = 0
        Dim iNbLignesOk% = 0
        Dim iNbPrenomsTot% = 0
        Dim iNbPrenomsTotOk% = 0
        Dim iNbPrenomsIgnores% = 0
        Dim iNbPrenomsIgnoresDate% = 0

        AnalyseFichierINSEE(asLignes,
            dicoCorrectionsPrenoms,
            dicoCorrectionsPrenomsUtil,
            dicoDefinitionsPrenomsMixtesHomophones,
            dicoDefinitionsPrenomsMixtesHomophonesUtil,
            dicoDefinitionsPrenomsSimilaires,
            dicoDefinitionsPrenomsSimilairesUtil,
            dicoE, dicoH, dicoS, bTest,
            iNbLignes, iNbLignesOk, iNbPrenomsTot, iNbPrenomsTotOk,
            iNbPrenomsIgnores, iNbPrenomsIgnoresDate)

        Dim sbCPMD As New StringBuilder
        Dim sbHPMD As New StringBuilder
        Dim sbSPMD As New StringBuilder

        If bTest Then GoTo Export

        DetectionAnomalies(sDossierAppli,
            dicoCorrectionsPrenoms,
            dicoDefinitionsPrenomsMixtesHomophones,
            dicoDefinitionsPrenomsSimilaires,
            dicoE, sbCPMD, sbHPMD, sbSPMD)

        FiltrerPrenomMixteEpicene(dicoE, iNbPrenomsTot,
            iSeuilMinPrenomsEpicenes, rSeuilFreqRelPrenomsEpicenes,
            iNbPrenomsTotOk, iNbPrenomsIgnores, iNbPrenomsIgnoresDate)

        FiltrerPrenomMixteHomophone(dicoH, dicoE, iNbPrenomsTot,
            iNbPrenomsTotOk, iNbPrenomsIgnores, iNbPrenomsIgnoresDate)

        FiltrerPrenomSimilaire(dicoS, dicoE, dicoH, iNbPrenomsTot,
            iNbPrenomsTotOk, iNbPrenomsIgnores, iNbPrenomsIgnoresDate)

        FiltrerPrenomUnigenre(dicoE, dicoH, dicoS, iNbPrenomsTot,
            iNbPrenomsTotOk, iNbPrenomsIgnores, iNbPrenomsIgnoresDate)

Export:
        If bExporter Then
            Exporter(sDossierAppli, asLignes, dicoE,
                dicoCorrectionsPrenoms,
                dicoCorrectionsPrenomsUtil,
                dicoDefinitionsPrenomsMixtesHomophones,
                dicoDefinitionsPrenomsMixtesHomophonesUtil,
                dicoDefinitionsPrenomsSimilaires,
                dicoDefinitionsPrenomsSimilairesUtil, bTest)
            GoTo Fin
        End If

        AfficherSyntheses(sDossierAppli,
            dicoCorrectionsPrenoms,
            dicoCorrectionsPrenomsUtil,
            dicoDefinitionsPrenomsMixtesHomophones,
            dicoDefinitionsPrenomsMixtesHomophonesUtil,
            dicoDefinitionsPrenomsSimilaires,
            dicoDefinitionsPrenomsSimilairesUtil,
            dicoE, dicoH, dicoS,
            iNbLignes, iNbLignesOk, iNbPrenomsTot, iNbPrenomsTotOk,
            iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            sbCPMD, sbHPMD, sbSPMD,
            sCheminCorrectionsPrenoms,
            sCheminDefPrenomsMixtesHomophones, sCheminDefPrenomsSimilaires)

        AnalysePrenomsSimilaires(sDossierAppli, dicoDefinitionsPrenomsSimilaires)

Fin:
        If Not bTest Then MsgBox("Terminé !", MsgBoxStyle.Information, sTitreAppli)

    End Sub

    Private Sub AnalyseFichierINSEE(asLignes$(),
            dicoCorrectionsPrenoms As DicoTri(Of String, String),
            dicoCorrectionsPrenomsUtil As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophones As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophonesUtil As DicoTri(Of String, String),
            dicoDefinitionsPrenomsSimilaires As DicoTri(Of String, String),
            dicoDefinitionsPrenomsSimilairesUtil As DicoTri(Of String, String),
            dicoE As DicoTri(Of String, clsPrenom),
            dicoH As DicoTri(Of String, clsPrenom),
            dicoS As DicoTri(Of String, clsPrenom),
            bTest As Boolean,
            ByRef iNbLignes%, ByRef iNbLignesOk%,
            ByRef iNbPrenomsTot%, ByRef iNbPrenomsTotOk%,
            ByRef iNbPrenomsIgnores%, ByRef iNbPrenomsIgnoresDate%)

        For Each sLigne As String In asLignes

            iNbLignes += 1
            If iNbLignes = 1 Then Continue For ' Entête

            Dim prenom As New clsPrenom
            If Not bAnalyserPrenom(sLigne$, prenom,
                dicoCorrectionsPrenoms,
                dicoCorrectionsPrenomsUtil,
                dicoDefinitionsPrenomsMixtesHomophones,
                dicoDefinitionsPrenomsMixtesHomophonesUtil,
                dicoDefinitionsPrenomsSimilaires,
                dicoDefinitionsPrenomsSimilairesUtil) Then Continue For
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

            prenom.rAnneeTot = prenom.iAnnee * prenom.iNbOcc
            If prenom.bMasc Then prenom.rAnneeTotMasc = prenom.iAnnee * prenom.iNbOcc
            If prenom.bFem Then prenom.rAnneeTotFem = prenom.iAnnee * prenom.iNbOcc

            Dim sCle$ = prenom.sPrenom
            If dicoE.ContainsKey(sCle) Then
                Dim prenom0 = dicoE(sCle)
                prenom0.Ajouter(prenom)
            Else
                dicoE.Add(sCle, prenom)
            End If

            If bTest Then Continue For
            'If bDebug Then Continue For

            ' Dico des prénoms homophones
            Dim prenomH = prenom.Clone() ' Il faut faire une copie pour que l'objet soit distinct
            If Not prenomH.dicoVariantesH.ContainsKey(prenomH.sPrenom) Then
                prenomH.dicoVariantesH.Add(prenomH.sPrenom, prenom)
            End If
            Dim sCleH$ = prenomH.sPrenomHomophone
            If dicoH.ContainsKey(sCleH) Then
                Dim prenom0 = dicoH(sCleH)
                prenom0.Ajouter(prenom)
                For Each kvp In prenomH.dicoVariantesH
                    If Not prenom0.dicoVariantesH.ContainsKey(kvp.Key) Then
                        prenom0.dicoVariantesH.Add(kvp.Key, prenom)
                    End If
                Next
            Else
                dicoH.Add(sCleH, prenomH)
            End If

            ' Dico des prénoms similaires
            Dim prenomS = prenom.Clone()
            If Not prenomS.dicoVariantesS.ContainsKey(prenomS.sPrenom) Then
                prenomS.dicoVariantesS.Add(prenomS.sPrenom, prenom)
            End If
            Dim sCleS$ = prenomS.sPrenomSimilaire
            If dicoS.ContainsKey(sCleS) Then
                Dim prenom0 = dicoS(sCleS)
                prenom0.Ajouter(prenom)
                For Each kvp In prenomS.dicoVariantesS
                    If Not prenom0.dicoVariantesS.ContainsKey(kvp.Key) Then
                        prenom0.dicoVariantesS.Add(kvp.Key, prenom)
                    End If
                Next
            Else
                dicoS.Add(sCleS, prenomS)
            End If

        Next

    End Sub

    Private Sub DetectionAnomalies(sDossierAppli$,
        dicoCorrectionsPrenoms As DicoTri(Of String, String),
        dicoDefinitionsPrenomsMixtesHomophones As DicoTri(Of String, String),
        dicoDefinitionsPrenomsSimilaires As DicoTri(Of String, String),
        dicoE As DicoTri(Of String, clsPrenom),
        ByRef sbCPMD As StringBuilder, ByRef sbHPMD As StringBuilder, ByRef sbSPMD As StringBuilder)

        Dim sdCP As New SortedDictionary(Of String, String) ' Correction de prénoms potentiels
        Dim sdHP As New SortedDictionary(Of String, String) ' Prénoms homophones potentiels
        Dim sdSP As New SortedDictionary(Of String, String) ' Prénoms similaires potentiels

        For Each prenom In dicoE.Trier("")

            ' Détection des corrections d'accent potentielles restantes
            Dim sPrenomSansAccent$ = sEnleverAccents(prenom.sPrenom, bMinuscule:=False)
            If sPrenomSansAccent <> prenom.sPrenom Then
                If dicoE.ContainsKey(sPrenomSansAccent) AndAlso
                   Not dicoCorrectionsPrenoms.ContainsKey(sPrenomSansAccent) AndAlso
                    Not sdCP.ContainsKey(sPrenomSansAccent) Then
                    sdCP.Add(sPrenomSansAccent, prenom.sPrenom)
                End If
            End If

            ' Détection des prénoms homophones potentiels restants
            Dim sPrenomF1$ = prenom.sPrenom & "e" ' Ex.: Renée : René
            AjoutPrenomsHomophonesPotentielsRestants(
                sPrenomF1, prenom.sPrenom, dicoE,
                dicoDefinitionsPrenomsMixtesHomophones, sdHP)
            Dim sPrenomF2$ = prenom.sPrenom & "le" ' Ex.: Gabrielle : Gabriel
            AjoutPrenomsHomophonesPotentielsRestants(
                sPrenomF2, prenom.sPrenom, dicoE,
                dicoDefinitionsPrenomsMixtesHomophones, sdHP)

            ' Détection des prénoms similaires potentiels restants
            Dim sPrenomF3$ = prenom.sPrenom & "tte" ' Ex.: Antoinette : Antoine
            AjoutPrenomsSimilairesPotentielsRestants(
                sPrenomF3, prenom.sPrenom, dicoE,
                dicoDefinitionsPrenomsSimilaires, sdSP)

            Dim sPrenomF4$ = prenom.sPrenom & "ne" ' Ex.: Fabien : Fabienne
            AjoutPrenomsSimilairesPotentielsRestants(
                sPrenomF4, prenom.sPrenom, dicoE,
                dicoDefinitionsPrenomsSimilaires, sdSP)

            Dim sPrenomF5$ = prenom.sPrenom & "ia" ' Ex.: Victor : Victoria
            AjoutPrenomsSimilairesPotentielsRestants(
                sPrenomF5, prenom.sPrenom, dicoE,
                dicoDefinitionsPrenomsSimilaires, sdSP)

        Next

        Dim sbCP As New StringBuilder("Liste des corrections potentielles d'accent")
        sbCPMD = New StringBuilder("Liste des corrections potentielles d'accent")
        sbCP.AppendLine()
        sbCPMD.AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine()
        For Each kvp In sdCP
            Dim sPrenom = kvp.Key
            Dim sPrenomC = kvp.Value
            sbCP.AppendLine(sPrenomC.ToLower & ";" & sPrenom.ToLower)
            sbCPMD.AppendLine(sPrenom & " : " & sPrenomC).AppendLine()
        Next
        Dim sCheminCP$ = sDossierAppli & "\CorrectionsPotentielles.txt"
        EcrireFichier(sCheminCP, sbCP)

        Dim sbHP As New StringBuilder("Liste des prénoms homophones potentiels restants")
        sbHPMD = New StringBuilder("Liste des prénoms homophones potentiels restants")
        sbHP.AppendLine()
        sbHPMD.AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine()
        For Each kvp In sdHP
            Dim sPrenom = kvp.Key
            Dim sPrenomC = kvp.Value
            Dim iNbOcc% = 0, iNbOccC% = 0, iNbOccM% = 0
            If dicoE.ContainsKey(sPrenom) Then iNbOcc = dicoE(sPrenom).iNbOcc
            If dicoE.ContainsKey(sPrenomC) Then iNbOccC = dicoE(sPrenomC).iNbOcc
            iNbOccM = Math.Max(iNbOcc, iNbOccC)
            If iNbOccM < iSeuilMinPrenomsHomophonesPotentiels Then Continue For
            sbHP.AppendLine(sPrenomC.ToLower & ";" & sPrenom.ToLower)
            sbHPMD.AppendLine(sPrenom & " : " & sPrenomC).AppendLine()
        Next
        Dim sCheminHP$ = sDossierAppli & "\PrenomsMixtesHomophonesPotentiels.txt"
        EcrireFichier(sCheminHP, sbHP)

        Dim sbSP As New StringBuilder("Liste des prénoms similaires potentiels restants")
        sbSPMD = New StringBuilder("Liste des prénoms similaires potentiels restants")
        sbSP.AppendLine()
        sbSPMD.AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine()
        For Each kvp In sdSP
            Dim sPrenom = kvp.Key
            Dim sPrenomC = kvp.Value
            sbSP.AppendLine(sPrenomC.ToLower & ";" & sPrenom.ToLower)
            sbSPMD.AppendLine(sPrenom & " : " & sPrenomC).AppendLine()
        Next
        Dim sCheminSP$ = sDossierAppli & "\PrenomsSimilairesPotentielsRestants.txt"
        EcrireFichier(sCheminSP, sbSP)

    End Sub

    Private Sub AjoutPrenomsHomophonesPotentielsRestants(
            sPrenomPot$, sPrenom$,
            dicoE As DicoTri(Of String, clsPrenom),
            dicoDefinitionsPrenomsMixtesHomophones As DicoTri(Of String, String),
            sdHO As SortedDictionary(Of String, String))

        Dim sPrenomPotMin = sPrenomPot.ToLower
        If dicoE.ContainsKey(sPrenomPot) AndAlso
           Not dicoDefinitionsPrenomsMixtesHomophones.ContainsKey(sPrenomPotMin) AndAlso
           Not dicoDefinitionsPrenomsMixtesHomophones.ContainsValue(sPrenomPotMin) AndAlso
           Not sdHO.ContainsKey(sPrenomPot) Then
            sdHO.Add(sPrenomPot, sPrenom)
        End If

    End Sub

    Private Sub AjoutPrenomsSimilairesPotentielsRestants(
            sPrenomPot$, sPrenom$,
            dicoE As DicoTri(Of String, clsPrenom),
            dicoDefinitionsPrenomsSimilaires As DicoTri(Of String, String),
            sdSP As SortedDictionary(Of String, String))

        Dim sPrenomPotMin = sPrenomPot.ToLower
        If dicoE.ContainsKey(sPrenomPot) AndAlso
           Not dicoDefinitionsPrenomsSimilaires.ContainsKey(sPrenomPotMin) AndAlso
           Not dicoDefinitionsPrenomsSimilaires.ContainsValue(sPrenomPotMin) AndAlso
           Not sdSP.ContainsKey(sPrenomPot) Then
            sdSP.Add(sPrenomPot, sPrenom)
        End If

    End Sub

    Private Sub AnalysePrenomsSimilaires(sDossierAppli$,
            dicoDefinitionsPrenomsSimilaires As DicoTri(Of String, String))

        Dim sbPS As New StringBuilder(
            "Dictionnaire des prénoms similaires + mixtes homophones")
        sbPS.AppendLine()

        Dim asTable = asTrierDicoStringString(dicoDefinitionsPrenomsSimilaires)
        For Each sPrenom In asTable
            sbPS.AppendLine(dicoDefinitionsPrenomsSimilaires(sPrenom).ToLower & ";" & sPrenom.ToLower)
        Next

        ' Vérifier que les clés sont bien dans le même ordre lors de la fusion des 2 dico
        ' -> Choisir la même clé de regroupement dans ces 2 dico pour résoudre ces problèmes

        ' Exemple, si on a :
        ' ----------------

        ' DefinitionsPrenomsMixtesHomophones.csv :
        ' pascal;pascale
        ' pascal;pasquale

        ' DefinitionsPrenomsSimilaires.csv :
        ' pascaline;pascal

        ' Alors pascal est signalé comme étant une clé inversée dans le dico fusionné
        ' -> Changer en pascal;pascaline dans DefinitionsPrenomsSimilaires.csv
        ' ----------------

        Dim hsCles As New HashSet(Of String)
        For Each kvp In dicoDefinitionsPrenomsSimilaires
            hsCles.Add(kvp.Key)
        Next
        For Each kvp In dicoDefinitionsPrenomsSimilaires
            If hsCles.Contains(kvp.Value) Then
                MsgBox("Fusion des prénoms similaires + mixtes homophones :" & vbLf &
                    "clé inversée : " & kvp.Value, MsgBoxStyle.Exclamation, sTitreAppli)
                sbPS.AppendLine(
                    "Fusion des prénoms similaires + mixtes homophones : clé inversée : " &
                    kvp.Value)
            End If
        Next

        Dim sCheminPS$ = sDossierAppli & "\PrenomsSimilairesEtMixtesHomophones.txt"
        EcrireFichier(sCheminPS, sbPS)

    End Sub

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

        Dim iVerif% = iNbPrenomsVerif + iNbPrenomsIgnores + iNbPrenomsIgnoresDate
        If iVerif <> iNbPrenomsTot Then
            Debug.WriteLine(sFormaterNum(iNbPrenomsVerif) & " <> " & sFormaterNum(iNbPrenomsTotOk))
            MsgBox("Décompte faux : " & iVerif & " <> " & iNbPrenomsTot,
                MsgBoxStyle.Exclamation, sTitreAppli)
        End If
        Dim iVerifMF% = iNbPrenomsVerifMF + iNbPrenomsIgnores + iNbPrenomsIgnoresDate
        If iVerifMF <> iNbPrenomsTot Then
            Debug.WriteLine(sFormaterNum(iVerifMF) & " <> " & sFormaterNum(iNbPrenomsTot))
            MsgBox("Décompte faux : " & iVerifMF & " <> " & iNbPrenomsTot,
                MsgBoxStyle.Exclamation, sTitreAppli)
        End If

    End Sub

    Private Sub FiltrerPrenomMixteHomophone(
            dicoH As DicoTri(Of String, clsPrenom),
            dicoE As DicoTri(Of String, clsPrenom),
            iNbPrenomsTot%, iNbPrenomsTotOk%, iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%)

        Dim iNbPrenomsVerif% = 0
        Dim iNbPrenomsVerifMF% = 0
        Dim aPrenomsH = dicoH.Trier("")
        For Each prenom In aPrenomsH

            If dicoE.ContainsKey(prenom.sPrenom) Then
                Dim prenom0 = dicoE(prenom.sPrenom)
                If prenom0.bMixteEpicene Then prenom.bMixteEpicene = prenom0.bMixteEpicene
            End If

            prenom.Calculer(iNbPrenomsTot)
            iNbPrenomsVerif += prenom.iNbOcc
            iNbPrenomsVerifMF += prenom.iNbOccMasc + prenom.iNbOccFem

            If prenom.dicoVariantesH.Count > 1 Then
                prenom.bMixteHomophone = True
                ' Marquer aussi l'original pour l'export
                'If dicoE.ContainsKey(prenom.sPrenomHomophone) Then
                '    Dim prenomH = dicoE(prenom.sPrenomHomophone)
                '    prenomH.bMixteHomophone = True
                'End If
                ' Marquer l'ensemble des variantes directement dans le dico épicène
                For Each prenom0 In prenom.dicoVariantesH.Trier()
                    Dim sPrenomE = prenom0.sPrenom
                    If dicoE.ContainsKey(sPrenomE) Then
                        Dim prenom1 = dicoE(sPrenomE)
                        prenom1.bMixteHomophone = True
                    End If
                Next
            End If

        Next

        Dim iVerif% = iNbPrenomsVerif + iNbPrenomsIgnores + iNbPrenomsIgnoresDate
        If iVerif <> iNbPrenomsTot Then
            Debug.WriteLine(sFormaterNum(iNbPrenomsVerif) & " <> " & sFormaterNum(iNbPrenomsTotOk))
            MsgBox("Décompte faux : " & iVerif & " <> " & iNbPrenomsTot,
                MsgBoxStyle.Exclamation, sTitreAppli)
        End If
        Dim iVerifMF% = iNbPrenomsVerifMF + iNbPrenomsIgnores + iNbPrenomsIgnoresDate
        If iVerifMF <> iNbPrenomsTot Then
            Debug.WriteLine(sFormaterNum(iVerifMF) & " <> " & sFormaterNum(iNbPrenomsTot))
            MsgBox("Décompte faux : " & iVerifMF & " <> " & iNbPrenomsTot,
                MsgBoxStyle.Exclamation, sTitreAppli)
        End If

        ' Détecter et retirer les variantes sous le seuil de fréquence relative min.

        ' Détecter (calculer la fréquence avec les variantes)
        For Each prenom In aPrenomsH
            If prenom.dicoVariantesH.Count <= 1 Then Continue For
            For Each kvp In prenom.dicoVariantesH
                Dim prenomH = kvp.Value
                prenomH.rFreqRelativeVarianteH = prenomH.iNbOcc / prenom.iNbOcc
                If prenomH.rFreqRelativeVarianteH < rSeuilFreqRelVariante Then
                    prenomH.bVarianteDecompteeH = True
                End If
            Next
        Next

        ' Retirer les variantes trop minoritaires
        ' Autre solution : cumuler sur une variante "Autres", si on ne veut pas les décompter
        Dim lstPrenomsRetires As New List(Of clsPrenom)
        For Each prenom In aPrenomsH
            If prenom.dicoVariantesH.Count <= 1 Then Continue For
            Dim lst As New List(Of String)
            For Each kvp In prenom.dicoVariantesH
                Dim prenomH = kvp.Value
                If prenomH.bVarianteDecompteeH Then
                    prenom.Retirer(prenomH)
                    prenom.Calculer(iNbPrenomsTot)
                    lst.Add(kvp.Key)
                    lstPrenomsRetires.Add(prenomH)
                End If
            Next
            For Each sCle In lst
                prenom.dicoVariantesH.Remove(sCle)
            Next
            If prenom.dicoVariantesH.Count <= 1 Then prenom.bMixteHomophone = False
        Next

        ' Vérifier le nouveau calcul
        iNbPrenomsVerif = 0
        iNbPrenomsVerifMF = 0
        For Each prenom In aPrenomsH
            iNbPrenomsVerif += prenom.iNbOcc
            iNbPrenomsVerifMF += prenom.iNbOccMasc + prenom.iNbOccFem
        Next
        For Each prenom In lstPrenomsRetires
            iNbPrenomsVerif += prenom.iNbOcc
            iNbPrenomsVerifMF += prenom.iNbOccMasc + prenom.iNbOccFem
        Next

        iVerif = iNbPrenomsVerif + iNbPrenomsIgnores + iNbPrenomsIgnoresDate
        If iVerif <> iNbPrenomsTot Then
            Debug.WriteLine(sFormaterNum(iNbPrenomsVerif) & " <> " & sFormaterNum(iNbPrenomsTotOk))
            MsgBox("Décompte faux : " & iVerif & " <> " & iNbPrenomsTot,
                MsgBoxStyle.Exclamation, sTitreAppli)
        End If
        iVerifMF = iNbPrenomsVerifMF + iNbPrenomsIgnores + iNbPrenomsIgnoresDate
        If iVerifMF <> iNbPrenomsTot Then
            Debug.WriteLine(sFormaterNum(iVerifMF) & " <> " & sFormaterNum(iNbPrenomsTot))
            MsgBox("Décompte faux : " & iVerifMF & " <> " & iNbPrenomsTot,
                MsgBoxStyle.Exclamation, sTitreAppli)
        End If

    End Sub

    Private Sub FiltrerPrenomSimilaire(
            dicoS As DicoTri(Of String, clsPrenom),
            dicoE As DicoTri(Of String, clsPrenom),
            dicoH As DicoTri(Of String, clsPrenom),
            iNbPrenomsTot%, iNbPrenomsTotOk%, iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%)

        Dim iNbPrenomsVerif% = 0
        Dim iNbPrenomsVerifMF% = 0
        Dim aPrenomsS = dicoS.Trier("")
        For Each prenom In aPrenomsS

            If dicoE.ContainsKey(prenom.sPrenom) Then
                Dim prenom0 = dicoE(prenom.sPrenom)
                If prenom0.bMixteEpicene Then prenom.bMixteEpicene = prenom0.bMixteEpicene
            End If

            If dicoH.ContainsKey(prenom.sPrenomHomophone) Then
                Dim prenom0 = dicoH(prenom.sPrenomHomophone)
                prenom.bMixteHomophone = prenom0.bMixteHomophone
            End If

            prenom.Calculer(iNbPrenomsTot)
            iNbPrenomsVerif += prenom.iNbOcc
            iNbPrenomsVerifMF += prenom.iNbOccMasc + prenom.iNbOccFem

            'If prenom.dicoVariantesS.Count > 1 Then prenom.bSimilaire = True
            ' Définition de similaire : au moins une variante non homophone ni épicène
            If prenom.dicoVariantesS.Count > 1 Then
                Dim bTousbMixtesHOuE = True
                For Each prenom0 In prenom.dicoVariantesS.Trier()
                    Dim sPrenomE = prenom0.sPrenom
                    If dicoE.ContainsKey(sPrenomE) Then
                        Dim prenom1 = dicoE(sPrenomE)
                        If Not prenom1.bMixteHomophone AndAlso
                           Not prenom1.bMixteEpicene Then bTousbMixtesHOuE = False : Exit For
                    End If
                Next
                If Not bTousbMixtesHOuE Then
                    prenom.bSimilaire = True
                    ' Marquer l'ensemble des variantes directement, dans le dico épicène
                    For Each prenom0 In prenom.dicoVariantesS.Trier()
                        Dim sPrenomE = prenom0.sPrenom
                        If dicoE.ContainsKey(sPrenomE) Then
                            Dim prenom1 = dicoE(sPrenomE)
                            prenom1.bSimilaire = True
                        End If
                    Next
                End If
            End If

        Next

        Dim iVerif% = iNbPrenomsVerif + iNbPrenomsIgnores + iNbPrenomsIgnoresDate
        If iVerif <> iNbPrenomsTot Then
            Debug.WriteLine(sFormaterNum(iNbPrenomsVerif) & " <> " & sFormaterNum(iNbPrenomsTotOk))
            MsgBox("Décompte faux : " & iVerif & " <> " & iNbPrenomsTot,
                MsgBoxStyle.Exclamation, sTitreAppli)
        End If
        Dim iVerifMF% = iNbPrenomsVerifMF + iNbPrenomsIgnores + iNbPrenomsIgnoresDate
        If iVerifMF <> iNbPrenomsTot Then
            Debug.WriteLine(sFormaterNum(iVerifMF) & " <> " & sFormaterNum(iNbPrenomsTot))
            MsgBox("Décompte faux : " & iVerifMF & " <> " & iNbPrenomsTot,
                MsgBoxStyle.Exclamation, sTitreAppli)
        End If

        ' Détecter et retirer les variantes sous le seuil de fréquence relative min.

        ' Détecter (calculer la fréquence avec les variantes)
        For Each prenom In aPrenomsS
            If prenom.dicoVariantesS.Count <= 1 Then Continue For
            For Each kvp In prenom.dicoVariantesS
                Dim prenomH = kvp.Value
                prenomH.rFreqRelativeVarianteS = prenomH.iNbOcc / prenom.iNbOcc
                If prenomH.rFreqRelativeVarianteS < rSeuilFreqRelVariante Then
                    prenomH.bVarianteDecompteeS = True
                End If
            Next
        Next

        ' Retirer les variantes trop minoritaires
        Dim lstPrenomsRetires As New List(Of clsPrenom)
        For Each prenom In aPrenomsS
            If prenom.dicoVariantesS.Count <= 1 Then Continue For
            Dim lst As New List(Of String)
            For Each kvp In prenom.dicoVariantesS
                Dim prenomH = kvp.Value
                If prenomH.bVarianteDecompteeS Then
                    prenom.Retirer(prenomH)
                    prenom.Calculer(iNbPrenomsTot)
                    lst.Add(kvp.Key)
                    lstPrenomsRetires.Add(prenomH)
                End If
            Next
            For Each sCle In lst
                prenom.dicoVariantesS.Remove(sCle)
            Next
            If prenom.dicoVariantesS.Count <= 1 Then prenom.bSimilaire = False
        Next

        ' Vérifier le nouveau calcul
        iNbPrenomsVerif = 0
        iNbPrenomsVerifMF = 0
        For Each prenom In aPrenomsS
            iNbPrenomsVerif += prenom.iNbOcc
            iNbPrenomsVerifMF += prenom.iNbOccMasc + prenom.iNbOccFem
        Next
        For Each prenom In lstPrenomsRetires
            iNbPrenomsVerif += prenom.iNbOcc
            iNbPrenomsVerifMF += prenom.iNbOccMasc + prenom.iNbOccFem
        Next

        iVerif = iNbPrenomsVerif + iNbPrenomsIgnores + iNbPrenomsIgnoresDate
        If iVerif <> iNbPrenomsTot Then
            Debug.WriteLine(sFormaterNum(iNbPrenomsVerif) & " <> " & sFormaterNum(iNbPrenomsTotOk))
            MsgBox("Décompte faux : " & iVerif & " <> " & iNbPrenomsTot,
                MsgBoxStyle.Exclamation, sTitreAppli)
        End If
        iVerifMF = iNbPrenomsVerifMF + iNbPrenomsIgnores + iNbPrenomsIgnoresDate
        If iVerifMF <> iNbPrenomsTot Then
            Debug.WriteLine(sFormaterNum(iVerifMF) & " <> " & sFormaterNum(iNbPrenomsTot))
            MsgBox("Décompte faux : " & iVerifMF & " <> " & iNbPrenomsTot,
                MsgBoxStyle.Exclamation, sTitreAppli)
        End If

    End Sub

    Private Sub FiltrerPrenomUnigenre(
            dicoE As DicoTri(Of String, clsPrenom),
            dicoH As DicoTri(Of String, clsPrenom),
            dicoS As DicoTri(Of String, clsPrenom),
            iNbPrenomsTot%, iNbPrenomsTotOk%, iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%)

        Dim iNbPrenomsVerif% = 0
        Dim iNbPrenomsVerifMF% = 0
        Dim aPrenomsE = dicoE.Trier("")
        For Each prenom In aPrenomsE

            iNbPrenomsVerif += prenom.iNbOcc
            iNbPrenomsVerifMF += prenom.iNbOccMasc + prenom.iNbOccFem

            ' Définition de unigenre : ni épicène, ni homophone, ni similaire
            prenom.bUnigenre = True
            If prenom.bMixteEpicene OrElse
               prenom.bMixteHomophone OrElse
               prenom.bSimilaire Then
                prenom.bUnigenre = False
            End If

        Next

        Dim iVerif% = iNbPrenomsVerif + iNbPrenomsIgnores + iNbPrenomsIgnoresDate
        If iVerif <> iNbPrenomsTot Then
            Debug.WriteLine(sFormaterNum(iNbPrenomsVerif) & " <> " & sFormaterNum(iNbPrenomsTotOk))
            MsgBox("Décompte faux : " & iVerif & " <> " & iNbPrenomsTot,
                MsgBoxStyle.Exclamation, sTitreAppli)
        End If
        Dim iVerifMF% = iNbPrenomsVerifMF + iNbPrenomsIgnores + iNbPrenomsIgnoresDate
        If iVerifMF <> iNbPrenomsTot Then
            Debug.WriteLine(sFormaterNum(iVerifMF) & " <> " & sFormaterNum(iNbPrenomsTot))
            MsgBox("Décompte faux : " & iVerifMF & " <> " & iNbPrenomsTot,
                MsgBoxStyle.Exclamation, sTitreAppli)
        End If

    End Sub

    Private Sub AfficherSyntheses(sDossierAppli$,
            dicoCorrectionsPrenoms As DicoTri(Of String, String),
            dicoCorrectionsPrenomsUtil As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophones As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophonesUtil As DicoTri(Of String, String),
            dicoDefinitionsPrenomsSimilaires As DicoTri(Of String, String),
            dicoDefinitionsPrenomsSimilairesUtil As DicoTri(Of String, String),
            dicoE As DicoTri(Of String, clsPrenom),
            dicoH As DicoTri(Of String, clsPrenom),
            dicoS As DicoTri(Of String, clsPrenom),
            iNbLignes%, iNbLignesOk%,
            iNbPrenomsTot%, iNbPrenomsTotOk%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            sbCPMD As StringBuilder, sbHPMD As StringBuilder, sbSPMD As StringBuilder,
            sCheminCorrectionsPrenoms$,
            sCheminDefPrenomsMixtesHomophones$,
            sCheminDefPrenomsSimilaires$)

        Dim sbBilan As New StringBuilder

        AfficherSynthesePrenomFrequentEtUnigenre(sDossierAppli, dicoE, iNbPrenomsTotOk,
            iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMinPrenomsFrequents, rSeuilFreqRelPrenomsFrequents,
            iNbLignesMaxPrenoms)
        ' Pour le bilan général, conserver l'ordre alphab. pour vérifier la non régression
        AfficherSynthesePrenomFrequentEtUnigenre(sDossierAppli, dicoE, iNbPrenomsTotOk,
            iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMinPrenomsFrequents, rSeuilFreqRelPrenomsFrequents,
            iNbLignesMaxPrenoms, sbBilan, bTriAlphab:=True)

        AfficherSynthesePrenomFrequentEtUnigenre(sDossierAppli, dicoE, iNbPrenomsTotOk,
            iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMinPrenomsUnigenre, rSeuilFreqRelPrenomsFrequents,
            iNbLignesMaxPrenoms, bUnigenre:=True)
        AfficherSynthesePrenomFrequentEtUnigenre(sDossierAppli, dicoE, iNbPrenomsTotOk,
            iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMinPrenomsUnigenre, rSeuilFreqRelPrenomsFrequents,
            iNbLignesMaxPrenoms, sbBilan, bTriAlphab:=True, bUnigenre:=True)

        AfficherSyntheseEpicene(sDossierAppli, dicoE,
            iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMinPrenomsEpicenes, rSeuilFreqRelPrenomsEpicenes,
            iNbLignesMaxPrenoms,
            dicoCorrectionsPrenoms, dicoCorrectionsPrenomsUtil)
        AfficherSyntheseEpicene(sDossierAppli, dicoE,
            iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMinPrenomsEpicenes, rSeuilFreqRelPrenomsEpicenes,
            iNbLignesMaxPrenoms,
            dicoCorrectionsPrenoms, dicoCorrectionsPrenomsUtil, sbBilan, bTriAlphab:=True)

        AfficherSyntheseHomophone(sDossierAppli, dicoH, dicoE,
            iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMinPrenomsHomophones, rSeuilFreqRelPrenomsHomophones,
            iNbLignesMaxPrenoms,
            dicoDefinitionsPrenomsMixtesHomophones,
            dicoDefinitionsPrenomsMixtesHomophonesUtil)
        AfficherSyntheseHomophone(sDossierAppli, dicoH, dicoE,
            iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMinPrenomsHomophones, rSeuilFreqRelPrenomsHomophones,
            iNbLignesMaxPrenoms,
            dicoDefinitionsPrenomsMixtesHomophones,
            dicoDefinitionsPrenomsMixtesHomophonesUtil, sbBilan, bTriAlphab:=True)

        AfficherSyntheseSimilaire(sDossierAppli,
            dicoS, dicoE,
            iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMinPrenomsSimilaires, rSeuilFreqRelPrenomsSimilaires,
            iNbLignesMaxPrenoms,
            dicoDefinitionsPrenomsSimilaires,
            dicoDefinitionsPrenomsSimilairesUtil) ' dicoH,
        AfficherSyntheseSimilaire(sDossierAppli,
            dicoS, dicoE,
            iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMinPrenomsSimilaires, rSeuilFreqRelPrenomsSimilaires,
            iNbLignesMaxPrenoms,
            dicoDefinitionsPrenomsSimilaires,
            dicoDefinitionsPrenomsSimilairesUtil, sbBilan, bTriAlphab:=True) 'dicoH,

        sbBilan.AppendLine("Liste des corrections d'accent").
            AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine()
        sbBilan.Append(sbLireFichier(sCheminCorrectionsPrenoms, bDoublerRAL:=True)).
            AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()

        sbBilan.AppendLine("Liste des définitions de prénoms mixtes homophones").
            AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine()
        sbBilan.Append(sbLireFichier(sCheminDefPrenomsMixtesHomophones, bDoublerRAL:=True)).
            AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()

        sbBilan.AppendLine("Liste des définitions de prénoms similaires").
            AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine()
        sbBilan.Append(sbLireFichier(sCheminDefPrenomsSimilaires, bDoublerRAL:=True)).
            AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()

        sbBilan.Append(sbCPMD).AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()
        sbBilan.Append(sbHPMD).AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()
        sbBilan.Append(sbSPMD)

        Dim sCheminBilan$ = sDossierAppli & "\Bilan.md"
        EcrireFichier(sCheminBilan, sbBilan)

    End Sub

    Private Sub AfficherSynthesePrenomFrequentEtUnigenre(sDossierAppli$,
            dicoE As DicoTri(Of String, clsPrenom),
            iNbPrenomsTotOk%, iNbPrenomsTot%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, iNbLignesMax%,
            Optional sbBilan As StringBuilder = Nothing,
            Optional bTriAlphab As Boolean = False,
            Optional ByVal bUnigenre As Boolean = False)

        ' Produire la synthèse statistique des prénoms fréquents

        Dim sb As New StringBuilder
        AfficherInfo(sb, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante)
        Dim sbMD As New StringBuilder ' Syntaxe MarkDown
        Dim sTitre$ = "Synthèse statistique des prénoms fréquents"
        If bUnigenre Then sTitre = "Synthèse statistique des prénoms unigenres"
        sbMD.AppendLine(sTitre)
        sbMD.AppendLine()
        AfficherInfo(sbMD, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante, bDoublerRAL:=True)
        sbMD.AppendLine(sEnteteMarkDown())

        Dim sbWK As New StringBuilder ' Syntaxe Wiki
        AfficherInfo(sbWK, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante, bDoublerRAL:=True)
        sbWK.AppendLine(sEnteteWiki(sTitre))

        Dim iNbPrenoms% = 0
        Dim sTri$ = "iNbOcc desc"
        If bUnigenre Then sTri = "bUnigenre desc, rFreqTotale desc"
        If bTriAlphab Then sTri = "sPrenom"
        For Each prenom In dicoE.Trier(sTri)

            If iSeuilMin > 0 AndAlso prenom.iNbOcc < iSeuilMin Then Continue For

            If bUnigenre AndAlso Not prenom.bUnigenre Then Continue For
            prenom.bSelect = True

            iNbPrenoms += 1
            If iNbLignesMax > 0 AndAlso iNbPrenoms > iNbLignesMax Then Exit For

            sb.AppendLine(sLigneDebug(prenom, prenom.sPrenom, iNbPrenoms, sFormatFreq))

            Dim bGras = False
            Dim bItalique = False
            Dim iNumVariante% = -1
            If prenom.bMixteEpicene Then bGras = True
            If prenom.bMixteHomophone Then bItalique = True
            sbMD.AppendLine(sLigneMarkDown(prenom, prenom.sPrenom, iNbPrenoms, sFormatFreq,
                iNumVariante, bGras, bItalique))

            sbWK.AppendLine(sLigneWiki(prenom, prenom.sPrenom, iNbPrenoms, sFormatFreq,
                iNumVariante, bGras, bItalique))

        Next
        sbWK.AppendLine("|}")

        Dim sSuffixe$ = ""
        If bTriAlphab Then
            sbBilan.Append(sbMD).AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()
            sSuffixe = "Alphab"
        End If
        Dim sFichier$ = "PrenomsFrequents"
        If bUnigenre Then sFichier = "PrenomsUnigenres"
        Dim sChemin$ = sDossierAppli & "\" & sFichier & sSuffixe & ".md"
        EcrireFichier(sChemin, sbMD)
        Dim sCheminWK$ = sDossierAppli & "\" & sFichier & sSuffixe & ".wiki"
        EcrireFichier(sCheminWK, sbWK)

    End Sub

    Private Sub AfficherSyntheseEpicene(sDossierAppli$,
            dicoE As DicoTri(Of String, clsPrenom),
            iNbPrenomsTotOk%, iNbPrenomsTot%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, iNbLignesMax%,
            dicoCorrectionsPrenoms As DicoTri(Of String, String),
            dicoCorrectionsPrenomsUtil As DicoTri(Of String, String),
            Optional sbBilan As StringBuilder = Nothing, Optional bTriAlphab As Boolean = False)

        ' Produire la synthèse statistique des prénoms mixtes épicènes

        Dim sb As New StringBuilder
        AfficherInfo(sb, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, 0)

        Dim sbMD As New StringBuilder ' Syntaxe MarkDown
        sbMD.AppendLine("Synthèse statistique des prénoms mixtes épicènes")
        sbMD.AppendLine()
        AfficherInfo(sbMD, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, 0, bDoublerRAL:=True)
        sbMD.AppendLine(sEnteteMarkDown())

        Dim sbWK As New StringBuilder ' Syntaxe Wiki
        AfficherInfo(sbWK, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, 0, bDoublerRAL:=True)
        sbWK.AppendLine(sEnteteWiki("Synthèse statistique des prénoms mixtes épicènes"))

        Dim iNbPrenomsMixtes% = 0
        Dim sTri$ = "bMixteEpicene desc, rFreqTotale desc"
        If bTriAlphab Then sTri = "sPrenom"
        For Each prenom In dicoE.Trier(sTri)

            If Not prenom.bMixteEpicene Then Continue For

            iNbPrenomsMixtes += 1
            If iNbLignesMax > 0 AndAlso iNbPrenomsMixtes > iNbLignesMax Then Exit For

            prenom.bSelect = True

            sb.AppendLine(sLigneDebug(prenom, prenom.sPrenom, iNbPrenomsMixtes, sFormatFreq))

            sbMD.AppendLine(sLigneMarkDown(prenom, prenom.sPrenom, iNbPrenomsMixtes, sFormatFreq))

            sbWK.AppendLine(sLigneWiki(prenom, prenom.sPrenom, iNbPrenomsMixtes, sFormatFreq))

        Next
        sbWK.AppendLine("|}")

        sb.AppendLine()
        sbMD.AppendLine()
        sbWK.AppendLine()
        For Each kvp In dicoCorrectionsPrenoms
            If Not dicoCorrectionsPrenomsUtil.ContainsKey(kvp.Key) Then
                Dim sLigne$ = "Correction de prénom non trouvée : " & kvp.Key
                Debug.WriteLine(sLigne)
                sb.AppendLine(sLigne)
                sbMD.AppendLine(sLigne)
                sbWK.AppendLine(sLigne)
            End If
        Next

        'Debug.WriteLine(sb.ToString)

        If bTriAlphab Then
            sbBilan.Append(sbMD).AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()
        Else
            Dim sCheminMD$ = sDossierAppli & "\PrenomsMixtesEpicenes.md"
            EcrireFichier(sCheminMD, sbMD)
            Dim sCheminWK$ = sDossierAppli & "\PrenomsMixtesEpicenes.wiki"
            EcrireFichier(sCheminWK, sbWK)
        End If

    End Sub

    Private Sub AfficherSyntheseHomophone(sDossierAppli$,
            dicoH As DicoTri(Of String, clsPrenom),
            dicoE As DicoTri(Of String, clsPrenom),
            iNbPrenomsTotOk%, iNbPrenomsTot%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, iNbLignesMax%,
            dicoDefinitionsPrenomsMixtesHomophones As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophonesUtil As DicoTri(Of String, String),
            Optional sbBilan As StringBuilder = Nothing, Optional bTriAlphab As Boolean = False)

        ' Produire la synthèse statistique des prénoms mixtes homophones

        Dim sb As New StringBuilder
        AfficherInfo(sb, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante)

        Dim sbMD As New StringBuilder ' Syntaxe MarkDown
        sbMD.AppendLine("Synthèse statistique des prénoms mixtes homophones")
        sbMD.AppendLine()
        AfficherInfo(sbMD, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante, bDoublerRAL:=True)
        sbMD.AppendLine(sEnteteMarkDown(bColonneFreqVariante:=True))

        Dim sbWK As New StringBuilder ' Syntaxe Wiki
        AfficherInfo(sbWK, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante, bDoublerRAL:=True)
        sbWK.AppendLine(sEnteteWiki("Synthèse statistique des prénoms mixtes homophones",
            bColonneFreqVariante:=True))

        Dim iNbPrenomsMixtes% = 0
        Dim sTri$ = "bMixteHomophone desc, rFreqTotale desc"
        If bTriAlphab Then sTri = "sPrenom"
        For Each prenom In dicoH.Trier(sTri)

            If Not prenom.bMixteHomophone Then Continue For
            If iSeuilMin > 0 AndAlso prenom.iNbOcc < iSeuilMin Then Continue For

            iNbPrenomsMixtes += 1
            If iNbLignesMax > 0 AndAlso iNbPrenomsMixtes > iNbLignesMax Then Exit For

            prenom.bSelect = True

            Dim sPrenom$ = prenom.sPrenomHomophone
            Dim sPrenomMD$ = sPrenom
            Dim sPrenomWiki$ = sPrenom
            Dim bVariantes = False
            If prenom.dicoVariantesH.Count > 1 Then
                bVariantes = True
                Dim lst = prenom.dicoVariantesH.ToList

                ' Mettre en gras le plus fréquent
                'Dim sPrenomMajoritaire$ = ""
                'For Each prenomV In prenom.dicoVariantesH.Trier("iNbOcc desc")
                '    sPrenomMajoritaire = prenomV.sPrenom
                '    Exit For
                'Next
                'sPrenomMD = sListerCleTxt(lst, sPrenomMajoritaire, sGras)
                'sPrenomWiki = sListerCleTxt(lst, sPrenomMajoritaire, "'''")

                ' Mêmes conditions que pour la liste des prénoms fréquents :
                ' Gras : épicène
                ' Italique : homophone
                ' Gras+Italique : épicène + homophone
                sPrenomMD = sListerCleTxtDico(lst, dicoE, bWiki:=False, bHomophoneEnItalique:=False)
                sPrenomWiki = sListerCleTxtDico(lst, dicoE, bWiki:=True, bHomophoneEnItalique:=False)

            End If

            sb.AppendLine(sLigneDebug(prenom, sPrenom, iNbPrenomsMixtes, sFormatFreq))
            sbMD.AppendLine(sLigneMarkDown(prenom, sPrenomMD, iNbPrenomsMixtes, sFormatFreq,
                iNumVariante:=0, bSuffixeNumVariante:=True))
            sbWK.AppendLine(sLigneWiki(prenom, sPrenomWiki, iNbPrenomsMixtes, sFormatFreq,
                iNumVariante:=0, bSuffixeNumVariante:=True))

            If bVariantes Then
                Dim iNumVariante% = 0
                For Each prenomV In prenom.dicoVariantesH.Trier("iNbOcc desc")
                    If prenomV.bVarianteDecompteeH Then Continue For
                    iNumVariante += 1
                    sb.AppendLine(sLigneDebug(prenomV, prenomV.sPrenom, iNbPrenomsMixtes, sFormatFreq))
                    Dim bGras = False
                    'If iNumVariante = 1 Then bGras = True ' Le plus fréquent
                    If dicoE.ContainsKey(prenomV.sPrenom) Then
                        Dim prenomE = dicoE(prenomV.sPrenom)
                        If prenomE.bMixteEpicene Then bGras = True
                    End If
                    Dim bItalique = False
                    ' Inutile, car dans la synthèse homophone, tous les prénoms sont au moins homophones
                    'If dicoE.ContainsKey(prenomV.sPrenom) Then
                    '    Dim prenomH = dicoE(prenomV.sPrenom)
                    '    If prenomH.bMixteHomophone Then bItalique = True
                    'End If
                    sbMD.AppendLine(sLigneMarkDown(prenomV, prenomV.sPrenom, iNbPrenomsMixtes,
                        sFormatFreq, iNumVariante, bGras, bItalique,
                        bSuffixeNumVariante:=True, bColonneFreqVarianteH:=True))
                    sbWK.AppendLine(sLigneWiki(prenomV, prenomV.sPrenom, iNbPrenomsMixtes,
                        sFormatFreq, iNumVariante, bGras, bItalique,
                        bSuffixeNumVariante:=True, bColonneFreqVarianteH:=True))
                Next
            End If

        Next
        sbWK.AppendLine("|}")

        sb.AppendLine()
        sbMD.AppendLine()
        sbWK.AppendLine()
        For Each kvp In dicoDefinitionsPrenomsMixtesHomophones
            If Not dicoDefinitionsPrenomsMixtesHomophonesUtil.ContainsKey(kvp.Key) Then
                Dim sLigne$ = "Correction de prénom (liste homophone) non trouvée : " & kvp.Key
                Debug.WriteLine(sLigne)
                sb.AppendLine(sLigne)
                sbMD.AppendLine(sLigne)
                sbWK.AppendLine(sLigne)
            End If
        Next

        'Debug.WriteLine(sb.ToString)

        If bTriAlphab Then
            sbBilan.Append(sbMD).AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()
        Else
            Dim sCheminMD$ = sDossierAppli & "\PrenomsMixtesHomophones.md"
            EcrireFichier(sCheminMD, sbMD)
            Dim sCheminWK$ = sDossierAppli & "\PrenomsMixtesHomophones.wiki"
            EcrireFichier(sCheminWK, sbWK)
        End If

    End Sub

    Private Sub AfficherSyntheseSimilaire(sDossierAppli$,
            dicoS As DicoTri(Of String, clsPrenom),
            dicoE As DicoTri(Of String, clsPrenom),
            iNbPrenomsTotOk%, iNbPrenomsTot%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, iNbLignesMax%,
            dicoDefinitionsPrenomsSimilaires As DicoTri(Of String, String),
            dicoDefinitionsPrenomsSimilairesUtil As DicoTri(Of String, String),
            Optional sbBilan As StringBuilder = Nothing, Optional bTriAlphab As Boolean = False)

        ' Produire la synthèse statistique des prénoms similaires

        Dim sb As New StringBuilder
        AfficherInfo(sb, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante)

        Dim sbMD As New StringBuilder ' Syntaxe MarkDown
        sbMD.AppendLine("Synthèse statistique des prénoms similaires")
        sbMD.AppendLine()
        AfficherInfo(sbMD, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante, bDoublerRAL:=True)
        sbMD.AppendLine(sEnteteMarkDown(bColonneFreqVariante:=True))

        Dim sbWK As New StringBuilder ' Syntaxe Wiki
        AfficherInfo(sbWK, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante, bDoublerRAL:=True)
        sbWK.AppendLine(sEnteteWiki("Synthèse statistique des prénoms similaires",
            bColonneFreqVariante:=True))

        Dim iNbPrenomsSimilaires% = 0
        Dim sTri$ = "bSimilaire desc, rFreqTotale desc"
        If bTriAlphab Then sTri = "sPrenom"
        For Each prenom In dicoS.Trier(sTri)

            'If dicoE.ContainsKey(prenom.sPrenom) Then
            '    If Not dicoE(prenom.sPrenom).bSimilaire Then Continue For
            'End If
            If Not prenom.bSimilaire Then Continue For

            If iSeuilMin > 0 AndAlso prenom.iNbOcc < iSeuilMin Then Continue For

            iNbPrenomsSimilaires += 1
            If iNbLignesMax > 0 AndAlso iNbPrenomsSimilaires > iNbLignesMax Then Exit For

            prenom.bSelect = True

            Dim sPrenom$ = prenom.sPrenomSimilaire
            Dim sPrenomMD$ = sPrenom
            Dim sPrenomWiki$ = sPrenom
            Dim bVariantes = False
            If prenom.dicoVariantesS.Count > 1 Then
                bVariantes = True
                Dim lst = prenom.dicoVariantesS.ToList
                ' Mêmes conditions que pour la liste des prénoms fréquents :
                ' Gras : épicène
                ' Italique : homophone
                ' Gras+Italique : épicène + homophone
                sPrenomMD = sListerCleTxtDico(lst, dicoE, bWiki:=False, bHomophoneEnItalique:=True)
                sPrenomWiki = sListerCleTxtDico(lst, dicoE, bWiki:=True, bHomophoneEnItalique:=True)
            End If

            sb.AppendLine(sLigneDebug(prenom, sPrenom, iNbPrenomsSimilaires, sFormatFreq))
            sbMD.AppendLine(sLigneMarkDown(prenom, sPrenomMD, iNbPrenomsSimilaires, sFormatFreq,
                iNumVariante:=0, bSuffixeNumVariante:=True))
            sbWK.AppendLine(sLigneWiki(prenom, sPrenomWiki, iNbPrenomsSimilaires, sFormatFreq,
                iNumVariante:=0, bSuffixeNumVariante:=True))

            If bVariantes Then
                Dim iNumVariante% = 0
                For Each prenomV In prenom.dicoVariantesS.Trier("iNbOcc desc")
                    If prenomV.bVarianteDecompteeS Then Continue For
                    iNumVariante += 1
                    sb.AppendLine(sLigneDebug(prenomV, prenomV.sPrenom, iNbPrenomsSimilaires, sFormatFreq))
                    Dim bGras = False
                    If dicoE.ContainsKey(prenomV.sPrenom) Then
                        Dim prenomE = dicoE(prenomV.sPrenom)
                        If prenomE.bMixteEpicene Then bGras = True
                    End If
                    Dim bItalique = False
                    'If dicoH.ContainsKey(prenomV.sPrenom) Then
                    '    Dim prenomH = dicoH(prenomV.sPrenom)
                    '    If prenomH.bMixteHomophone Then bItalique = True
                    'End If
                    If dicoE.ContainsKey(prenomV.sPrenom) Then
                        Dim prenomH = dicoE(prenomV.sPrenom)
                        If prenomH.bMixteHomophone Then bItalique = True
                    End If
                    sbMD.AppendLine(sLigneMarkDown(prenomV, prenomV.sPrenom, iNbPrenomsSimilaires,
                        sFormatFreq, iNumVariante, bGras, bItalique,
                        bSuffixeNumVariante:=True, bColonneFreqVarianteS:=True))
                    sbWK.AppendLine(sLigneWiki(prenomV, prenomV.sPrenom, iNbPrenomsSimilaires,
                        sFormatFreq, iNumVariante, bGras, bItalique,
                        bSuffixeNumVariante:=True, bColonneFreqVarianteS:=True))
                Next
            End If

        Next
        sbWK.AppendLine("|}")

        sb.AppendLine()
        sbMD.AppendLine()
        sbWK.AppendLine()
        For Each kvp In dicoDefinitionsPrenomsSimilaires
            If Not dicoDefinitionsPrenomsSimilairesUtil.ContainsKey(kvp.Key) Then
                Dim sLigne$ = "Correction de prénom (liste des prénoms similaires) non trouvée : " & kvp.Key
                Debug.WriteLine(sLigne)
                sb.AppendLine(sLigne)
                sbMD.AppendLine(sLigne)
                sbWK.AppendLine(sLigne)
            End If
        Next

        'Debug.WriteLine(sb.ToString)

        If bTriAlphab Then
            sbBilan.Append(sbMD).AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()
        Else
            Dim sCheminMD$ = sDossierAppli & "\PrenomsSimilaires.md"
            EcrireFichier(sCheminMD, sbMD)
            Dim sCheminWK$ = sDossierAppli & "\PrenomsSimilaires.wiki"
            EcrireFichier(sCheminWK, sbWK)
        End If

    End Sub

    Private Function sEnteteMarkDown$(Optional bColonneFreqVariante As Boolean = False)

        Dim sColonneFreqVariante$ = ""
        Dim sColonneFreqVariante2$ = ""
        If bColonneFreqVariante Then
            sColonneFreqVariante = "Fréq. rel. var.|"
            sColonneFreqVariante2 = "--:|"
        End If

        Dim s$ = ""
        s &= vbLf & "|n° |Occurrences|Prénom|Année moyenne|Année moyenne masc.|Année moyenne fém.|Occurrences masc.|Occurrences fém.|Fréq.|Fréq. rel. masc.|Fréq. rel. fém.|" & sColonneFreqVariante
        s &= vbLf & "|--:|--:|:--|:-:|:-:|:-:|--:|--:|--:|--:|--:|" & sColonneFreqVariante2
        Return s

    End Function

    Private Function sEnteteWiki$(sTitre$, Optional bColonneFreqVariante As Boolean = False)

        ' https://fr.wikipedia.org/wiki/Aide:Insérer_un_tableau_(wikicode,_avancé)

        Dim sColonneFreqVariante$ = ""
        If bColonneFreqVariante Then
            sColonneFreqVariante = vbLf & "! scope='col' | Fréq. rel. var."
        End If

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
        s &= sColonneFreqVariante
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
            ", freq. rel. m. " & sGenre & prenom.rFreqRelativeMasc.ToString(sFormatFreqRel) &
            ", freq. rel. f. " & sGenre & prenom.rFreqRelativeFem.ToString(sFormatFreqRel) &
            ", mixte épicène=" & prenom.bMixteEpicene
        Return s

    End Function

    Private Function sLigneMarkDown$(prenom As clsPrenom, sPrenom$, iNumPrenom%, sFormatFreq$,
            Optional iNumVariante% = -1,
            Optional bGras As Boolean = False,
            Optional bItalique As Boolean = False,
            Optional bSuffixeNumVariante As Boolean = False,
            Optional bColonneFreqVarianteH As Boolean = False,
            Optional bColonneFreqVarianteS As Boolean = False)

        Dim sMiseEnForme$ = ""
        Dim sNumVariante$ = ""
        If bSuffixeNumVariante AndAlso iNumVariante >= 0 Then sNumVariante = "." & iNumVariante
        If bGras Then sMiseEnForme = sGrasMD ' Gras
        If bItalique Then sMiseEnForme = sItaliqueMD ' Italique
        If bGras AndAlso bItalique Then sMiseEnForme = sItaliqueGrasMD ' Italique en gras
        Dim sColonneFreqVariante$ = ""
        If bColonneFreqVarianteH Then
            sColonneFreqVariante = "|" & prenom.rFreqRelativeVarianteH.ToString(sFormatFreqRelVariante)
        End If
        If bColonneFreqVarianteS Then
            sColonneFreqVariante = "|" & prenom.rFreqRelativeVarianteS.ToString(sFormatFreqRelVariante)
        End If

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
            "|" & prenom.rFreqRelativeMasc.ToString(sFormatFreqRel) &
            "|" & prenom.rFreqRelativeFem.ToString(sFormatFreqRel) &
            sColonneFreqVariante

        Return s

    End Function

    Private Function sLigneWiki$(prenom As clsPrenom, sPrenom$, iNumPrenom%, sFormatFreq$,
            Optional iNumVariante% = -1,
            Optional bGras As Boolean = False,
            Optional bItalique As Boolean = False,
            Optional bSuffixeNumVariante As Boolean = False,
            Optional bColonneFreqVarianteH As Boolean = False,
            Optional bColonneFreqVarianteS As Boolean = False)

        Dim sMiseEnForme$ = ""
        Dim sNumVariante$ = ""
        If bSuffixeNumVariante AndAlso iNumVariante >= 0 Then sNumVariante = "." & iNumVariante
        If bGras Then sMiseEnForme = sGrasWiki ' Gras
        If bItalique Then sMiseEnForme = sItaliqueWiki ' Italique
        If bGras AndAlso bItalique Then sMiseEnForme = sItaliqueGrasWiki ' Italique en gras
        Dim sColonneFreqVariante$ = ""
        If bColonneFreqVarianteH Then
            sColonneFreqVariante = "||" & prenom.rFreqRelativeVarianteH.ToString(sFormatFreqRelVariante)
        End If
        If bColonneFreqVarianteS Then
            sColonneFreqVariante = "||" & prenom.rFreqRelativeVarianteS.ToString(sFormatFreqRelVariante)
        End If

        ' Ne rien afficher si 0, car cela perturbe le tri dans le wiki
        ' (le vide est trié aussi en 1er, pareil pour -, 9999 est bien trié en dernier, mais bon...)
        Dim sAnneeMoyMasc$ = ""
        Dim sAnneeMoyFem$ = ""
        If prenom.rAnneeMoyMasc > 0 Then sAnneeMoyMasc = prenom.rAnneeMoyMasc.ToString("0")
        If prenom.rAnneeMoyFem > 0 Then sAnneeMoyFem = prenom.rAnneeMoyFem.ToString("0")

        Dim s$ = "|-" & vbLf &
                "|" & iNumPrenom & sNumVariante &
                "|| align='right' | " & sFormaterNumWiki(prenom.iNbOcc) &
                "||" & sMiseEnForme & sPrenom & sMiseEnForme &
                "||" & prenom.rAnneeMoy.ToString("0") &
                "||" & sAnneeMoyMasc &
                "||" & sAnneeMoyFem &
                "|| align='right' | " & sFormaterNumWiki(prenom.iNbOccMasc) &
                "|| align='right' | " & sFormaterNumWiki(prenom.iNbOccFem) &
                "||" & prenom.rFreqTotale.ToString(sFormatFreq) &
                "||" & prenom.rFreqRelativeMasc.ToString(sFormatFreqRel) &
                "||" & prenom.rFreqRelativeFem.ToString(sFormatFreqRel) &
                sColonneFreqVariante
        Return s

    End Function

    Private Function bAnalyserPrenom(sLigne$, prenom As clsPrenom,
            dicoCorrectionsPrenoms As DicoTri(Of String, String),
            dicoCorrectionsPrenomsUtil As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophones As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophonesUtil As DicoTri(Of String, String),
            dicoDefinitionsPrenomsSimilaires As DicoTri(Of String, String),
            dicoDefinitionsPrenomsSimilairesUtil As DicoTri(Of String, String)) As Boolean

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

        If dicoCorrectionsPrenoms.ContainsKey(sPrenom) Then
            Dim sPrenomCorrige$ = dicoCorrectionsPrenoms(sPrenom)
            If Not dicoCorrectionsPrenomsUtil.ContainsKey(sPrenom) Then
                dicoCorrectionsPrenomsUtil.Add(sPrenom, sPrenomCorrige)
            End If
            sPrenom = sPrenomCorrige
        End If

        Dim sPrenomHomophone = sPrenom
        If dicoDefinitionsPrenomsMixtesHomophones.ContainsKey(sPrenom) Then
            Dim sPrenomH$ = dicoDefinitionsPrenomsMixtesHomophones(sPrenom)
            If Not dicoDefinitionsPrenomsMixtesHomophonesUtil.ContainsKey(sPrenom) Then
                dicoDefinitionsPrenomsMixtesHomophonesUtil.Add(sPrenom, sPrenomH)
            End If
            sPrenomHomophone = sPrenomH
        End If

        ' Prénoms similaires (par ex.: antoinette : féminin de antoine)
        Dim sPrenomSimilaires = sPrenom
        If dicoDefinitionsPrenomsSimilaires.ContainsKey(sPrenom) Then
            Dim sPrenomS$ = dicoDefinitionsPrenomsSimilaires(sPrenom)
            If Not dicoDefinitionsPrenomsSimilairesUtil.ContainsKey(sPrenom) Then
                dicoDefinitionsPrenomsSimilairesUtil.Add(sPrenom, sPrenomS)
            End If
            sPrenomSimilaires = sPrenomS
        End If

        prenom.sPrenom = FirstCharToUpper(sPrenom)
        prenom.sPrenomHomophone = FirstCharToUpper(sPrenomHomophone)
        prenom.sPrenomSimilaire = FirstCharToUpper(sPrenomSimilaires)
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

    Private Sub Exporter(sDossierAppli$, asLignes$(),
        dicoE As DicoTri(Of String, clsPrenom),
        dicoCorrectionsPrenoms As DicoTri(Of String, String),
        dicoCorrectionsPrenomsUtil As DicoTri(Of String, String),
        dicoDefinitionsPrenomsMixtesHomophones As DicoTri(Of String, String),
        dicoDefinitionsPrenomsMixtesHomophonesUtil As DicoTri(Of String, String),
        dicoDefinitionsPrenomsSimilaires As DicoTri(Of String, String),
        dicoDefinitionsPrenomsSimilairesUtil As DicoTri(Of String, String),
        bTestPrenomOrig As Boolean)

        ' Génération d'un nouveau fichier csv filtré ou pas

        ' bTestPrenomOrig : Vérifier si le traitement appliqué préserve entièrement le fichier d'origine
        Const bFiltrerPrenomEpicene = False

        Dim sb As New StringBuilder
        Dim iNbLignes = 0
        Dim sAjoutEntete$ = ""
        If Not bTestPrenomOrig Then sAjoutEntete =
            ";Prénom d'origine;Prénom épicène;Prénom homophone;Prénom similaire;Prénom unigenre"
        For Each sLigne As String In asLignes

            iNbLignes += 1
            If iNbLignes = 1 Then sb.AppendLine(sLigne & sAjoutEntete) : Continue For

            Dim prenom As New clsPrenom
            If Not bAnalyserPrenom(sLigne$, prenom,
                dicoCorrectionsPrenoms,
                dicoCorrectionsPrenomsUtil,
                dicoDefinitionsPrenomsMixtesHomophones,
                dicoDefinitionsPrenomsMixtesHomophonesUtil,
                dicoDefinitionsPrenomsSimilaires,
                dicoDefinitionsPrenomsSimilairesUtil) Then Continue For

            ConvertirPrenom(prenom)

            If Not bTestPrenomOrig Then
                If prenom.iAnnee < iDateMinExport Then Continue For
                If prenom.iAnnee > iDateMaxExport Then Continue For
            End If

            Dim sAjout$ = ""
            Dim sPrenom$ = prenom.sPrenom
            If bTestPrenomOrig Then
                ' Si on remet en majuscule et qu'on rétablit les accents corrigés
                '  on doit retrouver exactement le fichier d'origine
                sPrenom = sPrenom.ToUpper
                If prenom.sPrenomOrig <> sPrenom Then
                    sPrenom = sEnleverAccents(sPrenom, bMinuscule:=False)

                    ' Ex.: jérôme corrigé en jérome
                    Dim sPrenomOrigSansAccent$ = sEnleverAccents(prenom.sPrenomOrig, bMinuscule:=False)
                    If sPrenomOrigSansAccent <> prenom.sPrenomOrig Then
                        'Debug.WriteLine(prenom.sPrenomOrig)
                        ' Pour les corrections spéciales, on est obligé de rétablir tel quel
                        sPrenom = prenom.sPrenomOrig.ToUpper
                    End If

                    If sPrenomOrigSansAccent <> sPrenom Then
                        ' Vérifier si tous les accents sont bien retirés
                        Debug.WriteLine(sPrenom)
                        Stop ' Provoque l'arrêt du test, comme une exception
                    End If
                End If
                If prenom.sPrenomOrig = clsPrenom.sPrenomRare Then GoTo Suite
                If prenom.sAnnee = clsPrenom.sDateXXXX Then GoTo Suite
            End If

            Dim sPrenomE$ = ""
            Dim sPrenomH$ = ""
            Dim sPrenomS$ = ""
            Dim sPrenomU$ = ""
            Dim sCleE$ = prenom.sPrenom
            If dicoE.ContainsKey(sCleE) Then
                Dim prenom0 = dicoE(sCleE)
                If bFiltrerPrenomEpicene AndAlso Not prenom0.bSelect Then Continue For
                If prenom0.bMixteEpicene Then sPrenomE = "1"
                If prenom0.bMixteHomophone Then sPrenomH = prenom0.sPrenomHomophone
                If prenom0.bSimilaire Then sPrenomS = prenom0.sPrenomSimilaire
                If prenom0.bUnigenre Then sPrenomU = "1"
            Else
                Continue For
            End If

            If Not bTestPrenomOrig Then sAjout = ";" & _
                prenom.sPrenomOrig & ";" & sPrenomE & ";" & sPrenomH & ";" & _
                sPrenomS & ";" & sPrenomU

Suite:
            Dim sLigneC$ = prenom.sCodeSexe & ";" & sPrenom & ";" & prenom.sAnnee & ";" & _
                prenom.iNbOcc & sAjout
            sb.AppendLine(sLigneC)

        Next

        Dim sCheminOut$ = sDossierAppli & "\" & sFichierPrenomsInseeCorrige
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
            iNbPrenomsTotOk%, iNbPrenomsTot%, iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, rSeuilFreqRelVariante!,
            Optional bDoublerRAL As Boolean = False)

        sb.AppendLine("Date début = 1900")
        If bDoublerRAL Then sb.AppendLine("")
        sb.AppendLine("Date fin   = " & iDateFin)
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
            sb.AppendLine("Fréquence relative min. genre = " &
                rSeuilFreqRel.ToString(sFormatFreqRel))
            If bDoublerRAL Then sb.AppendLine("")
        End If
        If rSeuilFreqRelVariante > 0 Then
            sb.AppendLine("Fréquence relative min. variante = " &
                rSeuilFreqRelVariante.ToString(sFormatFreqRelVariante))
            If bDoublerRAL Then sb.AppendLine("")
        End If

    End Sub

    Private Function LireFichier(sChemin$) As DicoTri(Of String, String)

        ' Lire un fichier de corrections de prénoms et retourner un dictionnaire

        Dim dico As New DicoTri(Of String, String)
        If Not IO.File.Exists(sChemin) Then
            MsgBox("Impossible de trouver le fichier suivant :" & vbLf &
                sChemin, MsgBoxStyle.Exclamation, sTitreAppli)
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
                    Case 1 : sValeurCorrigee = sChamp.Trim
                    Case 2 : sValeurOrig = sChamp.Trim
                End Select
                If sValeurOrig.Contains("'") Then ' Commentaire à la fin de la ligne
                    Dim iPosQuote% = sValeurOrig.IndexOf("'")
                    sValeurOrig = sValeurOrig.Substring(0, iPosQuote)
                    sValeurOrig = sValeurOrig.TrimEnd
                End If

                ' Vérifier la casse : pas de majuscule dans ces fichiers
                Dim bMajuscule = False
                For Each cCar In sValeurOrig
                    If Char.IsUpper(cCar) Then bMajuscule = True : Exit For
                Next
                For Each cCar In sValeurCorrigee
                    If Char.IsUpper(cCar) Then bMajuscule = True : Exit For
                Next
                If bMajuscule Then
                    MsgBox("Majuscule : " & sLigne & vbLf & IO.Path.GetFileName(sChemin),
                        MsgBoxStyle.Information, "Prénom Mixte")
                End If

            Next

            ' Vérifier les doublons
            Dim sCle$ = sValeurCorrigee & ";" & sValeurOrig
            If hsDoublons.Contains(sCle) Then
                Dim sMsg$ = "Doublon : " & sCle & vbLf & IO.Path.GetFileName(sChemin)
                Debug.WriteLine(sMsg)
                MsgBox(sMsg, MsgBoxStyle.Information, "Prénom Mixte")
            Else
                hsDoublons.Add(sCle)
            End If

            If String.IsNullOrEmpty(sValeurCorrigee) OrElse
               String.IsNullOrEmpty(sValeurOrig) Then Continue For
            If Not dico.ContainsKey(sValeurOrig) Then dico.Add(sValeurOrig, sValeurCorrigee)
        Next

        Return dico

    End Function

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
            sPrenomMEF$, sMEF$,
            Optional iNbMax% = 0)

        Dim sb As New StringBuilder("")
        Dim iNumOcc% = 0
        For Each kvp In lstTxt
            If sb.Length > 0 Then sb.Append(", ")
            Dim sPrenom$ = kvp.Key
            If sPrenom = sPrenomMEF Then sPrenom = sMEF & sPrenom & sMEF
            sb.Append(sPrenom)
            iNumOcc += 1
            If iNbMax > 0 Then
                If iNumOcc >= iNbMax Then sb.Append("...") : Exit For
            End If
        Next

        Return sb.ToString

    End Function

    Private Function sListerCleTxtDico$(
            lstTxt As List(Of KeyValuePair(Of String, clsPrenom)),
            dicoE As DicoTri(Of String, clsPrenom),
            bWiki As Boolean, bHomophoneEnItalique As Boolean)

        Dim sb As New StringBuilder("")
        Dim iNumOcc% = 0
        For Each kvp In lstTxt
            If sb.Length > 0 Then sb.Append(", ")
            Dim sPrenom$ = kvp.Key

            Dim bGras = False
            Dim bItalique = False
            If dicoE.ContainsKey(sPrenom) Then
                Dim prenomE = dicoE(sPrenom)
                If prenomE.bMixteEpicene Then bGras = True
                If prenomE.bMixteHomophone AndAlso bHomophoneEnItalique Then bItalique = True
            End If

            Dim sMEF$ = ""
            If bWiki Then
                If bGras Then sMEF = sGrasWiki
                If bItalique Then sMEF = sItaliqueWiki
                If bItalique AndAlso bGras Then sMEF = sItaliqueGrasWiki
            Else
                If bGras Then sMEF = sGrasMD
                If bItalique Then sMEF = sItaliqueMD
                If bItalique AndAlso bGras Then sMEF = sItaliqueGrasMD
            End If
            sPrenom = sMEF & sPrenom & sMEF

            sb.Append(sPrenom)
            iNumOcc += 1
        Next

        Return sb.ToString

    End Function

End Module