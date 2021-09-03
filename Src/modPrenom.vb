
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
    Public Const sDateVersionAppli$ = "03/09/2021"

    Public ReadOnly sVersionAppli$ =
        My.Application.Info.Version.Major & "." &
        My.Application.Info.Version.Minor &
        My.Application.Info.Version.Build

    Public Const sFichierPrenomsInsee$ = "nat2019.csv"
    Public Const sFichierPrenomsInseeCorrige$ = "nat2019_corrige.csv"
    Public Const sFichierPrenomsInseeZip$ = "nat2019_csv.zip"

    ' Seuils de fréquence relative min.
    'Const rSeuilFreqRel# = 0.001 ' 0.1% (par exemple 0.1% de masc. et 99.9% de fém.)
    Const rSeuilFreqRel# = 0.01 ' 1% (par exemple 1% de masc. et 99% de fém.)
    'Const rSeuilFreqRel# = 0.02 ' 2% (par exemple 2% de masc. et 98% de fém.)

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
    Const iSeuilMinPrenomsSpecifiquementGenres% = 20000

    ' Seuil min. pour la détection des prénoms homophones potentiels
    Const iSeuilMinPrenomsHomophonesPotentiels% = 10000
    Const iNbLignesMaxPrenoms% = 0 ' 32346 prénoms en tout (reste quelques accents à corriger)
    Const iDateMinExport% = 1900
    Const iDateMaxExport% = 2019

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

        Dim sCheminDefPrenomsGenres$ = sDossierAppli &
            "\DefinitionsPrenomsSpecifiquementGenres.csv"
        Dim dicoDefinitionsPrenomsGenres = LireFichier(sCheminDefPrenomsGenres)
        Dim dicoDefinitionsPrenomsGenresUtil As New DicoTri(Of String, String)

        ' Ajouter les définitions de prénoms mixtes homophones aux
        '  définitions de prénoms spécifiquement genrés
        ' (le prénom pivot doit être le même, le cas échéant, sinon une alerte sera générée)
        For Each kvp In dicoDefinitionsPrenomsMixtesHomophones
            If Not dicoDefinitionsPrenomsGenres.ContainsKey(kvp.Key) Then
                dicoDefinitionsPrenomsGenres.Add(kvp.Key, kvp.Value)
            End If
        Next

        Dim dicoE As New DicoTri(Of String, clsPrenom) ' épicène
        Dim dicoH As New DicoTri(Of String, clsPrenom) ' homophone
        Dim dicoG As New DicoTri(Of String, clsPrenom) ' masc. ou fém. (spécifiquement genrés)

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
            dicoDefinitionsPrenomsGenres,
            dicoDefinitionsPrenomsGenresUtil,
            dicoE, dicoH, dicoG, bTest,
            iNbLignes, iNbLignesOk, iNbPrenomsTot, iNbPrenomsTotOk,
            iNbPrenomsIgnores, iNbPrenomsIgnoresDate)

        Dim sbCPMD As New StringBuilder
        Dim sbHPMD As New StringBuilder
        Dim sbGPMD As New StringBuilder

        If bTest Then GoTo Export

        DetectionAnomalies(sDossierAppli,
            dicoCorrectionsPrenoms,
            dicoDefinitionsPrenomsMixtesHomophones,
            dicoDefinitionsPrenomsGenres,
            dicoE, sbCPMD, sbHPMD, sbGPMD)

        FiltrerPrenomMixteEpicene(dicoE, iNbPrenomsTot, iSeuilMinPrenomsEpicenes, rSeuilFreqRel,
            iNbPrenomsTotOk, iNbPrenomsIgnores, iNbPrenomsIgnoresDate)

        FiltrerPrenomMixteHomophone(dicoH, dicoE, iNbPrenomsTot,
            iNbPrenomsTotOk, iNbPrenomsIgnores, iNbPrenomsIgnoresDate)

        FiltrerPrenomSpecifiquementGenre(dicoG, dicoE, dicoH, iNbPrenomsTot,
            iNbPrenomsTotOk, iNbPrenomsIgnores, iNbPrenomsIgnoresDate)

Export:
        If bExporter Then
            EcrireFichierFiltre(sDossierAppli, asLignes, dicoE, dicoH, dicoG,
                dicoCorrectionsPrenoms,
                dicoCorrectionsPrenomsUtil,
                dicoDefinitionsPrenomsMixtesHomophones,
                dicoDefinitionsPrenomsMixtesHomophonesUtil,
                dicoDefinitionsPrenomsGenres,
                dicoDefinitionsPrenomsGenresUtil, bTest)
            GoTo Fin
        End If

        Syntheses(sDossierAppli,
            dicoCorrectionsPrenoms,
            dicoCorrectionsPrenomsUtil,
            dicoDefinitionsPrenomsMixtesHomophones,
            dicoDefinitionsPrenomsMixtesHomophonesUtil,
            dicoDefinitionsPrenomsGenres,
            dicoDefinitionsPrenomsGenresUtil,
            dicoE, dicoH, dicoG,
            iNbLignes, iNbLignesOk, iNbPrenomsTot, iNbPrenomsTotOk,
            iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            sbCPMD, sbHPMD, sbGPMD,
            sCheminCorrectionsPrenoms,
            sCheminDefPrenomsMixtesHomophones, sCheminDefPrenomsGenres)

        AnalysePrenomsGenres(sDossierAppli, dicoDefinitionsPrenomsGenres)

Fin:
        If Not bTest Then MsgBox("Terminé !", MsgBoxStyle.Information, sTitreAppli)

    End Sub

    Private Sub AnalyseFichierINSEE(asLignes$(),
            dicoCorrectionsPrenoms As DicoTri(Of String, String),
            dicoCorrectionsPrenomsUtil As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophones As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophonesUtil As DicoTri(Of String, String),
            dicoDefinitionsPrenomsGenres As DicoTri(Of String, String),
            dicoDefinitionsPrenomsGenresUtil As DicoTri(Of String, String),
            dicoE As DicoTri(Of String, clsPrenom),
            dicoH As DicoTri(Of String, clsPrenom),
            dicoG As DicoTri(Of String, clsPrenom),
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
                dicoDefinitionsPrenomsGenres,
                dicoDefinitionsPrenomsGenresUtil) Then Continue For
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

            ' Dico des prénoms masc. ou fém. (spécifiquement genrés)
            Dim prenomG = prenom.Clone()
            If Not prenomG.dicoVariantesG.ContainsKey(prenomG.sPrenom) Then
                prenomG.dicoVariantesG.Add(prenomG.sPrenom, prenom)
            End If
            Dim sCleG$ = prenomG.sPrenomSpecifiquementGenre
            If dicoG.ContainsKey(sCleG) Then
                Dim prenom0 = dicoG(sCleG)
                prenom0.Ajouter(prenom)
                For Each kvp In prenomG.dicoVariantesG
                    If Not prenom0.dicoVariantesG.ContainsKey(kvp.Key) Then
                        prenom0.dicoVariantesG.Add(kvp.Key, prenom)
                    End If
                Next
            Else
                dicoG.Add(sCleG, prenomG)
            End If

        Next

    End Sub

    Private Sub DetectionAnomalies(sDossierAppli$,
        dicoCorrectionsPrenoms As DicoTri(Of String, String),
        dicoDefinitionsPrenomsMixtesHomophones As DicoTri(Of String, String),
        dicoDefinitionsPrenomsGenres As DicoTri(Of String, String),
        dicoE As DicoTri(Of String, clsPrenom),
        ByRef sbCPMD As StringBuilder, ByRef sbHPMD As StringBuilder, ByRef sbGPMD As StringBuilder)

        Dim sdCP As New SortedDictionary(Of String, String) ' Correction de prénoms potentiels
        Dim sdHP As New SortedDictionary(Of String, String) ' Prénoms homophones potentiels
        Dim sdGP As New SortedDictionary(Of String, String) ' Prénoms spécifiquement genrés potentiels

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
            Dim sPrenomF1Min$ = sPrenomF1.ToLower
            If dicoE.ContainsKey(sPrenomF1) AndAlso
               Not dicoDefinitionsPrenomsMixtesHomophones.ContainsKey(sPrenomF1Min) AndAlso
               Not dicoDefinitionsPrenomsMixtesHomophones.ContainsValue(sPrenomF1Min) AndAlso
               Not sdHP.ContainsKey(sPrenomF1) Then
                sdHP.Add(sPrenomF1, prenom.sPrenom)
            End If
            Dim sPrenomF2$ = prenom.sPrenom & "le" ' Ex.: Gabrielle : Gabriel
            Dim sPrenomF2Min$ = sPrenomF2.ToLower
            If dicoE.ContainsKey(sPrenomF2) AndAlso
                Not dicoDefinitionsPrenomsMixtesHomophones.ContainsKey(sPrenomF2Min) AndAlso
                Not dicoDefinitionsPrenomsMixtesHomophones.ContainsValue(sPrenomF2Min) AndAlso
                Not sdHP.ContainsKey(sPrenomF2) Then
                sdHP.Add(sPrenomF2, prenom.sPrenom)
            End If

            ' Détection des prénoms spécifiquement genrés potentiels restants
            Dim sPrenomF3$ = prenom.sPrenom & "tte" ' Ex.: Antoinette : Antoine
            Dim sPrenomF3Min = sPrenomF3.ToLower
            If dicoE.ContainsKey(sPrenomF3) AndAlso
                Not dicoDefinitionsPrenomsGenres.ContainsKey(sPrenomF3Min) AndAlso
                Not dicoDefinitionsPrenomsGenres.ContainsValue(sPrenomF3Min) AndAlso
                Not sdGP.ContainsKey(sPrenomF3) Then
                sdGP.Add(sPrenomF3, prenom.sPrenom)
            End If

            Dim sPrenomF4$ = prenom.sPrenom & "ne" ' Ex.: Fabien : Fabienne
            Dim sPrenomF4Min = sPrenomF4.ToLower
            If dicoE.ContainsKey(sPrenomF4) AndAlso
                Not dicoDefinitionsPrenomsGenres.ContainsKey(sPrenomF4Min) AndAlso
                Not dicoDefinitionsPrenomsGenres.ContainsValue(sPrenomF4Min) AndAlso
                Not sdGP.ContainsKey(sPrenomF4) Then
                sdGP.Add(sPrenomF4, prenom.sPrenom)
            End If
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

        Dim sbGP As New StringBuilder("Liste des prénoms spécifiquement genrés potentiels restants")
        sbGPMD = New StringBuilder("Liste des prénoms spécifiquement genrés potentiels restants")
        sbGP.AppendLine()
        sbGPMD.AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine()
        For Each kvp In sdGP
            Dim sPrenom = kvp.Key
            Dim sPrenomC = kvp.Value
            sbGP.AppendLine(sPrenomC.ToLower & ";" & sPrenom.ToLower)
            sbGPMD.AppendLine(sPrenom & " : " & sPrenomC).AppendLine()
        Next
        Dim sCheminGP$ = sDossierAppli & "\PrenomsSpecifiquementGenresPotentielsRestants.txt"
        EcrireFichier(sCheminGP, sbGP)

    End Sub

    Private Sub Syntheses(sDossierAppli$,
            dicoCorrectionsPrenoms As DicoTri(Of String, String),
            dicoCorrectionsPrenomsUtil As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophones As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophonesUtil As DicoTri(Of String, String),
            dicoDefinitionsPrenomsGenres As DicoTri(Of String, String),
            dicoDefinitionsPrenomsGenresUtil As DicoTri(Of String, String),
            dicoE As DicoTri(Of String, clsPrenom),
            dicoH As DicoTri(Of String, clsPrenom),
            dicoG As DicoTri(Of String, clsPrenom),
            iNbLignes%, iNbLignesOk%,
            iNbPrenomsTot%, iNbPrenomsTotOk%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            sbCPMD As StringBuilder, sbHPMD As StringBuilder, sbGPMD As StringBuilder,
            sCheminCorrectionsPrenoms$,
            sCheminDefPrenomsMixtesHomophones$,
            sCheminDefPrenomsGenres$)

        Dim sbBilan As New StringBuilder

        AfficherSynthesePrenomsFrequents(sDossierAppli, dicoE, iNbPrenomsTotOk,
            iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMinPrenomsFrequents, 0, iNbLignesMaxPrenoms)
        ' Pour le bilan général, conserver l'ordre alphab. pour vérifier la non régression
        AfficherSynthesePrenomsFrequents(sDossierAppli, dicoE, iNbPrenomsTotOk,
            iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMinPrenomsFrequents, 0, iNbLignesMaxPrenoms, sbBilan, bTriAlphab:=True)

        AfficherSyntheseEpicene(sDossierAppli, dicoE, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores,
            iNbPrenomsIgnoresDate, iSeuilMinPrenomsEpicenes, rSeuilFreqRel,
            iNbLignesMaxPrenoms,
            dicoCorrectionsPrenoms, dicoCorrectionsPrenomsUtil)
        AfficherSyntheseEpicene(sDossierAppli, dicoE, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores,
            iNbPrenomsIgnoresDate, iSeuilMinPrenomsEpicenes, rSeuilFreqRel,
            iNbLignesMaxPrenoms,
            dicoCorrectionsPrenoms, dicoCorrectionsPrenomsUtil, sbBilan, bTriAlphab:=True)

        AfficherSyntheseHomophone(sDossierAppli, dicoH, dicoE, iNbPrenomsTotOk, iNbPrenomsTot,
            iNbPrenomsIgnores, iNbPrenomsIgnoresDate, iSeuilMinPrenomsHomophones, 0,
            iNbLignesMaxPrenoms,
            dicoDefinitionsPrenomsMixtesHomophones,
            dicoDefinitionsPrenomsMixtesHomophonesUtil)
        AfficherSyntheseHomophone(sDossierAppli, dicoH, dicoE, iNbPrenomsTotOk, iNbPrenomsTot,
            iNbPrenomsIgnores, iNbPrenomsIgnoresDate, iSeuilMinPrenomsHomophones, 0,
            iNbLignesMaxPrenoms,
            dicoDefinitionsPrenomsMixtesHomophones,
            dicoDefinitionsPrenomsMixtesHomophonesUtil, sbBilan, bTriAlphab:=True)

        AfficherSyntheseSpecifiquementGenre(sDossierAppli,
            dicoG, dicoE, dicoH, iNbPrenomsTotOk, iNbPrenomsTot,
            iNbPrenomsIgnores, iNbPrenomsIgnoresDate, iSeuilMinPrenomsSpecifiquementGenres, 0,
            iNbLignesMaxPrenoms,
            dicoDefinitionsPrenomsGenres,
            dicoDefinitionsPrenomsGenresUtil)
        AfficherSyntheseSpecifiquementGenre(sDossierAppli,
            dicoG, dicoE, dicoH, iNbPrenomsTotOk, iNbPrenomsTot,
            iNbPrenomsIgnores, iNbPrenomsIgnoresDate, iSeuilMinPrenomsSpecifiquementGenres, 0,
            iNbLignesMaxPrenoms,
            dicoDefinitionsPrenomsGenres,
            dicoDefinitionsPrenomsGenresUtil, sbBilan, bTriAlphab:=True)

        sbBilan.AppendLine("Liste des corrections d'accent").
            AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine()
        sbBilan.Append(sbLireFichier(sCheminCorrectionsPrenoms, bDoublerRAL:=True)).
            AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()

        sbBilan.AppendLine("Liste des définitions de prénoms mixtes homophones").
            AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine()
        sbBilan.Append(sbLireFichier(sCheminDefPrenomsMixtesHomophones, bDoublerRAL:=True)).
            AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()

        sbBilan.AppendLine(
            "Liste des définitions de prénoms masculins ou féminins (spécifiquement genrés)").
            AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine()
        sbBilan.Append(sbLireFichier(sCheminDefPrenomsGenres, bDoublerRAL:=True)).
            AppendLine().AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()

        sbBilan.Append(sbCPMD).AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()
        sbBilan.Append(sbHPMD).AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()
        sbBilan.Append(sbGPMD)

        Dim sCheminBilan$ = sDossierAppli & "\Bilan.md"
        EcrireFichier(sCheminBilan, sbBilan)

    End Sub

    Private Sub AnalysePrenomsGenres(sDossierAppli$,
            dicoDefinitionsPrenomsGenres As DicoTri(Of String, String))

        Dim sbPG As New StringBuilder(
            "Dictionnaire des prénoms spécifiquement genrés + mixtes homophones")
        sbPG.AppendLine()

        Dim asTable = asTrierDicoStringString(dicoDefinitionsPrenomsGenres)
        For Each sPrenom In asTable
            sbPG.AppendLine(dicoDefinitionsPrenomsGenres(sPrenom).ToLower & ";" & sPrenom.ToLower)
        Next

        ' Vérifier que les clés sont bien dans le même ordre lors de la fusion des 2 dico
        ' -> Choisir la même clé de regroupement dans ces 2 dico pour résoudre ces problèmes

        ' Exemple, si on a :
        ' ----------------

        ' DefinitionsPrenomsMixtesHomophones.csv :
        ' pascal;pascale
        ' pascal;pasquale

        ' DefinitionsPrenomsGenres.csv :
        ' pascaline;pascal

        ' Alors pascal est signalé comme étant une clé inversée dans le dico fusionné
        ' -> Changer en pascal;pascaline dans DefinitionsPrenomsGenres.csv
        ' ----------------

        Dim hsCles As New HashSet(Of String)
        For Each kvp In dicoDefinitionsPrenomsGenres
            hsCles.Add(kvp.Key)
        Next
        For Each kvp In dicoDefinitionsPrenomsGenres
            If hsCles.Contains(kvp.Value) Then
                MsgBox("Fusion des prénoms spécifiquement genrés + mixtes homophones :" & vbLf &
                    "clé inversée : " & kvp.Value, MsgBoxStyle.Exclamation, sTitreAppli)
                sbPG.AppendLine(
                    "Fusion des prénoms spécifiquement genrés + mixtes homophones : clé inversée : " &
                    kvp.Value)
            End If
        Next

        Dim sCheminPG$ = sDossierAppli & "\PrenomsSpecifiquementGenresEtMixtesHomophones.txt"
        EcrireFichier(sCheminPG, sbPG)

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
                If dicoE.ContainsKey(prenom.sPrenomHomophone) Then
                    Dim prenomG = dicoE(prenom.sPrenomHomophone)
                    prenomG.bMixteHomophone = True
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

    Private Sub FiltrerPrenomSpecifiquementGenre(
            dicoG As DicoTri(Of String, clsPrenom),
            dicoE As DicoTri(Of String, clsPrenom),
            dicoH As DicoTri(Of String, clsPrenom),
            iNbPrenomsTot%, iNbPrenomsTotOk%, iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%)

        Dim iNbPrenomsVerif% = 0
        Dim iNbPrenomsVerifMF% = 0
        Dim aPrenomsG = dicoG.Trier("")
        For Each prenom In aPrenomsG

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

            'If prenom.dicoVariantesG.Count > 1 Then prenom.bSpecifiquementGenre = True
            ' Définition de spécifiquement genré : au moins une variante non homophone ni épicène
            If prenom.dicoVariantesG.Count > 1 Then
                Dim bTousbMixtesHOuE = True
                For Each prenom0 In prenom.dicoVariantesG.Trier()
                    Dim sPrenomH = prenom0.sPrenomHomophone
                    If dicoH.ContainsKey(sPrenomH) Then
                        Dim prenom1 = dicoH(sPrenomH)
                        If Not prenom1.bMixteHomophone AndAlso
                           Not prenom1.bMixteEpicene Then bTousbMixtesHOuE = False : Exit For
                    End If
                Next
                If Not bTousbMixtesHOuE Then
                    prenom.bSpecifiquementGenre = True
                    ' Marquer aussi l'original pour l'export
                    If dicoE.ContainsKey(prenom.sPrenomSpecifiquementGenre) Then
                        Dim prenomG = dicoE(prenom.sPrenomSpecifiquementGenre)
                        prenomG.bSpecifiquementGenre = True
                    End If
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
        For Each prenom In aPrenomsG
            If prenom.dicoVariantesG.Count <= 1 Then Continue For
            For Each kvp In prenom.dicoVariantesG
                Dim prenomH = kvp.Value
                prenomH.rFreqRelativeVarianteG = prenomH.iNbOcc / prenom.iNbOcc
                If prenomH.rFreqRelativeVarianteG < rSeuilFreqRelVariante Then
                    prenomH.bVarianteDecompteeG = True
                End If
            Next
        Next

        ' Retirer les variantes trop minoritaires
        Dim lstPrenomsRetires As New List(Of clsPrenom)
        For Each prenom In aPrenomsG
            If prenom.dicoVariantesG.Count <= 1 Then Continue For
            Dim lst As New List(Of String)
            For Each kvp In prenom.dicoVariantesG
                Dim prenomH = kvp.Value
                If prenomH.bVarianteDecompteeG Then
                    prenom.Retirer(prenomH)
                    prenom.Calculer(iNbPrenomsTot)
                    lst.Add(kvp.Key)
                    lstPrenomsRetires.Add(prenomH)
                End If
            Next
            For Each sCle In lst
                prenom.dicoVariantesG.Remove(sCle)
            Next
            If prenom.dicoVariantesG.Count <= 1 Then prenom.bSpecifiquementGenre = False
        Next

        ' Vérifier le nouveau calcul
        iNbPrenomsVerif = 0
        iNbPrenomsVerifMF = 0
        For Each prenom In aPrenomsG
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

    Private Sub AfficherSynthesePrenomsFrequents(sDossierAppli$,
            dicoE As DicoTri(Of String, clsPrenom),
            iNbPrenomsTotOk%, iNbPrenomsTot%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, iNbLignesMax%,
            Optional sbBilan As StringBuilder = Nothing, Optional bTriAlphab As Boolean = False)

        ' Produire la synthèse statistique des prénoms fréquents

        Dim sb As New StringBuilder
        AfficherInfo(sb, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante)
        Dim sbMD As New StringBuilder ' Syntaxe MarkDown
        sbMD.AppendLine("Synthèse statistique des prénoms fréquents")
        sbMD.AppendLine()
        AfficherInfo(sbMD, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante, bDoublerRAL:=True)
        sbMD.AppendLine(sEnteteMarkDown())

        Dim sbWK As New StringBuilder ' Syntaxe Wiki
        AfficherInfo(sbWK, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante, bDoublerRAL:=True)
        sbWK.AppendLine(sEnteteWiki("Synthèse statistique des prénoms fréquents"))

        Dim iNbPrenoms% = 0
        Dim sTri$ = "iNbOcc desc"
        If bTriAlphab Then sTri = "sPrenom"
        For Each prenom In dicoE.Trier(sTri)

            If iSeuilMin > 0 AndAlso prenom.iNbOcc < iSeuilMin Then Continue For

            iNbPrenoms += 1
            If iNbLignesMax > 0 AndAlso iNbPrenoms > iNbLignesMax Then Exit For

            prenom.bSelect = True

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
        Dim sChemin$ = sDossierAppli & "\PrenomsFrequents" & sSuffixe & ".md"
        EcrireFichier(sChemin, sbMD)
        Dim sCheminWK$ = sDossierAppli & "\PrenomsFrequents" & sSuffixe & ".wiki"
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
                sPrenomMD = sListerCleTxtDico(lst, dicoE, dicoH, bWiki:=False)
                sPrenomWiki = sListerCleTxtDico(lst, dicoE, dicoH, bWiki:=True)

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
                    If dicoH.ContainsKey(prenomV.sPrenom) Then
                        Dim prenomH = dicoH(prenomV.sPrenom)
                        If prenomH.bMixteHomophone Then bItalique = True
                    End If
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

    Private Sub AfficherSyntheseSpecifiquementGenre(sDossierAppli$,
            dicoG As DicoTri(Of String, clsPrenom),
            dicoE As DicoTri(Of String, clsPrenom),
            dicoH As DicoTri(Of String, clsPrenom),
            iNbPrenomsTotOk%, iNbPrenomsTot%,
            iNbPrenomsIgnores%, iNbPrenomsIgnoresDate%,
            iSeuilMin%, rSeuilFreqRel!, iNbLignesMax%,
            dicoDefinitionsPrenomsGenres As DicoTri(Of String, String),
            dicoDefinitionsPrenomsGenresUtil As DicoTri(Of String, String),
            Optional sbBilan As StringBuilder = Nothing, Optional bTriAlphab As Boolean = False)

        ' Produire la synthèse statistique des prénoms masculins ou féminins (spécifiquement genrés)

        Dim sb As New StringBuilder
        AfficherInfo(sb, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante)

        Dim sbMD As New StringBuilder ' Syntaxe MarkDown
        sbMD.AppendLine(
            "Synthèse statistique des prénoms masculins ou féminins (spécifiquement genrés)")
        sbMD.AppendLine()
        AfficherInfo(sbMD, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante, bDoublerRAL:=True)
        sbMD.AppendLine(sEnteteMarkDown(bColonneFreqVariante:=True))

        Dim sbWK As New StringBuilder ' Syntaxe Wiki
        AfficherInfo(sbWK, iNbPrenomsTotOk, iNbPrenomsTot, iNbPrenomsIgnores, iNbPrenomsIgnoresDate,
            iSeuilMin, rSeuilFreqRel, rSeuilFreqRelVariante, bDoublerRAL:=True)
        sbWK.AppendLine(sEnteteWiki(
            "Synthèse statistique des prénoms masculins ou féminins (spécifiquement genrés)",
            bColonneFreqVariante:=True))

        Dim iNbPrenomsGenres% = 0
        Dim sTri$ = "bSpecifiquementGenre desc, rFreqTotale desc"
        If bTriAlphab Then sTri = "sPrenom"
        For Each prenom In dicoG.Trier(sTri)

            If Not prenom.bSpecifiquementGenre Then Continue For
            If iSeuilMin > 0 AndAlso prenom.iNbOcc < iSeuilMin Then Continue For

            iNbPrenomsGenres += 1
            If iNbLignesMax > 0 AndAlso iNbPrenomsGenres > iNbLignesMax Then Exit For

            prenom.bSelect = True

            Dim sPrenom$ = prenom.sPrenomSpecifiquementGenre
            Dim sPrenomMD$ = sPrenom
            Dim sPrenomWiki$ = sPrenom
            Dim bVariantes = False
            If prenom.dicoVariantesG.Count > 1 Then
                bVariantes = True
                Dim lst = prenom.dicoVariantesG.ToList
                ' Mêmes conditions que pour la liste des prénoms fréquents :
                ' Gras : épicène
                ' Italique : homophone
                ' Gras+Italique : épicène + homophone
                sPrenomMD = sListerCleTxtDico(lst, dicoE, dicoH, bWiki:=False)
                sPrenomWiki = sListerCleTxtDico(lst, dicoE, dicoH, bWiki:=True)
            End If

            sb.AppendLine(sLigneDebug(prenom, sPrenom, iNbPrenomsGenres, sFormatFreq))
            sbMD.AppendLine(sLigneMarkDown(prenom, sPrenomMD, iNbPrenomsGenres, sFormatFreq,
                iNumVariante:=0, bSuffixeNumVariante:=True))
            sbWK.AppendLine(sLigneWiki(prenom, sPrenomWiki, iNbPrenomsGenres, sFormatFreq,
                iNumVariante:=0, bSuffixeNumVariante:=True))

            If bVariantes Then
                Dim iNumVariante% = 0
                For Each prenomV In prenom.dicoVariantesG.Trier("iNbOcc desc")
                    If prenomV.bVarianteDecompteeG Then Continue For
                    iNumVariante += 1
                    sb.AppendLine(sLigneDebug(prenomV, prenomV.sPrenom, iNbPrenomsGenres, sFormatFreq))
                    Dim bGras = False
                    If dicoE.ContainsKey(prenomV.sPrenom) Then
                        Dim prenomE = dicoE(prenomV.sPrenom)
                        If prenomE.bMixteEpicene Then bGras = True
                    End If
                    Dim bItalique = False
                    If dicoH.ContainsKey(prenomV.sPrenom) Then
                        Dim prenomH = dicoH(prenomV.sPrenom)
                        If prenomH.bMixteHomophone Then bItalique = True
                    End If
                    sbMD.AppendLine(sLigneMarkDown(prenomV, prenomV.sPrenom, iNbPrenomsGenres,
                        sFormatFreq, iNumVariante, bGras, bItalique,
                        bSuffixeNumVariante:=True, bColonneFreqVarianteG:=True))
                    sbWK.AppendLine(sLigneWiki(prenomV, prenomV.sPrenom, iNbPrenomsGenres,
                        sFormatFreq, iNumVariante, bGras, bItalique,
                        bSuffixeNumVariante:=True, bColonneFreqVarianteG:=True))
                Next
            End If

        Next
        sbWK.AppendLine("|}")

        sb.AppendLine()
        sbMD.AppendLine()
        sbWK.AppendLine()
        For Each kvp In dicoDefinitionsPrenomsGenres
            If Not dicoDefinitionsPrenomsGenresUtil.ContainsKey(kvp.Key) Then
                Dim sLigne$ = "Correction de prénom (liste spécifiquement genrée) non trouvée : " & kvp.Key
                sb.AppendLine(sLigne)
                sbMD.AppendLine(sLigne)
                sbWK.AppendLine(sLigne)
            End If
        Next

        'Debug.WriteLine(sb.ToString)

        If bTriAlphab Then
            sbBilan.Append(sbMD).AppendLine().AppendLine(sSautDeLigneMD).AppendLine(sSautDeLigneMD).AppendLine()
        Else
            Dim sCheminMD$ = sDossierAppli & "\PrenomsSpecifiquementGenres.md"
            EcrireFichier(sCheminMD, sbMD)
            Dim sCheminWK$ = sDossierAppli & "\PrenomsSpecifiquementGenres.wiki"
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
            ", freq. rel. m. " & sGenre & prenom.rFreqRelativeMasc.ToString("0%") &
            ", freq. rel. f. " & sGenre & prenom.rFreqRelativeFem.ToString("0%") &
            ", mixte épicène=" & prenom.bMixteEpicene
        Return s

    End Function

    Private Function sLigneMarkDown$(prenom As clsPrenom, sPrenom$, iNumPrenom%, sFormatFreq$,
            Optional iNumVariante% = -1,
            Optional bGras As Boolean = False,
            Optional bItalique As Boolean = False,
            Optional bSuffixeNumVariante As Boolean = False,
            Optional bColonneFreqVarianteH As Boolean = False,
            Optional bColonneFreqVarianteG As Boolean = False)

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
        If bColonneFreqVarianteG Then
            sColonneFreqVariante = "|" & prenom.rFreqRelativeVarianteG.ToString(sFormatFreqRelVariante)
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
            "|" & prenom.rFreqRelativeMasc.ToString("0%") &
            "|" & prenom.rFreqRelativeFem.ToString("0%") &
            sColonneFreqVariante

        Return s

    End Function

    Private Function sLigneWiki$(prenom As clsPrenom, sPrenom$, iNumPrenom%, sFormatFreq$,
            Optional iNumVariante% = -1,
            Optional bGras As Boolean = False,
            Optional bItalique As Boolean = False,
            Optional bSuffixeNumVariante As Boolean = False,
            Optional bColonneFreqVarianteH As Boolean = False,
            Optional bColonneFreqVarianteG As Boolean = False)

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
        If bColonneFreqVarianteG Then
            sColonneFreqVariante = "||" & prenom.rFreqRelativeVarianteG.ToString(sFormatFreqRelVariante)
        End If

        Dim s$ = "|-" & vbLf &
                "|" & iNumPrenom & sNumVariante &
                "|| align='right' | " & sFormaterNumWiki(prenom.iNbOcc) &
                "||" & sMiseEnForme & sPrenom & sMiseEnForme &
                "||" & prenom.rAnneeMoy.ToString("0") &
                "||" & prenom.rAnneeMoyMasc.ToString("0") &
                "||" & prenom.rAnneeMoyFem.ToString("0") &
                "|| align='right' | " & sFormaterNumWiki(prenom.iNbOccMasc) &
                "|| align='right' | " & sFormaterNumWiki(prenom.iNbOccFem) &
                "||" & prenom.rFreqTotale.ToString(sFormatFreq) &
                "||" & prenom.rFreqRelativeMasc.ToString("0%") &
                "||" & prenom.rFreqRelativeFem.ToString("0%") &
                sColonneFreqVariante
        Return s

    End Function

    Private Function bAnalyserPrenom(sLigne$, prenom As clsPrenom,
            dicoCorrectionsPrenoms As DicoTri(Of String, String),
            dicoCorrectionsPrenomsUtil As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophones As DicoTri(Of String, String),
            dicoDefinitionsPrenomsMixtesHomophonesUtil As DicoTri(Of String, String),
            dicoDefinitionsPrenomsGenres As DicoTri(Of String, String),
            dicoDefinitionsPrenomsGenresUtil As DicoTri(Of String, String)) As Boolean

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

        ' Prénoms spécifiquement genrés (par ex.: antoinette : féminin de antoine)
        Dim sPrenomSpecifiquementGenre = sPrenom
        If dicoDefinitionsPrenomsGenres.ContainsKey(sPrenom) Then
            Dim sPrenomG$ = dicoDefinitionsPrenomsGenres(sPrenom)
            If Not dicoDefinitionsPrenomsGenresUtil.ContainsKey(sPrenom) Then
                dicoDefinitionsPrenomsGenresUtil.Add(sPrenom, sPrenomG)
            End If
            sPrenomSpecifiquementGenre = sPrenomG
        End If

        prenom.sPrenom = FirstCharToUpper(sPrenom)
        prenom.sPrenomHomophone = FirstCharToUpper(sPrenomHomophone)
        prenom.sPrenomSpecifiquementGenre = FirstCharToUpper(sPrenomSpecifiquementGenre)
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

    Private Sub EcrireFichierFiltre(sDossierAppli$, asLignes$(),
        dicoE As DicoTri(Of String, clsPrenom),
        dicoH As DicoTri(Of String, clsPrenom),
        dicoG As DicoTri(Of String, clsPrenom),
        dicoCorrectionsPrenoms As DicoTri(Of String, String),
        dicoCorrectionsPrenomsUtil As DicoTri(Of String, String),
        dicoDefinitionsPrenomsMixtesHomophones As DicoTri(Of String, String),
        dicoDefinitionsPrenomsMixtesHomophonesUtil As DicoTri(Of String, String),
        dicoDefinitionsPrenomsGenres As DicoTri(Of String, String),
        dicoDefinitionsPrenomsGenresUtil As DicoTri(Of String, String),
        bTestPrenomOrig As Boolean)

        ' Génération d'un nouveau fichier csv filtré ou pas

        ' bTestPrenomOrig : Vérifier si le traitement appliqué préserve entièrement le fichier d'origine
        Const bFiltrerPrenomEpicene = False

        Dim sb As New StringBuilder
        Dim iNbLignes = 0
        Dim sAjoutEntete$ = ""
        If Not bTestPrenomOrig Then sAjoutEntete = ";Prénom d'origine;Prénom épicène;Prénom homophone;Prénom masc. ou fém."
        For Each sLigne As String In asLignes

            iNbLignes += 1
            If iNbLignes = 1 Then sb.AppendLine(sLigne & sAjoutEntete) : Continue For

            Dim prenom As New clsPrenom
            If Not bAnalyserPrenom(sLigne$, prenom,
                dicoCorrectionsPrenoms,
                dicoCorrectionsPrenomsUtil,
                dicoDefinitionsPrenomsMixtesHomophones,
                dicoDefinitionsPrenomsMixtesHomophonesUtil,
                dicoDefinitionsPrenomsGenres,
                dicoDefinitionsPrenomsGenresUtil) Then Continue For

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
                    If prenom.sPrenomOrig <> sPrenom Then
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
            Dim sPrenomG$ = ""
            Dim sCleE$ = prenom.sPrenom
            'Dim bMixteEpicene = False
            'Dim bMixteHomophone = False
            'Dim bSpecifiquementGenre = False
            If dicoE.ContainsKey(sCleE) Then
                Dim prenom0 = dicoE(sCleE)
                If bFiltrerPrenomEpicene AndAlso Not prenom0.bSelect Then Continue For
                If prenom0.bMixteEpicene Then sPrenomE = "1" ': bMixteEpicene = True
                If prenom0.bMixteHomophone Then sPrenomH = prenom0.sPrenomHomophone ': bMixteHomophone = True
                If prenom0.bSpecifiquementGenre Then
                    sPrenomG = prenom0.sPrenomSpecifiquementGenre ': bSpecifiquementGenre = True
                End If
            Else
                Continue For
            End If
            'If bMixteEpicene Then GoTo Ajout

            Dim sCleH$ = prenom.sPrenomHomophone
            If dicoH.ContainsKey(sCleH) Then
                Dim prenom0 = dicoH(sCleH)
                If prenom0.bMixteHomophone Then sPrenomH = prenom0.sPrenomHomophone ': bMixteHomophone = True
            End If
            'If bMixteHomophone Then GoTo Ajout

            Dim sCleG$ = prenom.sPrenomSpecifiquementGenre
            If dicoG.ContainsKey(sCleG) Then
                Dim prenom0 = dicoG(sCleG)
                If prenom0.bSpecifiquementGenre Then
                    sPrenomG = prenom0.sPrenomSpecifiquementGenre
                    ' bSpecifiquementGenre = True
                End If
            End If

Ajout:
            If Not bTestPrenomOrig Then sAjout = ";" & prenom.sPrenomOrig & ";" & sPrenomE & ";" & sPrenomH & ";" & sPrenomG

Suite:
            Dim sLigneC$ = prenom.sCodeSexe & ";" & sPrenom & ";" & prenom.sAnnee & ";" & prenom.iNbOcc & sAjout
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
            sb.AppendLine("Fréquence relative min. genre = " & rSeuilFreqRel.ToString("0%"))
            If bDoublerRAL Then sb.AppendLine("")
        End If
        If rSeuilFreqRelVariante > 0 Then
            sb.AppendLine("Fréquence relative min. variante = " &
                rSeuilFreqRelVariante.ToString(sFormatFreqRelVariante))
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
            dicoH As DicoTri(Of String, clsPrenom),
            bWiki As Boolean)

        Dim sb As New StringBuilder("")
        Dim iNumOcc% = 0
        For Each kvp In lstTxt
            If sb.Length > 0 Then sb.Append(", ")
            Dim sPrenom$ = kvp.Key

            Dim bGras = False
            If dicoE.ContainsKey(sPrenom) AndAlso dicoE(sPrenom).bMixteEpicene Then bGras = True
            Dim bItalique = False
            If dicoH.ContainsKey(sPrenom) AndAlso dicoH(sPrenom).bMixteHomophone Then bItalique = True

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