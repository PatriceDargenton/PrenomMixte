﻿
' modUtil.vb
' ----------

Imports System.Text

Public Module modUtil

    Public Function sbLireFichier(sChemin$, Optional bDoublerRAL As Boolean = False) As StringBuilder

        Dim asLignes$() = IO.File.ReadAllLines(sChemin, Encoding.UTF8)
        If IsNothing(asLignes) Then Return New StringBuilder
        Dim sb As New StringBuilder
        For Each sLigne As String In asLignes
            sb.AppendLine(sLigne)
            If bDoublerRAL Then sb.AppendLine()
        Next
        Return sb

    End Function

    Public Function sLireFichier$(sChemin$)

        Return sbLireFichier(sChemin).ToString

    End Function

    Public Function sEnleverAccents$(sChaine$, Optional bMinuscule As Boolean = True)

        ' Enlever les accents

        If sChaine.Length = 0 Then Return ""

        Dim sTexteSansAccents$ = sRemoveDiacritics(sChaine)
        If bMinuscule Then Return sTexteSansAccents.ToLower
        Return sTexteSansAccents

    End Function

    Public Function sRemoveDiacritics$(sTexte$)

        Dim sb As StringBuilder = sbRemoveDiacritics(sTexte)
        Dim sTexteDest$ = sb.ToString
        Return sTexteDest

    End Function

    Public Function sbRemoveDiacritics(sTexte$) As StringBuilder

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

    Public Function asTrierDicoStringString(dico As DicoTri(Of String, String)) As String()

        Dim asTable$(0 To dico.Count - 1)
        dico.Keys.CopyTo(asTable, 0)
        Array.Sort(asTable)

        Return asTable

    End Function

End Module
