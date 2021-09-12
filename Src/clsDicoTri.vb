
' clsDicoTri.vb : Classe Dictionary triable
' -------------

Imports System.Runtime.Serialization

<Serializable> _
Public Class DicoTri(Of TKey, TValue) : Inherits Dictionary(Of TKey, TValue)

Sub New()
End Sub
Protected Sub New(info As SerializationInfo, context As StreamingContext)
    MyBase.New(info, context)
End Sub

Public Function Trier(Optional sOrdreTri$ = "") As TValue()

    ' Trier la Dico et renvoyer le tableau des éléments triés

    If String.IsNullOrEmpty(sOrdreTri) Then sOrdreTri = ""

    Dim iNbLignes% = Me.Count
    Dim arrayTvalue(iNbLignes - 1) As TValue
    Dim iNumLigne% = 0
    For Each line As KeyValuePair(Of TKey, TValue) In Me
        arrayTvalue(iNumLigne) = line.Value
        iNumLigne += 1
    Next

    ' Si pas de tri demandé, retourner simplement le tableau tel quel
    If sOrdreTri.Length = 0 Then Return arrayTvalue

    ' Tri des éléments
    Dim comp As New UniversalComparer(Of TValue)(sOrdreTri)
    Array.Sort(Of TValue)(arrayTvalue, comp)
    Return arrayTvalue

End Function

End Class



