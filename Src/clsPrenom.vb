
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