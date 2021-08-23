
' clsPrenom.vb
' ------------

Public Class clsPrenom : Implements ICloneable

    Public Const sPrenomRare$ = "_PRENOMS_RARES"
    Public Const sDateXXXX$ = "XXXX"

    Private Function IClone() As Object Implements ICloneable.Clone
        Return MemberwiseClone()
    End Function
    Public Function Clone() As clsPrenom
        Return DirectCast(Me.IClone(), clsPrenom)
    End Function

    Public sPrenom$, sPrenomOrig$, sPrenomHomophone$
    ' Spécifiquement genré (masc. ou fém., par ex.: antoine, antoinette)
    Public sPrenomSpecifiquementGenre$
    Public sAnnee$, sCodeSexe$, sNbOcc$
    Public bMasc As Boolean
    Public bFem As Boolean
    Public bMixteEpicene As Boolean
    Public bMixteHomophone As Boolean
    Public bSpecifiquementGenre As Boolean
    Public iNbOccMasc%, iNbOccFem%, iNbOcc%
    Public rFreqRelative#, rFreqRelativeMasc#, rFreqRelativeFem#
    Public rFreqTotale#, rFreqTotaleMasc#, rFreqTotaleFem#
    Public iAnnee%, rAnneeMoy#, rAnneeMoyMasc#, rAnneeMoyFem#
    Public bSelect As Boolean
    Public dicoVariantesH As New DicoTri(Of String, clsPrenom)
    Public dicoVariantesG As New DicoTri(Of String, clsPrenom)

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

    Public Sub Ajouter(prenom1 As clsPrenom)
        If Me.bMasc AndAlso prenom1.bFem Then Me.bFem = True
        If Me.bFem AndAlso prenom1.bMasc Then Me.bMasc = True
        Me.iNbOccFem += prenom1.iNbOccFem
        Me.iNbOccMasc += prenom1.iNbOccMasc
        Me.iNbOcc += prenom1.iNbOcc
        Me.rAnneeMoy += prenom1.rAnneeMoy
        Me.rAnneeMoyMasc += prenom1.rAnneeMoyMasc
        Me.rAnneeMoyFem += prenom1.rAnneeMoyFem
    End Sub

End Class