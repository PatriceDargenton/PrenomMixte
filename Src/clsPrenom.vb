
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

    Public sPrenom$, sPrenomOrig$, sPrenomHomophone$, sPrenomSimilaire$
    Public sAnnee$, sCodeSexe$, sNbOcc$
    Public bMasc As Boolean
    Public bFem As Boolean
    Public bMixteEpicene As Boolean
    Public bMixteHomophone As Boolean
    Public bSimilaire As Boolean
    Public bUnigenre As Boolean
    Public iNbOccMasc%, iNbOccFem%, iNbOcc%
    Public rFreqRelative#, rFreqRelativeMasc#, rFreqRelativeFem#

    ' Fréquence relative de la variante (homophone ou similaire)
    '  par rapport à la somme des variantes
    Public rFreqRelativeVarianteH#, rFreqRelativeVarianteS#
    Public bVarianteDecompteeH, bVarianteDecompteeS As Boolean

    Public rFreqTotale#, rFreqTotaleMasc#, rFreqTotaleFem#
    Public iAnnee%
    Public rAnneeMoy#, rAnneeMoyMasc#, rAnneeMoyFem#
    Public rAnneeTot#, rAnneeTotMasc#, rAnneeTotFem#
    Public bSelect As Boolean
    Public dicoVariantesH As New DicoTri(Of String, clsPrenom)
    Public dicoVariantesS As New DicoTri(Of String, clsPrenom)

    Public Sub Calculer(iNbPrenomsTot%)
        If iNbPrenomsTot > 0 Then
            Me.rFreqTotale = Me.iNbOcc / iNbPrenomsTot
            Me.rFreqTotaleMasc = Me.iNbOccMasc / iNbPrenomsTot
            Me.rFreqTotaleFem = Me.iNbOccFem / iNbPrenomsTot
        End If
        If Me.iNbOccMasc > 0 Then Me.rAnneeMoyMasc = Me.rAnneeTotMasc / Me.iNbOccMasc
        If Me.iNbOccFem > 0 Then Me.rAnneeMoyFem = Me.rAnneeTotFem / Me.iNbOccFem
        If Me.iNbOcc > 0 Then
            Me.rAnneeMoy = Me.rAnneeTot / Me.iNbOcc
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
        Me.rAnneeTot += prenom1.rAnneeTot
        Me.rAnneeTotMasc += prenom1.rAnneeTotMasc
        Me.rAnneeTotFem += prenom1.rAnneeTotFem
    End Sub

    Public Sub Retirer(prenom1 As clsPrenom)
        Me.iNbOccFem -= prenom1.iNbOccFem
        Me.iNbOccMasc -= prenom1.iNbOccMasc
        Me.iNbOcc -= prenom1.iNbOcc
        Me.rAnneeTot -= prenom1.rAnneeTot
        Me.rAnneeTotMasc -= prenom1.rAnneeTotMasc
        Me.rAnneeTotFem -= prenom1.rAnneeTotFem
    End Sub

End Class