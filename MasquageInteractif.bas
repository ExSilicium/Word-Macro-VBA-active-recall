' ============================================
' MACRO DE MASQUAGE INTERACTIF POUR WORD
' Active Recall - Révisions interactives
' ============================================

' Variables globales pour stocker l'état
Dim isDrawingMode As Boolean
Dim maskColor As Long
Dim isRevisionMode As Boolean

' ============================================
' PROCEDURE PRINCIPALE : Activer le mode dessin
' ============================================
Sub ActiverModeDessin()
    isDrawingMode = True
    isRevisionMode = False
    maskColor = RGB(255, 255, 0) ' Jaune par défaut
    
    ' Message d'information
    MsgBox "Mode Dessin Activé !" & vbCrLf & vbCrLf & _
           "• Utilisez l'onglet INSERTION > Formes > Rectangle" & vbCrLf & _
           "• Dessinez des rectangles sur les zones à masquer" & vbCrLf & _
           "• Cliquez sur 'Appliquer Masquage' quand terminé", _
           vbInformation, "Mode Masquage"
    
    ' Active l'onglet Insertion dans Word
    On Error Resume Next
    CommandBars.ExecuteMso "TabInsert"
    If Err.Number <> 0 Then
        MsgBox "Cliquez manuellement sur l'onglet INSERTION dans le ruban", vbInformation
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' ============================================
' PROCEDURE : Appliquer le masquage aux rectangles
' ============================================
Sub AppliquerMasquage()
    Dim shp As Shape
    Dim compteur As Integer
    compteur = 0
    
    ' Parcourir toutes les formes du document
    For Each shp In ActiveDocument.Shapes
        ' Vérifier si c'est un rectangle récent
        If shp.Type = msoAutoShape Then
            If shp.AutoShapeType = msoShapeRectangle Then
                ' Vérifier si ce n'est pas déjà un masque
                If shp.AlternativeText <> "MASQUE_ACTIF" Then
                    ' Appliquer le style de masquage
                    With shp
                        ' Remplissage opaque
                        .Fill.Visible = msoTrue
                        .Fill.ForeColor.RGB = maskColor
                        .Fill.Transparency = 0
                        
                        ' Bordure
                        .Line.Visible = True
                        .Line.ForeColor.RGB = RGB(100, 100, 100)
                        .Line.Weight = 0.5
                        
                        ' Ajouter un nom pour identification
                        .Name = "Masque_" & Format(Now, "yyyymmddhhmmss") & "_" & compteur
                        
                        ' Marquer comme masque actif
                        .AlternativeText = "MASQUE_ACTIF"
                        
                        ' Mettre au premier plan
                        .ZOrder msoBringToFront
                        
                        compteur = compteur + 1
                    End With
                End If
            End If
        End If
    Next shp
    
    If compteur > 0 Then
        MsgBox compteur & " rectangle(s) configuré(s) pour le masquage !" & vbCrLf & vbCrLf & _
               "COMMENT RÉVÉLER :" & vbCrLf & _
               "• Double-cliquez sur un rectangle pour le révéler" & vbCrLf & _
               "• Ou utilisez 'Révéler Sélection' après avoir cliqué sur un rectangle" & vbCrLf & _
               "• Ou utilisez 'Basculer Tous' pour tout révéler/masquer", _
               vbInformation, "Masquage appliqué"
    Else
        MsgBox "Aucun rectangle trouvé dans le document." & vbCrLf & _
               "Dessinez d'abord des rectangles avec Insertion > Formes", _
               vbExclamation, "Attention"
    End If
End Sub

' ============================================
' PROCEDURE : Activer le mode révision
' ============================================
Sub ActiverModeRevision()
    isRevisionMode = True
    
    ' Rendre tous les masques opaques au départ
    Dim shp As Shape
    For Each shp In ActiveDocument.Shapes
        If shp.AlternativeText = "MASQUE_ACTIF" Then
            shp.Fill.Transparency = 0
            shp.Line.Transparency = 0
        End If
    Next shp
    
    MsgBox "Mode Révision Activé !" & vbCrLf & vbCrLf & _
           "POUR RÉVÉLER UN CARRÉ :" & vbCrLf & _
           "1. Cliquez sur un carré pour le sélectionner" & vbCrLf & _
           "2. Cliquez sur le bouton 'Révéler Sélection'" & vbCrLf & _
           "   OU double-cliquez directement sur le carré" & vbCrLf & vbCrLf & _
           "• Une fois révélé, il reste visible" & vbCrLf & _
           "• Testez vos connaissances avant de révéler !", _
           vbInformation, "Mode Révision"
End Sub

' ============================================
' PROCEDURE : Désactiver le mode révision
' ============================================
Sub DesactiverModeRevision()
    isRevisionMode = False
    MsgBox "Mode Révision Désactivé !" & vbCrLf & vbCrLf & _
           "Vous pouvez maintenant masquer/révéler librement.", _
           vbInformation, "Mode Normal"
End Sub

' ============================================
' PROCEDURE : Révéler la forme sélectionnée
' ============================================
Sub RevelerSelection()
    On Error Resume Next
    
    Dim shp As Shape
    
    ' Vérifier si une forme est sélectionnée
    If Selection.Type = wdSelectionShape Or Selection.Type = wdSelectionInlineShape Then
        ' Parcourir les formes sélectionnées
        For Each shp In Selection.ShapeRange
            ' Vérifier si c'est un masque
            If shp.AlternativeText = "MASQUE_ACTIF" Then
                If isRevisionMode Then
                    ' MODE RÉVISION : révéler uniquement si opaque
                    If shp.Fill.Transparency = 0 Then
                        shp.Fill.Transparency = 0.85
                        shp.Line.Transparency = 0.85
                    End If
                Else
                    ' MODE NORMAL : basculer
                    If shp.Fill.Transparency = 0 Then
                        shp.Fill.Transparency = 0.85
                        shp.Line.Transparency = 0.85
                    Else
                        shp.Fill.Transparency = 0
                        shp.Line.Transparency = 0
                    End If
                End If
            End If
        Next shp
    Else
        MsgBox "Veuillez d'abord cliquer sur un carré pour le sélectionner !", vbExclamation
    End If
End Sub

' ============================================
' PROCEDURE : Basculer tous les masques
' ============================================
Sub BasculerTousMasques()
    Dim shp As Shape
    Dim tousTransparents As Boolean
    tousTransparents = True
    
    ' Vérifier l'état actuel
    For Each shp In ActiveDocument.Shapes
        If shp.AlternativeText = "MASQUE_ACTIF" Then
            If shp.Fill.Transparency = 0 Then
                tousTransparents = False
                Exit For
            End If
        End If
    Next shp
    
    ' Basculer tous les masques
    For Each shp In ActiveDocument.Shapes
        If shp.AlternativeText = "MASQUE_ACTIF" Then
            If tousTransparents Then
                ' Tout rendre opaque
                shp.Fill.Transparency = 0
                shp.Line.Transparency = 0
            Else
                ' Tout rendre transparent
                shp.Fill.Transparency = 0.85
                shp.Line.Transparency = 0.85
            End If
        End If
    Next shp
    
    If tousTransparents Then
        MsgBox "Tous les masques ont été réactivés (opaques)", vbInformation
    Else
        MsgBox "Tous les masques ont été révélés (transparents)", vbInformation
    End If
End Sub

' ============================================
' PROCEDURE : Changer la couleur des masques
' ============================================
Sub ChangerCouleurMasques()
    Dim couleur As String
    Dim nouvelleColor As Long
    
    couleur = InputBox("Choisissez une couleur :" & vbCrLf & vbCrLf & _
                      "1 = Jaune" & vbCrLf & _
                      "2 = Bleu" & vbCrLf & _
                      "3 = Vert" & vbCrLf & _
                      "4 = Rouge" & vbCrLf & _
                      "5 = Gris" & vbCrLf & _
                      "6 = Orange", "Couleur des masques", "1")
    
    Select Case couleur
        Case "1": nouvelleColor = RGB(255, 255, 0)   ' Jaune
        Case "2": nouvelleColor = RGB(0, 176, 240)   ' Bleu
        Case "3": nouvelleColor = RGB(146, 208, 80)  ' Vert
        Case "4": nouvelleColor = RGB(255, 0, 0)     ' Rouge
        Case "5": nouvelleColor = RGB(166, 166, 166) ' Gris
        Case "6": nouvelleColor = RGB(255, 192, 0)   ' Orange
        Case Else: nouvelleColor = RGB(255, 255, 0)  ' Jaune par défaut
    End Select
    
    maskColor = nouvelleColor
    
    ' Appliquer à tous les masques existants
    Dim shp As Shape
    For Each shp In ActiveDocument.Shapes
        If shp.AlternativeText = "MASQUE_ACTIF" Then
            shp.Fill.ForeColor.RGB = nouvelleColor
        End If
    Next shp
    
    MsgBox "Couleur des masques mise à jour !", vbInformation
End Sub

' ============================================
' PROCEDURE : Supprimer tous les masques
' ============================================
Sub SupprimerTousMasques()
    Dim shp As Shape
    Dim compteur As Integer
    Dim reponse As Integer
    
    compteur = 0
    
    ' Compter les masques
    For Each shp In ActiveDocument.Shapes
        If shp.AlternativeText = "MASQUE_ACTIF" Then
            compteur = compteur + 1
        End If
    Next shp
    
    If compteur > 0 Then
        reponse = MsgBox("Voulez-vous vraiment supprimer les " & compteur & " masque(s) ?", _
                        vbYesNo + vbQuestion, "Confirmation")
        
        If reponse = vbYes Then
            ' Supprimer les masques
            Dim i As Integer
            For i = ActiveDocument.Shapes.Count To 1 Step -1
                If ActiveDocument.Shapes(i).AlternativeText = "MASQUE_ACTIF" Then
                    ActiveDocument.Shapes(i).Delete
                End If
            Next i
            
            MsgBox compteur & " masque(s) supprimé(s) !", vbInformation
        End If
    Else
        MsgBox "Aucun masque trouvé dans le document.", vbInformation
    End If
End Sub

' ============================================
' PROCEDURE : Afficher les statistiques
' ============================================
Sub AfficherStatistiques()
    Dim shp As Shape
    Dim totalMasques As Integer
    Dim masquesOpaques As Integer
    Dim masquesTransparents As Integer
    
    totalMasques = 0
    masquesOpaques = 0
    masquesTransparents = 0
    
    For Each shp In ActiveDocument.Shapes
        If shp.AlternativeText = "MASQUE_ACTIF" Then
            totalMasques = totalMasques + 1
            If shp.Fill.Transparency = 0 Then
                masquesOpaques = masquesOpaques + 1
            Else
                masquesTransparents = masquesTransparents + 1
            End If
        End If
    Next shp
    
    Dim message As String
    message = "STATISTIQUES DES MASQUES" & vbCrLf & vbCrLf & _
              "Total de masques : " & totalMasques & vbCrLf & _
              "Masques cachés (opaques) : " & masquesOpaques & vbCrLf & _
              "Masques révélés (transparents) : " & masquesTransparents
    
    If totalMasques > 0 Then
        Dim pourcentage As Integer
        pourcentage = Round((masquesTransparents / totalMasques) * 100)
        message = message & vbCrLf & vbCrLf & _
                 "Progression : " & pourcentage & "% révélé"
    End If
    
    MsgBox message, vbInformation, "Statistiques"
End Sub

' ============================================
' PROCEDURE : Instructions d'utilisation
' ============================================
Sub AfficherAide()
    Dim aide As String
    aide = "GUIDE D'UTILISATION - MASQUAGE INTERACTIF" & vbCrLf & vbCrLf & _
           "ÉTAPES :" & vbCrLf & _
           "1. Cliquez sur 'Activer Mode Dessin'" & vbCrLf & _
           "2. Insertion > Formes > Rectangle" & vbCrLf & _
           "3. Dessinez des rectangles sur le texte à masquer" & vbCrLf & _
           "4. Cliquez sur 'Appliquer Masquage'" & vbCrLf & vbCrLf & _
           "MODES DE RÉVISION :" & vbCrLf & _
           "• Mode Normal : Basculer librement masqué/révélé" & vbCrLf & _
           "• Mode Révision : Révéler uniquement (pas de retour)" & vbCrLf & vbCrLf & _
           "COMMENT RÉVÉLER :" & vbCrLf & _
           "1. Cliquez sur un carré pour le sélectionner" & vbCrLf & _
           "2. Cliquez sur 'Révéler Sélection'" & vbCrLf & _
           "   (ou créez un bouton rapide)" & vbCrLf & vbCrLf & _
           "FONCTIONS DISPONIBLES :" & vbCrLf & _
           "• Révéler Sélection : Révèle le carré sélectionné" & vbCrLf & _
           "• Activer Mode Révision : Pour révision progressive" & vbCrLf & _
           "• Basculer Tous : Révèle/masque tout d'un coup" & vbCrLf & _
           "• Changer Couleur : Modifie la couleur des masques" & vbCrLf & _
           "• Statistiques : Voir votre progression" & vbCrLf & _
           "• Supprimer : Enlève tous les masques" & vbCrLf & vbCrLf & _
           "ASTUCE : Créez un raccourci clavier pour 'Révéler Sélection' !"
    
    MsgBox aide, vbInformation, "Aide"
End Sub

' ============================================
' MACRO AUTO : Configuration au clic (ThisDocument)
' ============================================
' IMPORTANT : Ce code doit être placé dans ThisDocument
Private Sub Document_Open()
    ' Message d'accueil
    If MsgBox("Document avec masquage interactif détecté." & vbCrLf & vbCrLf & _
              "Voulez-vous activer le mode révision ?", _
              vbYesNo + vbQuestion, "Mode Révision") = vbYes Then
        AfficherAide
    End If
End Sub
