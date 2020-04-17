' ===================================================
' Version VBA du célèbre jeu 2048
' adapté par :
' - Julien Constant
' - Elios Cama
' - Tristant Chalut Nadal
' - Arthur Muller
' v1.0 - TC1 - A2019
' ===================================================
Option Explicit

Private Sub CommandButton1_Click()

End Sub

' Déplacement avec boutons :
Private Sub Down_button_Click()
    Deplacer_Vers_Le_Bas
End Sub

Private Sub Left_button_Click()
    Deplacer_Vers_La_Gauche
End Sub

Private Sub Right_button_Click()
    Deplacer_Vers_La_Droite
End Sub

Private Sub Upper_button_Click()
    Deplacer_Vers_Le_Haut
End Sub

' Démarrer une nouvelle partie :
Private Sub NouvellePartie_button_Click()
    Demarrer_Nouvelle_Partie
End Sub

' Reset le classement
Private Sub ResetClassement_Click()
    Classement_Reset
End Sub

' Annuler le coup joué :
Private Sub Annuler_Click()
    Annuler_Coup
End Sub

' Déplacement et actions avec les touches du clavier :
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Cells(1, 7).Select
    Application.OnKey "{LEFT}", "Deplacer_Vers_La_Gauche"
    Application.OnKey "{UP}", "Deplacer_Vers_Le_Haut"
    Application.OnKey "{RIGHT}", "Deplacer_Vers_La_Droite"
    Application.OnKey "{DOWN}", "Deplacer_Vers_Le_Bas"
End Sub