' ===================================================
' Version VBA du célèbre jeu 2048
' adapté par :
' - Julien Constant
' - Elios Cama
' - Tristant Chalut Nadal
' - Arthur Muller
' v1.0 - TC1 - A2019
' ===================================================

' ---------------------------------------------------
' PRE-CODE
' ---------------------------------------------------
Option Explicit
Option Base 0

' Probabilité d'obtenir un 2 sur la grille :
' - fixée arbitrairement à 90%
' - sinon c'est un 4 qui sort
Public Const Proba_2 As Single = 0.9

' Définition de l'objectif
' - ici 2^11 = 2048 (les cases sont stockées en puissance de deux
Public Const Cible As Byte = 11

' Définition des vérifications :
' - Partie_Bloquee vérifie s'il y encore possibilité de bouger
' - Cible_Atteint vérifie la présence de l'objectif ou non
Public Partie_Bloquee As Boolean, Cible_Atteint As Boolean

' Création de la grille de jeu (4x4)
Public Grille_Principale(3, 3) As Byte

' Création de variables d'incrémentation du score, du nombre de mouvement et de la sauvegarde de la grille (Etats())
Public Etats() As Long, Mouvement As Long, Score As Long

' Définition des variables suivantes comme étant des plages de données
Dim Grille_Affichage As Range, Label_Score As Range, Label_Mouvement As Range

' Définition de la couleurs des cases
Dim Couleurs_R() As Variant
Dim Couleurs_G() As Variant
Dim Couleurs_B() As Variant

' Définition du classement
Dim Joueur As Variant
Dim Label_Joueur As Range
Dim Ligne As Integer
Dim nbr_Joueur As Variant
Dim Label_nbr_Joueur As Range

Dim Label_Classement As Range

' Effacement du classement
Dim Label_Efface As Range


' ---------------------------------------------------
' FONCTION POUR RESET ET DEMARRER UNE NOUVELLE PARTIE
' ---------------------------------------------------
Public Sub Demarrer_Nouvelle_Partie()
    Dim i As Byte, j As Byte
    
    If Ligne = 0 Then
        Ligne = 10
    End If
        nbr_Joueur = Ligne - 8
        With Worksheets("2048")
            Set Label_nbr_Joueur = .Range("D8")
            Label_nbr_Joueur = nbr_Joueur
        End With
    
    ' Initialisation des variables et des couleurs associées à chaque nombre :
    ' - le mouvement commence à moins 1 car la génération du premier 2 est considérée comme étant un mouvement
    Partie_Bloquee = False
    Cible_Atteint = False
    Mouvement = -1
    
    If Not Partie_Bloquee Then
        If MsgBox("Une partie est en cours, voulez-vous en lancer une autre?", vbYesNo, "2048") = vbNo Then
            Exit Sub
        Else
            Joueur = InputBox("Nouveau joueur :", "2048")
            With Worksheets("2048")
                
                ' Inscription du joueur et mémo du score précédent
                If .Range("B9") = "" Then
                    ' Inscription du 1er joueur
                    Set Label_Joueur = .Range("B9")
                Else
                    ' Inscription du n-ième joueur
                    ' Mémorisation du score du joueur précédent
                    Set Label_Joueur = .Range("B" & Ligne)
                    Set Label_Score = .Range("E" & Ligne - 1)
                    Label_Score = Score
                    Ligne = Ligne + 1
                End If
                
            End With
            Label_Joueur = Joueur
        End If
    End If

    '                    0    1    2    3     4    5    6    7    8    9   10   11   12   13    14    15    16     17
    '                    0    2    4    8    16   32   64  128  256  512 1024 2048 4096 8192 16384 32768 65536 131072
    Couleurs_R = Array(208, 153, 102, 51, 0, 0, 0, 0, 0, 0, 25, 51, 51, 102, 153, 204, 255, 255)
    Couleurs_G = Array(206, 204, 178, 153, 128, 102, 76, 51, 25, 0, 0, 0, 0, 0, 0, 0, 0, 51)
    Couleurs_B = Array(206, 255, 255, 255, 255, 204, 153, 102, 51, 51, 51, 51, 25, 51, 76, 102, 127, 153)
    
    Score = 0
    
    ' On décide de l'endroit ou placer les labels du score, de la grille, et du nombre de mouvements
    With Worksheets("2048")
        Set Grille_Affichage = .Range("B3:E6")
        Set Label_Mouvement = .Range("H2")
        Set Label_Score = .Range("G2")
        .Cells(10, 10).Select
    End With
    
    ' On vide la grille
    For i = 0 To 3
        For j = 0 To 3
            Grille_Principale(i, j) = 0
        Next j
    Next i
    
    Placer_Nombre
    Afficher_Grille
End Sub

' ---------------------------------------------------
' RESET DU CLASSEMENT
' ---------------------------------------------------
Public Sub Classement_Reset()
    Dim i As Byte, j As Byte
    
    For i = 9 To 21
        For j = 2 To 5
            Cells(i, j) = ""
        Next j
    Next i
    
    Demarrer_Nouvelle_Partie
    
End Sub

' ---------------------------------------------------
' CONVERSION ET AFFICHAGE DES PUISSANCE DE DEUX
' ---------------------------------------------------
' - Etants stockés en puissance de deux il faut les convertir en "vrai" nombre
' - Changement de couleur de la police en fonction de la couleur de la case
' ---------------------------------------------------
Public Sub Afficher_Grille()
    Dim i As Byte, j As Byte
    Dim Valeur As Byte
    Application.ScreenUpdating = False
    For i = 0 To 3
        For j = 0 To 3
            Valeur = Grille_Principale(i, j)
            With Grille_Affichage(i + 1, j + 1)
                If Valeur = 0 Then
                    .Value = ""
                Else
                    .Value = 2 ^ Valeur
                End If
                
                ' On choisit la couleur de la case et on change la couleur de la police en fonction
                .Interior.Color = RGB(Couleurs_R(Valeur), Couleurs_G(Valeur), Couleurs_B(Valeur))
                Select Case Valeur
                    Case 1, 2, 3
                        .Font.Color = vbBlack
                    Case Else
                        .Font.Color = vbWhite
                End Select
            End With
        Next j
    Next i
    
    ' On met à jour le score, le mouvement, et on réactive la mise à jour de "l'écran"
    Label_Score = "Score : " & Score
    Label_Mouvement = "Déplacements : " & Mouvement
    Application.ScreenUpdating = True

End Sub

' ---------------------------------------------------
' CHOIX DU NOMBRE ET DE LA CASE A REMPLIR AVEC
' ---------------------------------------------------
' - 90% de chance d'avoir un 2 / 10% d'avoir un 4
' - Un nombre apparait que si la grille évolue
'       (si le joueur se déplace contre un mur rien n'apparait)
' - On enregistre la grille dans un tableau
'       (afin de pouvoir annuler le dernier coup)
' ---------------------------------------------------
Public Sub Placer_Nombre()
    Dim i As Integer, j As Integer
    Randomize

    ' On prend une case de la grille au hasard
    i = Int(Rnd * 4)
    j = Int(Rnd * 4)

    While Grille_Principale(i, j) <> 0
        i = Int(Rnd * 4)
        j = Int(Rnd * 4)
    Wend

    If Rnd <= Proba_2 Then
        ' 2^1 = 2
        Grille_Principale(i, j) = 1
    Else
        ' 2^2 = 4
        Grille_Principale(i, j) = 2
    End If

    Mouvement = Mouvement + 1
    
    ' On enregistre l'état de la grille pour pouvoir revenir en arrière
    ReDim Preserve Etats(16, Mouvement)
    For i = 0 To 15
        Etats(i, Mouvement) = Grille_Principale((i - (i Mod 4)) / 4, i Mod 4)
    Next i
    Etats(16, Mouvement) = Score
    
    Afficher_Grille
End Sub

' ---------------------------------------------------
' REACTION LORSQUE CIBLE ATTEINTE
' ---------------------------------------------------
Private Sub Partie_Gagnee()
    If Cible_Atteint Then Exit Sub
    Cible_Atteint = True
    MsgBox "Bravo! Vous avez réussi à former le nombre " & 2 ^ Cible, vbExclamation, "2048"
End Sub

' ---------------------------------------------------
' REACTION LORSQUE CASE VIDE DISPONIBLE
' ---------------------------------------------------
' - Tant qu'il reste une case disponible, ou une fusion possible la partie n'est pas finie
' ---------------------------------------------------
Private Function Fin() As Boolean
    Dim i As Byte, j As Byte
    For i = 0 To 3
        For j = 0 To 3
            If Grille_Principale(i, j) = 0 Then Exit Function
            Select Case True
                Case i = 3 And j < 3
                    If Grille_Principale(i, j) = Grille_Principale(i, j + 1) Then Exit Function
                Case i < 3 And j = 3
                    If Grille_Principale(i, j) = Grille_Principale(i + 1, j) Then Exit Function
                Case i < 3 And j < 3
                    If Grille_Principale(i, j) = Grille_Principale(i + 1, j) Or Grille_Principale(i, j) = Grille_Principale(i, j + 1) Then Exit Function
            End Select
        Next j
    Next i
    
    With Worksheets("2048")
        Set Label_Score = .Range("E" & Ligne - 1)
        Label_Score = Score
    End With
    
    Fin = True
    Partie_Bloquee = True
End Function

' ---------------------------------------------------
' DEPLACEMENTS (BAS)
' ---------------------------------------------------
' - Les déplacements se font selon les règles du jeu:
'   - Deux cases adjacentes fusionnent si elle sont identiques et se déplacent dans le même sens
'   - On ne peut fusionner deux fois dans le même mouvement
Public Sub Deplacer_Vers_Le_Bas()
    Dim i As Integer, j As Integer, k As Integer
    Dim Evolution As Boolean
    Dim Fusion(3, 3) As Boolean
    For i = 2 To 0 Step -1
        For j = 0 To 3
            If Grille_Principale(i, j) <> 0 Then
                For k = i To 2
                    If Grille_Principale(k + 1, j) = Grille_Principale(k, j) And Fusion(k + 1, j) = False Then
                        Grille_Principale(k + 1, j) = CByte(1 + Grille_Principale(k + 1, j))
                        Score = Score + 2 ^ Grille_Principale(k + 1, j)
                        If Grille_Principale(k + 1, j) = Cible Then Partie_Gagnee
                        Grille_Principale(k, j) = 0
                        Fusion(k + 1, j) = True
                        Evolution = True
                        Exit For
                    ElseIf Grille_Principale(k + 1, j) = 0 Then
                        Grille_Principale(k + 1, j) = Grille_Principale(k, j)
                        Grille_Principale(k, j) = 0
                        Evolution = True
                    Else
                        Exit For
                    End If
                Next k
            End If
        Next j
    Next i
    Coup_Suivant Evolution
End Sub

' ---------------------------------------------------
' DEPLACEMENTS (DROITE)
' ---------------------------------------------------
Public Sub Deplacer_Vers_La_Droite()
    Dim i As Byte
    Dim j As Integer, k As Integer
    Dim Evolution As Boolean
    Dim Fusion(3, 3) As Boolean
    For i = 0 To 3
        For j = 2 To 0 Step -1
            If Grille_Principale(i, j) <> 0 Then
                For k = j To 2
                    If Grille_Principale(i, k + 1) = Grille_Principale(i, k) And Fusion(i, k + 1) = False Then
                        Grille_Principale(i, k + 1) = CByte(1 + Grille_Principale(i, k + 1))
                        Score = Score + 2 ^ Grille_Principale(i, k + 1)
                        If Grille_Principale(i, k + 1) = Cible Then Partie_Gagnee
                        Grille_Principale(i, k) = 0
                        Fusion(i, k + 1) = True
                        Evolution = True
                        Exit For
                    ElseIf Grille_Principale(i, k + 1) = 0 Then
                        Grille_Principale(i, k + 1) = Grille_Principale(i, k)
                        Grille_Principale(i, k) = 0
                        Evolution = True
                    Else
                        Exit For
                    End If
                Next k
            End If
        Next j
    Next i
    Coup_Suivant Evolution

End Sub

' ---------------------------------------------------
' DEPLACEMENTS (GAUCHE)
' ---------------------------------------------------
Public Sub Deplacer_Vers_La_Gauche()
    Dim i As Byte, j As Byte
    Dim k As Integer
    Dim Evolution As Boolean
    Dim Fusion(3, 3) As Boolean
    For i = 0 To 3
        For j = 1 To 3
            If Grille_Principale(i, j) <> 0 Then
                For k = j To 1 Step -1
                    If Grille_Principale(i, k - 1) = Grille_Principale(i, k) And Fusion(i, k - 1) = False Then
                        Grille_Principale(i, k - 1) = CByte(1 + Grille_Principale(i, k - 1))
                        Score = Score + 2 ^ Grille_Principale(i, k - 1)
                        If Grille_Principale(i, k - 1) = Cible Then Partie_Gagnee
                        Grille_Principale(i, k) = 0
                        Fusion(i, k - 1) = True
                        Evolution = True
                        Exit For
                    ElseIf Grille_Principale(i, k - 1) = 0 Then
                        Grille_Principale(i, k - 1) = Grille_Principale(i, k)
                        Grille_Principale(i, k) = 0
                        Evolution = True
                    Else
                        Exit For
                    End If
                Next k
            End If
        Next j
    Next i
    Coup_Suivant Evolution
End Sub

' ---------------------------------------------------
' DEPLACEMENTS (HAUT)
' ---------------------------------------------------
Public Sub Deplacer_Vers_Le_Haut()
    Dim i As Byte, j As Byte
    Dim k As Integer
    Dim Evolution As Boolean
    Dim Fusion(3, 3) As Boolean
    For i = 1 To 3
        For j = 0 To 3
            If Grille_Principale(i, j) <> 0 Then
                For k = i To 1 Step -1
                    If Grille_Principale(k - 1, j) = Grille_Principale(k, j) And Fusion(k - 1, j) = False Then
                        Grille_Principale(k - 1, j) = CByte(1 + Grille_Principale(k - 1, j))
                        Score = Score + 2 ^ Grille_Principale(k - 1, j)
                        If Grille_Principale(k - 1, j) = Cible Then Partie_Gagnee
                        Grille_Principale(k, j) = 0
                        Fusion(k - 1, j) = True
                        Evolution = True
                        Exit For
                    ElseIf Grille_Principale(k - 1, j) = 0 Then
                        Grille_Principale(k - 1, j) = Grille_Principale(k, j)
                        Grille_Principale(k, j) = 0
                        Evolution = True
                    Else
                        Exit For
                    End If
                Next k
            End If
        Next j
    Next i
    Coup_Suivant Evolution
End Sub

' ---------------------------------------------------
' ON DETERMINE SI LA GRILLE EVOLUE OU NON
' ---------------------------------------------------
Private Sub Coup_Suivant(Evolution As Boolean)
    If Evolution Then
        Placer_Nombre
    ElseIf Fin Then
        MsgBox "Partie terminée, votre score est de : " & Score, vbInformation, "2048"
    End If
End Sub

' ---------------------------------------------------
' ANNULATION / RETOUR EN ARRIERE
' ---------------------------------------------------
' - Permet de remettre la grille dans l'état exact du tour d'avant
' ---------------------------------------------------
Public Sub Annuler_Coup()
    Dim i As Integer
    ' On ne peut retourner en arrière si le joueur n'a pas encore joué
    If Mouvement = 0 Then
        MsgBox "Plus aucun déplacement à annuler", vbCritical, "2048"
        Exit Sub
    End If
    
    Mouvement = Mouvement - 1
    ReDim Preserve Etats(16, Mouvement)
    
    For i = 0 To 15
        Grille_Principale((i - (i Mod 4)) / 4, i Mod 4) = Etats(i, Mouvement)
    Next i
    Score = Etats(16, Mouvement)
    Afficher_Grille
    ' Permet de continuer de jouer après le dernier coup effectué; si celui-ci à mis fin à la partie
    Partie_Bloquee = False
End Sub