VERSION 5.00
Begin VB.Form FormPavTactileTotal 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PAVTactileTotalCDButton 
      Caption         =   "TOTAL"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "FormPavTactileTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***                      (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - PAVER TOTAL                                          ***
'******************************************************************************
'***    FONCTION:- INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES      ***
'***             - RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU         ***
'***               MODULES PRINCIPALE MICRO POINT DE VENTE                  ***
'******************************************************************************
'***    PROGRAMMEUR: Royer Jean-Laurent                                     ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:                                                          ***
'***    FONCTION:                                                           ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                             ***
'******************************************************************************
'***    Version 1.0 : 30/10/2000                                            ***
'******************************************************************************
    Option Explicit                                                            ' Je doit etre sur que mes variables on ete declarer

'******************************************************************************
'***    Declaration de variable internes                                    ***
'******************************************************************************

'******************************************************************************
'***    Declaration des procedures internes                                 ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:FormPavTactileTotalInitialise()                           ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormPavTactileTotalInitialise()                           ***
'******************************************************************************
Private Sub FormPavTactileTotalInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTactileTotalCDButtonInitialise()                       ***
'******************************************************************************
'***    FONCTION: initialiser - Le texte des boutons                        ***
'***                          - Les images des boutons                      ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTactileTotalCDButtonInitialise()                       ***
'******************************************************************************
Private Sub PAVTactileTotalCDButtonInitialise()
    ' Je charge l'image du bouton
    PAVTactileTotalCDButton.Picture = RessourceIconLoad(720)
    ' Je charge le texte du bouton
    PAVTactileTotalCDButton.Caption = RessourceStringLoad(720)
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:FormPavTactileTotalPosition()                             ***
'******************************************************************************
'***    FONCTION: Positionner la feuille a l'ecran                          ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormPavTactileTotalPosition()                             ***
'******************************************************************************
Private Sub FormPavTactileTotalPosition()
    ' Je positionne la feuille a l'ecran
    ControlInitialise FormPavTactileTotal, Screen.Width * 16.5 / 21, Screen.Height * 9 / 15, Screen.Width * 4.5 / 21, Screen.Height * 1 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTactileTotalCDButtonPosition()                         ***
'******************************************************************************
'***    FONCTION: Repositionner les boutons sur la form                     ***
'***              Redimensionner la taille des fontes                       ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTactileTotalCDButtonPosition()                         ***
'******************************************************************************
Private Sub PAVTactileTotalCDButtonPosition()
    Dim CDButtonLef As Single                                                   ' La position gauche du bouton courant
    Dim CDButtonTop As Single                                                   ' La hauteur du bouton courant
    Dim CDButtonLgr As Single                                                   ' La longueur du bouton courant
    Dim CDButtonLar As Single                                                   ' La largueur du bouton courant
    ' Je calcule la position gauche du boutons
    CDButtonLef = 0
    ' Je calcule la hauteur du boutons
    CDButtonTop = 0
    ' Je calcule la longueur du boutons
    CDButtonLgr = FormPavTactileTotal.ScaleWidth
    ' Je calcule la largueur du boutons
    CDButtonLar = FormPavTactileTotal.ScaleHeight
    ' Je repositionne le bouton et ajuste la taille des fontes
    ControlInitialise PAVTactileTotalCDButton, CDButtonLef, CDButtonTop, CDButtonLgr, CDButtonLar
    ' Fin
End Sub

'******************************************************************************
'***    Evenement internes                                                  ***
'******************************************************************************

'******************************************************************************
'***    EVENEMENT:  Form_Load()                                             ***
'******************************************************************************
'***    FONCTION:   Initialiser les variables des forms et des controles    ***
'***                Positionner la feuille sur l'ecran                      ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub Form_Load()
    ' J'initialise la form
    FormPavTactileTotalInitialise
    ' J'initialise les boutons
    PAVTactileTotalCDButtonInitialise
    ' Je positionne la form a l'ecran
    FormPavTactileTotalPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  Form_Resize()                                           ***
'******************************************************************************
'***    FONCTION:   repositionner les controles sur la form                 ***
'***                                                                        ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub Form_Resize()
    ' Je repositionne les boutons
    PAVTactileTotalCDButtonPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PAVTactileTotalCDButton_Click()                         ***
'******************************************************************************
'***    FONCTION:   L'user a appuyer sur la touche total                    ***
'***                                                                        ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PAVTactileTotalCDButton_Click()
    ' Je renvoie que l'user a appuyer sur la touche total
    TotalValidation
    ' Fin
End Sub

