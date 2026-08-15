VERSION 5.00
Begin VB.Form PAVTactileValidationForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PAVTactileValidationCDButton 
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton PAVTactileValidationCDButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton PAVTactileValidationCDButton 
      Caption         =   "ANNULE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "PAVTactileValidationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. 1.0 VB )                  (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - PAVER VALIDATION  ANNULATION   MENU                      ***
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
'***    PROCEDURE:PAVTactileValidationFormInitialise()                      ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTactileValidationFormInitialise()                      ***
'******************************************************************************
Private Sub PAVTactileValidationFormInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTactileValidationCDButtonInitialise()                  ***
'******************************************************************************
'***    FONCTION: Initialiser le texte des boutons                          ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTactileValidationCDButtonInitialise()                  ***
'******************************************************************************
Private Sub PAVTactileValidationCDButtonInitialise()
    Dim I As Integer                                                            ' Le numero du bouton courant
    ' Je fait defiler les boutons
    For I = 0 To PAVTactileValidationCDButton.Count - 1
        ' Je recupere l'image du bouton
        PAVTactileValidationCDButton(I).Picture = RessourceIconLoad(780 + I)
        ' Je recupere le texte du bouton
        PAVTactileValidationCDButton(I).Caption = RessourceStringLoad(780 + I)
    ' Je ne sort que si j'ai fini de faire defiler les boutons
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTactileValidationFormPosition()                        ***
'******************************************************************************
'***    FONCTION: Pour positionner la feuille a l'ecran                     ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTactileValidationFormPosition()                        ***
'******************************************************************************
Private Sub PAVTactileValidationFormPosition()
    ' Je positionne la feuille a l'ecran
    ControlInitialise PAVTactileValidationForm, Screen.Width * 16.5 / 21, Screen.Height * 7 / 15, Screen.Width * 4.5 / 21, Screen.Height * 1 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTactileValidationCDButtonPosition()                    ***
'******************************************************************************
'***    FONCTION: Pour repositionner les boutons sur la form                ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTactileValidationCDButtonPosition()                    ***
'******************************************************************************
Private Sub PAVTactileValidationCDButtonPosition()
    Dim I As Integer                                                            ' Le numero du bouton courant
    Dim CDButtonLef As Single                                                   ' La position gauche du bouton courant
    Dim CDButtonTop As Single                                                   ' La hauteur du bouton courant
    Dim CDButtonLgr As Single                                                   ' La longueur d'un bouton
    Dim CDButtonLar As Single                                                   ' La largeur d'un bouton
    ' Je fait defiler tous les boutons
    For I = 0 To PAVTactileValidationCDButton.Count - 1
        ' Je calcule la position gauche du boutons
        CDButtonLef = (I Mod 3) * PAVTactileValidationForm.ScaleWidth / 3
        ' Je calcule la hauteur du boutons
        CDButtonTop = 0
        ' Je calcule la longueur du bouton
        CDButtonLgr = PAVTactileValidationForm.ScaleWidth / 3
        ' Je calcule la largeur
        CDButtonLar = PAVTactileValidationForm.ScaleHeight
        ' Je repositionne les boutons et ajuste la taille des fontes
        ControlInitialise PAVTactileValidationCDButton(I), CDButtonLef, CDButtonTop, CDButtonLgr, CDButtonLar
    ' Je ne sort que si j'ai fini de faire defiler les boutons
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    Evenement internes                                                  ***
'******************************************************************************

'******************************************************************************
'***    EVENEMENT:  Form_Load()                                             ***
'******************************************************************************
'***    FONCTION:-  Initialiser les variables des forms et des controles    ***
'***             -  Positionner la feuille sur l'ecran                      ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub Form_Load()
    ' J'initialise la form
    PAVTactileValidationFormInitialise
    ' J'initialise les boutons
    PAVTactileValidationCDButtonInitialise
    ' Je positionne la form a l'ecran
    PAVTactileValidationFormPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  Form_Resize()                                           ***
'******************************************************************************
'***    FONCTION:   Repositionner les controles sur la form                 ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub Form_Resize()
    ' Je repositionne les boutons
    PAVTactileValidationCDButtonPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:PVValidationCDButton_Click(Index As Integer)              ***
'******************************************************************************
'***    FONCTION: Pour  - Retourner au menu principale                      ***
'***                    - Validation d'une ligne de saissie                 ***
'***                    - Annulation d'une zone de saissie                  ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PAVTactileValidationCDButton_Click(Index As Integer)
    ' Index:repere le N° du bouton appuyer
    ' Je regarde sur quel bouton a appuyer l'user
    Select Case Index
        Case ButtonMenuPrincipale
            ' C'est le button du menu principale
            ' J'envoie le menu principale
            MenuPrincipale
        Case ButtonZoneAnnulation
            ' C'est le button d'annulation d'une zone de saissie
            ' J'envoie l'annulation de la saisie
            ZoneAnnulation
        Case ButtonLigneValidation
            ' C'est le button de validation d'une ligne de saissie
            ' J'envoie la validation de la saisie
            LigneValidation
    End Select
    ' Fin
End Sub
