VERSION 5.00
Begin VB.Form PAVValidationForm 
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PAVValidationCDButton 
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
   Begin VB.CommandButton PAVValidationCDButton 
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
   Begin VB.CommandButton PAVValidationCDButton 
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
Attribute VB_Name = "PAVValidationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. VB )                                   ***
'******************************************************************************
'***    FEUILLE POUR - LA VALIDATION D'UNE LIGNE DE SAISSIE                 ***
'***                 - RETOUR AU MENU                                       ***
'***                 - ANNULATION DE LA ZONE SAISSIE                        ***
'******************************************************************************
'***    FONCTION:                                                           ***
'***    INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES                 ***
'***    RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU                    ***
'***    MODULES PRINCIPALES DU MICRO POINT DE VENTE                         ***
'***                                                                        ***
'******************************************************************************

Option Explicit                                                                 ' Je doit etre sur que mes variables on ete declarer

'******************************************************************************
'***    Declaration de variable internes                                    ***
'******************************************************************************

'******************************************************************************
'***    Declaration des procedures internes                                 ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:PAVValidationFormInit()                                  ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVValidationFormInit()                                  ***
'******************************************************************************
Private Sub PAVValidationFormInit()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVValidationCDButtonInit()                              ***
'******************************************************************************
'***    FONCTION: Initialiser le texte des boutons                          ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVValidationCDButtonInit()                              ***
'******************************************************************************
Private Sub PAVValidationCDButtonInit()
    Dim I As Integer                                                            ' Pour memoriser le numero du bouton courant
    ' Je fait defiler les boutons
    For I = 0 To PAVValidationCDButton.Count - 1
        ' Je recupere l'image du bouton
        PAVValidationCDButton(I).Picture = LoadRessourceIcon(780 + I)
        ' Je recupere le texte du bouton
        PAVValidationCDButton(I).Caption = LoadRessourceString(780 + I)
    ' Je ne sort que si j'ai fini de faire defiler les boutons
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVValidationFormPos()                                   ***
'******************************************************************************
'***    FONCTION: Pour positionner la feuille a l'ecran                     ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVValidationFormPos()                                   ***
'******************************************************************************
Private Sub PAVValidationFormPos()
    ' Je positionne la feuille a l'ecran
    PAVValidationForm.Move Screen.Width * 16.5 / 21, Screen.Height * 7 / 15, Screen.Width * 4.5 / 21, Screen.Height * 1 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVValidationCDButtonPos()                               ***
'******************************************************************************
'***    FONCTION: Pour repositionner les boutons sur la form                ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVValidationCDButtonPos()                               ***
'******************************************************************************
Private Sub PAVValidationCDButtonPos()
    Dim I As Integer                                                            ' Pour memoriser le numero du bouton courant
    Dim CDButtonLef As Single                                                   ' Pour memoriser la position gauche du bouton courant
    Dim CDButtonTop As Single                                                   ' Pour memoriser la hauteur du bouton courant
    Dim CDButtonLgr As Single                                                   ' Pour memoriser la longueur d'un bouton
    Dim CDButtonLar As Single                                                   ' Pour memoriser la largeur d'un bouton
    ' Je fait defiler tous les boutons
    For I = 0 To PAVValidationCDButton.Count - 1
        ' Je calcule la position gauche du boutons
        CDButtonLef = (I Mod 3) * PAVValidationForm.ScaleWidth / 3
        ' Je calcule la hauteur du boutons
        CDButtonTop = 0
        ' Je calcule la longueur du bouton
        CDButtonLgr = PAVValidationForm.ScaleWidth / 3
        ' Je calcule la largeur
        CDButtonLar = PAVValidationForm.ScaleHeight
        ' Je repositionne les boutons et ajuste la taille des fontes
        ControlInit PAVValidationCDButton(I), CDButtonLef, CDButtonTop, CDButtonLgr, CDButtonLar
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
Private Sub Form_Load()
    ' J'initialise la form
    PAVValidationFormInit
    ' J'initialise les boutons
    PAVValidationCDButtonInit
    ' Je positionne la form a l'ecran
    PAVValidationFormPos
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  Form_Resize()                                           ***
'******************************************************************************
'***    FONCTION:   Repositionner les controles sur la form                 ***
'******************************************************************************
Private Sub Form_Resize()
    ' Je repositionne les boutons
    PAVValidationCDButtonPos
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:PVValidationCDButton_Click(Index As Integer)              ***
'******************************************************************************
'***    FONCTION: Pour  - Retourner au menu principale                      ***
'***                    - Validation d'une ligne de saissie                 ***
'***                    - Annulation d'une zone de saissie                  ***
'******************************************************************************
Private Sub PAVValidationCDButton_Click(Index As Integer)
    ' Index:repere le N° du bouton appuyer
    ' Je regarde sur quel bouton a appuyer l'user
    Select Case Index
        Case ButtonMenuPrincipale
            ' C'est le button du menu principale
            ' J'envoie le menu principale
            MenuPrincipale
        Case ButtonAnnulationZone
            ' C'est le button d'annulation d'une zone de saissie
            ' J'envoie l'annulation de la saisie
            AnnulationZone
        Case ButtonValidationLigne
            ' C'est le button de validation d'une ligne de saissie
            ' J'envoie la validation de la saisie
            ValidationLigne
    End Select
    ' Fin
End Sub
