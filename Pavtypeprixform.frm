VERSION 5.00
Begin VB.Form PAVTypePrixForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PAVTypePrixCDButton 
      Caption         =   "TYPEPRIX 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Index           =   2
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   2145
   End
   Begin VB.CommandButton PAVTypePrixCDButton 
      Caption         =   "TYPEPRIX 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   2145
   End
   Begin VB.CommandButton PAVTypePrixCDButton 
      Caption         =   "TYPEPRIX 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1905
   End
End
Attribute VB_Name = "PAVTypePrixForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. 1.0 VB )                  (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - PAVER DES TYPES DE PRIX                                  ***
'******************************************************************************
'***    FONCTION:- INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES      ***
'***             - RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU         ***
'***               MODULES SAISSIE DES DONNES D'UN TICKET                   ***
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
'***    PROCEDURE:PAVTypePrixFormInitialise()                               ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTypePrixFormInitialise()                               ***
'******************************************************************************
Private Sub PAVTypePrixFormInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTypePrixCDButtonInitialise()                           ***
'******************************************************************************
'***    FONCTION: initialiser - Le texte des boutons                        ***
'***                          - Les images des boutons                      ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTypePrixCDButtonInitialise()                           ***
'******************************************************************************
Private Sub PAVTypePrixCDButtonInitialise()
    Dim I As Integer                                                            ' Le numero du bouton courant
    ' Je fait defiler tous les boutons
    For I = 0 To PAVTypePrixCDButton.Count - 1
        ' Je recupere l'image du bouton
        PAVTypePrixCDButton(I).Picture = RessourceIconLoad(730 + I)
        ' Je recupere le texte du bouton
        PAVTypePrixCDButton(I).Caption = RessourceStringLoad(730 + I)
    ' Je ne sort que si j'ai fini de faire defiler les boutons
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTypePrixFormPosition()                                 ***
'******************************************************************************
'***    FONCTION: Positionner la form a l'ecran                             ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTypePrixFormPosition()                                 ***
'******************************************************************************
Private Sub PAVTypePrixFormPosition()
    ' Je positionne la form a l'ecran
    ControlInitialise PAVTypePrixForm, Screen.Width * 7.5 / 21, Screen.Height * 10 / 15, Screen.Width * 4.5 / 21, Screen.Height * 1 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTypePrixCDButtonPosition()                             ***
'******************************************************************************
'***    FONCTION: Repositionner les boutons sur la form                     ***
'***              Redimensionner la taille des fontes                       ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTypePrixCDButtonPosition()                             ***
'******************************************************************************
Private Sub PAVTypePrixCDButtonPosition()
    Dim I As Integer                                                            ' Le numero du bouton courant
    Dim CDButtonLef As Single                                                   ' La position gauche du bouton courant
    Dim CDButtonTop As Single                                                   ' La hauteur du bouton courant
    Dim CDButtonLgr As Single                                                   ' La longueur du bouton courant
    Dim CDButtonLar As Single                                                   ' La largueur du bouton courant
    ' Je fait defiler tous les boutons
    For I = 0 To PAVTypePrixCDButton.Count - 1
        ' Je calcule la position gauche du boutons
        CDButtonLef = (I Mod 3) * PAVTypePrixForm.ScaleWidth / 3
        ' Je calcule la hauteur du boutons
        CDButtonTop = 0
        ' Je calcule la longueur du boutons
        CDButtonLgr = PAVTypePrixForm.ScaleWidth / 3
        ' Je calcule la largueur du boutons
        CDButtonLar = PAVTypePrixForm.ScaleHeight
        ' Je repositionne le bouton et ajuste la taille des fontes
        ControlInitialise PAVTypePrixCDButton(I), CDButtonLef, CDButtonTop, CDButtonLgr, CDButtonLar
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
'***    FONCTION:   Initialiser les variables des forms et des controles    ***
'***                Positionner la feuille sur l'ecran                      ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub Form_Load()
    ' J'initialise la form
    PAVTypePrixFormInitialise
    ' J'initialise les boutons
    PAVTypePrixCDButtonInitialise
    ' Je positionne la form a l'ecran
    PAVTypePrixFormPosition
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
    ' Je repositionne les boutons sur la form
    PAVTypePrixCDButtonPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PAVTypePrixCDButton_Click(Index As Integer)             ***
'******************************************************************************
'***    FONCTION:   l'user a appuyer un bouton du type de mouvement         ***
'***                                                                        ***
'******************************************************************************
'***    ENTREE:     index - Numero du bouton appuyer                        ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PAVTypePrixCDButton_Click(Index As Integer)
    ' Je regarde sur quel bouton je suis
    Select Case Index
        Case ButtonPrixNormal
            ' C'est le boutton du prix normal
            ' Je change le type de prix
            TicketDonneTypePrixChange (PrixNormal)
        Case ButtonPrixSolde
            ' C'est le boutton du prix normal
            ' Je change le type de prix
            TicketDonneTypePrixChange (PrixSolde)
        Case ButtonPrixDemarque
            ' C'est le boutton du prix normal
            ' Je change le type de prix
            TicketDonneTypePrixChange (PrixDemarque)
    End Select
    ' Je rend le focus au champs de saissie courants
    Zone.SetFocus
    ' Fin
End Sub
