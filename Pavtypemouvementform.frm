VERSION 5.00
Begin VB.Form PAVTypeMouvementForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7080
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   7980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PAVTypeMouvementCDButton 
      Caption         =   "MVT 9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Index           =   8
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   2625
   End
   Begin VB.CommandButton PAVTypeMouvementCDButton 
      Caption         =   "MVT 8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Index           =   7
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   2505
   End
   Begin VB.CommandButton PAVTypeMouvementCDButton 
      Caption         =   "MVT 7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Index           =   6
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5520
      Width           =   2505
   End
   Begin VB.CommandButton PAVTypeMouvementCDButton 
      Caption         =   "MVT 6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Index           =   5
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   2625
   End
   Begin VB.CommandButton PAVTypeMouvementCDButton 
      Caption         =   "MVT 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Index           =   4
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   2505
   End
   Begin VB.CommandButton PAVTypeMouvementCDButton 
      Caption         =   "MVT 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Index           =   3
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3255
      Width           =   2505
   End
   Begin VB.CommandButton PAVTypeMouvementCDButton 
      Caption         =   "MVT 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Index           =   2
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   975
      Width           =   2625
   End
   Begin VB.CommandButton PAVTypeMouvementCDButton 
      Caption         =   "MVT 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   2505
   End
   Begin VB.CommandButton PAVTypeMouvementCDButton 
      Caption         =   "MVT 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   975
      Width           =   2505
   End
End
Attribute VB_Name = "PAVTypeMouvementForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. 1.0 VB )                  (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - PAVER DES TYPE DE MOUVEMENTS                             ***
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
    Option Explicit                                                             ' Je doit etre sur que mes variables on ete declarer

'******************************************************************************
'***    Declaration de variable internes                                    ***
'******************************************************************************

'******************************************************************************
'***    Declaration des procedures internes                                 ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:PAVTypeMouvementFormInitialise()                          ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTypeMouvementFormInitialise()                          ***
'******************************************************************************
Private Sub PAVTypeMouvementFormInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTypeMouvementCDButtonInitialise()                      ***
'******************************************************************************
'***    FONCTION: Initialiser - Le texte des boutons                        ***
'***                          - Les images des boutons                      ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTypeMouvementCDButtonInitialise()                      ***
'******************************************************************************
Private Sub PAVTypeMouvementCDButtonInitialise()
    Dim I As Integer                                                            ' Le numero du bouton courant
    ' Je fait defiler tous les bouton
    For I = 0 To PAVTypeMouvementCDButton.Count - 1
        ' Je charge l'image du bouton
        PAVTypeMouvementCDButton(I).Picture = RessourceIconLoad(790 + I)
        ' Je charge le texte du bouton
        PAVTypeMouvementCDButton(I).Caption = RessourceStringLoad(790 + I)
    ' Je ne sort que si j'ai fini de faire defiler les boutons
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTypeMouvementFormPosition()                            ***
'******************************************************************************
'***    FONCTION: Positionner la form a l'ecran                             ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTypeMouvementFormPosition()                            ***
'******************************************************************************
Private Sub PAVTypeMouvementFormPosition()
    ' Je positionne la feuille a l'ecran
    ControlInitialise PAVTypeMouvementForm, Screen.Width * 16.5 / 21, Screen.Height * 4 / 15, Screen.Width * 4.5 / 21, Screen.Height * 3 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTypeMouvementCDButtonPosition()                        ***
'******************************************************************************
'***    FONCTION: Repositionner les boutons sur la form                     ***
'***              Redimensionner la taille des fontes                       ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTypeMouvementCDButtonPosition()                        ***
'******************************************************************************
Private Sub PAVTypeMouvementCDButtonPosition()
    Dim I As Integer                                                            ' Le numero du bouton courant
    Dim CDButtonLef As Single                                                   ' La position gauche du bouton courant
    Dim CDButtonTop As Single                                                   ' La hauteur du bouton courant
    Dim CDButtonLgr As Single                                                   ' La longueur du bouton courant
    Dim CDButtonLar As Single                                                   ' La largueur du bouton courant
    ' Je fait defiler tous les boutons
    For I = 0 To PAVTypeMouvementCDButton.Count - 1
        ' Je calcule la position gauche du boutons
        CDButtonLef = (I Mod 3) * PAVTypeMouvementForm.ScaleWidth / 3
        ' Je calcule la hauteur du boutons
        CDButtonTop = (Int(I / 3)) * PAVTypeMouvementForm.ScaleHeight / 3
        ' Je calcule la longueur du boutons
        CDButtonLgr = PAVTypeMouvementForm.ScaleWidth / 3
        ' Je calcule la largueur du boutons
        CDButtonLar = PAVTypeMouvementForm.ScaleHeight / 3
        ' Je repositionne le bouton et ajuste la taille des fontes
        ControlInitialise PAVTypeMouvementCDButton(I), CDButtonLef, CDButtonTop, CDButtonLgr, CDButtonLar
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
    ' J'initialise la form a l'ecran
    PAVTypeMouvementFormInitialise
    ' J'initialise les boutons
    PAVTypeMouvementCDButtonInitialise
    ' Je positionne la form a l'ecran
    PAVTypeMouvementFormPosition
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
    PAVTypeMouvementCDButtonPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PAVTypeMouvementCDButton_Click(Index As Integer)        ***
'******************************************************************************
'***    FONCTION:   l'user a appuyer un bouton du type de mouvement         ***
'***                                                                        ***
'******************************************************************************
'***    ENTREE:     index - Numero du bouton appuyer                        ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PAVTypeMouvementCDButton_Click(Index As Integer)
    ' Je regarde sur quel bouton l'user a appuyer
    Select Case Index
        Case ButtonSpecial
            ' C'est le bouton des vente special
            ' Je met a jour la ligne de la vente pour dire qu'elle est special ou non
            TicketDonneSpecialChange
        Case ButtonMouvementY
            ' C'est le bouton du mouvement y
            ' Je change le type de mouvement
            TicketDonneTypeMouvementChange (MouvementYPositif)
        Case ButtonMouvementZ
            ' C'est le bouton du mouvement z
            ' Je change le type de mouvement
            TicketDonneTypeMouvementChange (MouvementZPositif)
        Case ButtonMouvementReservation
            ' C'est le bouton du mouvement reservation
            ' Je change le type de mouvement
            TicketDonneTypeMouvementChange (MouvementReservation)
        Case ButtonMouvementArrhes
            ' C'est le bouton du mouvement arrhes
            ' Je change le type de mouvement
            TicketDonneTypeMouvementChange (MouvementArrhes)
        Case ButtonMouvementArrhesArticle
            ' C'est le bouton du mouvement arrhes article
            ' Je change le type de mouvement
            TicketDonneTypeMouvementChange (MouvementArrhesArticle)
        Case ButtonMouvementRendu
            ' C'est le bouton du mouvement rendu
            ' Je change le type de mouvement
            TicketDonneTypeMouvementChange (MouvementRendu)
        Case ButtonMouvementVente
            ' C'est le bouton du mouvement des ventes
            ' Je change le type de mouvement
            TicketDonneTypeMouvementChange (MouvementVente)
    End Select
    ' Je rend le focus au champs de saissie courants
    Zone.SetFocus
    ' Fin
End Sub
