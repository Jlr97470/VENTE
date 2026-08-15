VERSION 5.00
Begin VB.Form PAVMouvementForm 
   AutoRedraw      =   -1  'True
   ClientHeight    =   7080
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   7980
   ControlBox      =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PAVMouvementCDButton 
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
   Begin VB.CommandButton PAVMouvementCDButton 
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
   Begin VB.CommandButton PAVMouvementCDButton 
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
   Begin VB.CommandButton PAVMouvementCDButton 
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
   Begin VB.CommandButton PAVMouvementCDButton 
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
   Begin VB.CommandButton PAVMouvementCDButton 
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
   Begin VB.CommandButton PAVMouvementCDButton 
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
   Begin VB.CommandButton PAVMouvementCDButton 
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
   Begin VB.CommandButton PAVMouvementCDButton 
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
Attribute VB_Name = "PAVMouvementForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. VB )                                   ***
'******************************************************************************
'***    FEUILLE DU PAVER DES MOUVEMENTS                                     ***
'******************************************************************************
'***    FONCTION:                                                           ***
'***    INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES                 ***
'***    RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU                    ***
'***    MODULES SAISSIE DES DONNES D'UN TICKET                              ***
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
'***    PROCEDURE:PAVMouvementForm_Init()                                   ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVMouvementForm_Init()                                   ***
'******************************************************************************
Private Sub PAVMouvementForm_Init()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVMouvementCDButton_Init()                               ***
'******************************************************************************
'***    FONCTION: initialiser - Le texte des boutons                        ***
'***                          - Les images des boutons                      ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVMouvementCDButton_Init()                               ***
'******************************************************************************
Private Sub PAVMouvementCDButton_Init()
    Dim I As Integer                                                            ' Pour memoriser le numero du bouton courant
    ' Je fait defiler tous les bouton
    For I = 0 To PAVMouvementCDButton.Count - 1
        ' Je charge l'image du bouton
        PAVMouvementCDButton(I).Picture = LoadRessourceIcon(790 + I)
        ' Je charge le texte du bouton
        PAVMouvementCDButton(I).Caption = LoadRessourceString(790 + I)
    ' Je ne sort que si j'ai fini de faire defiler les boutons
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVMouvementForm_Pos()                                    ***
'******************************************************************************
'***    FONCTION: positionner la form a l'ecran                             ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVMouvementForm_Pos()                                    ***
'******************************************************************************
Private Sub PAVMouvementForm_Pos()
    ' Je positionne la feuille a l'ecran
    PAVMouvementForm.Move Screen.Width * 15 / 20 + Screen.TwipsPerPixelX, Screen.Height * 5 / 15 + Screen.TwipsPerPixelY, Screen.Width * 5 / 20 - Screen.TwipsPerPixelX * 2, Screen.Height * 4 / 15 - Screen.TwipsPerPixelY * 2
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVMouvementCDButton_Pos()                                ***
'******************************************************************************
'***    FONCTION: repositionner les boutons sur la form                     ***
'***              redimensionner la taille des fontes                       ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVMouvementCDButton_Pos()                                ***
'******************************************************************************
Private Sub PAVAlphaCDButton_Pos()
    Dim I As Integer                                                            ' Pour memoriser le numero du bouton courant
    Dim CDButtonLgr As Single                                                   ' Pour memoriser la longueur du bouton courant
    Dim CDButtonLar As Single                                                   ' Pour memoriser la largueur du bouton courant
    ' Je fait defiler tous les boutons
    For I = 0 To PAVMouvementCDButton.Count - 1
        ' Je calcule la longueur du boutons
        CDButtonLgr = PAVMouvementForm.ScaleWidth / 3 - Screen.TwipsPerPixelX * 2
        ' Je regarde la longueur est valide
        Select Case CDButtonLgr
            Case Is < 0
                ' La longueur n'est pas valide
                ' Je defini la longueur
                CDButtonLgr = 0
            Case Else
                ' La longueur est valide
        End Select
        ' Je calcule la largueur du boutons
        CDButtonLar = PAVMouvementForm.ScaleHeight / 3 - Screen.TwipsPerPixelY * 2
        ' Je regarde la largueur est valide
        Select Case CDButtonLar
            Case Is < 0
                ' La largueur n'est pas valide
                ' Je defini la longueur
                CDButtonLar = 0
            Case Else
                ' La largueur est valide
        End Select
        ' Je repositionne le bouton et ajuste la taille des fontes
        ControlInit PAVMouvementCDButton(I), (I Mod 3) * PAVMouvementForm.ScaleWidth / 3 + Screen.TwipsPerPixelX, (Int(I / 3)) * PAVMouvementForm.ScaleHeight / 3 + Screen.TwipsPerPixelY, CDButtonLgr, CDButtonLar, PAVMouvementCDButton(I).Caption
    ' Je ne sort que si j'ai fini de faire defiler les boutons
    Next I
    ' Fin
End Sub

' Evenement

Private Sub Form_Load()
    PVTypeMouvementForm_Init
End Sub

Private Sub Form_Resize()
    PVTypeMouvementCDButton_Init
End Sub

Private Sub PVTypeMouvementCDButton_Click(Index As Integer)
' index : numero du boutton appuyer
    Select Case Index
        Case ButtonSpecial            'bouton vente special
            ' je met a jour la ligne de la vente pour dire qu'elle est special ou non
            ChangeTicketDonneSpecial
        Case ButtonMouvementArrhes    'bouton mouvement arrhes
            ' je change le type de mouvement en mouvement des arrhes
            ChangeTicketDonneTypeMouvement (MouvementArrhes)
        Case ButtonMouvementRendu
            ' je change le type de mouvement en mouvement des rendu
            ChangeTicketDonneTypeMouvement (MouvementRendu)
        Case ButtonMouvementReservation
            ' je change le type de mouvement en mouvement des reservation
            ChangeTicketDonneTypeMouvement (MouvementReservation)
        Case ButtonMouvementY
            ' je change le type de mouvement en mouvement des mvts y
            ChangeTicketDonneTypeMouvement (MouvementY)
        Case ButtonMouvementZ
            ' je change le type de mouvement en mouvement des mvts z
            ChangeTicketDonneTypeMouvement (MouvementZ)
    End Select
    ' je rend le focus au champs de saisie
    Zone.SetFocus
End Sub
