VERSION 5.00
Begin VB.Form PAVClavierForm 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6045
   ClientLeft      =   240
   ClientTop       =   4905
   ClientWidth     =   7815
   ControlBox      =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.Label PAVClavierLabel 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3675
   End
   Begin VB.Image PAVClavierImage 
      Height          =   7200
      Left            =   0
      Picture         =   "PAVClavierForm.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "PAVClavierForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. VB )                                   ***
'******************************************************************************
'***    FEUILLE DU CLAVIER NUMERIQUE                                        ***
'******************************************************************************
'***    FONCTION:                                                           ***
'***    INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES                 ***
'***    RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU                    ***
'***    MODULES CLAVIER TACTILE                                             ***
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
'***    PROCEDURE:PAVClavierForm_Init()                                     ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVClavierForm_Init()                                     ***
'******************************************************************************
Private Sub PAVClavierForm_Init()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVClavierImage_Init()                                    ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVClavierImage_Init()                                    ***
'******************************************************************************
Private Sub PAVClavierImage_Init()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVClavierLabel_Init()                                    ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVClavierLabel_Init()                                    ***
'******************************************************************************
Private Sub PAVClavierLabel_Init()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVClavierForm_Pos()                                      ***
'******************************************************************************
'***    FONCTION: Positionner la feuille sur l'ecran                        ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVClavierForm_Pos()                                      ***
'******************************************************************************
Private Sub PAVClavierForm_Pos()
    ' Je positionne la form a l'ecran
    PAVClavierForm.Move 0, Screen.Height * 9 / 15, Screen.Width * 12 / 21, Screen.Height * 1 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVClavierImage_Pos()                                     ***
'******************************************************************************
'***    FONCTION: Positionner l'image sur la form                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVClavierImage_Pos()                                     ***
'******************************************************************************
Private Sub PAVClavierImage_Pos()
    ' Je positionne l'image sur la form
    PAVClavierImage.Move 0, 0, PAVClavierForm.ScaleWidth, PAVClavierForm.ScaleHeight
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVClavierLabel_Pos()                                     ***
'******************************************************************************
'***    FONCTION: Positionner le lable sur la form                          ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVClavierLabel_Pos()                                     ***
'******************************************************************************
Private Sub PAVClavierLabel_Pos()
    Dim LabelLgr As Single                                                   ' Pour memoriser la longueur du label
    Dim LabelLar As Single                                                   ' Pour memoriser la largueur du label
    ' Je calcule la longueur du label
    LabelLgr = PAVClavierForm.ScaleWidth
    ' Je calcule la largueur du label
    LabelLar = PAVClavierForm.ScaleHeight
    ' Je repositionne le label et ajuste les fontes
    ControlInit PAVClavierLabel, 0, 0, LabelLgr, LabelLar
    ' Fin
End Sub

'******************************************************************************
'***    Evenement internes                                                  ***
'******************************************************************************

'******************************************************************************
'***    EVENEMENT:  Form_Load()                                             ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
'***    FONCTION:   Initialiser les variables des forms et des controles    ***
'***                Positionner la feuille sur l'ecran                      ***
'******************************************************************************
Private Sub Form_Load()
    ' J'initialise la form
    PAVClavierForm_Init
    ' J'initialise l'image
    PAVClavierImage_Init
    ' J'initialise le label
    PAVClavierLabel_Init
    ' Je positionne la feuille a l'ecran
    PAVClavierForm_Pos
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  Form_Resize()                                           ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
'***    FONCTION:   repositionner les controles sur la form                 ***
'***                                                                        ***
'******************************************************************************
Private Sub Form_Resize()
    ' Je repositionne l'image
    PAVClavierImage_Pos
    'Jerepositionne le label
    PAVClavierLabel_Pos
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PANTicketLabel_Click(Index As Integer)                  ***
'******************************************************************************
'***    ENTREE:     index - le numero du label appuyer                      ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
'***    FONCTION:   Pour redimensionner la taille des fontes                ***
'******************************************************************************
Private Sub PAVClavierLabel_Change()
    ' Je recalcule la taille des font
    FontScale PAVClavierLabel
    ' Fin
End Sub

