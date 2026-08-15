VERSION 5.00
Begin VB.Form PAVTactileLabelForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   225
   ClientTop       =   4890
   ClientWidth     =   7815
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.Label PAVTactileLabelLabel 
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
   Begin VB.Image PAVTactileLabelImage 
      Height          =   7200
      Left            =   0
      Picture         =   "PAVTactileLabelForm.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "PAVTactileLabelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. 1.0 VB )                  (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - PAVER LABEL DU CLAVIER                                   ***
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
'***    PROCEDURE:PAVTactileLabelFormInitialise()                           ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTactileLabelFormInitialise()                           ***
'******************************************************************************
Private Sub PAVTactileLabelFormInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTactileLabelImageInitialise()                          ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTactileLabelImageInitialise()                          ***
'******************************************************************************
Private Sub PAVTactileLabelImageInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTactileLabelLabelInitialise()                          ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTactileLabelLabelInitialise()                          ***
'******************************************************************************
Private Sub PAVTactileLabelLabelInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTactileLabelFormPosition()                             ***
'******************************************************************************
'***    FONCTION: Positionner la feuille sur l'ecran                        ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTactileLabelFormPosition()                             ***
'******************************************************************************
Private Sub PAVTactileLabelFormPosition()
    ' Je positionne la form a l'ecran
    ControlInitialise PAVTactileLabelForm, 0, Screen.Height * 9 / 15, Screen.Width * 12 / 21, Screen.Height * 1 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTactileLabelImagePosition()                            ***
'******************************************************************************
'***    FONCTION: Positionner l'image sur la form                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTactileLabelImagePosition()                            ***
'******************************************************************************
Private Sub PAVTactileLabelImagePosition()
    ' Je positionne l'image sur la form
    ControlInitialise PAVTactileLabelImage, 0, 0, PAVTactileLabelForm.ScaleWidth, PAVTactileLabelForm.ScaleHeight
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTactileLabelLabelPosition()                            ***
'******************************************************************************
'***    FONCTION: Positionner le lable sur la form                          ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTactileLabelLabelPosition()                            ***
'******************************************************************************
Private Sub PAVTactileLabelLabelPosition()
    ' Je repositionne le label et ajuste les fontes
    ControlInitialise PAVTactileLabelLabel, 0, 0, PAVTactileLabelForm.ScaleWidth, PAVTactileLabelForm.ScaleHeight
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
    PAVTactileLabelFormInitialise
    ' J'initialise l'image
    PAVTactileLabelImageInitialise
    ' J'initialise le label
    PAVTactileLabelLabelInitialise
    ' Je positionne la feuille a l'ecran
    PAVTactileLabelFormPosition
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
    ' Je repositionne l'image
    PAVTactileLabelImagePosition
    'Jerepositionne le label
    PAVTactileLabelLabelPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PANTicketLabel_Click(Index As Integer)                  ***
'******************************************************************************
'***    FONCTION:   Pour redimensionner la taille des fontes                ***
'******************************************************************************
'***    ENTREE:     index - le numero du label appuyer                      ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PAVTactileLabelLabel_Change()
    ' Je recalcule la taille des font
    FontAjuste PAVTactileLabelLabel
    ' Fin
End Sub

