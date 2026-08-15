VERSION 5.00
Begin VB.Form PANFondForm 
   BorderStyle     =   0  'None
   ClientHeight    =   7080
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   9210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.Image PANFondImage 
      Height          =   7200
      Left            =   -720
      Picture         =   "PANFondForm.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "PANFondForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. 1.0 VB )                  (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - FOND DU PROGRAMME                                        ***
'******************************************************************************
'***    FONCTION:- INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES      ***
'***             - RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU         ***
'***               MODULES PRINCIPALES MICRO POINT DE VENTE                 ***
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
    Option Explicit                                                            ' Je veut etre sur que toute met variable sont declarer

'******************************************************************************
'***    Declaration de variable internes                                    ***
'******************************************************************************

'******************************************************************************
'***    Declaration des procedures internes                                 ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:  PANFondFormInitialise()                                 ***
'******************************************************************************
'***    FONCTION:   Initialiser                                             ***
'******************************************************************************
'***    ENTREE:     néant                                                   ***
'***                                                                        ***
'***    SORTIE:     néant                                                   ***
'******************************************************************************
'***    EXEMPLE:    PANFondFormInitialise                                   ***
'******************************************************************************
Private Sub PANFondFormInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE: PANFondImageInitialise()                                 ***
'******************************************************************************
'***    FONCTION:  Initialiser                                              ***
'******************************************************************************
'***    ENTREE:    néant                                                    ***
'***                                                                        ***
'***    SORTIE:    néant                                                    ***
'******************************************************************************
'***    EXEMPLE:   PANFondImageInitialise                                   ***
'******************************************************************************
Private Sub PANFondImageInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE: PANFondFormPosition()                                    ***
'******************************************************************************
'***    FONCTION:  positionner la form a l'ecran                            ***
'******************************************************************************
'***    ENTREE:    néant                                                    ***
'***                                                                        ***
'***    SORTIE:    néant                                                    ***
'******************************************************************************
'***    EXEMPLE:   PANFondFormPosition                                      ***
'******************************************************************************
Private Sub PANFondFormPosition()
    ' Je positionne la form a l'ecran
    ControlInitialise PANFondForm, 0, 0, Screen.Width, Screen.Height
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE: PANFondImagePosition()                                   ***
'******************************************************************************
'***    FONCTION:  Positionner l'image sur la form                          ***
'******************************************************************************
'***    ENTREE:    Néant                                                    ***
'***                                                                        ***
'***    SORTIE:    Néant                                                    ***
'******************************************************************************
'***    EXEMPLE:   PANFondImagePosition                                     ***
'******************************************************************************
Private Sub PANFondImagePosition()
    ' Je positionne l'image sur la form
    ControlInitialise PANFondImage, 0, 0, PANFondForm.ScaleWidth, PANFondForm.ScaleHeight
    ' Fin
End Sub

'******************************************************************************
'***    Evenement internes                                                  ***
'******************************************************************************

'******************************************************************************
'***    EVENEMENT: Form_Load()                                              ***
'******************************************************************************
'***    FONCTION:  Initialiser les variables des forms et des controles     ***
'***               Positionner la feuille sur l'ecran                       ***
'******************************************************************************
'***    ENTREE:    néant                                                    ***
'***                                                                        ***
'***    SORTIE:    néant                                                    ***
'******************************************************************************
Private Sub Form_Load()
    ' J'initialise la form
    PANFondFormInitialise
    ' J'initialise l'image
    PANFondImageInitialise
    ' Je positionne la feuille a l'ecran
    PANFondFormPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT: Form_Resize()                                            ***
'******************************************************************************
'***    FONCTION:  Repositionner les controles sur la form                  ***
'******************************************************************************
'***    ENTREE:    néant                                                    ***
'***                                                                        ***
'***    SORTIE:    néant                                                    ***
'******************************************************************************
Private Sub Form_Resize()
    ' Je repositionne l'image
    PANFondImagePosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT: PANFondImage_Click()                                     ***
'******************************************************************************
'***    FONCTION:  Validation de la Zone De Saisie                          ***
'******************************************************************************
'***    ENTREE:    néant                                                    ***
'***                                                                        ***
'***    SORTIE:    néant                                                    ***
'******************************************************************************
Private Sub PANFondImage_Click()
    ' Validation de la Zone De Saisie
    ZoneValidation
    ' Fin
End Sub

