VERSION 5.00
Begin VB.Form FormPavTypeDevise 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5580
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PAVTypeDeviseCDButton 
      Caption         =   "DEVISE 1"
      Height          =   3615
      Index           =   1
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton PAVTypeDeviseCDButton 
      Caption         =   "DEVISE 0"
      Height          =   3615
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "FormPavTypeDevise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***                      (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - PAVER DES DEVISES                                        ***
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
'***    PROCEDURE:FormPavTypeDeviseInitialise()                             ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormPavTypeDeviseInitialise()                             ***
'******************************************************************************
Private Sub FormPavTypeDeviseInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTypeDeviseCDButtonInitialise()                         ***
'******************************************************************************
'***    FONCTION: initialiser - Le texte des boutons                        ***
'***                          - Les images des boutons                      ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTypeDeviseCDButtonInit()                               ***
'******************************************************************************
Private Sub PAVTypeDeviseCDButtonInitialise()
    Dim I As Integer                                                            ' Le numero du bouton courant
    ' Je fait defiler tous les boutons
    For I = 0 To PAVTypeDeviseCDButton.Count - 1
        ' Je recupere l'image du bouton
        PAVTypeDeviseCDButton(I).Picture = RessourceIconLoad(1200 + I)
        ' Je recupere le texte du bouton
        PAVTypeDeviseCDButton(I).Caption = RessourceStringLoad(1200 + I)
    ' Je ne sort que si j'ai fini de faire defiler les boutons
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:FormPavTypeDevisePosition()                               ***
'******************************************************************************
'***    FONCTION: Positionner la form a l'ecran                             ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormPavTypeDevisePosition()                               ***
'******************************************************************************
Private Sub FormPavTypeDevisePosition()
    ' Je positionne la form a l'ecran
    ControlInitialise FormPavTypeDevise, Screen.Width * 7.5 / 21, Screen.Height * 10 / 15, Screen.Width * 4.5 / 21, Screen.Height * 1 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTypeDeviseCDButtonPosition()                           ***
'******************************************************************************
'***    FONCTION: Repositionner les boutons sur la form                     ***
'***              redimensionner la taille des fontes                       ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTypeDeviseCDButtonPosition()                           ***
'******************************************************************************
Private Sub PAVTypeDeviseCDButtonPosition()
    Dim I As Integer                                                            ' Le numero du bouton courant
    Dim CDButtonLef As Single                                                   ' La position gauche du bouton courant
    Dim CDButtonTop As Single                                                   ' La hauteur du bouton courant
    Dim CDButtonLgr As Single                                                   ' La longueur du bouton courant
    Dim CDButtonLar As Single                                                   ' La largueur du bouton courant
    ' Je fait defiler tous les boutons
    For I = 0 To PAVTypeDeviseCDButton.Count - 1
        ' Je calcule la position gauche du boutons
        CDButtonLef = (I Mod 2) * FormPavTypeDevise.ScaleWidth / 2
        ' Je calcule la hauteur du boutons
        CDButtonTop = 0
        ' Je calcule la longueur du boutons
        CDButtonLgr = FormPavTypeDevise.ScaleWidth / 2
        ' Je calcule la largueur du boutons
        CDButtonLar = FormPavTypeDevise.ScaleHeight
        ' Je repositionne le bouton et ajuste la taille des fontes
        ControlInitialise PAVTypeDeviseCDButton(I), CDButtonLef, CDButtonTop, CDButtonLgr, CDButtonLar
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
    FormPavTypeDeviseInitialise
    ' J'initialise les boutons
    PAVTypeDeviseCDButtonInitialise
    ' Je positionne la form a l'ecran
    FormPavTypeDevisePosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  Form_Resize()                                           ***
'******************************************************************************
'***    FONCTION:   Repositionner les controles sur la form                 ***
'***                                                                        ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub Form_Resize()
    ' Je repositionne les boutons sur la form
    PAVTypeDeviseCDButtonPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PAVTypeDeviseCDButton_Click(Index As Integer)           ***
'******************************************************************************
'***    FONCTION:   l'user a appuyer un bouton de devise                    ***
'***                                                                        ***
'******************************************************************************
'***    ENTREE:     Index - Numero du bouton appuyer                        ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PAVTypeDeviseCDButton_Click(Index As Integer)
    ' Je rend le focus au champs de saissie courants
    Zone.SetFocus
    ' Fin
End Sub

