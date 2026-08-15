VERSION 5.00
Begin VB.Form FormPavMenuPrincipale 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4635
   ClientLeft      =   3330
   ClientTop       =   3330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton PAVMenuPrincipaleCDButton 
      Caption         =   "MENU 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   0
      Left            =   0
      MaskColor       =   &H80000010&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton PAVMenuPrincipaleCDButton 
      Caption         =   "MENU 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   1
      Left            =   1590
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1500
   End
   Begin VB.CommandButton PAVMenuPrincipaleCDButton 
      Caption         =   "MENU 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   2
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1500
   End
   Begin VB.CommandButton PAVMenuPrincipaleCDButton 
      Caption         =   "MENU 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   3
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1500
   End
   Begin VB.CommandButton PAVMenuPrincipaleCDButton 
      Caption         =   "MENU 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   4
      Left            =   1590
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1590
      Width           =   1500
   End
   Begin VB.CommandButton PAVMenuPrincipaleCDButton 
      Caption         =   "MENU 8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   5
      Left            =   3180
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3150
      Width           =   1500
   End
   Begin VB.CommandButton PAVMenuPrincipaleCDButton 
      Caption         =   "MENU 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   6
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1500
   End
   Begin VB.CommandButton PAVMenuPrincipaleCDButton 
      Caption         =   "MENU 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   7
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3150
      Width           =   1500
   End
   Begin VB.CommandButton PAVMenuPrincipaleCDButton 
      Caption         =   "MENU 7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   8
      Left            =   1590
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3150
      Width           =   1500
   End
End
Attribute VB_Name = "FormPavMenuPrincipale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***                      (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - PAVER DU MENU PRINCIPALE                                 ***
'******************************************************************************
'***    FONCTION:- INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES      ***
'***             - RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU         ***
'***               MODULES PRINCIPALE DU MICRO POINT VENTE                  ***
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
'***    PROCEDURE:FormPAVMenuPrincipaleInitialise()                         ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormPAVMenuPrincipaleInitialise                           ***
'******************************************************************************
Private Sub FormPAVMenuPrincipaleInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVMenuPrincipaleCDButtonInitialise()                     ***
'******************************************************************************
'***    FONCTION: Initialiser - Le texte des boutons                        ***
'***                          - Les images des boutons                      ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVMenuPrincipaleCDButtonInitialise                       ***
'******************************************************************************
Private Sub PAVMenuPrincipaleCDButtonInitialise()
    Dim I As Integer                                                            ' Le numero du bouton courant
    ' Je fait defiler tous les bouton
    For I = 0 To PAVMenuPrincipaleCDButton.Count - 1
        ' Je charge l'image du bouton
        PAVMenuPrincipaleCDButton(I).Picture = RessourceIconLoad(610 + I)
        ' Je charge le texte du bouton
        PAVMenuPrincipaleCDButton(I).Caption = RessourceStringLoad(610 + I)
    ' Je ne sort que si j'ai fini de faire defiler les boutons
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:FormPAVMenuPrincipalePosition()                           ***
'******************************************************************************
'***    FONCTION: Positionner la form a l'ecran                             ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormPAVMenuPrincipalePosition                             ***
'******************************************************************************
Private Sub FormPAVMenuPrincipalePosition()
    ' Je positionne la feuille sur l'ecran
    ControlInitialise FormPavMenuPrincipale, Screen.Width * 2 / 21, Screen.Height * 2 / 15, Screen.Width * 15 / 21, Screen.Height * 11 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVMenuPrincipaleCDButtonPosition()                       ***
'******************************************************************************
'***    FONCTION: Repositionner les boutons sur la form                     ***
'***              Redimensionner la taille des fontes                       ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVMenuPrincipaleCDButtonPosition                         ***
'******************************************************************************
Private Sub PAVMenuPrincipaleCDButtonPosition()
    Dim I As Integer                                                            ' Le numero du bouton courant
    Dim CDButtonLef As Single                                                   ' La position gauche du bouton courant
    Dim CDButtonTop As Single                                                   ' La hauteur du bouton courant
    Dim CDButtonLgr As Single                                                   ' La longueur du bouton courant
    Dim CDButtonLar As Single                                                   ' La largueur du bouton courant
    ' Je fait defiler tous les boutons
    For I = 0 To PAVMenuPrincipaleCDButton.Count - 1
        ' Je calcule la position gauche du boutons
        CDButtonLef = (I Mod 3) * FormPavMenuPrincipale.ScaleWidth / 3
        ' Je calcule la hauteur du boutons
        CDButtonTop = (Int(I / 3)) * FormPavMenuPrincipale.ScaleHeight / 3
        ' Je calcule la longueur du boutons
        CDButtonLgr = FormPavMenuPrincipale.ScaleWidth / 3
        ' Je calcule la largueur du boutons
        CDButtonLar = FormPavMenuPrincipale.ScaleHeight / 3
        ' Je repositionne le bouton et ajuste la taille des fontes
        ControlInitialise PAVMenuPrincipaleCDButton(I), CDButtonLef, CDButtonTop, CDButtonLgr, CDButtonLar
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
    FormPAVMenuPrincipaleInitialise
    ' J'initialise les boutons
    PAVMenuPrincipaleCDButtonInitialise
    ' Je positionne la form a l'ecran
    FormPAVMenuPrincipalePosition
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
    PAVMenuPrincipaleCDButtonPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PAVMenuPrincipaleCDButton_Click(Index As Integer)       ***
'******************************************************************************
'***    FONCTION:   L'user a appuyer un bouton du clavier alpha             ***
'***                                                                        ***
'******************************************************************************
'***    ENTREE:     Index - Numero du bouton appuyer                        ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PAVMenuPrincipaleCDButton_Click(Index As Integer)
    ' Je regarde sur quel bouton a appuyer l'user
    Select Case Index
        Case 0
            ' C'est le Bouton Saisie Des Ventes
            ' Je lance la saissie des ventes
            MenuSaisieDesVentes
    End Select
    ' Fin
End Sub
