VERSION 5.00
Begin VB.Form FormPavTypeReglement 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PAVTypeReglementCDButton 
      Caption         =   "REG 9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Index           =   8
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4215
      Width           =   1665
   End
   Begin VB.CommandButton PAVTypeReglementCDButton 
      Caption         =   "REG 8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Index           =   7
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4215
      Width           =   1305
   End
   Begin VB.CommandButton PAVTypeReglementCDButton 
      Caption         =   "REG 7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Index           =   6
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   1185
   End
   Begin VB.CommandButton PAVTypeReglementCDButton 
      Caption         =   "REG 6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Index           =   5
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1665
   End
   Begin VB.CommandButton PAVTypeReglementCDButton 
      Caption         =   "REG 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Index           =   4
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2535
      Width           =   1305
   End
   Begin VB.CommandButton PAVTypeReglementCDButton 
      Caption         =   "REG 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Index           =   3
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1185
   End
   Begin VB.CommandButton PAVTypeReglementCDButton 
      Caption         =   "REG 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Index           =   2
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1785
   End
   Begin VB.CommandButton PAVTypeReglementCDButton 
      Caption         =   "REG 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Index           =   1
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1065
   End
   Begin VB.CommandButton PAVTypeReglementCDButton 
      Caption         =   "REG1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1185
   End
End
Attribute VB_Name = "FormPavTypeReglement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***                      (31/10/2000) ***
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
'***    PROCEDURE:FormPavTypeReglementInitialise()                          ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormPavTypeReglementInitialise()                          ***
'******************************************************************************
Private Sub FormPavTypeReglementInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTypeReglementCDButtonInitialise()                      ***
'******************************************************************************
'***    FONCTION: Initialiser - Le texte des boutons                        ***
'***                          - Les images des boutons                      ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTypeReglementCDButtonInitialise()                      ***
'******************************************************************************
Private Sub PAVTypeReglementCDButtonInitialise()
    Dim I As Integer                                                            ' Le numero du bouton courant
    ' Je fait defiler tous les bouton
    For I = 0 To PAVTypeReglementCDButton.Count - 1
        ' Je charge l'image du bouton
        PAVTypeReglementCDButton(I).Picture = RessourceIconLoad(760 + I)
        ' Je charge le texte du bouton
        PAVTypeReglementCDButton(I).Caption = RessourceStringLoad(760 + I)
    ' Je ne sort que si j'ai fini de faire defiler les boutons
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:FormPavTypeReglementPosition()                            ***
'******************************************************************************
'***    FONCTION: positionner la form a l'ecran                             ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormPavTypeReglementPosition()                            ***
'******************************************************************************
Private Sub FormPavTypeReglementPosition()
    ' Je positionne la feuille a l'ecran
    ControlInitialise FormPavTypeReglement, Screen.Width * 16.5 / 21, Screen.Height * 4 / 15, Screen.Width * 4.5 / 21, Screen.Height * 3 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PAVTypeReglementCDButtonPosition()                        ***
'******************************************************************************
'***    FONCTION: Repositionner les boutons sur la form                     ***
'***              Redimensionner la taille des fontes                       ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PAVTypeReglementCDButtonPosition()                        ***
'******************************************************************************
Private Sub PAVTypeReglementCDButtonPosition()
    Dim I As Integer                                                            ' Le numero du bouton courant
    Dim CDButtonLef As Single                                                   ' La position gauche du bouton courant
    Dim CDButtonTop As Single                                                   ' La hauteur du bouton courant
    Dim CDButtonLgr As Single                                                   ' La longueur du bouton courant
    Dim CDButtonLar As Single                                                   ' La largueur du bouton courant
    ' Je fait defiler tous les boutons
    For I = 0 To PAVTypeReglementCDButton.Count - 1
        ' Je calcule la position gauche du boutons
        CDButtonLef = (I Mod 3) * FormPavTypeReglement.ScaleWidth / 3
        ' Je calcule la hauteur du boutons
        CDButtonTop = (Int(I / 3)) * FormPavTypeReglement.ScaleHeight / 3
        ' Je calcule la longueur du boutons
        CDButtonLgr = FormPavTypeReglement.ScaleWidth / 3
        ' Je calcule la largueur du boutons
        CDButtonLar = FormPavTypeReglement.ScaleHeight / 3
        ' Je repositionne le bouton et ajuste la taille des fontes
        ControlInitialise PAVTypeReglementCDButton(I), CDButtonLef, CDButtonTop, CDButtonLgr, CDButtonLar
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
    FormPavTypeReglementInitialise
    ' J'initialise les boutons
    PAVTypeReglementCDButtonInitialise
    ' Je positionne la form a l'ecran
    FormPavTypeReglementPosition
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
    PAVTypeReglementCDButtonPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PAVTypeReglementCDButton_Click(Index As Integer)        ***
'******************************************************************************
'***    FONCTION:   l'user a appuyer un bouton du type de reglement         ***
'***                                                                        ***
'******************************************************************************
'***    ENTREE:     Index - Numero du bouton appuyer                        ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PAVTypeReglementCDButton_Click(Index As Integer)
    ' Je regarde sur quel touche l'user a appuyer
    Select Case Index
        Case ButtonReglementCheque
            ' C'est le bouton reglement par cheque
            ' Je change le type de reglement en reglement par cheque
            TicketReglementTypeReglementChange (ReglementCheque)
        Case ButtonReglementEspece
            ' C'est le bouton reglement par espece
            ' Je change le type de reglement en reglement par espece
            TicketReglementTypeReglementChange (ReglementEspece)
        Case ButtonReglementCarte
            ' C'est le bouton reglement par carte
            ' Je change le type de reglement en reglement par carte
            TicketReglementTypeReglementChange (ReglementCarte)
        Case ButtonReglementUtilisationAvoir
            ' C'est le bouton reglement par utilisation d'avoir
            ' Je change le type de reglement en reglement par utilisation d'avoir
            TicketReglementTypeReglementChange (ReglementUtilisationAvoir)
        Case ButtonReglementDiver1
            ' C'est le bouton reglement par diver1
            ' Je change le type de reglement en reglement par diver1
            TicketReglementTypeReglementChange (ReglementDiver1)
        Case ButtonReglementDiver2
            ' c'est le bouton reglement par diver2
            ' Je change le type de reglement en reglement par diver2
            TicketReglementTypeReglementChange (ReglementDiver2)
    End Select
    ' Fin
End Sub

