VERSION 5.00
Begin VB.Form PANTicketDonneTotalForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label PANTicketDonneTotalLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "PRIXNET2"
      DataField       =   "TICKETDONNETOTALNET"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "# ##0,00 "
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   2
      EndProperty
      DataMember      =   "SelTicketDonneTotalDevise2Cmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   765
      Index           =   7
      Left            =   3840
      TabIndex        =   8
      Top             =   2400
      Width           =   2280
   End
   Begin VB.Label PANTicketDonneTotalLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "DEVISE2"
      DataField       =   "DEVISECAR"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "# ##0,00 "
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   0
      EndProperty
      DataMember      =   "SelTicketDonneTotalDevise2Cmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   765
      Index           =   8
      Left            =   6360
      TabIndex        =   7
      Top             =   2280
      Width           =   2040
   End
   Begin VB.Label PANTicketDonneTotalLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "DEVISE2"
      DataField       =   "DEVISECAR"
      DataMember      =   "SelTicketDonneTotalDevise2Cmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   735
      Index           =   3
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   1755
   End
   Begin VB.Label PANTicketDonneTotalLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRIXBRUT2"
      DataField       =   "TICKETDONNETOTALBRUT"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "# ##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   1
      EndProperty
      DataMember      =   "SelTicketDonneTotalDevise2Cmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   645
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1680
   End
   Begin VB.Label PANTicketDonneTotalLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "DEVISE1"
      DataField       =   "DEVISECAR"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "# ##0,00 "
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   0
      EndProperty
      DataMember      =   "SelTicketDonneTotalDevise1Cmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   765
      Index           =   6
      Left            =   6360
      TabIndex        =   4
      Top             =   1560
      Width           =   2040
   End
   Begin VB.Label PANTicketDonneTotalLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "PRIXNET1"
      DataField       =   "TICKETDONNETOTALNET"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "# ##0,00 "
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   2
      EndProperty
      DataMember      =   "SelTicketDonneTotalDevise1Cmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   765
      Index           =   5
      Left            =   3840
      TabIndex        =   3
      Top             =   1680
      Width           =   2280
   End
   Begin VB.Label PANTicketDonneTotalLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTALAPAYER"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "# ##0,00 "
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   0
      EndProperty
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   1365
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   3480
   End
   Begin VB.Label PANTicketDonneTotalLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRIXBRUT1"
      DataField       =   "TICKETDONNETOTALBRUT"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "# ##0,00 "
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   2
      EndProperty
      DataMember      =   "SelTicketDonneTotalDevise1Cmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   645
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label PANTicketDonneTotalLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "DEVISE1"
      DataField       =   "DEVISECAR"
      DataMember      =   "SelTicketDonneTotalDevise1Cmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   735
      Index           =   1
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1755
   End
   Begin VB.Image PANTicketDonneTotalImage 
      Height          =   4320
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8640
   End
End
Attribute VB_Name = "PANTicketDonneTotalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. 1.0 VB )                  (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - AFFICHAGE DU TOTAL DES DONNES DU TICKET EN COURS         ***
'******************************************************************************
'***    FONCTION:- INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES      ***
'***             - RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU         ***
'***               MODULES PRINCIPALES DE LA GESTION DES DONNER DU TICKET   ***
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
'***    PROCEDURE:PANTicketDonneTotalFormInitialise()                       ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketDonneTotalFormInitialise                         ***
'******************************************************************************
Private Sub PANTicketDonneTotalFormInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketDonneTotalImageInitialise()                      ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketDonneTotalImageInitialise                        ***
'******************************************************************************
Private Sub PANTicketDonneTotalImageInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketDonneTotalLabelInitialise()                      ***
'******************************************************************************
'***    FONCTION: Initialiser  - Les textes des labels                      ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketDonneTotalLabelInitialise                        ***
'******************************************************************************
Private Sub PANTicketDonneTotalLabelInitialise()
    Dim I As Integer                                                            ' Le numero de label courant
    ' Je fais defiler les labels
    For I = 0 To PANTicketDonneTotalLabel.Count - 1
        ' Je Regarde sur quel zone je suis
        Select Case I
            Case ZoneAPayer
                ' Je suis sur le label a payer
                ' Je recupere le texte du label
                PANTicketDonneTotalLabel(I).Caption = RessourceStringLoad(770 + I)
            Case Else
                ' Je suis sur une autre zone
        End Select
    ' Je ne sort que si j'ai fini de faire defiler les labels
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketDonneTotalFormPosition()                         ***
'******************************************************************************
'***    FONCTION: Positionner la form sur l'ecran                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketDonneTotalFormPosition                           ***
'******************************************************************************
Private Sub PANTicketDonneTotalFormPosition()
    ' Je positionne la form a l'ecran
    ControlInitialise PANTicketDonneTotalForm, Screen.Width * 12 / 21, Screen.Height * 9 / 15, Screen.Width * 9 / 21, Screen.Height * 3 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketDonneTotalImagePosition()                        ***
'******************************************************************************
'***    FONCTION: Pour positionner l'image sur la form                      ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketDonneTotalImagePosition                          ***
'******************************************************************************
Private Sub PANTicketDonneTotalImagePosition()
    ' Je positionne l'image sur la form
    ControlInitialise PANTicketDonneTotalImage, 0, 0, PANTicketDonneTotalForm.ScaleWidth, PANTicketDonneTotalForm.ScaleHeight
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketDonneTotalLabelPosition()                        ***
'******************************************************************************
'***    FONCTION: Pour positionner les labels sur la form                   ***
'***              redimensionnement des fontes                              ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketDonneTotalLabelPosition                          ***
'******************************************************************************
Private Sub PANTicketDonneTotalLabelPosition()
    Dim I As Integer                                                            ' Le numero du label courant
    Dim LabelLef As Single                                                      ' La position gauche du label courant
    Dim LabelTop As Single                                                      ' La position haut du label courant
    Dim LabelLgr As Single                                                      ' La longueur du label courant
    Dim LabelLar As Single                                                      ' La largueur du label courant
    ' Je fait defiler tous les labels
    For I = 0 To PANTicketDonneTotalLabel.Count - 1
        ' Je regarde sur quel label je suis
        Select Case I
            Case ZoneTotalBrut1
                ' Je suis dans la zone total brut
                ' Je calcule la position gauche du label
                LabelLef = 0
                ' Je calcule la position haut du label
                LabelTop = 0
                ' Je calcule la longueur du label
                LabelLgr = PANTicketDonneTotalForm.ScaleWidth * 8 / 18
                ' Je calcule la largueur du label
                LabelLar = PANTicketDonneTotalForm.ScaleHeight / 3
            Case ZoneTotalBrutDevise1
                ' Je suis dans la zone total brut devise
                ' Je calcule la position gauche du label
                LabelLef = PANTicketDonneTotalForm.ScaleWidth * 8 / 18
                ' Je calcule la position haut du label
                LabelTop = 0
                ' Je calcule la longueur du label
                LabelLgr = PANTicketDonneTotalForm.ScaleWidth * 1 / 18
                ' Je calcule la largueur du label
                LabelLar = PANTicketDonneTotalForm.ScaleHeight / 3
            Case ZoneTotalBrut2
                ' Je suis dans la zone total brut
                ' Je calcule la position gauche du label
                LabelLef = 0
                ' Je calcule la position haut du label
                LabelTop = PANTicketDonneTotalForm.ScaleHeight / 3
                ' Je calcule la longueur du label
                LabelLgr = PANTicketDonneTotalForm.ScaleWidth * 8 / 18
                ' Je calcule la largueur du label
                LabelLar = PANTicketDonneTotalForm.ScaleHeight / 6
            Case ZoneTotalBrutDevise2
                ' Je suis dans la zone total brut devise
                ' Je calcule la position gauche du label
                LabelLef = PANTicketDonneTotalForm.ScaleWidth * 8 / 18
                ' Je calcule la position haut du label
                LabelTop = PANTicketDonneTotalForm.ScaleHeight / 3
                ' Je calcule la longueur du label
                LabelLgr = PANTicketDonneTotalForm.ScaleWidth * 1 / 18
                ' Je calcule la largueur du label
                LabelLar = PANTicketDonneTotalForm.ScaleHeight / 6
            Case ZoneAPayer
                ' Je suis dans la zone a payer
                ' Je calcule la position gauche du label
                LabelLef = 0
                ' Je calcule la position haut du label
                LabelTop = PANTicketDonneTotalForm.ScaleHeight * 2 / 3
                ' Je calcule la longueur du label
                LabelLgr = PANTicketDonneTotalForm.ScaleWidth / 2
                ' Je calcule la largueur du label
                LabelLar = PANTicketDonneTotalForm.ScaleHeight / 3
            Case ZoneTotalNet1
                ' Je suis dans la zone total net
                ' Je calcule la position gauche du label
                LabelLef = PANTicketDonneTotalForm.ScaleWidth / 2
                ' Je calcule la position haut du label
                LabelTop = PANTicketDonneTotalForm.ScaleHeight * 2 / 3
                ' Je calcule la longueur du label
                LabelLgr = PANTicketDonneTotalForm.ScaleWidth * 8 / 18
                ' Je calcule la largueur du label
                LabelLar = PANTicketDonneTotalForm.ScaleHeight / 6
            Case ZoneTotalNetDevise1
                ' Je suis dans la zone total net devise
                ' Je calcule la position gauche du label
                LabelLef = PANTicketDonneTotalForm.ScaleWidth * 17 / 18
                ' Je calcule la position haut du label
                LabelTop = PANTicketDonneTotalForm.ScaleHeight * 2 / 3
                ' Je calcule la longueur du label
                LabelLgr = PANTicketDonneTotalForm.ScaleWidth * 1 / 18
                ' Je calcule la largueur du label
                LabelLar = PANTicketDonneTotalForm.ScaleHeight / 6
            Case ZoneTotalNet2
                ' Je suis dans la zone total net
                ' Je calcule la position gauche du label
                LabelLef = PANTicketDonneTotalForm.ScaleWidth / 2
                ' Je calcule la position haut du label
                LabelTop = PANTicketDonneTotalForm.ScaleHeight * 5 / 6
                ' Je calcule la longueur du label
                LabelLgr = PANTicketDonneTotalForm.ScaleWidth * 8 / 18
                ' Je calcule la largueur du label
                LabelLar = PANTicketDonneTotalForm.ScaleHeight / 6
            Case ZoneTotalNetDevise2
                ' Je suis dans la zone total net devise
                ' Je calcule la position gauche du label
                LabelLef = PANTicketDonneTotalForm.ScaleWidth * 17 / 18
                ' Je calcule la position haut du label
                LabelTop = PANTicketDonneTotalForm.ScaleHeight * 5 / 6
                ' Je calcule la longueur du label
                LabelLgr = PANTicketDonneTotalForm.ScaleWidth * 1 / 18
                ' Je calcule la largueur du label
                LabelLar = PANTicketDonneTotalForm.ScaleHeight / 6
        End Select
        ' Je repositionne le controle et modifie la taille des polices
        ControlInitialise PANTicketDonneTotalLabel(I), LabelLef, LabelTop, LabelLgr, LabelLar
    ' Je m'arrete une fois que j'ai fini de faire defiler les labels
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
    PANTicketDonneTotalFormInitialise
    ' J'initialise l'image
    PANTicketDonneTotalImageInitialise
    ' J'initialise les labels
    PANTicketDonneTotalLabelInitialise
    ' Je positionne la feuille sur l'ecran
    PANTicketDonneTotalFormPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  Form_Resize()                                           ***
'******************************************************************************
'***    FONCTION:   Repositionner les controles sur la form                 ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Public Sub Form_Resize()
    ' Je repositionne l'image
    PANTicketDonneTotalImagePosition
    ' Je repositionne les labels
    PANTicketDonneTotalLabelPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PANTicketDonneTotalLabel_Change(Index As Integer)       ***
'******************************************************************************
'***    FONCTION:   Pour redimensionner la taille des fontes                ***
'******************************************************************************
'***    ENTREE:     index - Le numero du label changer                      ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PANTicketDonneTotalLabel_Change(Index As Integer)
    ' Je redimensionne les fontes
    FontAjuste PANTicketDonneTotalLabel(Index)
    ' Fin
End Sub

