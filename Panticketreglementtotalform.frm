VERSION 5.00
Begin VB.Form PANTicketReglementTotalForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label PANTicketReglementTotalLabel 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
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
      DataMember      =   "SelTicketReglementTotalDevise2Cmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   4
      Left            =   2520
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label PANTicketReglementTotalLabel 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "PRIX2"
      DataField       =   "TICKETREGLEMENTRESTEAPAYER"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "# ##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   1
      EndProperty
      DataMember      =   "SelTicketReglementTotalDevise2Cmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label PANTicketReglementTotalLabel 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
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
      DataMember      =   "SelTicketReglementTotalDevise1Cmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label PANTicketReglementTotalLabel 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "PRIX1"
      DataField       =   "TICKETREGLEMENTRESTEAPAYER"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "# ##0,00 "
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   2
      EndProperty
      DataMember      =   "SelTicketReglementTotalDevise1Cmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label PANTicketReglementTotalLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "RESTEAPAYER"
      ForeColor       =   &H000000FF&
      Height          =   855
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image PANTicketReglementTotalImage 
      Height          =   3135
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "PANTicketReglementTotalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. 1.0 VB )                  (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - AFFICHAGE DU TICKET EN COURS                             ***
'******************************************************************************
'***    FONCTION:- INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES      ***
'***             - RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU         ***
'***               MODULES PRINCIPALES DE LA GESTION DES REGLEMENTS DU      ***
'***               TICKETS                                                  ***
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
'***    PROCEDURE:PANTicketReglementTotalFormInitialise()                   ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketReglementTotalFormInitialise                     ***
'******************************************************************************
Private Sub PANTicketReglementTotalFormInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketReglementTotalImageInitialise()                  ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketReglementTotalImageInitialise                    ***
'******************************************************************************
Private Sub PANTicketReglementTotalImageInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketReglementTotalLabelInitialise()                  ***
'******************************************************************************
'***    FONCTION: Initialiser  - Les textes des labels                      ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketReglementTotalLabelInitialise                    ***
'******************************************************************************
Private Sub PANTicketReglementTotalLabelInitialise()
    Dim I As Integer                                                            ' Le numero de label courant
    ' Je fais defiler les labels
    For I = 0 To PANTicketReglementTotalLabel.Count - 1
        ' Je recupere le texte du label
        PANTicketReglementTotalLabel(I).Caption = RessourceStringLoad(1190 + I)
    ' Je ne sort que si j'ai fini de faire defiler les labels
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketReglementTotalFormPosition()                     ***
'******************************************************************************
'***    FONCTION: Positionner la form sur l'ecran                           ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketReglementTotalFormPosition                       ***
'******************************************************************************
Private Sub PANTicketReglementTotalFormPosition()
    ' Je positionne la form a l'ecran
    ControlInitialise PANTicketReglementTotalForm, Screen.Width * 12 / 21, Screen.Height * 12 / 15, Screen.Width * 9 / 21, Screen.Height * 3 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketReglementTotalImagePosition()                    ***
'******************************************************************************
'***    FONCTION: pour positionner l'image sur la form                      ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketReglementTotalImagePosition                      ***
'******************************************************************************
Private Sub PANTicketReglementTotalImagePosition()
    ' Je positionne l'image sur la form
    ControlInitialise PANTicketReglementTotalImage, 0, 0, PANTicketReglementTotalForm.ScaleWidth, PANTicketReglementTotalForm.ScaleHeight
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketReglementTotalLabelPosition()                    ***
'******************************************************************************
'***    FONCTION: Pour positionner les labels sur la form                   ***
'***              Redimensionnement des fontes                              ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketReglementTotalLabelPosition                      ***
'******************************************************************************
Private Sub PANTicketReglementTotalLabelPosition()
    Dim I As Integer                                                            ' Le numero du label courant
    Dim LabelLef As Single                                                      ' La position gauche du label courant
    Dim LabelTop As Single                                                      ' La position haut du label courant
    Dim LabelLgr As Single                                                      ' La longueur du label courant
    Dim LabelLar As Single                                                      ' La largueur du label courant
    ' Je fait defiler tous les labels
    For I = 0 To PANTicketReglementTotalLabel.Count - 1
        ' Je regarde sur quel zone je suis
        Select Case I
            Case ZoneResteAPayer
                ' Je suis sur la zone reste a payer
                ' Je calcule la position gauche du label
                LabelLef = 0
                ' Je calcule la position haut du label
                LabelTop = 0
                ' Je calcule la longueur du label
                LabelLgr = PANTicketReglementTotalForm.ScaleWidth
                ' Je calcule la largueur du label
                LabelLar = PANTicketReglementTotalForm.ScaleHeight / 4
                ' Je repositionne le controle et modifie la taille des polices
            Case ZoneTotalReste1
                ' Je suis dans la zone total reste devise 1
                ' Je calcule la position gauche du label
                LabelLef = 0
                ' Je calcule la position haut du label
                LabelTop = PANTicketReglementTotalForm.ScaleHeight / 4
                ' Je calcule la longueur du label
                LabelLgr = PANTicketReglementTotalForm.ScaleWidth * 16 / 18
                ' Je calcule la largueur du label
                LabelLar = PANTicketReglementTotalForm.ScaleHeight / 2
            Case ZoneTotalResteDevise1
                ' Je suis dans la zone total reste devise 1
                ' Je calcule la position gauche du label
                LabelLef = PANTicketReglementTotalForm.ScaleWidth * 16 / 18
                ' Je calcule la position haut du label
                LabelTop = PANTicketReglementTotalForm.ScaleHeight / 4
                ' Je calcule la longueur du label
                LabelLgr = PANTicketReglementTotalForm.ScaleWidth * 2 / 18
                ' Je calcule la largueur du label
                LabelLar = PANTicketReglementTotalForm.ScaleHeight / 4
            Case ZoneTotalReste2
                ' Je suis dans la zone total reste devise 2
                ' Je calcule la position gauche du label
                LabelLef = 0
                ' Je calcule la position haut du label
                LabelTop = PANTicketReglementTotalForm.ScaleHeight * 3 / 4
                ' Je calcule la longueur du label
                LabelLgr = PANTicketReglementTotalForm.ScaleWidth * 16 / 18
                ' Je calcule la largueur du label
                LabelLar = PANTicketReglementTotalForm.ScaleHeight / 4
            Case ZoneTotalResteDevise2
                ' Je suis dans la zone total reste devise 2
                ' Je calcule la position gauche du label
                LabelLef = PANTicketReglementTotalForm.ScaleWidth * 16 / 18
                ' Je calcule la position haut du label
                LabelTop = PANTicketReglementTotalForm.ScaleHeight * 3 / 4
                ' Je calcule la longueur du label
                LabelLgr = PANTicketReglementTotalForm.ScaleWidth * 2 / 18
                ' Je calcule la largueur du label
                LabelLar = PANTicketReglementTotalForm.ScaleHeight / 4
                
        End Select
        ' Jinitialise le controle
        ControlInitialise PANTicketReglementTotalLabel(I), LabelLef, LabelTop, LabelLgr, LabelLar
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
    PANTicketReglementTotalFormInitialise
    ' J'initialise l'image
    PANTicketReglementTotalImageInitialise
    ' J'initialise les labels
    PANTicketReglementTotalLabelInitialise
    ' Je positionne la feuille sur l'ecran
    PANTicketReglementTotalFormPosition
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
    PANTicketReglementTotalImagePosition
    ' Je repositionne les labels
    PANTicketReglementTotalLabelPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PANTicketReglementTotalLabel_Change(Index As Integer)   ***
'******************************************************************************
'***    FONCTION:   Pour redimensionner la taille des fontes                ***
'******************************************************************************
'***    ENTREE:     index - Le numero du label changer                      ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PANTicketReglementTotalLabel_Change(Index As Integer)
    ' Je redimensionne les fontes
    FontAjuste PANTicketReglementTotalLabel(Index)
    ' Fin
End Sub

