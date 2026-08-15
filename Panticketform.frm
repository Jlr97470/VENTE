VERSION 5.00
Begin VB.Form PANTicketForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.Timer PANTicketTimer 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label PANTicketLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HEURE:"
      DataField       =   "TICKETDONNEHEURE"
      DataMember      =   "InsTicketDonneCmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   6600
      TabIndex        =   9
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label PANTicketLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TICKETDATE:"
      DataField       =   "TICKETDATE"
      DataMember      =   "InsTicketCmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   4800
      TabIndex        =   8
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label PANTicketLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TICKETNUMERO:"
      DataField       =   "TICKETNUMERO"
      DataMember      =   "InsTicketCmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   5760
      TabIndex        =   7
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label PANTicketLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CAISSELIB:"
      DataField       =   "CAISSELIBELLE"
      DataMember      =   "SelMagasinCaisseCmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   3480
      TabIndex        =   6
      Top             =   960
      Width           =   885
   End
   Begin VB.Label PANTicketLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CAISSECAR:"
      DataField       =   "CAISSECAR"
      DataMember      =   "SelMagasinCaisseCmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label PANTicketLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAGASINLIB:"
      DataField       =   "MAGASINLIBELLE"
      DataMember      =   "SelMagasinCaisseCmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   600
      TabIndex        =   4
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label PANTicketLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAGASINCAR:"
      DataField       =   "MAGASINCAR"
      DataMember      =   "SelMagasinCaisseCmd"
      DataSource      =   "DAEVenteDSR"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label PANTicketLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TICKET:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   645
   End
   Begin VB.Label PANTicketLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAGASIN:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   795
   End
   Begin VB.Label PANTicketLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CAISSE:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   645
   End
   Begin VB.Image PANTicketImage 
      Height          =   7200
      Left            =   0
      Picture         =   "PANTicketForm.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "PANTicketForm"
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
'***               MODULES PRINCIPALES POUR LA GESTION DU TICKET            ***
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
'***    PROCEDURE:PANTicketFormInitialise()                                 ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketFormInitialise()                                 ***
'******************************************************************************
Private Sub PANTicketFormInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketImageInitialise()                                ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketImageInitialise()                                ***
'******************************************************************************
Private Sub PANTicketImageInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketLabelInitialise()                                ***
'******************************************************************************
'***    FONCTION: Initialiser  - Les textes des labels                      ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketLabelInitialise()                                ***
'******************************************************************************
Private Sub PANTicketLabelInitialise()
    Dim I As Integer                                                            ' Le numero du label courant
    ' Je fait defiler les labels
    For I = 0 To PANTicketLabel.Count - 1
        ' Je regarde sur quel label je suis
        Select Case I
            Case ZoneMagasin, ZoneCaisse, ZoneTicket
                ' Je recupere le texte du labels
                PANTicketLabel(I).Caption = RessourceStringLoad(600 + I)
            Case Else
                ' Je suis sur une autre zone
        End Select
    ' Je ne sort que si j'ai fini de faire defiler les labels
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketFormPosition()                                   ***
'******************************************************************************
'***    FONCTION: Positionner la form sur l'ecran                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketFormPosition()                                   ***
'******************************************************************************
Private Sub PANTicketFormPosition()
    ' Je positionne la feuille a l'ecran
    ControlInitialise PANTicketForm, 0, 0, Screen.Width * 16.5 / 21, Screen.Height * 2 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketImagePosition()                                  ***
'******************************************************************************
'***    FONCTION: Positionner l'image sur la form                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketImagePosition()                                  ***
'******************************************************************************
Private Sub PANTicketImagePosition()
    ' Je positionne l'image sur la form
    ControlInitialise PANTicketImage, 0, 0, PANTicketForm.ScaleWidth, PANTicketForm.ScaleHeight
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANTicketLabelPosition()                                  ***
'******************************************************************************
'***    FONCTION: Pour positionner les labels sur la form                   ***
'***              redimensionnement des fontes                              ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANTicketLabelPos()                                       ***
'******************************************************************************
Private Sub PANTicketLabelPosition()
    Dim I As Integer                                                            ' Le numero du label courant
    Dim LabelLef As Single                                                      ' La position gauche du label courant
    Dim LabelTop As Single                                                      ' La position haut du label courant
    Dim LabelLgr As Single                                                      ' La longueur du label courant
    Dim LabelLar As Single                                                      ' La largueur du label courant
    ' Je fait defiler tous les labels
    For I = 0 To PANTicketLabel.Count - 1
        ' Je regarde sur quel label je suis
        Select Case I
            Case ZoneTicketDate
                ' Je suis dans la zone date du ticket
                ' Je calcule la position gauche du label
                LabelLef = PANTicketForm.ScaleWidth * 2 / 3
                ' Je calcule la position haut du label
                LabelTop = PANTicketForm.ScaleHeight * 2 / 3
                ' Je calcule la longueur du label
                LabelLgr = PANTicketForm.ScaleWidth / 6
                ' Je calcule la largueur du label
                LabelLar = PANTicketForm.ScaleHeight / 3
            Case ZoneTicketDonneHeure
                ' Je suis dans la zone heure des donnees du ticket
                ' Je calcule la position gauche du label
                LabelLef = PANTicketForm.ScaleWidth * 5 / 6
                ' Je calcule la position haut du label
                LabelTop = PANTicketForm.ScaleHeight * 2 / 3
                ' Je calcule la longueur du label
                LabelLgr = PANTicketForm.ScaleWidth / 6
                ' Je calcule la largueur du label
                LabelLar = PANTicketForm.ScaleHeight / 3
            Case Else
                ' Je calcule la position gauche du label
                LabelLef = (I Mod 3) * PANTicketForm.ScaleWidth / 3
                ' Je calcule la position haut du label
                LabelTop = (Int(I / 3)) * PANTicketForm.ScaleHeight / 3
                ' Je calcule la longueur du label
                LabelLgr = PANTicketForm.ScaleWidth / 3
                ' Je calcule la largueur du label
                LabelLar = PANTicketForm.ScaleHeight / 3
        End Select
        ' Je repositionne le controle et modifie la taille des polices
        ControlInitialise PANTicketLabel(I), LabelLef, LabelTop, LabelLgr, LabelLar
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
    PANTicketFormInitialise
    ' J'initialise l'image
    PANTicketImageInitialise
    ' J'initialise les labels
    PANTicketLabelInitialise
    ' Je positionne la feuille a l'ecran
    PANTicketFormPosition
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
    PANTicketImagePosition
    ' Je repositionne les labels
    PANTicketLabelPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PANTicketLabel_Change(Index As Integer)                 ***
'******************************************************************************
'***    FONCTION:   Pour redimensionner la taille des fontes                ***
'******************************************************************************
'***    ENTREE:     index - le numero du label changer                      ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PANTicketLabel_Change(Index As Integer)
    ' Je recalcule la taille des font
    FontAjuste PANTicketLabel(Index)
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PANTicketLabel_Click(Index As Integer)                  ***
'******************************************************************************
'***    FONCTION:   Pour appeller une liste deroulante                      ***
'******************************************************************************
'***    ENTREE:     index - le numero du label appuyer                      ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PANTicketLabel_Click(Index As Integer)
    ' Je demande a avoir une liste de choix
    ListeValidation
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PANTicketLabel_Click(Index As Integer)                  ***
'******************************************************************************
'***    FONCTION:   Pour mettre a jour l'heure de la ligne du ticket        ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PANTicketTimer_Timer()
    ' Je met a jour l'affichage de l'heure
    PANTicketLabel(ZoneTicketDonneHeure).Caption = Format(Time, "")
    ' Fin
End Sub
