VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormLigTicketDonne 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7080
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   9480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSMask.MaskEdBox LIGTicketDonneMaskEdBox 
      Bindings        =   "FormLigTicketDonne.frx":0000
      DataField       =   "TICKETDONNESPECIAL"
      BeginProperty DataFormat 
         Type            =   5
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   "*"
         FalseValue      =   " "
         NullValue       =   ""
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   7
      EndProperty
      DataMember      =   "InsTicketDonneCmd"
      DataSource      =   "DsrDatVente"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox LIGTicketDonneMaskEdBox 
      Bindings        =   "FormLigTicketDonne.frx":004D
      DataField       =   "TYPEMOUVEMENTCAR"
      DataMember      =   "SelTypeVenteCmd"
      DataSource      =   "DsrDatVente"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox LIGTicketDonneMaskEdBox 
      Bindings        =   "FormLigTicketDonne.frx":0093
      DataField       =   "TICKETDONNEARTICLEREFERENCE"
      DataMember      =   "InsTicketDonneCmd"
      DataSource      =   "DsrDatVente"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox LIGTicketDonneMaskEdBox 
      Bindings        =   "FormLigTicketDonne.frx":00E9
      DataField       =   "TICKETDONNEARTICLELIBELLE"
      DataMember      =   "InsTicketDonneCmd"
      DataSource      =   "DsrDatVente"
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox LIGTicketDonneMaskEdBox 
      Bindings        =   "FormLigTicketDonne.frx":013D
      DataField       =   "TICKETDONNETAILLELIBELLE"
      DataMember      =   "InsTicketDonneCmd"
      DataSource      =   "DsrDatVente"
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   4
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox LIGTicketDonneMaskEdBox 
      Bindings        =   "FormLigTicketDonne.frx":0190
      DataField       =   "TICKETDONNEQUANTITE"
      DataMember      =   "InsTicketDonneCmd"
      DataSource      =   "DsrDatVente"
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   5
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox LIGTicketDonneMaskEdBox 
      Bindings        =   "FormLigTicketDonne.frx":01DE
      DataField       =   "TYPEPRIXCAR"
      DataMember      =   "SelTypeVenteCmd"
      DataSource      =   "DsrDatVente"
      Height          =   255
      Index           =   7
      Left            =   6240
      TabIndex        =   6
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox LIGTicketDonneMaskEdBox 
      Bindings        =   "FormLigTicketDonne.frx":0220
      DataField       =   "TICKETDONNEPRIX"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "# ##0,00 "
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   2
      EndProperty
      DataMember      =   "InsTicketDonneCmd"
      DataSource      =   "DsrDatVente"
      Height          =   255
      Index           =   8
      Left            =   6840
      TabIndex        =   7
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox LIGTicketDonneMaskEdBox 
      Bindings        =   "FormLigTicketDonne.frx":0265
      DataField       =   "TICKETDONNEVENDEURCAR"
      DataMember      =   "InsTicketDonneCmd"
      DataSource      =   "DsrDatVente"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox LIGTicketDonneMaskEdBox 
      Bindings        =   "FormLigTicketDonne.frx":02B0
      DataField       =   "DEVISECAR"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "# ##0,00 ""F"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   0
      EndProperty
      DataMember      =   "SelDeviseCmd"
      DataSource      =   "DsrDatVente"
      Height          =   255
      Index           =   9
      Left            =   8040
      TabIndex        =   9
      Top             =   960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "FormLigTicketDonne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***                      (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - SAISSIE D'UNE LIGNE DE DONNER DU TICKET                  ***
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
    Option Explicit                                                             ' Je veut etre sur que toute met variable sont declarer
    
'******************************************************************************
'***    Declaration de variable internes                                    ***
'******************************************************************************

'******************************************************************************
'***    Variable pour memoriser le nombre de caractere d'une ligne          ***
'******************************************************************************
    Dim LIGTicketDonneCarMax As Integer
    
'******************************************************************************
'***    Declaration des procedures internes                                 ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:FormLigTicketDonneInitialise()                            ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormLigTicketDonneInitialise                              ***
'******************************************************************************
Private Sub FormLigTicketDonneInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:LIGTicketDonneMaskEdBoxInitialise()                       ***
'******************************************************************************
'***    FONCTION: Initialiser - Les masques des champs de saissies          ***
'***                          - La longueur des champs de saissies          ***
'***                          - Le nombre de caractere a afficher sur       ***
'***                            la ligne de saissie                         ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LIGTicketDonneMaskEdBoxInitialise                         ***
'******************************************************************************
Private Sub LIGTicketDonneMaskEdBoxInitialise()
    Dim I As Integer                                                            ' Le numero de zone courant de saissie
    ' Je vais pas mal utiliser les commandes du dataenvironement
    With DsrDatVente
        ' Je mes a zero le nombre de caractere d'une ligne de saissie
        LIGTicketDonneCarMax = 0
        ' Je fais defiler les zones de saissies
        For I = 0 To LIGTicketDonneMaskEdBox.Count - 1
            ' Je regarde sur quel zone je suis
            Select Case I
                Case ZoneVendeurCar, ZoneArticleReference, ZoneArticleLibelle, ZoneTaille, ZonePrix
                    ' Je suis sur une zone de saissie
                    ' Je charge le mask de la zone de saissie
                    LIGTicketDonneMaskEdBox(I).Mask = RessourceStringLoad(800 + I)
                    ' Je charge le format de la zone de saissie
                    LIGTicketDonneMaskEdBox(I).Format = RessourceStringLoad(900 + I)
                Case Else
                    ' Je suis sur une autre zone
            End Select
            ' Je defini la longueur du champ
            LIGTicketDonneMaskEdBox(I).MaxLength = .Recordsets(LIGTicketDonneMaskEdBox(I).DataMember).Fields(LIGTicketDonneMaskEdBox(I).DataField).DefinedSize
            ' Je calcule la longueur de la ligne
            LIGTicketDonneCarMax = LIGTicketDonneCarMax + LIGTicketDonneMaskEdBox(I).MaxLength
        ' Je ne sort que si j'ai fini de faire defiler les zones
        Next I
    ' Je n'ai plus besoin du dataenvironnement
    End With
    ' Fin
End Sub
    
'******************************************************************************
'***    PROCEDURE:FormLigTicketDonnePosition()                              ***
'******************************************************************************
'***    FONCTION: Positionner la form sur l'ecran                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormLigTicketDonnePosition                                ***
'******************************************************************************
Private Sub FormLigTicketDonnePosition()
    ' J'initialise la form sur l'ecran
    ControlInitialise FormLigTicketDonne, 0, Screen.Height * 8 / 15, Screen.Width * 16.5 / 21, Screen.Height * 1 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LIGTicketDonneMaskEdBoxPosition()                         ***
'******************************************************************************
'***    FONCTION: Positionner les champs de saissies sur la form            ***
'***              De la feuilles de saissie d'une ligne de donner du ticket ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LIGTicketDonneMaskEdBoxPosition                           ***
'******************************************************************************
Private Sub LIGTicketDonneMaskEdBoxPosition()
    Dim I As Integer                                                            ' Le numero de zone courant
    Dim MaskEdBoxPos As Single                                                  ' La position de la zone courant sur la form
    Dim MaskEdBoxLar As Single                                                  ' La largeur de la zone courant
    Dim MaskEdBoxLgr As Single                                                  ' La longueur de la zone courant
    ' Je fais defiler les zone pour les positionner sur la feuilles
    For I = 0 To LIGTicketDonneMaskEdBox.Count - 1
        ' Je calcule la longueur de la zone
        MaskEdBoxLgr = (FormLigTicketDonne.ScaleWidth * LIGTicketDonneMaskEdBox(I).MaxLength) / LIGTicketDonneCarMax
        ' Je calcule la largeur de la zone
        MaskEdBoxLar = FormLigTicketDonne.ScaleHeight
        ' Je positionne la zone saisie sur la feuille de saissie et redimensionne les fontes
        ControlInitialise LIGTicketDonneMaskEdBox(I), MaskEdBoxPos, 0, MaskEdBoxLgr, MaskEdBoxLar
        ' Je calcule la position de la zone suivantes
        MaskEdBoxPos = MaskEdBoxPos + MaskEdBoxLgr
    ' J'ai fini de faire defiler les zones
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    Evenement Internes                                                  ***
'******************************************************************************

'******************************************************************************
'***    EVENEMENT:Form_Load()                                               ***
'******************************************************************************
'***    FONCTION: Initialiser les variables des forms et des controles      ***
'***              Positionner la feuille sur l'ecran                        ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
Private Sub Form_Load()
    ' J'initialise la feuille
    FormLigTicketDonneInitialise
    ' J'initialise les zones de saissies
    LIGTicketDonneMaskEdBoxInitialise
    ' Je positionne la feuille sur l'ecran
    FormLigTicketDonnePosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:Form_Resize()                                             ***
'******************************************************************************
'***    FONCTION: Repositionnement de tous les controles                    ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
Private Sub Form_Resize()
    ' Je repositionne les zones de saissies
    LIGTicketDonneMaskEdBoxPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:LIGTicketDonneMaskEdBox_Change(Index As Integer)          ***
'******************************************************************************
'***    FONCTION: Redimensionner la taille des fontes                       ***
'******************************************************************************
'***    ENTREE:   index - le numero de la zone                              ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
Private Sub LIGTicketDonneMaskEdBox_Change(Index As Integer)
    ' Je recalcule la taille des font
    FontAjuste LIGTicketDonneMaskEdBox(Index)
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:LIGTicketDonneMaskEdBox_KeyDown(Index As Integer,         ***
'***              KeyCode As Integer, Shift As Integer)                     ***
'******************************************************************************
'***    FONCTION: Regarde si il y a un code douchette envoyer               ***
'******************************************************************************
'***    ENTREE:   Index   - Le numero de la zone textbox                    ***
'***              Keycode - Le code de la touche enfoncer                   ***
'***              Shift   - Si la touche shift a ete enfoncer               ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
Private Sub LIGTicketDonneMaskEdBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    ' Je regarde sur quel zone je suis
    Select Case Index
        Case ZoneArticleReference
            ' Je suis sur la zone reference
            ' Je regarde le code de la touche appuyer
            Select Case KeyCode
                Case 113
                    ' C'est le code du debut de la lecture de la douchette
                    ' Je defini le mask de la douchette
                    LIGTicketDonneMaskEdBox(ZoneArticleReference).Mask = RessourceStringLoad(813)
                    ' Je met le champ a vide
                    LIGTicketDonneMaskEdBox(ZoneArticleReference).Text = ""
                Case 114
                    ' C'est le code de fin de la lecture de la douchette
                    ' Je valide la douchette
                    DouchetteValidation
                    ' Je valide la zone reference
                    ZoneValidation
                    ' Je valide la zone taille
                    ZoneValidation
                    ' Je redefini le mask de la reference
                    LIGTicketDonneMaskEdBox(ZoneArticleReference).Mask = RessourceStringLoad(803)
                Case Else
                    ' C'est une autre touche appuyer
            End Select
        Case Else
            ' Je suis sur une autre zone
    End Select
    ' Fin
End Sub

