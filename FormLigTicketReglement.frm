VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormLigTicketReglement 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   9480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox LIGTicketReglementMaskEdBox 
      Bindings        =   "FormLigTicketReglement.frx":0000
      DataField       =   "TYPEREGLEMENTCAR"
      DataMember      =   "SelTypeReglementCmd"
      DataSource      =   "DsrDatVente"
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox LIGTicketReglementMaskEdBox 
      Bindings        =   "FormLigTicketReglement.frx":004F
      DataField       =   "TYPEREGLEMENTLIBELLE"
      DataMember      =   "SelTypeReglementCmd"
      DataSource      =   "DsrDatVente"
      Height          =   735
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox LIGTicketReglementMaskEdBox 
      Bindings        =   "FormLigTicketReglement.frx":006D
      DataField       =   "TICKETREGLEMENTPRIX"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   1
      EndProperty
      DataMember      =   "InsTicketReglementCmd"
      DataSource      =   "DsrDatVente"
      Height          =   735
      Index           =   2
      Left            =   4680
      TabIndex        =   2
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox LIGTicketReglementMaskEdBox 
      Bindings        =   "FormLigTicketReglement.frx":008D
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
      Height          =   735
      Index           =   3
      Left            =   6960
      TabIndex        =   3
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   0
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "FormLigTicketReglement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***                      (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - SAISSIE D'UNE LIGNE DE REGLEMENT DU TICKET               ***
'******************************************************************************
'***    FONCTION:- INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES      ***
'***             - RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU         ***
'***               MODULES PRINCIPALES DE LA GESTION DES REGLEMENT DU TICKET***
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
    Dim LIGTicketReglementCarMax As Integer
    
'******************************************************************************
'***    Declaration des procedures internes                                 ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:FormLigTicketReglementInitialise()                        ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormLigTicketReglementInitialise                          ***
'******************************************************************************
Private Sub FormLigTicketReglementInitialise()
End Sub
    
'******************************************************************************
'***    PROCEDURE:LIGTicketReglementMaskEdBoxInitialise ()                  ***
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
'***    EXEMPLE:  LIGTicketReglementMaskEdBoxInitialise                     ***
'******************************************************************************
Private Sub LIGTicketReglementMaskEdBoxInitialise()
    Dim I As Integer                                                            ' Le numero de zone courant de saissie
    ' Je vais pas mal utiliser les commandes du dataenvironement
    With DsrDatVente
        ' Je mes a zero le nombre de caractere d'une ligne de saissie
        LIGTicketReglementCarMax = 0
        ' Je fais defiler les zones de saissies
        For I = 0 To LIGTicketReglementMaskEdBox.Count - 1
            ' Je regarde sur quel zone je suis
            Select Case I
                Case ZoneValeur
                    ' Je suis sur la zone valeur du reglement
                    ' Je charge le mask de la zone de saissie
                    LIGTicketReglementMaskEdBox(I).Mask = RessourceStringLoad(820 + I)
                    ' Je defini le format du champ
                    LIGTicketReglementMaskEdBox(I).Format = RessourceStringLoad(920 + I)
                Case Else
                    ' Je suis sur une autre zone
            End Select
            ' Je defini la longueur du champ
            LIGTicketReglementMaskEdBox(I).MaxLength = .Recordsets(LIGTicketReglementMaskEdBox(I).DataMember).Fields(LIGTicketReglementMaskEdBox(I).DataField).DefinedSize
            ' Je calcule la longueur de la ligne
            LIGTicketReglementCarMax = LIGTicketReglementCarMax + LIGTicketReglementMaskEdBox(I).MaxLength
        ' Je ne sort que si j'ai fini de faire defiler les zones
        Next I
    ' Je n'ai plus besoin du dataenvironnement
    End With
    ' Fin
End Sub
    
'******************************************************************************
'***    PROCEDURE:FormLigTicketReglementPosition()                          ***
'******************************************************************************
'***    FONCTION: Positionner la form sur l'ecran                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormLigTicketReglementPosition                            ***
'******************************************************************************
Private Sub FormLigTicketReglementPosition()
    ' J'initialise la feuille a l'ecran
    ControlInitialise FormLigTicketReglement, 0, Screen.Height * 8 / 15, Screen.Width * 16.5 / 21, Screen.Height * 1 / 15
    ' Fin
End Sub
    
'******************************************************************************
'***    PROCEDURE:LIGTicketReglementMaskEdBoxPosition()                     ***
'******************************************************************************
'***    FONCTION: Positionner les champs de saissies sur la form            ***
'***              De la feuilles de saisie d'une ligne de reglement du      ***
'***              Ticket                                                    ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LIGTicketReglementMaskEdBoxPosition                       ***
'******************************************************************************
Private Sub LIGTicketReglementMaskEdBoxPosition()
    Dim I As Integer                                                            ' Le numero de zone courant
    Dim MaskEdBoxPos As Single                                                  ' La position de la zone courant sur la form
    Dim MaskEdBoxLar As Single                                                  ' La largeur de la zone courant
    Dim MaskEdBoxLgr As Single                                                  ' La longueur de la zone courant
    ' Je fais defiler les zone pour les positionner sur la feuilles
    For I = 0 To LIGTicketReglementMaskEdBox.Count - 1
        ' Je calcule la longueur de la zone
        MaskEdBoxLgr = (FormLigTicketReglement.ScaleWidth * LIGTicketReglementMaskEdBox(I).MaxLength) / LIGTicketReglementCarMax
        ' Le calcule la largeur de la zone
        MaskEdBoxLar = FormLigTicketReglement.ScaleHeight
        ' Je positionne la zone saissie sur la feuille de saissie et redimensionne les fontes
        ControlInitialise LIGTicketReglementMaskEdBox(I), MaskEdBoxPos, 0, MaskEdBoxLgr, MaskEdBoxLar
        ' Je calcule la position de la zone suivantes
        MaskEdBoxPos = MaskEdBoxPos + MaskEdBoxLgr
    ' J'ai fini de faire defiler les zones
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    Evenement internes                                                  ***
'******************************************************************************

'******************************************************************************
'***    EVENEMENT:Form_Load()                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    FONCTION: Initialiser les variables des forms et des controles      ***
'***              Positionner la feuille sur l'ecran                        ***
'******************************************************************************
Private Sub Form_Load()
    ' J'initialise la feuille
    FormLigTicketReglementInitialise
    ' J'initialise les zones de saissies
    LIGTicketReglementMaskEdBoxInitialise
    ' Je positionne la feuille sur l'ecran
    FormLigTicketReglementPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:Form_Resize()                                             ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    FONCTION: Repositionnement de tous les controles                    ***
'******************************************************************************
Private Sub Form_Resize()
    ' Je repositionne les zones de saissies
    LIGTicketReglementMaskEdBoxPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:LIGTicketReglementMaskEdBox_Change(Index As Integer)      ***
'******************************************************************************
'***    ENTREE:   Index - le numero de la zone                              ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    FONCTION: Redimensionner la taille des fontes                       ***
'******************************************************************************
Private Sub LIGTicketReglementMaskEdBox_Change(Index As Integer)
    ' Je recalcule la taille des font
    FontAjuste LIGTicketReglementMaskEdBox(Index)
    ' Fin
End Sub

