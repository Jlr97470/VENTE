VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form LSTTicketDonneForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7080
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
   ScaleHeight     =   7080
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid LSTTicketDonneDataGrid 
      Bindings        =   "LSTTicketDonneForm.frx":0000
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   10821
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "SelTicketDonneCmd"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "TICKETDONNEVENDEURCAR"
         Caption         =   "TICKETDONNEVENDEURCAR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "TICKETDONNESPECIAL"
         Caption         =   "TICKETDONNESPECIAL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "*"
            FalseValue      =   ""
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "TYPEMOUVEMENTCARLIBELLE"
         Caption         =   "TYPEMOUVEMENTCARLIBELLE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "TICKETDONNEARTICLEREFERENCE"
         Caption         =   "TICKETDONNEARTICLEREFERENCE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "TICKETDONNEARTICLELIBELLE"
         Caption         =   "TICKETDONNEARTICLELIBELLE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "TICKETDONNETAILLELIBELLE"
         Caption         =   "TICKETDONNETAILLELIBELLE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "TICKETDONNEQUANTITESIGNE"
         Caption         =   "TICKETDONNEQUANTITESIGNE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "TYPEPRIXCAR"
         Caption         =   "TYPEPRIXCAR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "TICKETDONNEPRIXDEVISE"
         Caption         =   "TICKETDONNEPRIXDEVISE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "# ##0,00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "DEVISECAR"
         Caption         =   "DEVISECAR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "TICKETDONNETOTALDEVISE"
         Caption         =   "TICKETDONNETOTALDEVISE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "# ##0,00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "DEVISECAR"
         Caption         =   "DEVISECAR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   0
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            Object.Visible         =   -1  'True
            ColumnWidth     =   2355,024
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1890,142
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1874,835
         EndProperty
         BeginProperty Column03 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2880
         EndProperty
         BeginProperty Column04 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2534,74
         EndProperty
         BeginProperty Column05 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2399,811
         EndProperty
         BeginProperty Column06 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2550,047
         EndProperty
         BeginProperty Column07 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column08 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2190,047
         EndProperty
         BeginProperty Column09 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column10 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2340,284
         EndProperty
         BeginProperty Column11 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   989,858
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "LSTTicketDonneForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. 1.0 VB )                  (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - LISTE DES DONNEES DU TICKET                              ***
'******************************************************************************
'***    FONCTION:- INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES      ***
'***             - RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU         ***
'***             - MODULES PRINCIPALES DE LA SAISSIE D'UNE LIGNE DE DONNER  ***
'***               DU TICKET                                                ***
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
'***    PROCEDURE:LSTTicketDonneFormInitialise()                            ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTTicketDonneFormInitialise                              ***
'******************************************************************************
Private Sub LSTTicketDonneFormInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:LSTTicketDonneDataGridInitialise()                        ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTTicketDonneDataGridInitialise                          ***
'******************************************************************************
Private Sub LSTTicketDonneDataGridInitialise()
    Dim I As Integer                                                            ' Le numero de la colonne courante
    ' Je fait defiler toutes les colonnes de la datagrid
    For I = 0 To LSTTicketDonneDataGrid.Columns.Count - 1
        ' Je recupere le nom de la colonne
        LSTTicketDonneDataGrid.Columns(I).Caption = RessourceStringLoad(700 + I)
    ' J'ai fini de faire defiler les colonnes
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LSTTicketDonneFormPosition()                              ***
'******************************************************************************
'***    FONCTION: Pour positionner la feuille a l'ecran                     ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTTicketDonneFormPosition                                ***
'******************************************************************************
Private Sub LSTTicketDonneFormPosition()
    ' J'initialise la forme sur l'ecran
    ControlInitialise LSTTicketDonneForm, 0, Screen.Height * 2 / 15, Screen.Width * 16.5 / 21, Screen.Height * 6 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LSTTicketDonneDataGridPosition()                          ***
'******************************************************************************
'***    FONCTION: Pour positionner le datagrid sur la form                  ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTTicketDonneDataGridPosition                            ***
'******************************************************************************
Private Sub LSTTicketDonneDataGridPosition()
    ' Je positionne la liste sur la form
    ControlInitialise LSTTicketDonneDataGrid, 0, 0, LSTTicketDonneForm.ScaleWidth, LSTTicketDonneForm.ScaleHeight
    ' Fin
End Sub

'******************************************************************************
'***    Evenement internes                                                  ***
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
    ' J'initialise la form
    LSTTicketDonneFormInitialise
    ' J'intialise la liste des donnees du ticket
    LSTTicketDonneDataGridInitialise
    ' Je positionne la feuille sur l'ecran
    LSTTicketDonneFormPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:Form_Resize()                                             ***
'******************************************************************************
'***    FONCTION: Repositionner les controles sur la form                   ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
Private Sub Form_Resize()
    ' Je repositionne la liste des donnes du ticket
    LSTTicketDonneDataGridPosition
    ' Fin
End Sub

