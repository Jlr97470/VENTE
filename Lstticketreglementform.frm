VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form LSTTicketReglementForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid LSTTicketReglementDataGrid 
      Bindings        =   "LSTTicketReglementForm.frx":0000
      Height          =   5775
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   10186
      _Version        =   393216
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
      DataMember      =   "SelTicketReglementCmd"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "TYPEREGLEMENTCAR"
         Caption         =   "TYPEREGLEMENTCAR"
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
         DataField       =   "TYPEREGLEMENTLIBELLE"
         Caption         =   "TYPEREGLEMENTLIBELLE"
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
      BeginProperty Column02 
         DataField       =   "TICKETREGLEMENTPRIXDEVISE"
         Caption         =   "TICKETREGLEMENTPRIXDEVISE"
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   540,284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   989,858
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "LSTTicketReglementForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. 1.0 VB )                  (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - LISTE DES REGLEMENTS DU TICKET                           ***
'******************************************************************************
'***    FONCTION:- INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES      ***
'***             - RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU         ***
'***               MODULES PRINCIPALES DE LA SAISSIE D'UNE LIGNE DE         ***
'***               REGLEMENT DU TICKET                                      ***
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
'***    PROCEDURE:LSTTicketReglementFormInitialise()                        ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTTicketReglementFormInitialise                          ***
'******************************************************************************
Private Sub LSTTicketReglementFormInitialise()
End Sub
    
'******************************************************************************
'***    PROCEDURE:LSTTicketReglementDataGridInitialise()                    ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTTicketReglementDataGridInitialise                      ***
'******************************************************************************
Private Sub LSTTicketReglementDataGridInitialise()
    Dim I As Integer                                                            ' Le numero de la colonne courante
    ' Je fait defiler toutes les colonnes de la datagrid
    For I = 0 To LSTTicketReglementDataGrid.Columns.Count - 1
        ' Je recupere le nom de la colonne
        LSTTicketReglementDataGrid.Columns(I).Caption = RessourceStringLoad(740 + I)
    ' J'ai fini de faire defiler les colonnes
    Next I
    ' Fin
End Sub
    
'******************************************************************************
'***    PROCEDURE:LSTTicketReglementFormPosition()                          ***
'******************************************************************************
'***    FONCTION: Pour positionner la feuille a l'ecran                     ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTTicketReglementFormPosition                            ***
'******************************************************************************
Private Sub LSTTicketReglementFormPosition()
    ' J'initialise la forme sur l'ecran
    ControlInitialise LSTTicketReglementForm, Screen.Width * 5 / 21, Screen.Height * 2 / 15, Screen.Width * 11.5 / 21, Screen.Height * 6 / 15
    ' Fin
End Sub
    
'******************************************************************************
'***    PROCEDURE:LSTTicketReglementDataGridPosition()                      ***
'******************************************************************************
'***    FONCTION: Pour positionner le datagrid sur la form                  ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTTicketReglementDataGridPosition                        ***
'******************************************************************************
Private Sub LSTTicketReglementDataGridPosition()
    ' Je positionne la liste sur la form
    ControlInitialise LSTTicketReglementDataGrid, 0, 0, LSTTicketReglementForm.ScaleWidth, LSTTicketReglementForm.ScaleHeight
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
    LSTTicketReglementFormInititialise
    ' J'intialise la liste des donnees du ticket
    LSTTicketReglementDataGridInitialise
    ' Je positionne la feuille sur l'ecran
    LSTTicketReglementFormPosition
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
    LSTTicketReglementDataGridPosition
    ' Fin
End Sub


