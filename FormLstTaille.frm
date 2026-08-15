VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormLstTaille 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton LSTTailleCDButton 
      Height          =   2535
      Index           =   2
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton LSTTailleCDButton 
      Height          =   2535
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton LSTTailleCDButton 
      Height          =   2535
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid LSTTailleDataGrid 
      Bindings        =   "FormLstTaille.frx":0000
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "SelTailleCmd"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "TAILLELIBELLE"
         Caption         =   "TAILLELIBELLE"
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
            ColumnWidth     =   1244,976
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FormLstTaille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***                      (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - LISTE DES TAILLES                                        ***
'******************************************************************************
'***    FONCTION:- INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES      ***
'***             - RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU         ***
'***               MODULES PRINCIPALES DE LA SAISSIE D'UNE LIGNE DE DONNER  ***
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
    Option Explicit                                                             ' Je veut etre sur que toute met variable sont declarer

'******************************************************************************
'***    Declaration de variable internes                                    ***
'******************************************************************************

'******************************************************************************
'***    Declaration des procedures internes                                 ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:LSTailleFormInitialise()                                  ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTailleFormInitialise                                    ***
'******************************************************************************
Private Sub formLSTtailleInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE: LSTTailleDataGridInitialise()                            ***
'******************************************************************************
'***    FONCTION:  Initialiser    - Le nom des colonnes                     ***
'***                              - Le nombre de caractere a afficher       ***
'***                                Sur une ligne                           ***
'******************************************************************************
'***    ENTREE:    Néant                                                    ***
'***                                                                        ***
'***    SORTIE:    Néant                                                    ***
'******************************************************************************
'***    EXEMPLE:   LSTTailleDataGridInitialise                              ***
'******************************************************************************
Private Sub LSTTailleDataGridInitialise()
    Dim I As Integer                                                            ' Le numero de la colonne courante
    ' Je fait defiler toutes les colonnes de la datagrid
    For I = 0 To LSTTailleDataGrid.Columns.Count - 1
        ' Je recupere le nom de la colonne
        LSTTailleDataGrid.Columns(I).Caption = RessourceStringLoad(1160 + I)
    ' J'ai fini de faire defiler les colonnes
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LSTTailleCDButtonInitialise()                             ***
'******************************************************************************
'***    FONCTION: Initialiser - Le nom du bouton                            ***
'***                          - L'icon du bouton                            ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE: LSTTailleCDButtonInitialise                                ***
'******************************************************************************
Private Sub LSTTailleCDButtonInitialise()
    Dim I As Integer                                                            ' Le numero du bouton courant
    ' Je fait defiler les boutons
    For I = 0 To LSTTailleCDButton.Count - 1
        ' Je recupere l'icon du bouton
        LSTTailleCDButton(I).Picture = RessourceIconLoad(1180 + I)
        ' Je recupere le texte du bouton
        LSTTailleCDButton(I).Caption = RessourceStringLoad(1180 + I)
    ' Je ne sort que si j'ai fini de faire defiler les bouton
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:formLSTtaillePosition()                                   ***
'******************************************************************************
'***    FONCTION: Positionner la form sur l'ecran                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  formLSTtaillePosition                                     ***
'******************************************************************************
Private Sub formLSTtaillePosition()
    ' J'initialise la form sur l'ecran
    ControlInitialise FormLstTaille, Screen.Width * 5 / 21, Screen.Height * 2 / 15, Screen.Width * 11.5 / 21, Screen.Height * 6 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LSTTailleDataGridPosition()                               ***
'******************************************************************************
'***    FONCTION: Repositionner la datagrid sur la form                     ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTTailleDataGridPosition                                 ***
'******************************************************************************
Private Sub LSTTailleDataGridPosition()
    ' Je positionne la liste sur la form
    ControlInitialise LSTTailleDataGrid, 0, 0, FormLstTaille.ScaleWidth, FormLstTaille.ScaleHeight * 2 / 3
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LSTTailleCDButtonPosition()                               ***
'******************************************************************************
'***    FONCTION: Repositionner les boutons sur la form                     ***
'***              Ajuster la taille des fontes des boutons                  ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTTailleCDButtonPosition                                 ***
'******************************************************************************
Private Sub LSTTailleCDButtonPosition()
    Dim I As Integer                                                            ' Le numero du bouton courant
    Dim CDButtonLef As Single                                                   ' La position gauche du bouton courant
    Dim CDButtonTop As Single                                                   ' La hauteur du bouton courant
    Dim CDButtonLgr As Single                                                   ' La longueur d'un bouton
    Dim CDButtonLar As Single                                                   ' La largeur d'un bouton
    ' Je fait defiler les boutons
    For I = 0 To LSTTailleCDButton.Count - 1
        ' Je calcule la position gauche du boutons
        CDButtonLef = I * FormLstTaille.ScaleWidth / 3
        ' Je calcule la hauteur du boutons
        CDButtonTop = FormLstTaille.ScaleHeight * 2 / 3
        ' Je calcule la longueur du bouton
        CDButtonLgr = FormLstTaille.ScaleWidth / 3
        ' Je calcule la largueur du bouton
        CDButtonLar = FormLstTaille.ScaleHeight / 3
        ' Je positionne les boutons sur la form et ajuste les fontes
        ControlInitialise LSTTailleCDButton(I), CDButtonLef, CDButtonTop, CDButtonLgr, CDButtonLar
    ' Je ne sort que si j'ai fini de faire defiler les bouton
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
    formLSTtailleInitialise
    ' J'initialise la liste des tailles
    LSTTailleDataGridInitialise
    ' J'initialise les boutons de la liste des tailles
    LSTTailleCDButtonInitialise
    ' Je positionne la feuille sur l'ecran
    formLSTtaillePosition
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
Private Sub Form_Resize()
    ' Je repositionne la liste des tailles
    LSTTailleDataGridPosition
    ' Je repositionne les boutons de la lsite des tailles
    LSTTailleCDButtonPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  LSTTailleCDButton_Click(Index As Integer)               ***
'******************************************************************************
'***    FONCTION:   - Repositionner les lignes de la liste des tailles      ***
'******************************************************************************
'***    ENTREE:     Index - numero du bouton appuyer                        ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub LSTTailleCDButton_Click(Index As Integer)
    ' Je regarde quel boutton a ete appuyer
    Select Case Index
        Case 0
            ' Je fais defiler la liste vers le haut
            LSTTailleDataGrid.Scroll 0, -LSTTailleDataGrid.VisibleRows
        Case 1
            ' Je fais defiler la liste vers le bas
            LSTTailleDataGrid.Scroll 0, LSTTailleDataGrid.VisibleRows - 1
        Case 2
    End Select
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  LSTTailleDataGrid_RowColChange(LastRow As Variant,      ***
'***                ByVal LastCol As Integer)                               ***
'******************************************************************************
'***    FONCTION: - Validation du choix du taille                           ***
'***                Dechargement de la liste des taille                     ***
'******************************************************************************
'***    ENTREE:     LastRow                                                 ***
'***                LastCol                                                 ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub LSTTailleDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ' Je decharge la feuille
    Unload Me
    ' Je valide le choix de la taille
    ZoneTailleListeValidation
    ' Je valide la zone courante
    ZoneValidation
    ' Fin
End Sub
