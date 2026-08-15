VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form LSTArticleForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton LSTArticleCDButton 
      Height          =   1815
      Index           =   2
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton LSTArticleCDButton 
      Height          =   1815
      Index           =   1
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton LSTArticleCDButton 
      Height          =   1815
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid LSTArticleDataGrid 
      Bindings        =   "LSTArticleForm.frx":0000
      Height          =   3135
      Left            =   0
      TabIndex        =   3
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
      DataMember      =   "SelArticleCmd"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "ARTICLEREFERENCE"
         Caption         =   "ARTICLEREFERENCE"
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
         DataField       =   "ARTICLELIBELLE"
         Caption         =   "ARTICLELIBELLE"
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
            ColumnWidth     =   1725,165
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1395,213
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "LSTArticleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. 1.0 VB )                  (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - LISTE DES ARTICLES                                       ***
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
'***    PROCEDURE:LSTArticleFormInitialise()                                ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTArticleFormInitialise                                  ***
'******************************************************************************
Private Sub LSTArticleFormInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:LSTArticleDataGridInitialise()                            ***
'******************************************************************************
'***    FONCTION: Initialiser    - Le nom des colonnes                      ***
'***                             - Le nombre de caractere a afficher        ***
'***                                sur une ligne                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTArticleDataGridInitialise                              ***
'******************************************************************************
Private Sub LSTArticleDataGridInitialise()
    Dim I As Integer                                                            ' Le numero de la colonne courante
    ' Je fait defiler toutes les colonnes de la datagrid
    For I = 0 To LSTArticleDataGrid.Columns.Count - 1
        ' Je recupere le nom de la colonne
        LSTArticleDataGrid.Columns(I).Caption = RessourceStringLoad(1150 + I)
    ' J'ai fini de faire defiler les colonnes
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LSTArticleCDButtonInitialise()                            ***
'******************************************************************************
'***    FONCTION: Initialiser - Le nom du bouton                            ***
'***                          - L'icon du bouton                            ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTArticleCDButtonInitialise                              ***
'******************************************************************************
Private Sub LSTArticleCDButtonInitialise()
    Dim I As Integer                                                            ' Le numero du bouton courant
    ' Je fait defiler les boutons
    For I = 0 To LSTArticleCDButton.Count - 1
        ' Je recupere l'icon du bouton
        LSTArticleCDButton(I).Picture = RessourceIconLoad(1180 + I)
        ' Je recupere le texte du bouton
        LSTArticleCDButton(I).Caption = RessourceStringLoad(1180 + I)
    ' Je ne sort que si j'ai fini de faire defiler les bouton
    Next I
    ' fin
End Sub

'******************************************************************************
'***    PROCEDURE:LSTArticleFormPosition()                                  ***
'******************************************************************************
'***    FONCTION: Positionner la form sur l'ecran                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTArticleFormPosition                                    ***
'******************************************************************************
Private Sub LSTArticleFormPosition()
    ' J'initialise la form sur la feuille
    ControlInitialise LSTArticleForm, Screen.Width * 5 / 21, Screen.Height * 2 / 15, Screen.Width * 11.5 / 21, Screen.Height * 6 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LSTArticleDataGridPosition()                              ***
'******************************************************************************
'***    FONCTION: Positionner la datagrid sur la form                       ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTArticleDataGridPosition                                ***
'******************************************************************************
Private Sub LSTArticleDataGridPosition()
    ' Je positionne la liste sur la form
    ControlInitialise LSTArticleDataGrid, 0, 0, LSTArticleForm.ScaleWidth, LSTArticleForm.ScaleHeight * 2 / 3
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LSTArticleCDButtonPosition()                              ***
'******************************************************************************
'***    FONCTION: Positionner les boutons sur la form                       ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTArticleCDButtonPos                                     ***
'******************************************************************************
Private Sub LSTArticleCDButtonPosition()
    Dim I As Integer                                                            ' Le numero du bouton courant
    Dim CDButtonLef As Single                                                   ' La position gauche du bouton courant
    Dim CDButtonTop As Single                                                   ' La hauteur du bouton courant
    Dim CDButtonLgr As Single                                                   ' La longueur d'un bouton
    Dim CDButtonLar As Single                                                   ' La largeur d'un bouton
    ' Je fait defiler les boutons
    For I = 0 To LSTArticleCDButton.Count - 1
        ' Je calcule la position gauche du boutons
        CDButtonLef = I * LSTArticleForm.ScaleWidth / 3
        ' Je calcule la hauteur du boutons
        CDButtonTop = LSTArticleForm.ScaleHeight * 2 / 3
        ' Je calcule la longueur du bouton
        CDButtonLgr = LSTArticleForm.ScaleWidth / 3
        ' Je calcule la largueur du bouton
        CDButtonLar = LSTArticleForm.ScaleHeight / 3
        ' Je positionne les boutons sur la form et ajuste les fontes
        ControlInitialise LSTArticleCDButton(I), CDButtonLef, CDButtonTop, CDButtonLgr, CDButtonLar
    ' Je ne sort que si j'ai fini de faire defiler les bouton
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
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
Private Sub Form_Load()
    ' J'initialise la form
    LSTArticleFormInitialise
    ' J'initialise la liste des Articles
    LSTArticleDataGridInitialise
    ' J'initialise les boutons de la liste des Articles
    LSTArticleCDButtonInitialise
    ' Je positionne la form a l'ecran
    LSTArticleFormPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:Form_Resize()                                             ***
'******************************************************************************
'***    FONCTION: Repositionnement des controles                            ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
Private Sub Form_Resize()
    ' Je repositionne la liste des Articles sur la form
    LSTArticleDataGridPosition
    ' Je repositionne les boutons sur la forms
    LSTArticleCDButtonPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:LSTArticleCDButton_Click(Index As Integer)                ***
'******************************************************************************
'***    FONCTION: - Repositionner les lignes de la liste des Article        ***
'******************************************************************************
'***    ENTREE:   Index - numero du bouton appuyer                          ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
Private Sub LSTArticleCDButton_Click(Index As Integer)
    ' Je regarde quel boutton a ete appuyer
    Select Case Index
        Case 0
            ' Je fais defiler la liste vers le haut
             LSTArticleDataGrid.Scroll 0, -LSTArticleDataGrid.VisibleRows
        Case 1
            ' Je fais defiler la liste vers le bas
             LSTArticleDataGrid.Scroll 0, LSTArticleDataGrid.VisibleRows - 1
        Case 2
    End Select
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  LSTArticleDataGrid_RowColChange(LastRow As Variant,     ***
'***                ByVal LastCol As Integer)                               ***
'******************************************************************************
'***    FONCTION: - Validation du choix du taille                           ***
'***                Dechargement de la liste des taille                     ***
'******************************************************************************
'***    ENTREE:     LastRow                                                 ***
'***                LastCol                                                 ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub LSTArticleDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ' Je decharge la feuille
    Unload Me
    ' Je valide le choix de la Article
    ZoneArticleReferenceListeValidation
    ' Je valide la zone courante
    ZoneValidation
End Sub
