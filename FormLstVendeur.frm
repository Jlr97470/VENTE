VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormLstVendeur 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton LSTVendeurCDButton 
      Height          =   975
      Index           =   2
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton LSTVendeurCDButton 
      Height          =   975
      Index           =   1
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton LSTVendeurCDButton 
      Height          =   975
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid LSTVendeurDataGrid 
      Bindings        =   "FormLstVendeur.frx":0000
      Height          =   2655
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
      DataMember      =   "SelVendeurCmd"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "CLECAISSENUMERO"
         Caption         =   "CLECAISSENUMERO"
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
         DataField       =   "VENDEURCAR"
         Caption         =   "VENDEURCAR"
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
         DataField       =   "VENDEURLIBELLE"
         Caption         =   "VENDEURLIBELLE"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   1665,071
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FormLstVendeur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***                      (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - LISTE DES VENDEURS                                       ***
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
'***    PROCEDURE:formLSTVendeurInitialise()                                ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  formLSTVendeurInitialise                                  ***
'******************************************************************************
Private Sub formLSTVendeurInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:LSTVendeurDataGridInitialise()                            ***
'******************************************************************************
'***    FONCTION: Initialiser    - Le nom des colonnes                      ***
'***                             - Le nombre de caractere a afficher        ***
'***                               Sur une ligne                            ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTVendeurDataGridInitialise                              ***
'******************************************************************************
Private Sub LSTVendeurDataGridInitialise()
    Dim I As Integer                                                            ' Le numero de la colonne courante
    ' Je fait defiler toutes les colonnes de la datagrid
    For I = 0 To LSTVendeurDataGrid.Columns.Count - 1
        ' Je recupere le nom de la colonne
        LSTVendeurDataGrid.Columns(I).Caption = RessourceStringLoad(1170 + I)
    ' J'ai fini de faire defiler les colonnes
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LSTVendeurCDButtonInitialise()                            ***
'******************************************************************************
'***    FONCTION: Initialiser - Le nom du bouton                            ***
'***                          - L'icon du bouton                            ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTVendeurCDButtonInitialise                              ***
'******************************************************************************
Private Sub LSTVendeurCDButtonInitialise()
    Dim I As Integer                                                            ' Le numero du bouton courant
    ' Je fait defiler les boutons
    For I = 0 To LSTVendeurCDButton.Count - 1
        ' Je recupere l'icon du bouton
        LSTVendeurCDButton(I).Picture = RessourceIconLoad(1180 + I)
        ' Je recupere le texte du bouton
        LSTVendeurCDButton(I).Caption = RessourceStringLoad(1180 + I)
    ' Je ne sort que si j'ai fini de faire defiler les bouton
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:formLSTVendeurPosition()                                  ***
'******************************************************************************
'***    FONCTION: Positionner la form sur l'ecran                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  formLSTVendeurPosition                                    ***
'******************************************************************************
Private Sub formLSTVendeurPosition()
    ' J'initialise la form sur l'ecran
    ControlInitialise FormLstVendeur, Screen.Width * 5 / 21, Screen.Height * 2 / 15, Screen.Width * 11.5 / 21, Screen.Height * 6 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LSTVendeurDataGridPosition()                              ***
'******************************************************************************
'***    FONCTION: Repositionner la datagrid sur la form                     ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTVendeurDataGridPosition                                ***
'******************************************************************************
Private Sub LSTVendeurDataGridPosition()
    ' Je positionne la liste sur la form
    ControlInitialise LSTVendeurDataGrid, 0, 0, FormLstVendeur.ScaleWidth, FormLstVendeur.ScaleHeight * 2 / 3
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LSTVendeurCDButtonPosition()                              ***
'******************************************************************************
'***    FONCTION: Repositionner les boutons sur la form                     ***
'***              Ajuster la taille des fontes                              ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LSTVendeurCDButtonPosition                                ***
'******************************************************************************
Private Sub LSTVendeurCDButtonPosition()
    Dim I As Integer                                                            ' Le numero du bouton courant
    Dim CDButtonLef As Single                                                   ' La position gauche du bouton courant
    Dim CDButtonTop As Single                                                   ' La hauteur du bouton courant
    Dim CDButtonLgr As Single                                                   ' La longueur d'un bouton
    Dim CDButtonLar As Single                                                   ' La largeur d'un bouton
    ' Je fait defiler les boutons
    For I = 0 To LSTVendeurCDButton.Count - 1
        ' Je calcule la position gauche du boutons
        CDButtonLef = I * FormLstVendeur.ScaleWidth / 3
        ' Je calcule la hauteur du boutons
        CDButtonTop = FormLstVendeur.ScaleHeight * 2 / 3
        ' Je calcule la longueur du bouton
        CDButtonLgr = FormLstVendeur.ScaleWidth / 3
        ' Je calcule la largueur du bouton
        CDButtonLar = FormLstVendeur.ScaleHeight / 3
        ' Je positionne les boutons sur la form et ajuste les fontes
        ControlInitialise LSTVendeurCDButton(I), CDButtonLef, CDButtonTop, CDButtonLgr, CDButtonLar
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
    ' J'initialise la feuille
    formLSTVendeurInitialise
    ' J'initialise la liste des vendeurs
    LSTVendeurDataGridInitialise
    ' J'initialise les boutons de la liste des vendeurs
    LSTVendeurCDButtonInitialise
    ' Je positionne la feuille a l'ecran
    formLSTVendeurPosition
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
    ' Je repositionne la liste des vendeurs
    LSTVendeurDataGridPosition
    ' Je repositionne les boutons de la liste des vendeurs
    LSTVendeurCDButtonPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  LSTVendeurCDButton_Click(Index As Integer)              ***
'******************************************************************************
'***    FONCTION:   - Repositionner les lignes de la liste des vendeurs     ***
'******************************************************************************
'***    ENTREE:     Index - numero du bouton appuyer                        ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub LSTVendeurCDButton_Click(Index As Integer)
    ' Je regarde quel boutton a ete appuyer
    Select Case Index
        Case 0
            ' Je fais defiler la liste vers le haut
            LSTVendeurDataGrid.Scroll 0, -LSTVendeurDataGrid.VisibleRows
        Case 1
            ' Je fais defiler la liste vers le bas
            LSTVendeurDataGrid.Scroll 0, LSTVendeurDataGrid.VisibleRows - 1
        Case 2
    End Select
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  LSTVendeurDataGrid_RowColChange(LastRow As Variant,     ***
'***                ByVal LastCol As Integer)                               ***
'******************************************************************************
'***    FONCTION:   - Validation du choix du vendeur                        ***
'***                  Dechargement de la liste des vendeurs                 ***
'******************************************************************************
'***    ENTREE:     LastRow                                                 ***
'***                LastCol                                                 ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub LSTVendeurDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ' Je decharge la feuille
    Unload Me
    ' Je valide le choix du vendeurs
    ZoneVendeurListeValidation
    ' Je valide la zone courante
    ZoneValidation
    ' Fin
End Sub
