VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FormPanArticle 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox PANArticleFileListBox 
      Height          =   2235
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Timer PANArticleTimer 
      Interval        =   5000
      Left            =   5040
      Top             =   4440
   End
   Begin WMPLibCtl.WindowsMediaPlayer PANArticleMediaPlayer 
      Height          =   735
      Left            =   4080
      TabIndex        =   1
      Top             =   1680
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   1296
   End
   Begin VB.Image PANArticleImage 
      Height          =   1815
      Left            =   240
      Stretch         =   -1  'True
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "FormPanArticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***                      (31/10/2000) ***
'******************************************************************************
'***    FEUILLE: - AFFICHAGE DES MEDIAS DE L'ARTICLE                        ***
'******************************************************************************
'***    FONCTION:- INITIALISER ET POSITIONNER LES DIFFERENTS CONTROLES      ***
'***             - RECUPERATION DES EVENEMENTS POUR LES ENVOYERS AU         ***
'***             - MODULES PRINCIPALES DE LA GESTION DES MEDIAS             ***
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
    Option Explicit                                                            ' Je veut etre sur que toute met variable sont declarer
  
'******************************************************************************
'***    Declaration de variable internes                                    ***
'******************************************************************************

'******************************************************************************
'***    Declaration des procedures internes                                 ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:FormPanArticleInitialise()                                ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormPanArticleInitialise                                  ***
'******************************************************************************
Private Sub FormPanArticleInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PANArticleFileListBoxInitialise()                         ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANArticleFileListBoxInitialise                           ***
'******************************************************************************
Private Sub PANArticleFileListBoxInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PANArticleImageInitialise()                               ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANArticleImageInitialise                                 ***
'******************************************************************************
Private Sub PANArticleImageInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PANArticleMediaPlayerInitialise()                         ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANArticleMediaPlayerInitialise                           ***
'******************************************************************************
Private Sub PANArticleMediaPlayerInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:PANArticleTimerInitialise()                               ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANArticleTimerInitialise                                 ***
'******************************************************************************
Private Sub PANArticleTimerInitialise()
End Sub

'******************************************************************************
'***    PROCEDURE:FormPanArticlePosition()                                  ***
'******************************************************************************
'***    FONCTION: Positionner la form sur l'ecran                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormPanArticlePosition                                    ***
'******************************************************************************
Private Sub FormPanArticlePosition()
    ' Je positionne ma form a l'ecran
    ControlInitialise FormPanArticle, Screen.Width * 16.5 / 21, 0, Screen.Width * 4.5 / 21, Screen.Height * 4 / 15
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANArticleFileListBoxPosition()                           ***
'******************************************************************************
'***    FONCTION: Positionner l'image sur la form sur l'ecran               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANArticleFileListBoxPosition                             ***
'******************************************************************************
Private Sub PANArticleFileListBoxPosition()
End Sub

'******************************************************************************
'***    PROCEDURE:PANArticleImagePosition()                                 ***
'******************************************************************************
'***    FONCTION: Positionner l'image sur la form sur l'ecran               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANArticleImagePosition                                   ***
'******************************************************************************
Private Sub PANArticleImagePosition()
    ' Je positionne l'image sur la feuille
    ControlInitialise PANArticleImage, 0, 0, FormPanArticle.ScaleWidth, FormPanArticle.ScaleHeight
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANArticleMediaPlayerPosition()                           ***
'******************************************************************************
'***    FONCTION: Positionner le mediaplayer sur la form sur l'ecran        ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANArticleMediaPlayerPosition                             ***
'******************************************************************************
Private Sub PANArticleMediaPlayerPosition()
    ' Je positionne l'MediaPlayer sur la feuille
    ControlInitialise PANArticleMediaPlayer, 0, 0, FormPanArticle.ScaleWidth, FormPanArticle.ScaleHeight
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANArticleTimerPosition()                                 ***
'******************************************************************************
'***    FONCTION: Positionner le mediaplayer sur la form sur l'ecran        ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANArticleTimerPosition                                   ***
'******************************************************************************
Private Sub PANArticleTimerPosition()
End Sub

'******************************************************************************
'***    PROCEDURE:FormPanArticle_Min()                                      ***
'******************************************************************************
'***    FONCTION: RePositionner la form sur l'ecran a sa position precedente***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormPanArticle_Min                                        ***
'******************************************************************************
Private Sub FormPanArticle_Min()
    ' Je positionne ma form a l'ecran
    FormPanArticle.Move Screen.Width * 16.5 / 21, 0, Screen.Width * 4.5 / 21, Screen.Height * 4 / 15
    ' Fin
    Zone.SetFocus
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:FormPanArticle_Max()                                      ***
'******************************************************************************
'***    FONCTION: Positionner la form sur l'ecran en plein ecran            ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FormPanArticle_Max                                        ***
'******************************************************************************
Private Sub FormPanArticle_Max()
    ' Je positionne ma form a l'ecran
    FormPanArticle.Move 0, 0, Screen.Width, Screen.Height
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
    FormPanArticleInitialise
    ' J'initialise la liste des fichiers
    PANArticleFileListBoxInitialise
    ' J'initialise l'image
    PANArticleImageInitialise
    ' J'initialise le media player
    PANArticleMediaPlayerInitialise
    ' J'initilise l'horloge
    PANArticleTimerInitialise
    ' Je positionne la form a l'ecran
    FormPanArticlePosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:Form_Resize()                                             ***
'******************************************************************************
'***    FONCTION: Eepositionnement des controles                            ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
Private Sub Form_Resize()
    ' Je repositionne la liste des fichiers
    PANArticleFileListBoxPosition
    ' Je repositionne l'image
    PANArticleImagePosition
    ' Je repositionne le media player
    PANArticleMediaPlayerPosition
    ' Je repositionne l'horloge
    PANArticleTimerPosition
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PANArticleImage_Click()                                 ***
'******************************************************************************
'***    FONCTION:   Repositionner la form du media de l'article             ***
'***                Dans le coin en haut a droite                           ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PANArticleImage_Click()
    ' Je met la feuille de l'article a sa position initiale
    FormPanArticle_Min
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PANArticleImage_DblClick()                              ***
'******************************************************************************
'***    FONCTION:   Repositionner la form du media de l'article             ***
'***                En plein ecran                                          ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PANArticleImage_DblClick()
    ' Je met en plein ecran la feuille de l'article
    FormPanArticle_Max
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PANArticleImage_Click()                                 ***
'******************************************************************************
'***    FONCTION:   Repositionner la form du media de l'article             ***
'***                Dans le coin en haut a droite                           ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
'Private Sub PANArticleMediaPlayer_Click(Button As Integer, ShiftState As Integer, x As Single, y As Single)
'    ' Je regarde sur quel button on a appuyer
'    Select Case Button
'        Case vbRightButton
'            ' C'est la bouton droit
'            ' Je met la feuille de l'article a sa position initiale
'            FormPanArticle_Min
'        Case Else
'            ' C'est un autre bouton
'    End Select
'    ' Fin
'End Sub

'******************************************************************************
'***    EVENEMENT:  PANArticleMediaPlayer_DblClick(Button As Integer,       ***
'***                ShiftState As Integer, x As Single, y As Single)        ***
'******************************************************************************
'***    FONCTION:   repositionner la form du media de l'article             ***
'***                en plein ecran                                          ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PANArticleMediaPlayer_DblClick(Button As Integer, ShiftState As Long, x As Single, y As Long)
    ' Je regarde sur quel button on a appuyer
    Select Case Button
        Case vbRightButton
            ' C'est la bouton droit
            ' Je met en plein ecran la feuille de l'article
            FormPanArticle_Max
        Case Else
            ' C'est un autre bouton
    End Select
    ' Fin
End Sub

'******************************************************************************
'***    EVENEMENT:  PANArticleTimer_Timer()                                 ***
'******************************************************************************
'***    FONCTION:   changer le media de l'article dans un intervalle de temp***
'***                determiner                                              ***
'******************************************************************************
'***    ENTREE:     Néant                                                   ***
'***                                                                        ***
'***    SORTIE:     Néant                                                   ***
'******************************************************************************
Private Sub PANArticleTimer_Timer()
    ' Je change le media afficher
    MediaChange
    ' Fin
End Sub


