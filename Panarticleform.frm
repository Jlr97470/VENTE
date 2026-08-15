VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form PANArticleForm 
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
   Begin VB.Image PANArticleImage 
      Height          =   1815
      Left            =   240
      Stretch         =   -1  'True
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MediaPlayerCtl.MediaPlayer PANArticleMediaPlayer 
      Height          =   4215
      Left            =   4800
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   4095
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   0   'False
      SendWarningEvents=   0   'False
      SendErrorEvents =   0   'False
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   -1  'True
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   0   'False
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "PANArticleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. 1.0 VB )                  (31/10/2000) ***
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
'***    PROCEDURE:PANArticleFormInitialise()                                ***
'******************************************************************************
'***    FONCTION: Initialiser                                               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANArticleFormInitialise                                  ***
'******************************************************************************
Private Sub PANArticleFormInitialise()
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
'***    PROCEDURE:PANArticleFormPosition()                                  ***
'******************************************************************************
'***    FONCTION: Positionner la form sur l'ecran                           ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANArticleFormPosition                                    ***
'******************************************************************************
Private Sub PANArticleFormPosition()
    ' Je positionne ma form a l'ecran
    ControlInitialise PANArticleForm, Screen.Width * 16.5 / 21, 0, Screen.Width * 4.5 / 21, Screen.Height * 4 / 15
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
    ControlInitialise PANArticleImage, 0, 0, PANArticleForm.ScaleWidth, PANArticleForm.ScaleHeight
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
    ControlInitialise PANArticleMediaPlayer, 0, 0, PANArticleForm.ScaleWidth, PANArticleForm.ScaleHeight
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
'***    PROCEDURE:PANArticleForm_Min()                                      ***
'******************************************************************************
'***    FONCTION: RePositionner la form sur l'ecran a sa position precedente***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANArticleForm_Min                                        ***
'******************************************************************************
Private Sub PANArticleForm_Min()
    ' Je positionne ma form a l'ecran
    PANArticleForm.Move Screen.Width * 16.5 / 21, 0, Screen.Width * 4.5 / 21, Screen.Height * 4 / 15
    ' Fin
    Zone.SetFocus
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:PANArticleForm_Max()                                      ***
'******************************************************************************
'***    FONCTION: Positionner la form sur l'ecran en plein ecran            ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PANArticleForm_Max                                        ***
'******************************************************************************
Private Sub PANArticleForm_Max()
    ' Je positionne ma form a l'ecran
    PANArticleForm.Move 0, 0, Screen.Width, Screen.Height
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
    PANArticleFormInitialise
    ' J'initialise la liste des fichiers
    PANArticleFileListBoxInitialise
    ' J'initialise l'image
    PANArticleImageInitialise
    ' J'initialise le media player
    PANArticleMediaPlayerInitialise
    ' J'initilise l'horloge
    PANArticleTimerInitialise
    ' Je positionne la form a l'ecran
    PANArticleFormPosition
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
    PANArticleForm_Min
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
    PANArticleForm_Max
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
Private Sub PANArticleMediaPlayer_Click(Button As Integer, ShiftState As Integer, x As Single, y As Single)
    ' Je regarde sur quel button on a appuyer
    Select Case Button
        Case vbRightButton
            ' C'est la bouton droit
            ' Je met la feuille de l'article a sa position initiale
            PANArticleForm_Min
        Case Else
            ' C'est un autre bouton
    End Select
    ' Fin
End Sub

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
Private Sub PANArticleMediaPlayer_DblClick(Button As Integer, ShiftState As Integer, x As Single, y As Single)
    ' Je regarde sur quel button on a appuyer
    Select Case Button
        Case vbRightButton
            ' C'est la bouton droit
            ' Je met en plein ecran la feuille de l'article
            PANArticleForm_Max
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


