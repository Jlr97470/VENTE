Attribute VB_Name = "GBLMedia"
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. VB )                                   ***
'******************************************************************************
'***    MODULE POUR LA GESTION DE L'AFFICHAGE D'UN ARTICLE                  ***
'******************************************************************************
'***    FONCTION:   - gerer le changement du repertoire des media           ***
'***                  de l'article                                          ***
'***                - gerer l'affichage des media dans la feuille           ***
'***                  de l'article                                          ***
'******************************************************************************
Option Explicit                                                                 ' Je veut etre sur que toute met variable sont declarer

'******************************************************************************
'***    Declaration des procedures global                                   ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:MediaPathChange(Reference As String)                      ***
'******************************************************************************
'***    FONCTION: Change le repertoire des medias courants                  ***
'******************************************************************************
'***    ENTREE:   Reference - le nom du repertoire                          ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  MediaPathChange("555555")                                 ***
'******************************************************************************
Public Sub MediaPathChange(Reference As String)
    Dim FileSystem As New FileSystemObject                                      ' Pour la gestion des disque et des repertoires
    ' je defini le repertoire des medias
    AppPath.MediaReference = AppPath.MediaReference + Reference + "\"
    ' je verifie l'existance du repertoire
    Select Case FileSystem.FolderExists(AppPath.MediaReference)
        Case True
            ' le repertoire existe
            ' j'utilise plusieur fois les proprietes de la liste des fichiers de la feuilles des articles
            With PANArticleForm.PANArticleFileListBox
                ' je me place dans le repertoires des medias
                .Path = AppPath.MediaReference
                ' je stop l'affichage du media player
                MediaPlayerStop
                ' je stop l'affichage de l'image
                ImageStop
                ' je regarde si le repertoire est vide
                Select Case .ListCount
                    'le repertoire est vide ou inextistant
                    Case 0
                        ' je ne fait rien
                    Case Else
                        ' je change le media afficher
                        MediaChange
                End Select
            ' je n'utilise plus les proprietes de la liste des fichiers
            End With
        Case False
            ' le repertoire n'existe pas
    End Select
    ' fin
End Sub

'******************************************************************************
'***    PROCEDURE:MediaChange()                                             ***
'******************************************************************************
'***    FONCTION: Pour changer le media de l'article en cours               ***
'******************************************************************************
'***    ENTREE:  néant                                                      ***
'***                                                                        ***
'***    SORTIE:  néant                                                      ***
'******************************************************************************
'***    EXEMPLE: MediaChange()                                              ***
'******************************************************************************
Public Sub MediaChange()
    ' je vais beaucoup utiliser les proprietes de la feuille de l'article
    With PANArticleForm
        ' je regarde si la liste des fichiers est vide
        Select Case .PANArticleFileListBox.ListCount
            ' le repertoire est vide ou inexistant
            Case 0
                ' je ne change rient
            Case Else
                ' je regarde sur quel index je suis
                Select Case .PANArticleFileListBox.ListIndex
                    Case Is >= 0
                        ' je regarde si je suis sur le dernier media
                        If .PANArticleFileListBox.ListIndex = .PANArticleFileListBox.ListCount - 1 Then
                            'je suis sur le dernier je me remets en premier
                            .PANArticleFileListBox.ListIndex = 0
                        Else
                            ' je ne suis pas sur le dernier je passe au suivants
                            .PANArticleFileListBox.ListIndex = .PANArticleFileListBox.ListIndex + 1
                        End If
                    Case Else
                        ' je suis hors index je me repositionne au premier
                        .PANArticleFileListBox.ListIndex = 0
                End Select
                ' je regarde quel type de media je doit afficher
                Select Case UCase(Right$(.PANArticleFileListBox.List(.PANArticleFileListBox.ListIndex), 3))
                    Case "BMP", "JPG"
                        ' Je doit afficher une images
                        ' je stop l'affichage du media player
                        MediaPlayerStop
                        ' je verifie qu'il y a plus d'un media a afficher
                        If (.PANArticleFileListBox.ListCount > 1) Or (.PANArticleFileListBox.ListCount = 1 And .PANArticleImage.Visible = False) Then
                            ' je charge l'image de l'article
                            .PANArticleImage.Picture = ImageCharge(AppPath.MediaReference + .PANArticleFileListBox.List(.PANArticleFileListBox.ListIndex))
                            ' Je rend visible le controle image visible
                            .PANArticleImage.Visible = True
                        End If
                    Case "MPG", "MOV", "AVI", "WAV", "MPEG2", "MPEG"
                        ' Je doit afficher une video
                        ' je stop l'affichage de l'image
                        ImageStop
                        ' je verifie qu'il y a plus d'un media a afficher
                        If (.PANArticleFileListBox.ListCount > 1) Or (.PANArticleFileListBox.ListCount = 1 And .PANArticleMediaPlayer.Visible = False) Then
                            ' je charge la video de l'article
                            .PANArticleMediaPlayer.Open AppPath.MediaReference + .PANArticleFileListBox.List(.PANArticleFileListBox.ListIndex)
                            ' Je rend visible le controle image
                            .PANArticleMediaPlayer.Visible = True
                        End If
                End Select
        End Select
    ' plus besoin des proprietes de la feuilles
    End With
    ' fin
End Sub

'******************************************************************************
'***    PROCEDURE:ImageStop()                                               ***
'******************************************************************************
'***    FONCTION: Pour arreter l'affichage de la photo de l'article en cours***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ImageStop()                                               ***
'******************************************************************************
Public Sub ImageStop()
    ' Je vais beaucoup utiliser les proprietes de la feuille de l'article
    With PANArticleForm
        ' Je verifie que l'image est invisible
        Select Case .PANArticleImage.Visible
            Case True
                ' L'image est visible
                ' Je met une image a blanc
                .PANArticleImage.Picture = New StdPicture
                ' Je la rend invisible
                .PANArticleImage.Visible = False
            Case False
                ' L'image est invisible
        End Select
    ' Plus besoin des proprietes de la feuilles
    End With
End Sub

'******************************************************************************
'***    PROCEDURE:MediaPlayerStop()                                         ***
'******************************************************************************
'***    FONCTION: Pour arreter l'affichage du media de l'article en cours   ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  MediaPlayerStop()                                         ***
'******************************************************************************
Public Sub MediaPlayerStop()
    ' je vais beaucoup utiliser les proprietes de la feuille de l'article
    With PANArticleForm
        ' je verifie que le media player est arreter
        Select Case .PANArticleMediaPlayer.Visible
            Case True
                ' le media player est en route
                ' je verifie que le media player est bien arreter
                .PANArticleMediaPlayer.Stop
                ' je repete tant que la video n'est pas arreter
                Do Until .PANArticleMediaPlayer.OpenState = mpClosed
                ' je ne sort que si la video est arreter
                Loop
                ' j'arrete la video
                .PANArticleMediaPlayer.FileName = ""
                ' je le rend invisible le media player
                .PANArticleMediaPlayer.Visible = False
                ' j'interdit le redimensionnement auto
                .PANArticleMediaPlayer.AllowChangeDisplaySize = False
                ' j'interdit le menu contextuel
                .PANArticleMediaPlayer.EnableContextMenu = False
                ' j'interdit le paver de control de position
                .PANArticleMediaPlayer.EnablePositionControls = False
                ' j'interdit le control
                .PANArticleMediaPlayer.EnableTracker = False
                ' j'interdit les message d'erreur
                .PANArticleMediaPlayer.SendErrorEvents = False
                ' j'autorise l'evenement
                .PANArticleMediaPlayer.SendMouseClickEvents = True
                ' j'interdit l'evenement
                .PANArticleMediaPlayer.SendOpenStateChangeEvents = False
                ' j'interdit l'evenement
                .PANArticleMediaPlayer.SendPlayStateChangeEvents = False
                ' j'interdit l'evenement
                .PANArticleMediaPlayer.SendWarningEvents = False
                ' je rend invisible le control du media player
                .PANArticleMediaPlayer.ShowControls = False
                ' je rend invisible le controle du volume
                .PANArticleMediaPlayer.ShowAudioControls = False
                ' je rend invisible le controle de position
                .PANArticleMediaPlayer.ShowPositionControls = False
                ' je rend invisible
                .PANArticleMediaPlayer.ShowTracker = False
            Case False
                ' le media player est arreter
        End Select
    ' plus besoin des proprietes de la feuilles
    End With
End Sub

