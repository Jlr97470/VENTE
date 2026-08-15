Attribute VB_Name = "BasModGraphic"
'******************************************************************************
'***                      (31/10/2000) ***
'******************************************************************************
'***    MODULE:     - Gestion Des Controles d'une feuille                   ***
'******************************************************************************
'***    FONCTION:   - Positionner,Ajuster les controles                     ***
'***                - Ajuster la taille des fontes                          ***
'***                - Pour gerer l'affichage d'une photo,video              ***
'******************************************************************************
'***    PROGRAMMEUR: Royer Jean-Laurent                                     ***
'******************************************************************************

'******************************************************************************
'***    FONCTION: ImageCharge(ByRef Fichier As String) As Variant:          ***
'***    FONCTION: Pour charger une image si elle existe                     ***
'******************************************************************************
'***    FONCTION: ControlAjuste(ByRef ObjControl As Variant) as boolean     ***
'***    FONCTION: Pour Ajuster les proprietes internes du controle          ***
'******************************************************************************
'***    PROCEDURE:ControlInitialise(ByRef ObjControl As Variant,            ***
'***              ObjLeft As Single,ObjTop As Single, ObjWidth As Single,   ***
'***              ObjHeight As Single) as boolean                           ***
'***    FONCTION: Pour initialiser la position d'un controle et ajuster     ***
'***              la taille des fontes                                      ***
'******************************************************************************
'***    PROCEDURE:FontInitialise(ByRef ObjControl As Variant)               ***
'***    FONCTION: Pour initialiser les fontes du controles                  ***
'******************************************************************************
'***    PROCEDURE:FontAjuste(ByRef ObjControl As Variant)                   ***
'***    FONCTION: Pour ajuster la taille des fontes des controles           ***
'******************************************************************************
'***    PROCEDURE:ImagePlay(ByRef ObjControl As Image,                      ***
'***                        ByRef Fichier As String)                        ***
'***    FONCTION: Pour lancer l'afficher une image photo                    ***
'******************************************************************************
'***    PROCEDURE:ImageStop(ByRef ObjControl As Image)                      ***
'***    FONCTION: Pour arreter l'affichage de la photo                      ***
'******************************************************************************
'***    PROCEDURE:MediaPlayerPlay(ByRef ObjControl As MediaPlayer,          ***
'***                              ByRef Fichier As String)                  ***
'***    FONCTION: Pour lancer l'affichage d'Un media                        ***
'******************************************************************************
'***    PROCEDURE:MediaPlayerStop(ByRef ObjControl As MediaPlayer)          ***
'***    FONCTION: Pour arreter l'affichage d'Un media                       ***
'******************************************************************************


'******************************************************************************
'***    MODIF :                                                             ***
'******************************************************************************
'***    Version 1.0 : 30/10/2000                                            ***
'******************************************************************************
    Option Explicit                                                            ' Je doit etre sur que mes variables on ete declarer

'******************************************************************************
'***    PROCEDURE:ControlInitialise(ByRef ObjControl As Control,            ***
'***                ObjLeft As Single, ObjTop As Single, ObjWidth As Single,***
'***                ObjHeight As Single)                                    ***
'******************************************************************************
'***    FONCTION: Pour initialiser la position d'un controle et ajuster la  ***
'***              taille des fontes                                         ***
'******************************************************************************
'***    ENTREE:   ObjControl - Le control concerne                          ***
'***              ObjLeft    - La position gauche du controle               ***
'***              ObjTop     - La position haute du controle                ***
'***              ObjWidth   - La longueur du controle                      ***
'***              ObjHeight  - La largeur gauche du controle                ***
'***    SORTIE:   ControleInitialise - Si j'ai bien initiliser le control   ***
'******************************************************************************
'***    EXEMPLE:  ControleInitialise(Label,10,10,100,100)                   ***
'******************************************************************************
Public Function ControlInitialise(ByRef ObjControl As Variant, ObjLeft As Single, ObjTop As Single, ObjWidth As Single, ObjHeight As Single) As Boolean
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ControlInitialise_Erreur
    ' Je verifie que le controle a bouger
    If (ObjControl.Left <> ObjLeft) Or (ObjControl.Top <> ObjTop) Or (ObjControl.Width <> ObjWidth) Or (ObjControl.Height <> ObjHeight) Then
        ' Le controle a bouger
        ' Je repositionne le controle
        ObjControl.Move ObjLeft, ObjTop, ObjWidth, ObjHeight
        ' Je regarde si le controle est un bouton
        If TypeOf ObjControl Is CommandButton Then
            ' C'est un bouton
            ' Je doit ajuster les fontes
            ControlInitialise = True
        End If
        ' Je regarde si le controle est un textbox
        If TypeOf ObjControl Is TextBox Then
            ' C'est un textbox
            ' Je doit ajuster les fontes
            ControlInitialise = True
        End If
        ' Je regarde si le controle est un label
        If TypeOf ObjControl Is Label Then
            ' C'est un label
            ' Je doit ajuster les fontes
            ControlInitialise = True
        End If
        ' Je regarde si le controle est un maskedbox
        If TypeOf ObjControl Is MaskEdBox Then
            ' C'est un maskedbox
            ' Je doit ajuster les fontes
            ControlInitialise = True
        End If
        ' Je regarde si le controle est un DBgrid
    '    If TypeOf ObjControl Is DBGrid Then
            ' C'est un dbgrid
            ' Je ne doit pas ajuster les fontes
         '   ControlInitialise = False
      '  End If
        ' Je regarde si je doit ajuster les fontes
        Select Case ControlInitialise
            Case True
                ' Je doit ajuster les fontes
                ' Je redimensionne la taille des fontes
                FontAjuste ObjControl
            Case False
                ' Je ne doit pas ajuster les fontes
        End Select
        ' J'ajuste le controle
        ControlAjuste ObjControl
        ' J'ai initialiser le controle
        ControlInitialise = True
    End If
    ' Fin
ControlInitialise_Exit:
    ' Je Sort De La Procedure
    Exit Function
    ' Fin
ControlInitialise_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "ControlInitialise"
    ' Je Continue
    Resume ControlInitialise_Exit
    ' Fin
End Function

'******************************************************************************
'***    PROCEDURE:ControlAjuste(ByRef ObjControl As Control)                ***
'******************************************************************************
'***    FONCTION: Pour Ajuster les proprietes internes du controle          ***
'******************************************************************************
'***    ENTREE:   ObjControl - Le control concerne                          ***
'***                                                                        ***
'***    SORTIE:   ControlAjuste - Si j'ai bien ajuster le control           ***
'******************************************************************************
'***    EXEMPLE:  ControlAjuste(Label)                                      ***
'******************************************************************************
Public Function ControlAjuste(ByRef ObjControl As Variant) As Boolean
    Dim ColoneNum As Integer                                                   ' Le numero de la colonne courante
    Dim ColoneNbrMax As Integer                                                ' Le nombre de colonne afficher
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ControlAjuste_Erreur
    ' Je regarde si le controle est un bouton
    If TypeOf ObjControl Is CommandButton Then
        ' C'est un bouton
        ' Il ne faut pas ajuster le controle
        ControlAjuste = False
    End If
    ' Je regarde si le controle est un textbox
    If TypeOf ObjControl Is TextBox Then
        ' C'est un textbox
        ' Il ne faut pas ajuster le controle
        ControlAjuste = False
    End If
    ' Je regarde si le controle est un label
    If TypeOf ObjControl Is Label Then
        ' C'est un label
        ' Il ne faut pas ajuster le controle
        ControlAjuste = False
    End If
    ' Je regarde si le controle est un maskedbox
    If TypeOf ObjControl Is MaskEdBox Then
        ' C'est un maskedbox
        ' Il ne faut pas ajuster le controle
        ControlAjuste = False
    End If
    ' Je regarde si le controle est un DBGrid
 '   If TypeOf ObjControl Is DBGrid Then
        ' C'est un dbgrid
        ' Il faut ajuster le controle
     '   ControlAjuste = True
 '   End If
    ' Je regarde si le controle est un DBGrid
    If TypeOf ObjControl Is DataGrid Then
        ' C'est un dbgrid
        ' Il faut ajuster le controle
        ControlAjuste = True
    End If
    ' Je regarde si il faut ajuster le controle
    Select Case ControlAjuste
        Case True
            ' Il faut ajuster le controle
            ' Je met a zero le nombre de caractere a afficher sur une ligne
            ColoneNbrMax = 0
            ' Je fait defiler toutes les colonnes de la DBgrid
            For ColoneNum = 0 To ObjControl.Columns.Count - 1
                ' Je regarde si la colonne est visible
                Select Case ObjControl.Columns(ColoneNum).Visible
                    Case True
                        ' La colonne est visible
                        ' Je me rapelle le nombre de colonne afficher
                        ColoneNbrMax = ColoneNbrMax + 1
                    Case False
                        ' La colonne est invisible
                End Select
            ' J'ai fini de faire defiler les colonnes
            Next ColoneNum
            ' Je fais defiler les colonnes de la liste des fichiers
            For ColoneNum = 0 To ObjControl.Columns.Count - 1
                ' Je regarde si la colonne est visible
                Select Case ObjControl.Columns(ColoneNum).Visible
                    Case True
                        ' La colonne est visible
                        ' Je defini la longueur
                        ObjControl.Columns(ColoneNum).Width = ObjControl.Width / ColoneNbrMax
                    Case False
                        ' La colonne est invisible
                End Select
                ' Je ne sort que si j'ai fini de faire defiler les colonnes
            Next ColoneNum
        Case False
            ' Il ne faut pas ajuster le controle
    End Select
    ' Fin
ControlAjuste_Exit:
    ' Je Sort De La Procedure
    Exit Function
    ' Fin
ControlAjuste_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "ControlAjuste"
    ' Je Continue
    Resume ControlAjuste_Exit
    ' Fin
End Function

'******************************************************************************
'***    PROCEDURE:FontInitialise(ByRef ObjControl As Control)               ***
'******************************************************************************
'***    FONCTION: Pour initialiser les fontes du controles                  ***
'******************************************************************************
'***    ENTREE:   ObjControl - Le control concerne                          ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FontInitialise(Label)                                     ***
'******************************************************************************
Public Sub FontInitialise(ByRef ObjControl As Variant)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo FontInitialise_Erreur
    ' Je defini le nom de la fonte
    ObjControl.Font.Name = "Arial Narrow"
    ' Je defini la fonte du controle en gras
    ObjControl.Font.Bold = True
    ' Je defini la taille de la fonte a 8
    ObjControl.Font.Size = 8
    ' Fin
FontInitialise_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
FontInitialise_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "FontInitialise"
    ' Je Continue
    Resume FontInitialise_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:FontAjuste(ByRef ObjControl As Control)                    ***
'******************************************************************************
'***    FONCTION: Pour ajuster la taille des fontes des controles           ***
'******************************************************************************
'***    ENTREE:   ObjControl - Le control concerne                          ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  FontAjuste(Label)                                         ***
'******************************************************************************
Public Sub FontAjuste(ByRef ObjControl As Variant)
    Dim LetX As Single                                                          ' La longueur d'un caractere
    Dim LetY As Single                                                          ' La largeur d'un caractere
    Dim ObjString As String                                                     ' Le texte du controle
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo FontAjuste_Erreur
    ' J'Initialise la fontes du control
    FontInitialise ObjControl
    ' Je regarde si le controle est un bouton
    If TypeOf ObjControl Is CommandButton Then
        ' C'est un bouton
        ' Je recupere le texte du bouton
        ObjString = ObjControl.Caption
    End If
    ' Je regarde si le controle est un textbox
    If TypeOf ObjControl Is TextBox Then
        ' C'est un textbox
        ' Je recupere le texte du textbox
        ObjString = ObjControl.Text
    End If
    ' Je regarde si le controle est un label
    If TypeOf ObjControl Is Label Then
        ' C'est un label
        ' Je recupere le texte du label
        ObjString = ObjControl.Caption
    End If
    ' Je regarde si le controle est un maskedbox
    If TypeOf ObjControl Is MaskEdBox Then
        ' C'est un maskedbox
        ' Je recupere le texte du maskedbox
        ObjString = ObjControl.Text
    End If
    ' Je regade la longueur du texte du controle
    Select Case Len(ObjString)
        Case 0
            ' Le texte a afficher est une chaine vide
        Case Else
            ' Le texte a afficher n'est pas une chaine vide
            ' Je calcule la longueur en pixel d'un caractere
            LetX = ObjControl.Parent.ScaleX(ObjControl.Width / Len(ObjString), vbTwips, vbPixels)
            ' Je calcule la largueur en pixel d'un caractere
            LetY = ObjControl.Parent.ScaleY(ObjControl.Height, vbTwips, vbPixels) * 5 / 8
            ' Je compare la largeur et la longueur d'un caractere
            Select Case LetX
                Case Is > LetY
                    ' La longueur est plus grande que la largeur
                    ' Je regarde si la valeur de la largeur
                    Select Case LetY
                        Case Is > 0
                            ' La valeur est positive
                            ' Je defini la taille de la fonte
                            ObjControl.Font.Size = LetY
                        Case Else
                            ' La valeur est nul
                    End Select
                Case Else
                    ' La largueur est plus grande que la longeur
                    ' Je regarde si la valeur de la longeur
                    Select Case LetX
                        Case Is > 0
                            ' La valeur est positive
                            ' Je defini la taille de la fonte
                            ObjControl.Font.Size = LetX
                        Case Else
                            ' La valeur est nul
                    End Select
            End Select
    End Select
    ' Fin
FontAjuste_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
FontAjuste_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "FontAjuste"
    ' Je Continue
    Resume FontAjuste_Exit
    ' Fin
End Sub

'******************************************************************************
'***    FONCTION: ImageCharge(ByRef Fichier As String) As Image             ***
'******************************************************************************
'***    FONCTION: Pour charger une image si elle existe                     ***
'******************************************************************************
'***    ENTREE:   Fichier - Le nom et le chemin de l'image                  ***
'***                                                                        ***
'***    SORTIE:   ImageCharge  - Une image                                  ***
'******************************************************************************
'***    EXEMPLE:  =ImageCharge("f:\image.bmp")                              ***
'******************************************************************************
Public Function ImageCharge(ByRef Fichier As String) As Variant
    ' Je gere l'erreur
    On Error Resume Next
    ' Je charge l'image
    Set ImageCharge = LoadPicture(Fichier)
    ' Je regarde si l'image a bien ete charger
    Select Case Err.Number
        Case 0
            ' L'image a bien ete charger
        Case 53, 76
            ' L'image n'exite pas ou est invalide
            ' Je met une image vide
            Set ImageCharge = New StdPicture
        Case Else
            ' Il y a une autre erreur je la gere
            MsgBox "Erreur No " + Str(Err) + Err.Description, , ""
    End Select
    ' J'efface les donnees de l'erreur
    Err.Clear
    ' Fin
End Function

'******************************************************************************
'***    PROCEDURE:ImageStop(ByRef ObjControl As Image)                      ***
'******************************************************************************
'***    FONCTION: Pour arreter l'affichage de la photo                      ***
'******************************************************************************
'***    ENTREE:   ObjControl - Le controle image                            ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ImageStop(ControlImage)                                   ***
'******************************************************************************
Public Sub ImageStop(ByRef ObjControl As Image)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ImageStop_Erreur
    ' Je verifie que l'image est visible
    Select Case ObjControl.Visible
        Case True
            ' L'image est visible
            ' Je met une image a blanc
            ObjControl.Picture = New StdPicture
            ' Je la rend invisible
            ObjControl.Visible = False
        Case False
            ' L'image est invisible
    End Select
    ' Fin
ImageStop_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
ImageStop_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "ImageStop"
    ' Je Continue
    Resume ImageStop_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:ImagePlay(ByRef ObjControl As Image,                      ***
'***                        ByRef Fichier As String)                        ***
'******************************************************************************
'***    FONCTION: Pour lancer l'afficher une image photo                    ***
'******************************************************************************
'***    ENTREE:   ObjControl - Le controle image                            ***
'***              Fichier    - Le nom +chemin du fichier                    ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ImagePlay(ControlImage,"c:\titi.jpg")                     ***
'******************************************************************************
Public Sub ImagePlay(ByRef ObjControl As Image, ByRef Fichier As String)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo ImagePlay_Erreur
    ' Je verifie que l'image est invisible
    Select Case ObjControl.Visible
        Case False
            ' L'image est invisible
            ' Je met une image a blanc
            ObjControl.Picture = ImageCharge(Fichier)
            ' Je la rend visible
            ObjControl.Visible = True
        Case True
            ' L'image est visible
    End Select
    ' Fin
ImagePlay_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
ImagePlay_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "ImagePlay"
    ' Je Continue
    Resume ImagePlay_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:MediaPlayerStop(ByRef ObjControl As MediaPlayer)          ***
'******************************************************************************
'***    FONCTION: Pour arreter l'affichage d'Un media                       ***
'******************************************************************************
'***    ENTREE:   ObjControl - Le control media player                      ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  MediaPlayerStop(ControlMediaPlayer)                       ***
'******************************************************************************
Public Sub MediaPlayerStop(ByRef ObjControl As WindowsMediaPlayer)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo MediaPlayerStop_Erreur
    ' Je verifie que le media player est visible
    Select Case ObjControl.Visible
        Case True
            ' Le media player est visible
            ' Je verifie que le media player est bien arreter
            ObjControl.Stop
            ' Je repete tant que le controle n'est pas pres
            Do Until ObjControl.ReadyState = mpReadyStateComplete
                ' Je laisse les autres evenements
                DoEvents
            ' Je ne sort que si la video est arreter
            Loop
            ' J'arrete la video
            ObjControl.FileName = ""
            ' Je le rend invisible le media player
            ObjControl.Visible = False
            ' J'interdit le demarrage automatique
            ObjControl.autoStart = False
            ' J'interdit le redimensionnement auto
            ObjControl.AllowChangeDisplaySize = False
            ' J'interdit le menu contextuel
            ObjControl.enableContextMenu = False
            ' J'interdit le paver de control de position
            ObjControl.EnablePositionControls = False
            ' J'interdit le control
            ObjControl.EnableTracker = False
            ' J'interdit les message d'erreur
            ObjControl.SendErrorEvents = False
            ' J'autorise l'evenement
            ObjControl.SendMouseClickEvents = True
            ' J'interdit l'evenement
            ObjControl.SendOpenStateChangeEvents = False
            ' J'interdit l'evenement
            ObjControl.SendPlayStateChangeEvents = False
            ' J'interdit l'evenement
            ObjControl.SendWarningEvents = False
            ' Je rend invisible le control du media player
            ObjControl.ShowControls = False
            ' Je rend invisible le controle du volume
            ObjControl.ShowAudioControls = False
            ' Je rend invisible le controle de position
            ObjControl.ShowPositionControls = False
            ' Je rend invisible
            ObjControl.ShowTracker = False
        Case False
            ' Je media player est invisible
    End Select
    ' Fin
MediaPlayerStop_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
MediaPlayerStop_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "MediaPlayerStop"
    ' Je Continue
    Resume MediaPlayerStop_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:MediaPlayerPlay(ByRef ObjControl As MediaPlayer,          ***
'***                              ByRef Fichier As String)                  ***
'******************************************************************************
'***    FONCTION: Pour lancer l'affichage d'Un media                        ***
'******************************************************************************
'***    ENTREE:   ObjControl  - Le controle media player                    ***
'***              Fichier     - Le nom+chemin du fichier                    ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  MediaPlayerPlay(ControlMediaPlayer,"c:\titi.mov")         ***
'******************************************************************************
Public Sub MediaPlayerPlay(ByRef ObjControl As WindowsMediaPlayer, ByRef Fichier As String)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo MediaPlayerPlay_Erreur
    ' Je verifie que le media player est invisible
    Select Case ObjControl.Visible
        Case False
            ' Le media player est invisible
            ' Je le rend visible le media player
            ObjControl.Visible = True
            ' J'interdit le demarrage automatique
            ObjControl.autoStart = False
            ' J'interdit le redimensionnement auto
            ObjControl.AllowChangeDisplaySize = False
            ' J'interdit le menu contextuel
            ObjControl.enableContextMenu = False
            ' J'interdit le paver de control de position
            ObjControl.EnablePositionControls = False
            ' J'interdit le control
            ObjControl.EnableTracker = False
            ' J'interdit les message d'erreur
            ObjControl.SendErrorEvents = False
            ' J'autorise l'evenement
            ObjControl.SendMouseClickEvents = True
            ' J'interdit l'evenement
            ObjControl.SendOpenStateChangeEvents = False
            ' J'interdit l'evenement
            ObjControl.SendPlayStateChangeEvents = False
            ' J'interdit l'evenement
            ObjControl.SendWarningEvents = False
            ' Je rend invisible le control du media player
            ObjControl.ShowControls = False
            ' Je rend invisible le controle du volume
            ObjControl.ShowAudioControls = False
            ' Je rend invisible le controle de position
            ObjControl.ShowPositionControls = False
            ' Je rend invisible
            ObjControl.ShowTracker = False
            ' Je charge la video de l'article
            ObjControl.Open Fichier
            ' Je repete tant que le controle n'est pas pres
            Do Until ObjControl.ReadyState = mpReadyStateComplete
                ' Je laisse les autres evenements
                DoEvents
            ' Je ne sort que si la video est arreter
            Loop
            ' Je demarre la video
            ObjControl.play
            ' Je repete tant que le controle n'est pas pres
            Do Until ObjControl.ReadyState = mpReadyStateComplete
                ' Je laisse les autres evenements
                DoEvents
            ' Je ne sort que si la video est arreter
            Loop
        Case True
            ' Le media player est visible
    End Select
    ' Fin
MediaPlayerPlay_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
MediaPlayerPlay_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "MediaPlayerPlay"
    ' Je Continue
    Resume MediaPlayerPlay_Exit
    ' Fin
End Sub



