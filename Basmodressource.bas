Attribute VB_Name = "BasModRessource"
Option Explicit

Public Function RessourceStringLoad(IdString As Integer) As String
    On Error Resume Next
    RessourceStringLoad = LoadResString(IdString)
    Select Case Err.Number
        Case 0
        Case 326
            RessourceStringLoad = ""
        Case Else
            MsgBox "Erreur No " + Str(Err) + Err.Description, , ""
    End Select
    Err.Clear   ' Efface les champs d'obLget Err
End Function

Public Function RessourceIconLoad(IdIcon As Integer) As Variant
    On Error Resume Next
    Set RessourceIconLoad = LoadResPicture(IdIcon, vbResIcon)
    Select Case Err.Number
        Case 0
        Case 326
            Set RessourceIconLoad = New StdPicture
        Case Else
            MsgBox "Erreur No " + Str(Err) + Err.Description, , ""
    End Select
    Err.Clear   ' Efface les champs d'obLget Err
End Function

Public Function RessourceCursorLoad(IdCursor As Integer) As Variant
    On Error Resume Next
    Set RessourceCursorLoad = LoadResPicture(IdCursor, vbResCursor)
    Select Case Err.Number
        Case 0
        Case 326
            Set RessourceCursorLoad = New StdPicture
        Case Else
            MsgBox "Erreur No " + Str(Err) + Err.Description, , ""
    End Select
    Err.Clear   ' Efface les champs d'obLget Err
End Function

Public Function RessourceDAImageLoad(IdDAImage As Integer) As Variant
    Dim TmpStr As String
    Const Tmp = "TMP000"
    On Error Resume Next

    TmpStr = StrConv(LoadResData(IdDAImage, "Image"), vbUnicode)
    Select Case Err.Number
        Case 0
        Case 326
            TmpStr = ""
        Case Else
            MsgBox "Erreur No " + Str(Err) + Err.Description, , ""
    End Select
    Err.Clear   ' Efface les champs d'obLget Err
    
    ' Ecriture De L'Image Dans Tampon
    Open Tmp For Output As #1
        Print #1, TmpStr
    Close #1
    
    ' Importation De L'image
    'Set LoadResourceDAImage = ImportImage(Tmp)
    
    'Kill (Tmp)
End Function








