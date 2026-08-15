Attribute VB_Name = "BasModInitialisation"
'******************************************************************************
'***                      (31/10/2000) ***
'******************************************************************************
'***    MODULE:  - Initialisation Programme                                 ***
'******************************************************************************
'***    FONCTION:- Pour Initialiser les chemins des fichiers du programme   ***
'***             - Pour Initialiser la base ODBC acces                      ***
'***             - Pour initialiser les connection a la base vente          ***
'******************************************************************************
'***    PROGRAMMEUR: Royer Jean-Laurent                                     ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:PathInitialise()                                          ***
'***    FONCTION: Pour memoriser tous les chemins des fichiers du programme ***
'******************************************************************************
'***    PROCEDURE:AccessInitialise()                                        ***
'***    FONCTION: Pour initialiser la base acces et le registre windows     ***
'******************************************************************************
'***    PROCEDURE:AccessDataEnvironementInitialise(DataEnv as DataEnvironement)
'***    FONCTION: Pour initialiser les connections du data environement     ***
'******************************************************************************
'***    PROCEDURE:CriteriaInitialise()                                      ***
'***    FONCTION: Pour initialiser la base criteria dans la base de registre***
'******************************************************************************
'***    PROCEDURE:CriteriaDataInitialise(DataEnv as Data)                   ***
'***    FONCTION: Pour initialiser les connections du controle data         ***
'******************************************************************************
'***    PROCEDURE:CriteriaDataBaseInitialise(DataEnv as Database)           ***
'***    FONCTION: Pour initialiser les connections de la data base          ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                             ***
'******************************************************************************
'***    Version 1.0 : 30/10/2000                                            ***
'******************************************************************************
    Option Explicit                                                             ' Je doit etre sur que mes variables on ete declarer
    
'******************************************************************************
'***    Declaration De Constante Global                                     ***
'******************************************************************************

'******************************************************************************
'***    Constante Qui Defini Si On Est En Mode Executable Ou Debug          ***
'******************************************************************************
    #Const DebugExe = True                                                      ' Mode D'Execution Du Programme
    
'******************************************************************************
'***    Constante Qui Defini Le Chemin Des Donnees Du Programme En Debug    ***
'******************************************************************************
    Global Const DebugPath = "C:\MES_SOURCES\VISUAL_BASIC\VENTE\"        ' Le Chemin Du Programme En Mode Debug
    
'******************************************************************************
'***    Constante Pour gerer les nom pour la base access                    ***
'******************************************************************************
    Const ACCESSFICHIER = "Vente.mdb"                                           ' Le nom du fichier de la base ACCESS
    Const ACCESSODBC = "MS Access Database"                                     ' Le nom de la source ODBC pour la base ACCESS
    Const ACCESSDRIVER = "Microsoft Access Driver (*.mdb)"                      ' Le nom du driver Ms Access pour ODBC
    
'******************************************************************************
'***    Constante Pour gerer les nom pour la base access                    ***
'******************************************************************************
    Const CRITERIAFICHIER = "WBouti.dbd"                                        ' Le nom du fichier de la base CRITERIA
    Const CRITERIAODBC = "Criteria"                                             ' Le nom de la source ODBC pour la base CRITERIA
    Const CRITERIADRIVER = "Prologue CRITERIA driver 32bit"                     ' Le nom du driver Ms Access pour ODBC
    
'******************************************************************************
'***    Constante Pour gerer le nom de l'utilisateur et le mot de passe     ***
'******************************************************************************
    Const USERNAME = "Admin"                                                    ' Le nom de l'utilisateur
    Const USERPASSWORD = ""                                                ' Le mot de passe
    
'******************************************************************************
'***    Constante qui defini le nom des bases                               ***
'******************************************************************************

'******************************************************************************
'***    Declaration De La Structure De La Variable Des Chemins Du Programme ***
'******************************************************************************
    Type PathType                                                               ' Je defini un nouveaux type de variables
        App As String                                                           ' Le chemin du programme
        Base As String                                                          ' Le chemin de la base access
        Media As String                                                         ' Le chemin des media
        MediaLogo As String                                                     ' Le chemin du logo
        MediaReference As String                                                ' Le chemin des media de l'article en cours
    End Type
    
'******************************************************************************
'***    Declaration De Variable Global                                      ***
'******************************************************************************

'******************************************************************************
'***    Variable Pour Gerer Les Chemins Du Programme                        ***
'******************************************************************************
    Global AppPath As PathType                                                  ' Les repertoires du programme
        
'******************************************************************************
'***    Declaration Des Procedures Global                                   ***
'******************************************************************************
    
'******************************************************************************
'***    PROCEDURE:PathInitialise()                                          ***
'******************************************************************************
'***    FONCTION: Pour memoriser tous les chemins des fichiers du programme ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  PathInitialise()                                          ***
'******************************************************************************
Public Sub PathInitialise()
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo PathInitialise_Erreur
    ' Je regarde si je suis en mode executable ou debug
    #If DebugExe Then
        ' Je suis en mode debug
        ' Je defini le repertoire de l'application
        AppPath.App = DebugPath
    #Else
        ' Je suis en mode exe
        ' Je defini le repertoire de l'application
        AppPath.App = App.Path + "\"
    #End If
    ' Je defini le chemin de la base
    AppPath.Base = AppPath.App + "DATA\"
    ' Je defini le chemin des media du programme
    AppPath.Media = AppPath.Base + "MEDIA\"
    ' Je defini le repertoire du logo
    AppPath.MediaLogo = AppPath.Media + "LOGO\"
    ' Je defini le repertoire de l'article par default le logo
    AppPath.MediaReference = AppPath.Media + "REFERENCE\"
    ' Fin
PathInitialise_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
PathInitialise_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "PathInitialise"
    ' Je Continue
    Resume PathInitialise_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:AccessInitialise()                                        ***
'******************************************************************************
'***    FONCTION: Pour initialiser la base acces dans la base de registre   ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  AccessInitialise()                                        ***
'******************************************************************************
Public Sub AccessInitialise()
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo AccessInitialise_Erreur
    ' Je defini la DSN Utilisateur dans la base de registre
    DBEngine.RegisterDatabase ACCESSODBC, ACCESSDRIVER, True, "DBQ=" + AppPath.Base + ACCESSFICHIER + ";"
    ' Fin
AccessInitialise_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
AccessInitialise_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "AccessInitialise"
    ' Je Continue
    Resume AccessInitialise_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:AccessDataEnvironementInitialise(DataEnv as DataEnvironement)
'******************************************************************************
'***    FONCTION: Pour initialiser les connections du data environement     ***
'******************************************************************************
'***    ENTREE:   DataEnv - Le data environement a definir les connections  ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  AccessDataEnvironementInitialise(DsrDatVente)             ***
'******************************************************************************
Public Sub AccessDataEnvironementInitialise(ByRef DataEnv As Variant)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo AccessDataEnvironementInitialise_Erreur
    ' Je defini la connection du dataenvironement
    DataEnv.DACDataEnvironement.ConnectionString = "ODBC;DBQ=" + AppPath.Base + ACCESSFICHIER + ";UID=" + USERNAME + ";PWD=" + USERPASSWORD + ";DSN=" + ACCESSODBC + ";"
    ' Je defini la connection du datareport
    DataEnv.DACDataReport.ConnectionString = "ODBC;DBQ=" + AppPath.Base + ACCESSFICHIER + ";UID=" + USERNAME + ";PWD=" + USERPASSWORD + ";DSN=" + ACCESSODBC + ";"
    ' Fin
AccessDataEnvironementInitialise_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
AccessDataEnvironementInitialise_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "AccessDataEnvironementInitialise"
    ' Je Continue
    Resume AccessDataEnvironementInitialise_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:CriteriaInitialise()                                      ***
'******************************************************************************
'***    FONCTION: Pour initialiser la base criteria dans la base de registre***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  CriteriaInitialise()                                      ***
'******************************************************************************
Public Sub CriteriaInitialise()
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo CriteriaInitialise_Erreur
    ' Je defini la DSN Utilisateur dans la base de registre
    DBEngine.RegisterDatabase CRITERIAODBC, CRITERIADRIVER, True, "DBQ=" + AppPath.Base + CRITERIAFICHIER + ";DESCRIPTION=;"
    ' Fin
CriteriaInitialise_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
CriteriaInitialise_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "CriteriaInitialise"
    ' Je Continue
    Resume CriteriaInitialise_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:CriteriaDataInitialise(DataEnv as Data)                   ***
'******************************************************************************
'***    FONCTION: Pour initialiser les connections du controle data         ***
'******************************************************************************
'***    ENTREE:   DataEnv - Le controle data a definir la connection        ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  CriteriaDataInitialise(DataEnv)                           ***
'******************************************************************************
Public Sub CriteriaDataInitialise(ByRef DataEnv As Data)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo CriteriaDataInitialise_Erreur
    ' Je defini la connection du data
    DataEnv.Connect = "ODBC;DBQ=" + AppPath.Base + CRITERIAFICHIER + ";UID=" + USERNAME + ";PWD=" + USERPASSWORD + ";DSN=" + CRITERIAODBC + ";"
    ' Fin
CriteriaDataInitialise_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
CriteriaDataInitialise_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "CriteriaDataInitialise"
    ' Je Continue
    Resume CriteriaDataInitialise_Exit
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:CriteriaDataBaseInitialise(DataEnv as database)           ***
'******************************************************************************
'***    FONCTION: Pour initialiser les connections de la data base          ***
'******************************************************************************
'***    ENTREE:   DataEnv - Le database a definir la connection             ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  CriteriaDataBaseInitialise(DataEnv)                       ***
'******************************************************************************
Public Sub CriteriaDataBaseInitialise(ByRef DataEnv As Database)
    ' En Cas D'Erreur Je Gere L'Erreur
    On Error GoTo CriteriaDataBaseInitialise_Erreur
    ' Je defini la connection du data
    Set DataEnv = OpenDatabase(CRITERIAODBC, dbDriverNoPrompt, False, "ODBC;DBQ=" + AppPath.Base + CRITERIAFICHIER + ";UID=" + USERNAME + ";PWD=" + USERPASSWORD + ";DSN=" + CRITERIAODBC + ";CODEPAGE=1252;")
    ' Fin
CriteriaDataBaseInitialise_Exit:
    ' Je Sort De La Procedure
    Exit Sub
    ' Fin
CriteriaDataBaseInitialise_Erreur:
    ' J'Affiche L'Erreur A L'user
    MsgBox Err.Description, , "CriteriaDataBaseInitialise"
    ' Je Continue
    Resume CriteriaDataBaseInitialise_Exit
    ' Fin
End Sub

