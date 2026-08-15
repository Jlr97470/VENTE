Attribute VB_Name = "BasManMicroPointVente"
'******************************************************************************
'***                                   ***
'******************************************************************************
'***    MODULE Principale Du Programme                                      ***
'******************************************************************************
'***    FONCTION: - Gerer les evenements,constantes,variables commun a tous ***
'***                Les Modules du programmes                               ***
'******************************************************************************
'***    PROGRAMMEUR: Royer Jean-Laurent                                     ***
'******************************************************************************
    Option Explicit                                                            ' Je doit etre sur que mes variables on ete declarer
    
'******************************************************************************
'***    Declaration De Constante Global                                     ***
'******************************************************************************

'******************************************************************************
'***    Constante qui defini si on est en mode executable ou debug          ***
'******************************************************************************
    #Const DebugExe = True                                                     ' Le Mode D'Execution
    
'******************************************************************************
'***    Constante qui defini le chemin des donnees du programme en debug    ***
'******************************************************************************
    Global Const DebugPath = "C:\MES_SOURCES\VISUAL_BASIC\VENTE\"       ' Le Chemin De Base Du Programme
             
'******************************************************************************
'***    Constante qui identifie les boutons de validation et d'annulation   ***
'***    D'une ligne du ticket                                               ***
'******************************************************************************
    Global Const ButtonLigneAnnulation = 1                                      ' Pour identifier le button d'annulation d'une ligne de saissie
    Global Const ButtonLigneValidation = 2                                      ' Pour identifier le button de validation d'une ligne de saissie
    
'******************************************************************************
'***    Constante qui identifie les boutons de validation et d'annulation   ***
'***    D'une zone du ticket                                                ***
'******************************************************************************
    Global Const ButtonZoneAnnulation = 1                                       ' Pour identifier le button d'annulation d'une zone de saissie
    Global Const ButtonZoneValidation = 2                                       ' Pour identifier le button de validation d'une zone de saissie
    
'******************************************************************************
'***    Constante qui identifie le boutons du menu principale               ***
'******************************************************************************
    Global Const ButtonMenuPrincipale = 0                                       ' Pour identifier le bouton du menu principale

'******************************************************************************
'***    Declaration de variable global au MPV                               ***
'******************************************************************************

'******************************************************************************
'***    Variable pour memoriser le numero du ticket courant                 ***
'******************************************************************************
    
'******************************************************************************
'***    Variable pour memoriser la zone sur la quel on ete                  ***
'******************************************************************************
    Global Zone As Variant                                                      ' Pour me rapeller la zone sur la quel on ete
    
'******************************************************************************
'***    PROCEDURE:Main()                                                    ***
'******************************************************************************
'***    FONCTION: Pour demarrer le programme des ventes                     ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  Main()                                                    ***
'******************************************************************************
Sub Main()
    ' J'initilise les variables des chemins et des volumes des fichiers du programme
    PathInitialise
    ' J'initilise la base ACCESS dans ODBC
    AccessInitialise
    ' J'initialise les connections a la base access
    AccessDataEnvironementInitialise DsrDatVente
    ' Je travaille sur les commandes de la base vente
    With DsrDatVente
        .rsInsTicketCmd.Open
    
        .rsInsTicketDonneCmd.Open
         
        .rsSelTicketDonneTotalDevise1Cmd.Open
        
        .rsSelTicketDonneTotalDevise2Cmd.Open
        
        .rsInsTicketReglementCmd.Open
         
        .rsSelTicketReglementTotalDevise1Cmd.Open
        
        .rsSelTicketReglementTotalDevise2Cmd.Open
        ' J'ouvre la commande de la liste des magasin et caisse
        .rsSelMagasinCaisseCmd.Open
        ' J'ouvre la commande de la liste des vendeurs
        .rsSelVendeurCmd.Open
        ' J'ouvre la commande de la liste des type de vente
        .rsSelTypeVenteCmd.Open
        ' J'ouvre la commande de la liste des articles
        .rsSelArticleCmd.Open
        ' J'ouvre la commande de la liste des tailles
        .rsSelTailleCmd.Open
        ' J'ouvre la commande de la liste des devises
        .rsSelDeviseCmd.Open
        ' J'ouvre la commande du tarifs
        .rsSelTarifCmd.Open
        
        .rsSelTicketReglementCmd.Open
        ' Je filtre les commandes avec le magasin et la caisse courante
        ' J'ouvre la commande de l'impression d'un ticket
        .rsDarTicketCmd.Open
        
        MagasinCaisseChange
    End With
    ' J'appelle le menu principale
    MenuPrincipale
    ' Fin
End Sub
   
'******************************************************************************
'***    PROCEDURE:MenuPrincipale()                                          ***
'******************************************************************************
'***    FONCTION: Pour lancer le menu principale                            ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  MenuPrincipale()                                          ***
'******************************************************************************
Public Sub MenuPrincipale()
    ' J'affiche la feuille du fond du mpv
    FormPanFond.Show
    ' J'affiche la feuille de l'identite du ticket
    FormPanTicket.Show vbModeless, FormPanFond
    ' Je decharge la photo
    Unload FormPanArticle
    ' Je decharge la feuille du total d'un ticket
    Unload FormPanTicketDonneTotal
    ' Je decharge la feuille du paver total des reglement
    Unload FormPanTicketReglementTotal
    ' Je decharge la feuille du paver de validation
    Unload FormPavTactileValidation
    ' Je decharge la feuille du paver label
    Unload FormPavTactileLabel
    ' Je decharge la feuille du paver total
    Unload FormPavTactileTotal
    ' Je decharge la feuille du paver alpaha
    Unload FormPavTactileAlpha
    ' Je decharge la feuille du paver num
    Unload FormPavTactileNum
    ' Je decharge la feuille du paver type de reglement
    Unload FormPavTypeReglement
    ' Je decharge la feuille du paver type de mouvement
    Unload FormPavTypeMouvement
    ' Je decharge la feuille du paver type de prix
    Unload FormPavTypePrix
    ' Je decharge la feuille du paver type de devise
    Unload FormPavTypeDevise
    ' J'affiche la feuille du menu prinbcipale
    FormPavMenuPrincipale.Show vbModeless, FormPanFond
    ' Je decharge la feuille de saisie d'une ligne du ticket
    Unload FormLigTicketDonne
    ' Je decharge la feuille de saisie d'un reglement
    Unload FormLigTicketReglement
    ' Je decharge la feuille de donne d'un ticket
    Unload FormLstTicketDonne
    ' Je decharge la feuille des donnes des reglements
    Unload FormLstTicketReglement
    ' Fin
End Sub
   
'******************************************************************************
'***    PROCEDURE:MenuSaisieDesVentes()                                     ***
'******************************************************************************
'***    FONCTION: Pour lancer la saissie des ventes                         ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  MenuSaisieDesVentes()                                     ***
'******************************************************************************
Public Sub MenuSaisieDesVentes()
    Dim TicketNumero As Integer                                              ' Pour me rappeller le numero du ticket courant
    ' Je travaille sur les commandes de la base vente
    With DsrDatVente
        ' J'affiche la feuille du fond du mpv
        FormPanFond.Show
        ' J'affiche la feuille de l'identite du ticket
        FormPanTicket.Show vbModeless, FormPanFond
        ' J'affiche la feuille de la photo
        FormPanArticle.Show vbModeless, FormPanFond
        ' J'affiche la feuille du total du ticket
        FormPanTicketDonneTotal.Show vbModeless, FormPanFond
        ' Je decharge la feuille du paver total des reglement
        Unload FormPanTicketReglementTotal
        ' J'affiche la feuille du paver de validation
        FormPavTactileValidation.Show vbModeless, FormPanFond
        ' J'affiche la feuille du clavier
        FormPavTactileLabel.Show vbModeless, FormPanFond
        ' J'affiche les feuilles du paver de total
        FormPavTactileTotal.Show vbModeless, FormPanTicketDonneTotal
        ' J'affiche la feuille du paver alpha numerique
        FormPavTactileAlpha.Show vbModeless, FormPavTactileTotal
        ' J'affiche la feuille du paver clavier numerique
        FormPavTactileNum.Show vbModeless, FormPanFond
        ' Je decharge la feuille du paver type de reglement
        Unload FormPavTypeReglement
        ' J'affiche la feuille du paver du type de mouvement
        FormPavTypeMouvement.Show vbModeless, FormPanFond
        ' J'affiche la feuille du paver de type de prix
        FormPavTypePrix.Show vbModeless, FormPanFond
        ' Je decharge la feuille du paver type de devise
        Unload FormPavTypeDevise
        ' Je decharge le menu
        Unload FormPavMenuPrincipale
        ' J'affiche la feuille de saisie d'une ligne du ticket
        FormLigTicketDonne.Show vbModeless, FormPanFond
        ' Je decharge la feuille de saisie d'un reglement
        Unload FormLigTicketReglement
        ' J'affiche la feuille des donnees de ticket donne
        FormLstTicketDonne.Show vbModeless, FormPanFond
        ' Je decharge la feuille des donnes des reglements
        Unload FormLstTicketReglement
        ' Je verifie si il a un ticket deja saisie dans le fichier
        Select Case .rsInsTicketCmd.RecordCount
            Case 0
                ' Il n'y a pas de ticket
                ' J'initialise le N° ticket
                TicketNumero = 1
            Case Else
                ' Sinon je cherche le dernier ticket
                .rsInsTicketCmd.MoveLast
                ' J'incremente le numero du ticket
                TicketNumero = .rsInsTicketCmd("TicketNumero").Value + 1
        End Select
        ' J'initialise le nouveau ticket dans la table des tickets
        .rsInsTicketCmd.AddNew
        ' Je renseigne le N° de la caisse du ticket
        .rsInsTicketCmd("CleCaisseNumero").Value = .rsSelMagasinCaisseCmd("CleCaisseNumero").Value
        ' Je renseigne la date du ticket
        .rsInsTicketCmd("TicketDate").Value = Date
        ' Je renseigne le numero du ticket
        .rsInsTicketCmd("TicketNumero").Value = TicketNumero
        ' Je met a jour la base
        .rsInsTicketCmd.Update
        ' Je met a jour les liens des commandes avec le ticket courant et la devise courante
        TicketDeviseChange
        ' Je creer une nouvelle ligne de vente
        TicketDonneNouvelle
        ' J'autorise la saisie sur le code vendeur
        TicketDonneZoneChange (ZoneVendeurCar)
    End With
    ' Fin
End Sub
   
'******************************************************************************
'***    PROCEDURE:MenuSaisieDesReglements()                                 ***
'******************************************************************************
'***    FONCTION: Pour lancer la saissie des reglements                     ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  MenuSaisieDesReglements()                                 ***
'******************************************************************************
Public Sub MenuSaisieDesReglements()
    ' J'affiche la feuille du fond du mpv
    FormPanFond.Show
    ' J'affiche la feuille de l'identite du ticket
    FormPanTicket.Show vbModeless, FormPanFond
    ' J'affiche la feuille de la photo
    FormPanArticle.Show vbModeless, FormPanFond
    ' J'affiche la feuille du total du ticket
    FormPanTicketDonneTotal.Show vbModeless, FormPanFond
    ' J'affiche la feuille du paver des total des reglements
    FormPanTicketReglementTotal.Show vbModeless, FormLigTicketReglement
    ' J'affiche la feuille du paver de validation
    FormPavTactileValidation.Show vbModeless, FormPanFond
    ' J'affiche la feuille du clavier
    FormPavTactileLabel.Show vbModeless, FormPanFond
    ' J'affiche les feuilles du paver de total
    FormPavTactileTotal.Show vbModeless, FormPanTicketDonneTotal
    ' Je decharge la feuille du paver alpaha
    Unload FormPavTactileAlpha
    ' J'affiche la feuille du paver clavier numerique
    FormPavTactileNum.Show vbModeless, FormPanFond
    ' J'affiche la feuilles du paver des types de reglements
    FormPavTypeReglement.Show vbModeless, FormPanTicketReglementTotal
    ' Je decharge la feuille du paver type de mouvement
    Unload FormPavTypeMouvement
    ' Je decharge la feuille du paver type de prix
    Unload FormPavTypePrix
    ' J'affiche la feuilles du paver des devises
    FormPavTypeDevise.Show vbModeless, FormPanTicketReglementTotal
    ' Je decharge la feuille du menu prinbcipale
    Unload FormPavMenuPrincipale
    ' Je decharge la feuille de saisie d'une ligne du ticket
    Unload FormLigTicketDonne
    ' J'affiche la feuille de saisie d'une ligne de reglement
    FormLigTicketReglement.Show vbModeless, FormLstTicketReglement
    ' J'affiche la feuille des donnees de ticket donne
    FormLstTicketDonne.Show vbModeless, FormPanFond
    ' J'affiche la feuille des donnees de ticket reglement
    FormLstTicketReglement.Show vbModeless, FormLstTicketDonne
    ' Je creer une nouvelle ligne de reglement
    TicketReglementNouvelle
    ' J'autorise la saisie sur la valeur du reglement
    TicketReglementZoneChange (ZoneValeur)
End Sub
                
'pour lancer l'edition du ticket
Public Sub MenuEditionTicket()
    ' je decharge le paver du bouton total
    Unload FormPavTactileTotal
    'je decharge la ligne de saisie des reglements
    Unload FormLigTicketReglement
    'je charge la feuille du menu secondaire
    FormPavMenuSecondaire.Show vbModeless, FormLstTicketReglement
    
End Sub
   
'******************************************************************************
'***    PROCEDURE:ListeValidation()                                         ***
'******************************************************************************
'***    FONCTION: Pour appeller une liste de choix                          ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ListeValidation()                                         ***
'******************************************************************************
Public Sub ListeValidation()
    ' Je regarde si une zone a ete selectionner
    Select Case IsObject(Zone)
        Case True
            ' Une zone a ete selectionner
            ' Je regarde le nom de la feuille sur laquelle etait l'user
            Select Case Zone.Parent.Name
                Case "FormLigTicketDonne"
                    ' Je suis sur la feuille de saissie d'une ligne du ticket
                    ' J'appelle la liste des choix des donnees du ticket
                    TicketDonneZoneListe
                Case Else
                    ' Je suis sur une autre feuille de saissie
            End Select
            ' Je rend le focus a la zone
            Zone.SetFocus
        Case False
            ' Auncunne zone a ete selectionner
    End Select
    ' Fin
End Sub
   
'******************************************************************************
'***    PROCEDURE:TotalValidation()                                         ***
'******************************************************************************
'***    FONCTION: Pour valider toutes les lignes saissies                   ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TotalValidation()                                         ***
'******************************************************************************
Public Sub TotalValidation()
    ' Je regarde si une zone a ete selectionner
    Select Case IsObject(Zone)
        Case True
            ' Une zone a ete selectionner
            ' Je regarde le nom de la feuille sur laquelle etait l'user
            Select Case Zone.Parent.Name
                Case "FormLigTicketDonne"
                    ' Je suis sur la feuille de saisie d'une ligne du ticket
                    ' Je valide les donnees du tickets
                    TicketDonneValidation
                Case "FormLigTicketReglement"
                    ' Je suis sur la feuille de saisie d'un reglement
                    ' Je valide les reglements du tickets
                    TicketReglementValidation
            End Select
        Case False
            ' Aucunne zone a ete selectionner
    End Select
    ' Fin
End Sub
   
'******************************************************************************
'***    PROCEDURE:LigneValidation()                                         ***
'******************************************************************************
'***    FONCTION: Pour valider une ligne de saissie                         ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LigneValidation()                                         ***
'******************************************************************************
Public Sub LigneValidation()
    ' Je regarde si une zone a ete selectionner
    Select Case IsObject(Zone)
        Case True
            ' Une zone a ete selectionner
            ' Je regarde le nom de la feuille sur laquelle etait l'user
            Select Case Zone.Parent.Name
                Case "FormLigTicketDonne"
                    ' Je suis sur la feuille de saisie d'une ligne d'un ticket
                    ' Je valide la ligne des donnees du ticket
                    TicketDonneLigneValidation
                Case "FormLigTicketReglement"
                    ' Je suis sur la feuille de saisie d'une ligne d'un reglement
                    ' Je valide la ligne des reglements du ticket
                    TicketReglementLigneValidation
            End Select
        Case False
            ' Aucunne zone a ete selectionner
    End Select
    ' Fin
End Sub
      
'******************************************************************************
'***    PROCEDURE:ZoneValidation()                                          ***
'******************************************************************************
'***    FONCTION: Pour valider une zone d'une ligne de saissie              ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ZoneValidation()                                          ***
'******************************************************************************
Public Sub ZoneValidation()
    ' Je regarde si une zone a ete selectionner
    Select Case IsObject(Zone)
        Case True
            ' Une zone a ete selectionner
            ' Je regarde le nom de la feuille sur laquelle etait l'user
            Select Case Zone.Parent.Name
                Case "FormLigTicketDonne"
                    ' Je suis sur la feuille de saissie d'une ligne du ticket
                    ' Je valide la zone de la ligne du ticket
                    TicketDonneZoneValidation
                Case "FormLigTicketReglement"
                    ' Je suis sur la feuille de saissie d'une ligne de reglement
            End Select
            ' Je rend le focus a la zone
            Zone.SetFocus
        Case False
            ' Aucunne zone a ete selectionner
    End Select
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:ToucheValidation(Touche As String)                        ***
'******************************************************************************
'***    FONCTION: Pour envoyer un caractere a la zone de saisie             ***
'******************************************************************************
'***    ENTREE:   Touche - Caractere de la touche utiliser                  ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ToucheValidation("A")                                     ***
'******************************************************************************
Public Sub ToucheValidation(Touche As String)
   Dim Focus_Index As Integer                                                   ' Pour memoriser l'emplacement du focus
   ' Je met en memoire l'emplacement du focus
   Focus_Index = Zone.SelStart
   ' Je met le caractere a la suite du focus
   Zone.Text = Left(Zone.Text, Focus_Index) + Touche
   ' Je met a jour l'emplacement du focus
   Zone.SelStart = Focus_Index + 1
   ' Je rend le focus au control
   Zone.SetFocus
   ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:ZoneAnnulation()                                          ***
'******************************************************************************
'***    FONCTION: Pour annuler une zone de saissie                          ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ZoneAnnulation()                                          ***
'******************************************************************************
Public Sub ZoneAnnulation()
    ' Je regare si la zone est vide
    Select Case Zone.Text
        Case ""
            ' La zone est vide
            ' Je met la ligne a vide
            LigneAnnulation
        Case Else
            ' La zone n'est pas vide
            ' Je met la zone a vide
            Zone.Text = ""
    End Select
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:LigneAnnulation()                                         ***
'******************************************************************************
'***    FONCTION: Pour annuler une ligne de saissie                         ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  LigneAnnulation()                                         ***
'******************************************************************************
Public Sub LigneAnnulation()
    ' Je regarde si une zone a ete selectionner
    Select Case IsObject(Zone)
        Case True
            ' Une zone a ete selectionner
            ' Je regarde le nom de la feuille sur laquelle etait l'user
            Select Case Zone.Parent.Name
                Case "FormLigTicketDonne"
                    ' Je suis sur la feuille de saisie d'une ligne d'un ticket
                    ' Je valide la ligne des donnees du ticket
                    TicketDonneLigneAnnulation
                Case "LGTicketReglementForm"
                    ' Je suis sur la feuille de saisie d'une ligne d'un reglement
                    ' Je valide la ligne des reglements du ticket
                  '  AnnulationTicketReglementLigne
            End Select
        Case False
            ' Aucunne zone a ete selectionner
    End Select
    ' Fin
End Sub
                
'******************************************************************************
'***    PROCEDURE:MagasinCaisseChange()                                     ***
'******************************************************************************
'***    FONCTION: Pour mettre a jour les liens des commandes                ***
'***              avec le magasin et la caisse                              ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  MagasinCaisseChange()                                     ***
'******************************************************************************
Public Sub MagasinCaisseChange()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je regarde le nombre d'enregistrement dans la commande magasin caisse
        Select Case .rsSelMagasinCaisseCmd.RecordCount
            Case 0
                ' Il y a aucun magasin et caisse declarer
            Case Else
                ' Il y des magasin est des caisse declarer
                ' J'applique le filtre sur la liste des vendeurs
                .rsSelVendeurCmd.Filter = "[CLECAISSENUMERO]='" + Str(.rsSelMagasinCaisseCmd("CleCaisseNumero").Value) + "'"
        End Select
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:TicketDeviseChange()                                      ***
'******************************************************************************
'***    FONCTION: Pour mettre a jour les liens des commandes                ***
'***              avec le ticket courant et la devise courante              ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDeviseChange()                                      ***
'******************************************************************************
Public Sub TicketDeviseChange()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je regarde le nombre d'enregistrement dans la commande insertion ticket
        Select Case .rsInsTicketCmd.RecordCount
            Case 0
                ' Il y a pas de ticket
            Case Else
                ' Il y a des  ticket
                ' Je n'affiche que les donnees de ticket donne concernant le ticket courant et la devise courante
                .rsSelTicketDonneCmd.Filter = "[CLETICKETNUMERO]='" + Str(.rsInsTicketCmd("CleTicketNumero").Value) + "' AND [CLEDEVISENUMERO]='" + Str(.rsSelDeviseCmd("CleDeviseNumero").Value) + "'"
                ' Je n'affiche que le total du ticket concernant le ticket courant
                .rsSelTicketDonneTotalDevise1Cmd.Filter = "[CLETICKETNUMERO]='" + Str(.rsInsTicketCmd("CleTicketNumero").Value) + "'"
                ' Je n'affiche que le total du ticket concernant le ticket courant
                .rsSelTicketDonneTotalDevise2Cmd.Filter = "[CLETICKETNUMERO]='" + Str(.rsInsTicketCmd("CleTicketNumero").Value) + "'"
                ' Je n'affiche que les reglements du ticket concernant le ticket courant et la devise courante
                .rsSelTicketReglementCmd.Filter = "[CLETICKETNUMERO]='" + Str(.rsInsTicketCmd("CleTicketNumero").Value) + "'"
                ' Je n'affiche que le total des reglements concernant le ticket courant et la devise courante
                .rsSelTicketReglementTotalDevise1Cmd.Filter = "[CLETICKETNUMERO]='" + Str(.rsInsTicketCmd("CleTicketNumero").Value) + "'"
                ' Je n'affiche que le total des reglements concernant le ticket courant et la devise courante
                .rsSelTicketReglementTotalDevise2Cmd.Filter = "[CLETICKETNUMERO]='" + Str(.rsInsTicketCmd("CleTicketNumero").Value) + "'"
                ' Je ne genere que les informations du ticket courant pour l'edition du ticket
                .rsDarTicketCmd.Filter = "[CLETICKETNUMERO]='" + Str(.rsInsTicketCmd("CleTicketNumero").Value) + "'"
        End Select
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:ReferenceChange()                                         ***
'******************************************************************************
'***    FONCTION: Pour mettre a jour les liens des commandes                ***
'***              Avec la reference                                         ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ReferenceChange()                                         ***
'******************************************************************************
Public Sub ReferenceChange()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je defini le filtre pour les tailles
        ' J'applique le filtre sur la liste des tailles
        .rsSelTailleCmd.Filter = "[ARTICLEREFERENCE]='" + Zone.Text + "'"
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub


