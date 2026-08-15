Attribute VB_Name = "GBLTicketReglement"
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. VB )                                   ***
'******************************************************************************
'***    MODULE POUR LA GESTION DES REGLEMENT DU TICKET                      ***
'******************************************************************************
'***    FONCTION:   - gerer l'affichage de l'interface de saissie           ***
'***                - gerer la validation des zones                         ***
'***                - gerer les modifications des donnes dans la table      ***
'******************************************************************************
    Option Explicit                                                             ' Je veut etre sur que toute met variable sont declarer
     
'******************************************************************************
'***    Declaration de constante global du MPV                              ***
'******************************************************************************

'******************************************************************************
'***    Constantes qui identifie la zone de saissie d'une ligne de reglement***
'******************************************************************************
    Global Const ZoneReglementCar = 0                                           ' Pour identifier la zone reglement car
    Global Const ZoneReglementLibelle = 1                                       ' Pour identifier la zone libelle du reglement
    Global Const ZoneValeur = 2                                                 ' Pour identifier la zone valeur du reglement
    Global Const ZoneDeviseCar = 3                                              ' Pour identifier la zone devise du reglement
    
'******************************************************************************
'***    Constante qui identifie les zone du total des reglements du ticket  ***
'******************************************************************************
    Global Const ZoneResteAPayer = 0                                            ' Pour identifier la zone a payer
    Global Const ZoneTotalReste1 = 1                                            ' Pour identifier la zone total reste dans la devise 1
    Global Const ZoneTotalResteDevise1 = 2                                      ' Pour identifier la zone devise 1 du total reste
    Global Const ZoneTotalReste2 = 3                                            ' Pour identifier la zone total reste dans la devise 2
    Global Const ZoneTotalResteDevise2 = 4                                      ' Pour identifier la zone devise 2 du total reste
    
'******************************************************************************
'***    Constantes qui identifie le bouton des types de reglements          ***
'******************************************************************************
    Global Const ButtonReglementCheque = 0                                      ' Pour identifier le bouton reglement par cheque
    Global Const ButtonReglementEspece = 1                                      ' Pour identifier le bouton reglement par espece
    Global Const ButtonReglementCarte = 2                                       ' Pour identifier le bouton reglement par carte
    Global Const ButtonReglementUtilisationAvoir = 3                            ' Pour identifier le bouton reglement par utilisation d'avoir
    Global Const ButtonReglementDiver1 = 4                                      ' Pour identifier le bouton reglement par diver 1
    Global Const ButtonReglementDiver2 = 5                                      ' Pour identifier le bouton reglement par diver 2
    
'******************************************************************************
'***    Constantes qui identifie le bouton des devises                      ***
'******************************************************************************
    Global Const ButtonDeviseFranc = 0                                          ' Pour identifier le bouton de la devise franc
    Global Const ButtonDeviseEuro = 1                                           ' Pour identifier le bouton de la devise euro
    
'******************************************************************************
'***    Constantes qui identifie le type de mouvement dans                  ***
'***    la table des type de mouvement                                      ***
'******************************************************************************
    Global Const ReglementCheque = "S"                                          ' Pour identifier le Reglement par cheque
    Global Const ReglementEspece = "E"                                          ' Pour identifier le Reglement par espece
    Global Const ReglementCarte = "C"                                           ' Pour identifier le Reglement par carte
    Global Const ReglementUtilisationAvoir = "U"                                ' Pour identifier le Reglement par util avoir
    Global Const ReglementDiver1 = "T"                                          ' Pour identifier le Reglement par div.1
    Global Const ReglementDiver2 = "O"                                          ' Pour identifier le Reglement par div.2

'******************************************************************************
'***    Declaration des procedures global                                   ***
'******************************************************************************
                        
'******************************************************************************
'***    PROCEDURE:TicketReglementAjoute()                                   ***
'******************************************************************************
'***    FONCTION: Pour ajouter une ligne dans la table ticket reglement     ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketReglementAjoute()                                   ***
'******************************************************************************
Public Sub TicketReglementAjoute()
    Dim I As Integer                                                            ' Pour memoriser le numero du label courant
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DAEVenteDSR
        ' Je met a jour le lien entre la commande ticket reglement et ticket
        .rsINTicketReglementCmd("CLETICKETNUMERO").Value = .rsINTicketCmd("CLETICKETNUMERO").Value
        ' Je met dans la ligne de reglement le type de reglement
        .rsINTicketReglementCmd("CLETYPEREGLEMENTNUMERO").Value = .rsSETypeReglementCmd("CLETYPEREGLEMENTNUMERO").Value
        ' Je met dans la ligne de reglement la devise
        .rsINTicketReglementCmd("CLEDEVISENUMERO").Value = .rsSEDeviseCmd("CLEDEVISENUMERO").Value
        ' Je rajoute ma ligne a la table ticket donne
        .rsINTicketReglementCmd.Update
        ' Je met a jour les informations de la commande qui affiche les reglements du ticket
        .rsSETicketReglementCmd.Requery
        ' Je met a jour les donnees du ticket affichee a l'ecran
        LSTTicketReglementForm.LSTTicketReglementDataGrid.DataMember = "SETicketReglementCmd"
        ' Je met a jour les informations de la commande qui affiche le total des reglements
        .rsSETicketReglementTotalDevise1Cmd.Requery
        ' Je met a jour les informations de la commande qui affiche le total des reglements
        .rsSETicketReglementTotalDevise2Cmd.Requery
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Je vais beaucoup utiliser les proprietes des labels du total du tickets
    With PANTicketReglementTotalForm.PANTicketReglementTotalLabel
        ' Je fait defiler tous les labels de la form ticket donne total
        For I = 0 To .Count - 1
            ' Je regarde sur quel zone je suis
            Select Case I
                Case ZoneResteAPayer
                    ' Je suis sur la zone total
                Case Else
                    ' Je suis sur une autre zone
                    ' Je met a jour l'affichage du total des reglements
                    .Item(I).DataField = ""
            End Select
        ' Je ne sort que si j'ai fini de faire defiler les boutons
        Next I
    ' Je n'ai plus besoin des proprites des labels du total du tickets
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:TicketReglementSupprime()                                 ***
'******************************************************************************
'***    FONCTION: Pour supprimer une ligne dans la table ticket reglement   ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketReglementSupprime()                                 ***
'******************************************************************************
Public Sub TicketReglementSupprime()

End Sub

'******************************************************************************
'***    PROCEDURE:TicketReglementEdite()                                    ***
'******************************************************************************
'***    FONCTION: Pour editer une ligne dans la table ticket reglement      ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketReglementEdite()                                    ***
'******************************************************************************
Public Sub TicketReglementEdite()

End Sub

'******************************************************************************
'***    PROCEDURE:TicketReglementNouvelle()                                 ***
'******************************************************************************
'***    FONCTION: Pour ajouter une nouvelle ligne dans la table ticket      ***
'***              Reglement                                                 ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketReglementNouvelle()                                 ***
'******************************************************************************
Public Sub TicketReglementNouvelle()
    With DAEVenteDSR 'comme d'hab
        'je creer une nouvelle ligne dans la commande ticketdonne
        .rsINTicketReglementCmd.AddNew
        'je me a jour le lien entre la commande ticket et ticketdonne
        .rsINTicketReglementCmd("CLETICKETNUMERO").Value = .rsINTicketReglementCmd("CLETICKETNUMERO").Value
    End With
End Sub

'******************************************************************************
'***    PROCEDURE:TicketReglementAnnule()                                   ***
'******************************************************************************
'***    FONCTION: Pour Annuler une ligne dans la table ticket reglement     ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketReglementAnnule()                                   ***
'******************************************************************************
Public Sub TicketReglementAnnule()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DAEVenteDSR
        ' J'annule la ligne
        .rsINTicketReglementCmd.CancelUpdate
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:TicketReglementInterfaceChange(Zone As Integer)           ***
'******************************************************************************
'***    FONCTION: Changer l'interface utilisateur selon la zone ou on est   ***
'******************************************************************************
'***    ENTREE:   Zone : repere le numero de la zone ou on est              ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketReglementInterfaceChange(0)                         ***
'******************************************************************************
Public Sub TicketReglementInterfaceChange(Zone As Integer)
    ' J'indique a l'user ou je suis
    PAVTactileLabelForm.PAVTactileLabelLabel = RessourceStringLoad(750 + Zone)
    ' Je regarde sur quel zone je suis
    Select Case Zone
        Case ZoneReglementCar
            ' Je suis sur la zone type reglement car
            ' J'active le paver total
            PAVTactileTotalForm.Enabled = False
            ' Je desactive le paver alpha numerique
            PAVTactileAlphaForm.Enabled = False
            ' J'active le paver numerique
            PAVTactileNumForm.Enabled = False
            ' J'active le paver du type des reglements
            PAVTypeReglementForm.Enabled = True
        Case ZoneReglementLibelle
            ' Je suis sur la zone type reglement libelle
            ' Je desactive le paver total
            PAVTactileTotalForm.Enabled = False
            ' Je desactive le paver alpha numerique
            PAVTactileAlphaForm.Enabled = False
            ' Je desactive le paver numerique
            PAVTactileNumForm.Enabled = False
            ' J'active le paver du type des reglements
            PAVTypeReglementForm.Enabled = True
        Case ZoneValeur
            ' Je suis sur la zone valeur reglement
            ' J'active le paver total
            PAVTactileTotalForm.Enabled = True
            ' Je desactive le paver alpha numerique
            PAVTactileAlphaForm.Enabled = False
            ' J'active le paver numerique
            PAVTactileNumForm.Enabled = True
            ' J'active le paver du type des reglements
            PAVTypeReglementForm.Enabled = True
    End Select
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:TicketReglementZoneChange(Index As Integer)               ***
'******************************************************************************
'***    FONCTION: Pour autorisée la saissie dans une seuls zone             ***
'***              Indiquer a l'user la zone active                          ***
'******************************************************************************
'***    ENTREE:   Index : repere qui indique la zone a autorisée en saisie  ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketReglementZoneChange(Index As Integer)               ***
'******************************************************************************
Public Sub TicketReglementZoneChange(Index As Integer)
    Dim I As Integer    ' pour pouvoir faire defiler les zones de saisies
    ' Je fait defiler tous les zones pour
    For I = 0 To LIGTicketReglementForm.LIGTicketReglementMaskEdBox.Count - 1
        ' Je regarde si c'est la zone a autoriser
        Select Case I
            Case Index
                ' J'ai trouver la zone a saisir
                ' Je change l'interface utilisateur
                TicketReglementInterfaceChange (Index)
                ' Je me rapelle que j'ai donner la main a cette zone
                Set Zone = LIGTicketReglementForm.LIGTicketReglementMaskEdBox(I)
                ' Je rend la saisie possible dans cette zone
                Zone.Enabled = True
                ' Je lui donne le focus
                Zone.SetFocus
                ' Je change la couleur du fond de la zone de saisie
                Zone.BackColor = vbYellow
            Case Else
                ' C'est une autre zone de saissie
                ' Je rend la saissie impossible dans cet zone
                LIGTicketReglementForm.LIGTicketReglementMaskEdBox(I).Enabled = False
                 ' Je change la couleur du fond de la zone de saisie
                LIGTicketReglementForm.LIGTicketReglementMaskEdBox(I).BackColor = vbWhite
        End Select
    ' Je ne sort que si j'ai fini de faire defiler les zones
    Next I
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:TicketReglementTypeReglementChange(TypeReglementCar       ***
'***              As String)                                                ***
'******************************************************************************
'***    FONCTION: Changer le type de reglement                              ***
'******************************************************************************
'***    ENTREE:   TypeReglementCar : repere le Type de Reglement            ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketReglementTypeReglementChange("C")                   ***
'******************************************************************************
Public Sub TicketReglementTypeReglementChange(TypeReglementCar As String)
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DAEVenteDSR
        ' Je selectionne le type reglement dans le table des type de reglement
        .rsSETypeReglementCmd.Filter = "[TYPEREGLEMENTCAR]='" + TypeReglementCar + "'"
        ' Je regarde si il existe des enregistrement pour ce type de mouvement
        Select Case .rsSETypeReglementCmd.RecordCount
            Case 0
                ' Il n'y a pas de type de reglement ou le reglement n'exsite pas
                MsgBox "Probleme Type Reglement"
            Case Else
                ' Il y a des enregistrement
        End Select
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:TicketReglementValidation()                               ***
'******************************************************************************
'***    FONCTION: Pour valider les reglements de la ligne du ticket courant ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketReglementValidation()                               ***
'******************************************************************************
Public Sub TicketReglementValidation()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DAEVenteDSR
        ' Je regarde si un reglement a ete saissie
        If .rsSETicketReglementCmd.RecordCount Then
            'des reglements on ete saisie
            'je regarde le reste a encaisser
            If .rsSETicketReglementTotalDevise1Cmd("TICKETREGLEMENTRESTEAPAYER").Value > 0 Then
                'je met le reste a payer dans la derniere ligne du ticket
                .rsINTicketReglementCmd("TICKETREGLEMENTPRIX").Value = .rsSETicketReglementTotalDevise1Cmd("TICKETREGLEMENTRESTEAPAYER").Value
                'je valide la ligne du ticket en cours
                TicketReglementAjoute
            End If '.rsSETicketReglementTotalCmd("TicketReglementResteAPayer").Value > 0
        Else
            'aucun reglement saisie
            'je met le total net dans la derniere ligne du ticket
            .rsINTicketReglementCmd("TICKETREGLEMENTPRIX").Value = .rsSelTicketDonneTotalDevise1Cmd("TICKETDONNETOTALNET").Value
            'je valide la ligne du ticket en cours
            TicketReglementAjoute
        End If
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Je passe a l'edition du ticket
    EditionTicket
    ' Je passe la saissie des ventes
    MenuSaisieDesVentes
End Sub

'******************************************************************************
'***    PROCEDURE:TicketReglementLigneValidation()                          ***
'******************************************************************************
'***    FONCTION: Pour valider la ligne courante du ticket                  ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketReglementLigneValidation()                          ***
'******************************************************************************
Public Sub TicketReglementLigneValidation()
    ' Je regarde si je suis bien sur la derniere zone ligne de la ligne du ticket
    Select Case Zone.Index
        Case ZoneValeur
            ' Je suis sur la zone reglement valeur
            ' Je regarde si la valeur est bonne
            Select Case ZoneValeurValidation
                Case True
                    ' La valeur est bonne
                    ' Je met a jour la ligne des reglements dans la table ticket reglements
                    TicketReglementAjoute
                    ' Je creer une nouvelle ligne dans la table ticket reglements
                    TicketReglementNouvelle
                    ' J'autorise la saissie sur la valeur du reglements
                    TicketReglementZoneChange (ZoneValeur)
                Case False
                    ' La valeur est mauvaise
            End Select
    End Select
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:TicketReglementZoneValidation()                           ***
'******************************************************************************
'***    FONCTION: Pour valider une zone de la ligne ticket reglement        ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketReglementZoneValidation()                           ***
'******************************************************************************
Public Sub TicketReglementZoneValidation()
    ' Je regarde sur quelle zone j'etait
    Select Case Zone.Index
        Case ZoneValeur
            ' Je suis en zone valeur du reglement
            ' Je regarde si la valeur du reglement est valide
            Select Case ZoneValeurValidation
                Case True
                    ' La valeur du reglement est valide
                Case False
                    ' Le valeur du reglement est invalide
            End Select
     End Select
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:ZoneValeurValidation() As Boolean                         ***
'******************************************************************************
'***    FONCTION: Pour verifier que la valeur du reglement est bon          ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   ValidationZonePrix - indique que la zone                  ***
'**               prix est valide                                           ***
'******************************************************************************
'***    EXEMPLE: =ZoneValeurValidation()                                    ***
'******************************************************************************
Public Function ZoneValeurValidation() As Boolean
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DAEVenteDSR
        ' Je verifie que la valeur saisie est superieur a 0
        Select Case Val(Zone.Text)
            Case Is > 0
                ' La valeur saisie est positive
                ' Je regarde si le reste a payer est nul
                Select Case .rsSETicketReglementTotalDevise1Cmd.RecordCount
                    Case 0
                        ' La valeur est null aucun reglement saissie
                        ' Je verifie que la valeur saissie est superieur a 0
                        Select Case Val(Zone.Text)
                            Case Is < .rsSelTicketDonneTotalDevise1Cmd("TICKETDONNETOTALNET").Value
                                ' La valeur saissie est inferieur au total net
                                ' Je met dans la ligne du ticket la valeur
                                .rsINTicketReglementCmd("TICKETREGLEMENTPRIX").Value = Val(Zone.Text)
                                ' Je renvoie que le prix est bon
                                ZoneValeurValidation = True
                            Case Else
                                ' La valeur saisie est null ou negative
                                ' Je renvoie que le prix est mauvais
                                ZoneValeurValidation = False
                        End Select
                    Case Else
                        ' La valeur est non null un reglement a ete saissie
                        ' Je verifie que la valeur saissie est superieur a 0
                        Select Case Val(Zone.Text)
                            Case Is < .rsSETicketReglementTotalDevise1Cmd("TICKETREGLEMENTRESTEAPAYER").Value
                                ' La valeur saissie est inferieur au reste a payer
                                ' Je met dans la ligne du ticket la valeur
                                .rsINTicketReglementCmd("TICKETREGLEMENTPRIX").Value = Val(Zone.Text)
                                ' Je renvoie que le prix est bon
                                ZoneValeurValidation = True
                            Case Else
                                ' La valeur saisie est null ou negative
                                ' Je renvoie que le prix est mauvais
                                ZoneValeurValidation = False
                        End Select
                End Select
            Case Else
                ' La valeur saisie est null ou negative
                ' Je renvoie que la valeur est mauvaise
                ZoneValeurValidation = False
        End Select
    ' Je n'ai plus besoin des commande de la base vente
    End With
    ' Fin
End Function



