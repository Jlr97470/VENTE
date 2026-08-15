Attribute VB_Name = "BasGblTicketDonne"
'******************************************************************************
'***                                   ***
'******************************************************************************
'***    MODULE POUR LA GESTION DES DONNER DU TICKET                         ***
'******************************************************************************
'***    FONCTION:   - gerer l'affichage de l'interface de saissie           ***
'***                - gerer la validation des zones                         ***
'***                - gerer les selection par liste                         ***
'***                - gerer les modifications des donnes dans la table      ***
'******************************************************************************
'***    PROGRAMMEUR: Royer Jean-Laurent                                     ***
'******************************************************************************
    Option Explicit                                                             ' Je veut etre sur que toute met variable sont declarer

'******************************************************************************
'***    Declaration de constante global du MPV                              ***
'******************************************************************************

'******************************************************************************
'***    Constante qui identifie la zone de saissie d'une ligne de donner du ***
'***    ticket                                                              ***
'******************************************************************************
    Global Const ZoneVendeurCar = 0                                             ' Pour identifier la zone vendeur car
    Global Const ZoneSpecial = 1                                                ' Pour identifier la zone vente special
    Global Const ZoneTypeMouvement = 2                                          ' Pour identifier la zone type de mouvement
    Global Const ZoneArticleReference = 3                                       ' Pour identifier la zone reference article
    Global Const ZoneArticleLibelle = 4                                         ' Pour identifier la zone libelle article
    Global Const ZoneTaille = 5                                                 ' Pour identifier la zone taille
    Global Const ZoneQuantite = 6                                               ' Pour identifier la zone quantite
    Global Const ZoneTypePrix = 7                                               ' Pour identifier la zone type de prix
    Global Const ZonePrix = 8                                                   ' Pour identifier la zone prix
    Global Const ZoneGuelte = 9                                                 ' Pour identifier la zone guelte
    
'******************************************************************************
'***    Constante qui identifie les zones du total des donners ticket       ***
'******************************************************************************
    Global Const ZoneTotalBrut1 = 0                                             ' Pour identifier la zone total brut dans la devise 1
    Global Const ZoneTotalBrutDevise1 = 1                                       ' Pour identifier la zone devise 1 du total brut
    Global Const ZoneTotalBrut2 = 2                                             ' Pour identifier la zone total brut dans la devise 2
    Global Const ZoneTotalBrutDevise2 = 3                                       ' Pour identifier la zone devise 2 du total brut
    Global Const ZoneAPayer = 4                                                 ' Pour identifier la zone a payer
    Global Const ZoneTotalNet1 = 5                                              ' Pour identifier la zone total net
    Global Const ZoneTotalNetDevise1 = 6                                        ' Pour identifier la zone total net devise
    Global Const ZoneTotalNet2 = 7                                              ' Pour identifier la zone total net
    Global Const ZoneTotalNetDevise2 = 8                                        ' Pour identifier la zone total net devise
        
'******************************************************************************
'***    Constantes qui identifie le bouton de type de mouvement             ***
'******************************************************************************
    Global Const ButtonSpecial = 0                                              ' Pour identifier le bouton vente special
    Global Const ButtonMouvementY = 1                                           ' Pour identifier le bouton mouvement y
    Global Const ButtonMouvementZ = 2                                           ' Pour identifier le bouton mouvement z
    Global Const ButtonMouvementReservation = 3                                 ' Pour identifier le bouton mouvement reservation
    Global Const ButtonMouvementArrhes = 4                                      ' Pour identifier le bouton mouvement arrhes
    Global Const ButtonMouvementArrhesArticle = 5                               ' Pour identifier le bouton mouvement arrhes article
    Global Const ButtonMouvementRendu = 6                                       ' Pour identifier le bouton mouvement rendu
    Global Const ButtonMouvementVente = 7                                       ' Pour identifier le bouton mouvement vente

'******************************************************************************
'***    Constante qui identifie le bouton type de prix                      ***
'******************************************************************************
    Global Const ButtonPrixNormal = 0                                           ' Pour identifier le bouton prix normal
    Global Const ButtonPrixSolde = 1                                            ' Pour identifier le bouton prix solde
    Global Const ButtonPrixDemarque = 2                                         ' Pour identifier le bouton prix demarque
    Global Const ButtonQuantite = 3                                             ' Pour identifier le bouton quantite
                           
'******************************************************************************
'***    Constantes qui identifie le type de mouvement dans la table         ***
'***    des type de mouvement                                               ***
'******************************************************************************
    Global Const MouvementArrhes = "A"                                          ' Pour identifier le mouvement arrhes
    Global Const MouvementArrhesArticle = "B"                                   ' Pour identifier le mouvement arrhes article
    Global Const MouvementRendu = "C"                                           ' Pour identifier le mouvement rendu
    Global Const MouvementReservation = "D"                                     ' Pour identifier le mouvement reservation
    Global Const MouvementVente = "E"                                           ' Pour identifier le mouvement vente
    Global Const MouvementYPositif = "F"                                        ' Pour identifier le mouvement y positif
    Global Const MouvementZPositif = "G"                                        ' Pour identifier le mouvement z positif
    Global Const MouvementYNegatif = "H"                                        ' Pour identifier le mouvement y negatif
    Global Const MouvementZNegatif = "I"                                        ' Pour identifier le mouvement z negatif
    
'******************************************************************************
'***    Constantes qui identifie le type de prix dans la table la           ***
'***    des type de prix                                                    ***
'******************************************************************************
    Global Const PrixNormal = "N"                                               ' Pour identifier le prix normal
    Global Const PrixSolde = "S"                                                ' Pour identifier le prix solde
    Global Const PrixDemarque = "D"                                             ' Pour identifier le prix demarque
                          
'******************************************************************************
'***    Declaration des procedures global                                   ***
'******************************************************************************
                        
'******************************************************************************
'***    PROCEDURE:TicketDonneAjoute()                                       ***
'******************************************************************************
'***    FONCTION: Pour ajouter une ligne dans la table ticket donne         ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneAjoute()                                       ***
'******************************************************************************
Public Sub TicketDonneAjoute()
    Dim I As Integer                                                            ' Pour memoriser le numero du label courant
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je met dans la ligne du ticket le type de vente
        .rsInsTicketDonneCmd("CLETYPEVENTENUMERO").Value = .rsSelTypeVenteCmd("CLETYPEVENTENUMERO").Value
        ' Je rajoute ma ligne a la table ticket donne
        .rsInsTicketDonneCmd.Update
        ' Je met a jour la commande de mise en form des donners du ticket
        .rsSelTicketDonneCmd.Requery
        ' Je met a jour les donnees du ticket affichee a l'ecran
        Set FormLstTicketDonne.LSTTicketDonneDataGrid.DataSource = DsrDatVente
        ' Je met a jour la commande de mise en form du total des donners du tickets de la devise 1
        .rsSelTicketDonneTotalDevise1Cmd.Requery
        ' Je met a jour la commande de mise en form du total des donners du tickets de la devise 2
        .rsSelTicketDonneTotalDevise2Cmd.Requery
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Je vais beaucoup utiliser les proprietes des labels du total du tickets
    With FormPanTicketDonneTotal.PANTicketDonneTotalLabel
        ' Je fait defiler tous les labels de la form ticket donne total
        For I = 0 To .Count - 1
            ' Je regarde si le labels et lier a une commande
            Select Case .Item(I).DataMember
                Case ""
                    ' Je label n'est lier a une commande
                Case Else
                    ' Je label est lier a une commande
                    ' Je met a jour l'affiche du total du ticket
                    Set .Item(I).DataSource = DsrDatVente
            End Select
        ' Je ne sort que si j'ai fini de faire defiler les boutons
        Next I
    ' Je n'ai plus besoin des proprites des labels du total du tickets
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:TicketDonneSupprime()                                     ***
'******************************************************************************
'***    FONCTION: Pour supprimer une ligne dans la table ticket donne       ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneSupprime()                                     ***
'******************************************************************************
Public Sub TicketDonneSupprime()
End Sub

'******************************************************************************
'***    PROCEDURE:TicketDonneEdite()                                        ***
'******************************************************************************
'***    FONCTION: Pour editer une ligne dans la table ticket donne          ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneEdite()                                        ***
'******************************************************************************
Public Sub TicketDonneEdite()
End Sub

'******************************************************************************
'***    PROCEDURE:TicketDonneNouvelle()                                     ***
'******************************************************************************
'***    FONCTION: Pour ajouter une nouvelle ligne dans la table ticket donne***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneNouvelle()                                     ***
'******************************************************************************
Public Sub TicketDonneNouvelle()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je creer une nouvelle ligne dans la commande ticket donne
        .rsInsTicketDonneCmd.AddNew
        ' Je me a jour le lien entre la commande ticket et ticket donne
        .rsInsTicketDonneCmd("CLETICKETNUMERO").Value = .rsInsTicketCmd("CLETICKETNUMERO").Value
        ' Par default ce une vente normal
        TicketDonneTypeMouvementChange (MouvementVente)
    ' Je n'est plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub
        
'******************************************************************************
'***    PROCEDURE:TicketDonneAnnule()                                       ***
'******************************************************************************
'***    FONCTION: Pour Annuler une ligne dans la table ticket donne         ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneAnnule()                                       ***
'******************************************************************************
Public Sub TicketDonneAnnule()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' J'annule la ligne
        .rsInsTicketDonneCmd.CancelUpdate
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub
        
'******************************************************************************
'***    PROCEDURE:TicketDonneInterfaceChange(Zone As Integer)               ***
'******************************************************************************
'***    FONCTION: Changer l'interface utilisateur selon la zone ou on est   ***
'******************************************************************************
'***    ENTREE:   Zone : repere le numero de la zone ou on est              ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneInterfaceChange(0)                             ***
'******************************************************************************
Public Sub TicketDonneInterfaceChange(Zone As Integer)
    ' J'indique a l'user sur quel zone on est
    FormPavTactileLabel.PAVTactileLabelLabel = RessourceStringLoad(930 + Zone)
    ' Je regarde sur quel zone je suis
    Select Case Zone
        Case ZoneVendeurCar
            ' Je suis sur la zone code vendeur
            ' Je desactive le paver du total
            FormPavTactileTotal.Enabled = False
            ' J'active le paver apha numerique
            FormPavTactileAlpha.Enabled = True
            ' J'active le paver numerique
            FormPavTactileNum.Enabled = True
            ' Je desactive le paver des mouvements
            FormPavTypeMouvement.Enabled = False
            ' Je desactive le paver du type de prix
            FormPavTypePrix.Enabled = False
        Case ZoneArticleReference
            ' Je suis sur la Zone reference article
            ' Je desactive le paver du total
            FormPavTactileTotal.Enabled = False
            ' Je desactive le paver alpha numerique
            FormPavTactileAlpha.Enabled = False
            ' J'active le paver numerique
            FormPavTactileNum.Enabled = True
            ' J'active le paver du type des mouvements
            FormPavTypeMouvement.Enabled = True
            ' Je desactive le paver du type de prix
            FormPavTypePrix.Enabled = False
        Case ZoneArticleLibelle
            ' Je suis sur la zone libelle de l'article
            ' Je desactive le paver total
            FormPavTactileTotal.Enabled = False
            ' J'active le paver apha numerique
            FormPavTactileAlpha.Enabled = True
            ' J'active le paver numerique
            FormPavTactileNum.Enabled = True
            ' Je desactive le paver du type de mouvement
            FormPavTypeMouvement.Enabled = False
            ' Je desactive le paver du type de prix
            FormPavTypePrix.Enabled = False
        Case ZoneTaille
            ' Je suis sur la Zone taille de l'article
            ' Je desactive le paver du total
            FormPavTactileTotal.Enabled = False
            ' J'active le paver apha numerique
            FormPavTactileAlpha.Enabled = True
            ' J'active le paver numerique
            FormPavTactileNum.Enabled = True
            ' J'active le paver du type de mouvement
            FormPavTypeMouvement.Enabled = False
            ' Je desactive le paver du type de prix
            FormPavTypePrix.Enabled = False
        Case ZoneQuantite
            ' Je suis sur la Zone quantite vendu
            ' Je desactive le paver total
            FormPavTactileTotal.Enabled = False
            ' Je desactive le paver alpha numerique
            FormPavTactileAlpha.Enabled = False
            ' J'active le paver numerique
            FormPavTactileNum.Enabled = True
            ' Je desactive le paver du type de mouvement
            FormPavTypeMouvement.Enabled = False
            ' Je desactive le paver du type de rpix
            FormPavTypePrix.Enabled = False
        Case ZonePrix
            ' Je suis sur la zone prix
            ' J'active la paver du total
            FormPavTactileTotal.Enabled = True
            ' Je desactive le paver alpha numerique
            FormPavTactileAlpha.Enabled = False
            ' J'active la paver numerique
            FormPavTactileNum.Enabled = True
            ' J'active le paver du type de mouvement
            FormPavTypeMouvement.Enabled = False
            ' J'active le paver du type de prix
            FormPavTypePrix.Enabled = True
    End Select
    ' Fin
End Sub
        
'******************************************************************************
'***    PROCEDURE:TicketDonneZoneChange(Index As Integer)                   ***
'******************************************************************************
'***    FONCTION: Pour autorisée la saissie dans une seuls zone             ***
'***              Indiquer a l'user la zone active                          ***
'******************************************************************************
'***    ENTREE:   Index : repere qui indique la zone a autorisée en saisie  ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneZoneChange(Index As Integer)                   ***
'******************************************************************************
Public Sub TicketDonneZoneChange(Index As Integer)
    Dim I As Integer                                                            ' Pour memoriser le numero de la zone courant
    ' Je fait defiler toutes les zones
    For I = 0 To FormLigTicketDonne.LIGTicketDonneMaskEdBox.Count - 1
        ' Je regarde sur quel zone je suis
        Select Case I
            Case Index
                ' J'ai trouver la zone a saisir
                ' Je change l'interface utilisateur
                TicketDonneInterfaceChange (Index)
                ' Je me rapelle que j'ai donner la main a cette zone
                Set Zone = FormLigTicketDonne.LIGTicketDonneMaskEdBox(Index)
                ' Je rend la saisie possible dans cette zone
                Zone.Enabled = True
                ' Je change la couleur du fond de la zone de saisie
                Zone.BackColor = vbYellow
                ' Je lui donne le focus
                Zone.SetFocus
            Case Else
                ' C'est une autre zone de saissie
                ' Je rend la saissie impossible dans cet zone
                FormLigTicketDonne.LIGTicketDonneMaskEdBox(I).Enabled = False
                ' Je change la couleur du fond de la zone de saisie
                FormLigTicketDonne.LIGTicketDonneMaskEdBox(I).BackColor = vbWhite
        End Select
    ' Je ne sort que si j'ai fini de faire defiler les zones
    Next I
    ' Fin
End Sub
        
'******************************************************************************
'***    PROCEDURE:TicketDonneZoneListe()                                        ***
'******************************************************************************
'***    FONCTION: Pour afficher une liste de choix de la zone en cours      ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneZoneListe()                                        ***
'******************************************************************************
Public Sub TicketDonneZoneListe()
    ' Je regarde la zone sur laquelle etait le l'user
    Select Case Zone.Index
        Case ZoneVendeurCar
            ' Je suis sur la zone vendeur
            ' J'affiche la liste des vendeurs
            ZoneVendeurListe
        Case ZoneSpecial
            ' Je suis sur la zone special
        Case ZoneTypeMouvement
            ' Je suis sur la zone mouvement
        Case ZoneArticleReference
            ' Je suis sur la zone reference article
            ' J'affiche la liste des articles
            ZoneArticleReferenceListe
        Case ZoneArticleLibelle
            ' Je suis sur la zone libelle article
        Case ZoneTaille
            ' Je suis sur la zone taille
            ' J'affiche la liste des tailles
            ZoneTailleListe
        Case ZoneQuantite
            ' Je suis sur la zone quantite
        Case ZoneTypePrix
            ' Je suis sur la zone type de prix
        Case ZonePrix
            ' Je suis sur la zone prix
    End Select
    ' Fin
End Sub
'******************************************************************************
'***    PROCEDURE:TicketDonneZoneValidation()                               ***
'******************************************************************************
'***    FONCTION: Pour valider une zone de la ligne ticket donne            ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneZoneValidation()                               ***
'******************************************************************************
Public Sub TicketDonneZoneValidation()
    ' Je regarde sur quelle zone j'etait
    Select Case Zone.Index
        Case ZoneVendeurCar
            ' Je suis en zone vendeur
            ' Je regarde si le code vendeur est valide
            Select Case ZoneVendeurValidation
                Case True
                    ' Le code vendeur est bon
                    ' Je passe a la zone reference
                    TicketDonneZoneChange (ZoneArticleReference)
                Case False
                    ' Le code vendeur est mauvais
            End Select
        Case ZoneSpecial
            ' Je suis en zone vente special
        Case ZoneTypeMouvement
            ' Je suis en zone type mouvement
        Case ZoneArticleReference
            ' Je suis en zone reference
            ' Je regarde si le reference de l'article est valide
            Select Case ZoneArticleReferenceValidation
                Case True
                    ' La reference est bonne
                    ' Je passe a la zone taille
                    TicketDonneZoneChange (ZoneTaille)
                Case False
                    ' La reference est mauvaise
            End Select
        Case ZoneArticleLibelle
            ' Je suis en zone article
            ' Je regarde si le libelle est correcte
            Select Case ZoneArticleLibelleValidation
                Case True
                    ' Le libelle est correcte
                    ' Je passe a la zone prix
                    TicketDonneZoneChange (ZonePrix)
                Case False
                    ' Le libelle est faux
            End Select
        Case ZoneTaille
            ' Je suis en zone taille
            ' Je regarde si la taille de l'article est valide
            Select Case ZoneTailleValidation
                Case True
                    ' La taille est bonne
                    ' Je guide l'user en lui impossent la saissie
                    TicketDonneZoneChange (ZonePrix)
                Case False
                    ' La taille est incorrecte
            End Select
        Case ZonePrix
            ' Je suis en zone prix
            Select Case ZonePrixValidation
                Case True
                    ' Le prix est bon
                Case False
                    ' Le prix est mauvais
            End Select
    End Select
    ' Fin
End Sub
        
'******************************************************************************
'***    PROCEDURE:TicketDonneTypeMouvementChange(TypeMouvementCar As String)***
'******************************************************************************
'***    FONCTION: Changer le type de mouvement                              ***
'******************************************************************************
'***    ENTREE:   TypeMouvementCar : repere le type de mouvement            ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneTypeMouvementChange("C")                       ***
'******************************************************************************
Public Sub TicketDonneTypeMouvementChange(TypeMouvementCar As String)
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je selectionne le type mouvement dans le table des type de vente
        .rsSelTypeVenteCmd.Filter = "[TYPEMOUVEMENTCAR]='" + TypeMouvementCar + "'"
        ' Je regarde si il existe des enregistrement pour ce type de mouvement
        Select Case .rsSelTypeVenteCmd.RecordCount
            Case 0
                ' Il n'y a pas de type de prix ou le mouvement n'exsite pas
                MsgBox "Probleme Type Mouvement"
            Case Else
                ' Il y a des enregistrement
                ' Je me mais en type de prix normal
                TicketDonneTypePrixChange (PrixNormal)
                ' Je regarde s'il faut saisir une reference article
                Select Case .rsSelTypeVenteCmd("TYPEMOUVEMENTARTICLEREFERENCE")
                    Case True
                        ' Il faut saissir une reference article
                        ' Je passe a la zone article reference
                        TicketDonneZoneChange (ZoneArticleReference)
                    Case False
                        ' Il ne faut pas saisir une reference article
                        ' Je mais une chaine vide dans la reference article
                        .rsInsTicketDonneCmd("TICKETDONNEARTICLEREFERENCE") = ""
                        ' Je passe a la zone article libelle
                        TicketDonneZoneChange (ZoneArticleLibelle)
                End Select
        End Select
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:TicketDonneTypePrixChange(TypePrixCar As String)          ***
'******************************************************************************
'***    FONCTION: Changer le type de prix                                   ***
'******************************************************************************
'***    ENTREE:   TypePrixCar : repere le type de prix                      ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneTypePrixChange("N")                            ***
'******************************************************************************
Public Sub TicketDonneTypePrixChange(TypePrixCar As String)
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je me place au debut de la table
        .rsSelTypeVenteCmd.MoveFirst
        ' Je cherche l'enregistrement concerner
        .rsSelTypeVenteCmd.Find "[TYPEPRIXCAR]='" + TypePrixCar + "'"
        ' Je regarde si j'ai trouver l'enregistrement concerner
        Select Case .rsSelTypeVenteCmd.EOF
            Case True
                ' Je ne l'ai pas trouver
            Case False
                ' Je l'ai trouver
        End Select
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub
        
'******************************************************************************
'***    PROCEDURE:TicketDonneSpecialChange()                                ***
'******************************************************************************
'***    FONCTION: Changer le flag de vente special                          ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneSpecialChange()                                ***
'******************************************************************************
Public Sub TicketDonneSpecialChange()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je met a jour la ligne du ticket
        .rsInsTicketDonneCmd.Fields("TICKETDONNESPECIAL").Value = Not .rsInsTicketDonneCmd("TICKETDONNESPECIAL").Value
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub
       
'******************************************************************************
'***    PROCEDURE:TicketDonnePrixChange()                                   ***
'******************************************************************************
'***    FONCTION: Pour changer la valeur du prix si il y a un prix          ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonnePrixChange()                                   ***
'******************************************************************************
Public Sub TicketDonnePrixChange()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je me met au debut de la liste des tarifs
        .rsSelTarifCmd.MoveFirst
        ' Je recherche le prix correspondant a l'article
        ' Je recherche le magasin dans la table tarif
        .rsSelTarifCmd.Find "[MAGASINCAR]='" + .rsSelMagasinCaisseCmd("MAGASINCAR").Value + "'"
        ' Je regarde si je suis a la fin de la table tarif
        Select Case .rsSelTarifCmd.EOF
            Case False
                ' Je ne suis pas a la fin
                ' Je recherche la reference dans la table tarif
                .rsSelTarifCmd.Find "[ARTICLEREFERENCE]='" + .rsInsTicketDonneCmd("TICKETDONNEARTICLEREFERENCE").Value + "'"
                ' Je regarde si je suis a la fin de la table tarif
                Select Case .rsSelTarifCmd.EOF
                    Case False
                        ' Je ne suis pas a la fin
                        ' Je recherche le type de prix
                        .rsSelTarifCmd.Find "[TYPEPRIXCAR]='" + .rsSelTypeVenteCmd("TYPEPRIXCAR").Value + "'"
                        ' Je regarde si je suis a la fin de la table tarif
                        Select Case .rsSelTarifCmd.EOF
                            Case False
                                ' Je ne suis pas a la fin
                                ' Je recherche la taille pour cette article
                                .rsSelTarifCmd.Find "[TAILLELIBELLE]='" + .rsInsTicketDonneCmd("TICKETDONNETAILLELIBELLE").Value + "'"
                                ' Je regarde si je l'est trouver
                                Select Case .rsSelTarifCmd.EOF
                                    Case False
                                        ' Je l'ai trouver
                                        ' Je rajoute le tarif a la ligne du ticket
                                        .rsInsTicketDonneCmd("TICKETDONNEPRIX").Value = .rsSelTarifCmd("TARIFPRIX").Value
                                    Case True
                                        ' Je ne l'ai pas trouver
                                        ' Je l'indique a l'user
                                        MsgBox "Attention !!! Aucun Tarif pour cette taille ?"
                                End Select
                            Case True
                                ' Je suis a la fin
                                ' Je l'indique a l'user
                                MsgBox "Attention !!! Aucun Tarif pour ce type de prix"
                        End Select
                    Case True
                        ' Je suis a la fin
                        ' Je l'indique a l'user
                        MsgBox "Attention !!! Aucun Tarif pour cette article"
                End Select
            Case True
                ' Je suis a la fin
                ' Je l'indique a l'user
                MsgBox "Attention !!! Aucun Tarif pour se magasin"
        End Select
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:TicketDonneValidation()                                   ***
'******************************************************************************
'***    FONCTION: Pour valider les donnees de la ligne du ticket courant    ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneValidation()                                   ***
'******************************************************************************
Public Sub TicketDonneValidation()
    ' Je regarde si le prix est correcte
    Select Case ZonePrixValidation
        Case True
            ' Le prix est correcte
            ' Je valide la ligne du ticket en cours
            TicketDonneAjoute
            ' Je passe au reglement
            MenuSaisieDesReglements
        Case False
            ' Le prix est incorrecte
    End Select
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:TicketDonneLigneValidation()                              ***
'******************************************************************************
'***    FONCTION: Pour valider la saissie d'une ligne du ticket             ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneLigneValidation()                              ***
'******************************************************************************
Public Sub TicketDonneLigneValidation()
    ' Je regarde si je suis bien sur la derniere zone de la ligne du ticket
    Select Case Zone.Index
        Case ZonePrix
            ' Je suis sur la zone prix
            ' Je regarde si le prix est bon
            Select Case ZonePrixValidation
                Case True
                    ' Le prix est bon
                    ' J'ajoute la ligne a la table ticket donne
                    TicketDonneAjoute
                    ' Je creer une nouvelle ligne dans la table ticket donne
                    TicketDonneNouvelle
                    ' J'autorise la saissie sur le code vendeur
                    TicketDonneZoneChange (ZoneVendeurCar)
                Case False
                    ' Le prix est mauvais
                    MsgBox "Prix Incorrecte"
            End Select
        Case Else
            ' Je suis sur une autre zone de la ligne du ticket
    End Select
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:DouchetteValidation()                                     ***
'******************************************************************************
'***    FONCTION: Pour valider les informations de la douchette             ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  DouchetteValidation()                                     ***
'******************************************************************************
Public Sub DouchetteValidation()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je met a jour le code vendeur de la ligne en cour
        ' Je recupere la taille
        .rsInsTicketDonneCmd("TICKETDONNETAILLELIBELLE").Value = Mid$(Zone.Text, 7, 3)
        ' Je recupere la reference
        .rsInsTicketDonneCmd("TICKETDONNEARTICLEREFERENCE").Value = Left$(Zone.Text, 6)
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:ZoneVendeurListeValidation()                                  ***
'******************************************************************************
'***    FONCTION: Pour valider le vendeur choisi par l'user dans la liste   ***
'***              des vendeurs                                              ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ZoneVendeurListeValidation()                                  ***
'******************************************************************************
Public Sub ZoneVendeurListeValidation()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je met a jour le code vendeur de la ligne en cour
        .rsInsTicketDonneCmd("TICKETDONNEVENDEURCAR") = .rsSelVendeurCmd("VENDEURCAR").Value
        ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:ZoneArticleReferenceListeValidation()                     ***
'******************************************************************************
'***    FONCTION: Pour valider le vendeur choisi par l'user dans la liste   ***
'***              des vendeurs                                              ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ZoneArticleReferenceListeValidation()                     ***
'******************************************************************************
Public Sub ZoneArticleReferenceListeValidation()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je met a jour le la reference de l'article de la ligne ne cour
        .rsInsTicketDonneCmd("TICKETDONNEARTICLEREFERENCE") = .rsSelArticleCmd("ARTICLEREFERENCE").Value
        ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:ZoneTailleListeValidation()                                   ***
'******************************************************************************
'***    FONCTION: Pour valider la taille choisi par l'user dans la liste des***
'***              Tailles                                                   ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ZoneTailleListeValidation()                                   ***
'******************************************************************************
Public Sub ZoneTailleListeValidation()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je met a jour la taille  de la ligne en cour
        .rsInsTicketDonneCmd("TICKETDONNETAILLELIBELLE") = .rsSelTailleCmd("TAILLELIBELLE").Value
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:ZoneVendeurValidation() As Boolean                        ***
'******************************************************************************
'***    FONCTION: Pour verifier que la zone vendeur est bonne               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   ZoneVendeurValidation - indique si la zone vendeur est    ***
'***              Valide                                                    ***
'******************************************************************************
'***    EXEMPLE:  =ValidationZoneVendeur                                    ***
'******************************************************************************
Public Function ZoneVendeurValidation() As Boolean
    ' Je vais beaucoup utiliser la base vente
    With DsrDatVente
        ' Je verifie que le code est bien sur deux caracteres
        Select Case Len(Zone.Text)
            Case 2
                ' Le code vendeur est bien sur deux caracteres
                ' Je verifie qu'il y a des vendeurs dans la table des vendeurs
                Select Case .rsSelVendeurCmd.RecordCount
                    Case 0
                        ' Il n'y a pas de vendeur dans la table des vendeurs
                        ' Je previent l'user
                        MsgBox "Attention !!! Aucun code vendeur n'a ete declarer ?"
                    Case Else
                        ' Je me place au debut de la liste
                        .rsSelVendeurCmd.MoveFirst
                        ' Je recherche le code vendeur qui a ete saisie
                        .rsSelVendeurCmd.Find "[VENDEURCAR]='" + Zone.Text + "'"
                        ' Je regarde si je l'ai trouver
                        Select Case .rsSelVendeurCmd.EOF
                            Case True
                                ' Je ne l'ai pas trouver
                                ' Je previent l'user
                                MsgBox "Attention !! Code vendeur non declarer ?"
                            Case False
                                ' Je l'ai trouver
                        End Select
                End Select
                ' Je met dans la ligne du ticket la valeur du vendeur
                .rsInsTicketDonneCmd("TICKETDONNEVENDEURCAR").Value = Zone.Text
                ' Je renvoie que le code vendeur est bon
                ZoneVendeurValidation = True
            Case Else
                ' Le code vendeur n'est pas sur deux caracteres
                ' Je renvoie que le code vendeur est faux
                ZoneVendeurValidation = False
        End Select
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Function

'******************************************************************************
'***    PROCEDURE:ZoneArticleReferenceValidation() As Boolean               ***
'******************************************************************************
'***    FONCTION: Pour verifier que la zone reference est bonne             ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   ZoneArticleReferenceValidation - indique que la zone      ***
'**               Reference article est valide                              ***
'******************************************************************************
'***    EXEMPLE: =ZoneArticleReferenceValidation()                          ***
'******************************************************************************
Public Function ZoneArticleReferenceValidation() As Boolean
    ' Je vais baucoup utiliser la base vente
    With DsrDatVente
        ' Je verifie que la reference est bien sur six caracteres
        Select Case Len(Zone.Text)
            Case 6
                ' La reference est bien sur 6 caracteres
                ' Je verifie qu'il y a des references dans la table des articles
                Select Case .rsSelArticleCmd.RecordCount
                    Case 0
                        ' Il n'y a pas de reference dans la table des articles
                        ' Je met a jour le libelle de l'article de la ligne du ticket
                        .rsInsTicketDonneCmd.Fields("TICKETDONNEARTICLELIBELLE").Value = "REFERENCE INCONNUE"
                        ' Je previent l'user
                        MsgBox "Attention !!! Aucun Article n'a ete declarer ?"
                    Case Else
                        ' Je me place au debut de la liste des articles
                        .rsSelArticleCmd.MoveFirst
                        ' Je recherche la reference qui a ete saisie
                        .rsSelArticleCmd.Find "[ARTICLEREFERENCE]='" + Zone.Text + "'"
                        ' Je regarde si je l'ai trouver
                        Select Case .rsSelArticleCmd.EOF
                            Case True
                                ' Je ne l'ai pas trouver
                                ' Je met a jour le libelle de l'article de la ligne du ticket
                                .rsInsTicketDonneCmd.Fields("TICKETDONNEARTICLELIBELLE").Value = "REFERENCE INCONNUE"
                                ' Je previent l'user
                                MsgBox "Attention !!! Article inconnue"
                            Case False
                                ' Je l'ai  trouver
                                ' Je met a jour le libelle de l'article de la ligne du ticket
                                .rsInsTicketDonneCmd.Fields("TICKETDONNEARTICLELIBELLE").Value = .rsSelArticleCmd.Fields("ARTICLELIBELLE").Value
                        End Select
                End Select
                ' Je me positionne dans le repertoire des media de la reference en cours
                MediaPathChange (Zone.Text)
                ' Je met a jour les liens des commandes
                ReferenceChange
                ' Je met dans la ligne du ticket la valeur de la reference article
                .rsInsTicketDonneCmd("TICKETDONNEARTICLEREFERENCE").Value = Zone.Text
                ' Je renvoie que la reference est vrai
                ZoneArticleReferenceValidation = True
            Case Else
                ' La reference n'est pas sur six caracteres
                ' Je renvoie que la reference est faux
                ZoneArticleReferenceValidation = False
        End Select
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Function

'******************************************************************************
'***    PROCEDURE:ZoneArticleLibelleValidation() As Boolean                 ***
'******************************************************************************
'***    FONCTION: Pour verifier que la zone libelle est bonne               ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   ZoneArticleLibelleValidation - indique que la zone        ***
'**               libelle est valide                                        ***
'******************************************************************************
'***    EXEMPLE: =ZoneArticleLibelleValidation()                            ***
'******************************************************************************
Public Function ZoneArticleLibelleValidation() As Boolean
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je verifie que la taille a bien un caracteres
        Select Case Len(Zone.Text)
            Case Is > 0
                ' Le libelle a au moins un caractere
                ' Je met la zone quantite a 1
                .rsInsTicketDonneCmd("TICKETDONNEQUANTITE").Value = "1"
                ' Je met une chaine vide dans la taille
                .rsInsTicketDonneCmd("TICKETDONNETAILLELIBELLE").Value = ""
                ' Je met dans la ligne du ticket la valeur du libelle de l'article
                .rsInsTicketDonneCmd("TICKETDONNEARTICLELIBELLE").Value = Zone.Text
                ' Je renvoie que le libelle est bonne
                ZoneArticleLibelleValidation = True
            Case Else
                ' Le libelle ne comporte pas au moins un caractere
                ' Je renvoie que le libelle est faux
                ZoneArticleLibelleValidation = False
        End Select
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Function

'******************************************************************************
'***    PROCEDURE:ZoneTailleValidation() As Boolean                         ***
'******************************************************************************
'***    FONCTION: Pour verifier que la zone taille est bonne                ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   ZoneTailleValidation - indique que la zone                ***
'**               taille est valide                                         ***
'******************************************************************************
'***    EXEMPLE: =ZoneTailleValidation()                                    ***
'******************************************************************************
Public Function ZoneTailleValidation() As Boolean
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je verifie que la taille a bien un caracteres
        Select Case Len(Zone.Text)
            Case Is > 0
                ' La taille a au moins un caractere
                ' Je verifie qu'il existe une taille pour la reference en cours
                Select Case .rsSelTailleCmd.RecordCount
                    Case 0
                        ' Il n'y a pas de taille pour cette article
                        ' Je previent l'user
                        MsgBox "Attention !!!! Aucunne taille pour cette article ?"
                    Case Else
                        ' Je me place au debut de la liste des tailles
                        .rsSelTailleCmd.MoveFirst
                        ' Je recherche la taille qui a ete saisie
                        .rsSelTailleCmd.Find "[TAILLELIBELLE]='" + Zone.Text + "'"
                        ' Je regarde si je l'ai trouver
                        Select Case .rsSelTailleCmd.EOF
                            Case True
                                ' Je ne l'ai pas trouver
                                ' Je previent l'user
                                MsgBox "Attention !!!! Taille inconnu pour cette article"
                            Case False
                                ' Je l'ai  trouver
                        End Select
                End Select
                ' Je met la zone quantite a 1
                .rsInsTicketDonneCmd("TICKETDONNEQUANTITE").Value = "1"
                ' Je change le prix de l'article
                TicketDonnePrixChange
                ' Je met dans la ligne du ticket la valeur de la taille
                .rsInsTicketDonneCmd("TICKETDONNETAILLELIBELLE").Value = Zone.Text
                ' Je renvoie que la taille est bonne
                ZoneTailleValidation = True
            Case Else
                ' La taille ne comporte pas au moins un caractere
                ' Je renvoie que la taille est faux
                ZoneTailleValidation = False
        End Select
    ' Je n'ai plus besoin des commande de la base vente
    End With
    ' Fin
End Function

'******************************************************************************
'***    PROCEDURE:ZonePrixValidation() As Boolean                           ***
'******************************************************************************
'***    FONCTION: Pour verifier que la zone prix est bonne                  ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   ValidationZonePrix - indique que la zone                  ***
'**               prix est valide                                           ***
'******************************************************************************
'***    EXEMPLE: =ZonePrixValidation()                                      ***
'******************************************************************************
Public Function ZonePrixValidation() As Boolean
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je verifie que la valeur saisie est superieur a 0
        Select Case Val(Zone.Text)
            Case Is > 0
                ' La valeur saisie est positive
                ' Je regarde dans quel type de mouvement je suis
                Select Case .rsSelTypeVenteCmd("TYPEMOUVEMENTTARIFPRIX")
                    Case True
                        ' C'est un mouvement avec tarif
                        ' Je me met au debut de la table
                        .rsSelTarifCmd.MoveFirst
                        ' Je recherche le magasin dans la table tarif
                        .rsSelTarifCmd.Find "[MAGASINCAR]='" + .rsSelMagasinCaisseCmd("MAGASINCAR").Value + "'"
                        ' Je recherche la reference dans la table tarif
                        .rsSelTarifCmd.Find "[ARTICLEREFERENCE]='" + .rsInsTicketDonneCmd("TICKETDONNEARTICLEREFERENCE").Value + "'"
                        ' Je recherche le type de prix
                        .rsSelTarifCmd.Find "[TYPEPRIXCAR]='" + .rsSelTypeVenteCmd("TYPEPRIXCAR").Value + "'"
                        ' Je recherche la taille pour cette article
                        .rsSelTarifCmd.Find "[TAILLELIBELLE]='" + .rsInsTicketDonneCmd("TICKETDONNETAILLELIBELLE").Value + "'"
                        ' Je regarde si il y a un prix
                        Select Case .rsSelTarifCmd.EOF
                            Case True
                                ' Je n'ai pas trouver de prix
                            Case False
                                ' J'ai trouver un prix
                                ' Je verifie le prix saisie avec le prix trouver
                                Select Case .rsSelTarifCmd("TARIFPRIX").Value
                                    Case Val(Zone.Text)
                                        ' Le prix saisie correspond au prix du tarif
                                    Case Else
                                        ' Le prix saisie est different du prix du tarif
                                        ' Je  previent l'user qu'il a saisie un prix different du tarif
                                        MsgBox " Attention !!! Le prix que vous avez saisie n'est pas le meme que le prix du tarif"
                                End Select
                        End Select
                    Case False
                        ' C'est un mouvement sans tarif
                End Select
                ' Je met dans la ligne du ticket la valeur du vendeur
                .rsInsTicketDonneCmd("TICKETDONNEPRIX").Value = Val(Zone.Text)
                ' Je renvoie que le prix est bon
                ZonePrixValidation = True
            Case Else
                ' La valeur saisie est null ou negative
                ' Je renvoie que le prix est mauvais
                ZonePrixValidation = False
        End Select
    ' Je n'ai plus besoin des commande de la base vente
    End With
    ' Fin
End Function

'******************************************************************************
'***    PROCEDURE:TicketDonneLigneAnnulation()                              ***
'******************************************************************************
'***    FONCTION: Pour annuler la saissie d'une ligne du ticket             ***
'******************************************************************************
'***    ENTREE:   neant                                                     ***
'***                                                                        ***
'***    SORTIE:   neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketDonneLigneAnnulation()                              ***
'******************************************************************************
Public Sub TicketDonneLigneAnnulation()
    ' J'annule la ligne en cours dans la commande d'insertion d'une ligne de ticket
    TicketDonneAnnule
    ' Je creer une nouvelle ligne dans la commande d'insertion d'une ligne de ticket
    TicketDonneNouvelle
    ' J'autorise la saissie sur le code vendeur
    TicketDonneZoneChange (ZoneVendeurCar)
    ' Fin
End Sub


'******************************************************************************
'***    PROCEDURE:ZoneVendeurListe()                                            ***
'******************************************************************************
'***    FONCTION: Pour afficher une liste de choix de la zone des vendeurs  ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ZoneVendeurListe()                                            ***
'******************************************************************************
Public Sub ZoneVendeurListe()
    ' Je rend la saissie dans la zone impossible
    Zone.Enabled = False
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je regarde le nombre d'enregistrement dans la table vendeur
        Select Case .rsSelVendeurCmd.RecordCount
            Case Is > 0
                ' Il y a des vendeurs
                ' Je me met au debut de la liste des vendeurs
                .rsSelVendeurCmd.MoveFirst
                ' Je regarde si un code vendeur a ete deja saisie
                Select Case IsNull(.rsInsTicketDonneCmd("TicketDonneVendeurCar").Value)
                    Case False
                        ' Un code vendeur a ete saisie
                        ' Je recherche le vendeur actuel
                        .rsSelVendeurCmd.Find "[VENDEURCAR]='" + Zone.Text + "'"
                        ' Je regarde si j'ai trouver le code vendeur
                        Select Case .rsSelVendeurCmd.EOF
                            Case True
                                ' Je n'est pas trouver le code vendeur
                                ' Je me positionne au debut de la liste
                                .rsSelVendeurCmd.MoveFirst
                            Case False
                                ' Je l'ai trouver
                        End Select
                    Case True
                        ' Aucun code vendeur n'a ete selectionner
                End Select
                ' J'affiche la feuille de la liste des vendeurs
                FormLstVendeur.Show vbModal, FormLstTicketDonne
            Case Else
                ' Il n'y a pas de vendeurs
                ' Je previent l'user
                MsgBox "Attention !!! Aucun vendeur n'a ete declarer ?"
        End Select
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:ZoneArticleReferenceListe()                               ***
'******************************************************************************
'***    FONCTION: Pour afficher une liste de choix de la zone des articles  ***
'******************************************************************************
'***    ENTREE:   Néant                                                     ***
'***                                                                        ***
'***    SORTIE:   Néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ZoneArticleReferenceListe()                               ***
'******************************************************************************
Public Sub ZoneArticleReferenceListe()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je regarde le nombre d'enregistrement dans la table article
        Select Case .rsSelArticleCmd.RecordCount
            Case Is > 0
                ' Il y a des articles
                ' Je me met au debut de la liste des articles
                .rsSelArticleCmd.MoveFirst
                ' Je regarde si une reference a ete saisie
                Select Case IsNull(.rsInsTicketDonneCmd("TICKETDONNEARTICLEREFERENCE").Value)
                    Case False
                        ' Une reference a ete saisie
                        ' Je recherche la reference actuel
                        .rsSelArticleCmd.Find "[ARTICLEREFERENCE]='" + Zone.Text + "'"
                        ' Je regarde si j'ai trouver la reference
                        Select Case .rsSelArticleCmd.EOF
                            Case True
                                ' Je n'est pas trouver la reference
                                ' Je me remet au debut de la liste
                                .rsSelArticleCmd.MoveFirst
                            Case False
                                ' Je l'ai trouver
                        End Select
                    Case True
                        ' Aucunne reference n'a ete trouver
                End Select
                ' J'affiche la feuille de la liste des references
                FormLstArticle.Show vbModal, FormLstTicketDonne
            Case Else
                ' Il n'y a pas d'article
                ' Je previent l'user
                MsgBox "Attention !!! Aucun Article n'a ete declarer ?"
        End Select
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:ZoneTailleListe()                                         ***
'******************************************************************************
'***    FONCTION: Pour afficher une liste de choix d'une taille             ***
'******************************************************************************
'***    ENTREE:   néant                                                     ***
'***                                                                        ***
'***    SORTIE:   néant                                                     ***
'******************************************************************************
'***    EXEMPLE:  ZoneTailleListe()                                         ***
'******************************************************************************
Public Sub ZoneTailleListe()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        ' Je ne boss que si le fichier est non vide
        Select Case .rsSelTailleCmd.RecordCount
            Case 0
                ' Le fichier est vide
                ' Je previent l'user
                MsgBox "Aucunne Taille Pour Cette Article"
            Case Else
                ' Le fichier n'est pas vide
                ' Je me met au debut de la liste des tailles
                .rsSelTailleCmd.MoveFirst
                ' Je regarde si une taille a ete saisie
                Select Case IsNull(.rsInsTicketDonneCmd("TICKETDONNETAILLELIBELLE").Value)
                    Case True
                        ' Aucunne taille a ete saissie
                    Case False
                        ' Une taille a ete saisie
                        ' Je recherche la taille actuel
                        .rsSelTailleCmd.Find "[TAILLELIBELLE]='" + Zone.Text + "'"
                        ' Je regarde si j'ai trouver la taille
                        Select Case .rsSelArticleCmd.EOF
                            Case True
                                ' Je n'est pas trouver la taille je revient au debut de la liste
                                .rsSelTailleCmd.MoveFirst
                            Case False
                                ' J'ai trouver la taille
                        End Select
                End Select
                ' J'affiche la feuille de la liste des tailles
                FormLstTaille.Show vbModal, FormLstTicketDonne
        End Select
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

