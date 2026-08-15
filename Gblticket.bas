Attribute VB_Name = "GBLTicket"
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. VB )                                   ***
'******************************************************************************
'***    MODULE POUR LA GESTION DU TICKET                                    ***
'******************************************************************************
'***    FONCTION:  - gerer les modifications des donnes dans la table       ***
'***                                                                        ***
'***                                                                        ***
'***                                                                        ***
'******************************************************************************
    Option Explicit                                                             ' Je doit etre sur que mes variables on ete declarer

'******************************************************************************
'***    Declaration de constante global du MPV                              ***
'******************************************************************************

'******************************************************************************
'***    Constante qui identifie les zones du panneau du ticket              ***
'******************************************************************************
    Global Const ZoneMagasin = 0                                                ' Pour identifier la zone libelle magasin
    Global Const ZoneCaisse = 1                                                 ' Pour identifier la zone libelle caisse
    Global Const ZoneTicket = 2                                                 ' Pour identifier la zone libelle ticket
    Global Const ZoneMagasinCar = 3                                             ' Pour identifier la Zone Magasin car
    Global Const ZoneCaisseCar = 4                                              ' Pour identifier la zone caisse car
    Global Const ZoneTicketNumero = 5                                           ' Pour identifier la zone numero de ticket
    Global Const ZoneMagasinLibelle = 6                                         ' Pour identifier la zone libelle du magasin
    Global Const ZoneCaisseLibelle = 7                                          ' Pour identifier la Zone libelle de la caisse
    Global Const ZoneTicketDate = 8                                             ' Pour identifier la zone date du ticket
    Global Const ZoneTicketDonneHeure = 9                                       ' Pour identifier la zone heure du ticket

'******************************************************************************
'***    Declaration des procedures global                                   ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:TicketAjoute()                                            ***
'******************************************************************************
'***    FONCTION: Pour ajouter un ticket dans la table ticket               ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketAjoute()                                            ***
'******************************************************************************
Public Sub TicketDonneAjoute()
    ' Fin
End Sub

'******************************************************************************
'***    PROCEDURE:TicketSupprime()                                          ***
'******************************************************************************
'***    FONCTION: Pour supprimer un ticket dans la table ticket             ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketSupprime()                                          ***
'******************************************************************************
Public Sub TicketSupprime()
End Sub

'******************************************************************************
'***    PROCEDURE:TicketEdite()                                             ***
'******************************************************************************
'***    FONCTION: Pour editer un ticket dans la table ticket                ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketEdite()                                             ***
'******************************************************************************
Public Sub TicketEdite()
End Sub

'******************************************************************************
'***    PROCEDURE:TicketNouvelle()                                          ***
'******************************************************************************
'***    FONCTION: Pour ajouter un nouveaux ticket dans la table ticket      ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketNouvelle()                                          ***
'******************************************************************************
Public Sub TicketNouvelle()
End Sub
        
'******************************************************************************
'***    PROCEDURE:TicketAnnule()                                            ***
'******************************************************************************
'***    FONCTION: Pour Annuler un ticket dans la table ticket               ***
'******************************************************************************
'***    ENTREE:   Neant                                                     ***
'***                                                                        ***
'***    SORTIE:   Neant                                                     ***
'******************************************************************************
'***    EXEMPLE:  TicketAnnule()                                            ***
'******************************************************************************
Public Sub TicketAnnule()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DAEVenteDSR
        ' J'annule les modification du ticket courant
        .rsINTicketCmd.CancelUpdate
    ' Je n'ai plus besoin des commandes de la base vente
    End With
    ' Fin
End Sub

