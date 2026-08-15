Attribute VB_Name = "GLMicroPointVente"
'******************************************************************************
'***    PROGICIEL BOUTISCOPE ( Vers. VB )                                   ***
'******************************************************************************
'***    MODULE Principale Du Programme                                      ***
'******************************************************************************
'***    FONCTION:                                                           ***
'***    Pour gerer les evenements,constantes,variables commun a toutes les  ***
'***    feuilles du programme                                               ***
'***    differentes etapes de la vente                                      ***
'******************************************************************************
    Option Explicit                                                             ' Je doit etre sur que mes variables on ete declarer
'******************************************************************************
'***    Declaration de variable global au MPV                               ***
'******************************************************************************
    Global TicketNumero As Integer                                              ' Pour me rappeller le numero du ticket
    Global ZoneTicketDonne_Index As Integer                                     ' Pour me rappeller la zone de saisie dans la feuille de saisie d'une ligne d'un ticket
    Global ZoneTicketReglement_Index As Integer                                 ' Pour me rappeller la zone de saisie dans la feuille de saisie d'une de reglement
'******************************************************************************
'***    Declaration de constante global du MPV                              ***
'******************************************************************************
'***    Constantes qui identifie la zone de saisie de ligne des reglements  ***
'******************************************************************************
    Global Const ZoneTypeReglementCar = 0                                       ' Zone caractere du reglement
    Global Const ZoneTypeReglementLibelle = 1                                   ' Zone libelle du reglement
    Global Const ZoneTicketReglementValeur = 2                                  ' Zone Valeur du reglement
'******************************************************************************
'***    Constantes qui identifie le bouton des types de reglements          ***
'******************************************************************************
    Global Const ButtonTypeReglementCheque = 0                                  ' Button reglement par cheque
    Global Const ButtonTypeReglementEspece = 1                                  ' Button reglement en espece
    Global Const ButtonTypeReglementCarte = 2                                   ' Button reglement par carte
    Global Const ButtonTypeReglementUtilisationAvoir = 3                        ' Button utilisation d'avoir
    Global Const ButtonTypeReglementDiver1 = 4                                  ' Button diver 1
    Global Const ButtonTypeReglementDiver2 = 5                                  ' Button diver 2
'******************************************************************************
'***    Constantes qui identifie le type de mouvement dans                  ***
'***    la table des type de mouvement                                      ***
'******************************************************************************
    Global Const ReglementChequeCar = "S"                                       ' Reglement par cheque
    Global Const ReglementEspeceCar = "E"                                       ' Reglement par espece
    Global Const ReglementCarteCar = "C"                                        ' Reglement par carte
    Global Const ReglementUtilisationAvoirCar = "U"                             ' Reglement par util avoir
    Global Const ReglementDiver1Car = "T"                                       ' Reglement par div.1
    Global Const ReglementDiver2Car = "O"                                       ' Reglement par div.2
'******************************************************************************
'***    Constante qui identifie la zone d'entete du ticket                  ***
'******************************************************************************
    Global Const ZoneMagasinCar = 3                                             ' Zone Magasin
    Global Const ZoneCaisseCar = 4                                              ' Zone caisse
    Global Const ZoneTicketNumero = 5                                           ' Zone Numero de ticket
    Global Const ZoneMagasinLibelle = 6                                         ' Zone Libelle Magasin
    Global Const ZoneCaisseLibelle = 7                                          ' Zone Libelle Caisse
    Global Const ZoneTicketDate = 8                                             ' Zone Date Du Ticket
    Global Const ZoneTicketDonneHeure = 9                                       ' Zone Heure Du Ticket
'******************************************************************************
'***                         Main()                                         ***
'******************************************************************************
'***    FONCTION: Pour demarrer le programme des ventes                     ***
'******************************************************************************
'***    ENTREE:  neant                                                      ***
'***                                                                        ***
'***    SORTIE:                                                             ***
'******************************************************************************
'*** EXEMPLE: Main()                                                        ***
'******************************************************************************
Sub Main()
    ' j'initilise les variables des chemins et des volumes des fichiers du programme
    InitPath
    ' Je travaille sur les commandes de la base vente
    With DEVente
        ' J'ouvre la table de la liste des type de vente
        .rsSETypeVenteCmd.Open
        ' J'ouvre la table de la liste des vendeurs
        .rsSEVendeurCmd.Open
        ' J'ouvre la table de la liste des articles
        .rsSEArticleCmd.Open
        ' J'ouvre la table de la liste des tailles
        .rsSETailleCmd.Open
        ' J'ouvre la table des donners du tickets
        .rsSETicketDonneTotalCmd.Open
        ' J'ouvre la table des reglements du tickets
        .rsSETicketReglementTotalCmd.Open
    End With
    
    MenuPrincipale
End Sub
   
   
'pour appeller une liste de choix
Public Sub ValidationListe()
    'je regarde si une zone a ete selectionner
    If IsObject(ValidationZone) Then
        'je regarde le nom de la zone sur laquelle
        'etait le l'user
        Select Case ValidationZone.Name
            Case "LGTicketDonneMaskEdBox"
                'je suis sur la feuille de saisie d'une ligne du ticket
                ValidationListeTicketDonne
        End Select
    End If
End Sub
   
'pour valider une ligne de saisie
Public Sub ValidationSaisie()
    'je regarde le nom de la feuille sur laquelle
    'etait l'user
    Select Case ValidationZone.Parent.Name
        Case "LGTicketDonneForm"
            'je suis sur la feuille de saisie d'une ligne d'un ticket
            ValidationTicketDonneLigne
        Case "LGTicketReglementForm"
            'je suis sur la feuille de saisie d'une ligne d'un reglement
            ValidationTicketReglementLigne
    End Select
End Sub
      
'pour valider une zone de saisie
Public Sub ValidationZoneSaisie()
    'je regarde le nom de la zone sur laquelle
    'etait le l'user
    'je regarde si une zone a ete selectionner
    If IsObject(ValidationZone) Then
        Select Case ValidationZone.Name
            Case "LGTicketDonneMaskEdBox"
                'je suis sur la feuille de saisie d'une ligne du ticket
                ValidationTicketDonneZone
        End Select
    End If
End Sub

'pour annuler une zone de saisie
Public Sub AnnulationZoneSaisie()
    'je met le texte de controle a 0
    ValidationZone.Text = ""
End Sub

'pour cloturer la vente
Public Sub ValidationTotal()
    'je regarde le nom de la feuille sur laquelle
    'etait l'user
    Select Case ValidationZone.Parent.Name
        Case "LGTicketDonneForm"
            'je suis sur la feuille de saisie d'une ligne du ticket
            'je valide les donnees du tickets
            ValidationTicketDonne
        Case "LGTicketReglementForm"
            'je suis sur la feuille de saisie d'un reglement
            'je valide les reglements du tickets
            ValidationTicketReglement
    End Select
End Sub

'pour lancer le menu principale
Public Sub MenuPrincipale()
    ' je decharge la feuille du paver de validation
    Unload PVValidationForm
    'je decharge la photo
    Unload PNMediaArticleForm
    'je decharge la feuille du paver total
    Unload PVTotalForm
    'je decharge la feuille du paver alpaha
    Unload PVAlphaForm
    ' je decharge la feuille du paver num
    Unload PVNumForm
    ' je decharge la feuille du paver type de reglement
    Unload PVTypeReglementForm
    ' je decharge la feuille du paver type de reglement
    Unload PVTypeMouvementForm
    ' je decharge la feuille du paver type de reglement
    Unload PVTypePrixForm
    ' je decharge la feuille de saisie d'une ligne du ticket
    Unload LGTicketDonneForm
    ' je decharge la feuille de donne d'un ticket
    Unload LSTicketDonneForm
    ' je decharge la feuille du total d'un ticket
    Unload PVTicketDonneTotalForm
    ' je decharge la feuille de saisie d'un reglement
    Unload LGTicketReglementForm
    ' je decharge la feuille des donnes des reglements
    Unload LSTicketReglementForm
    ' je decharge la feuille du paver total des reglement
    Unload PVTicketReglementTotalForm
    ' j'affiche la feuille du fond du mpv
    PNFondMpvForm.Show
    ' j'affiche la feuille de l'identite du ticket
    PNTicketIdentiteForm.Show vbModeless, PNFondMpvForm
    ' j'affiche la feuille du menu prinbcipale
    PVMenuPrincipaleForm.Show vbModeless, PNFondMpvForm
End Sub

'pour lancer la saisie des ventes
Public Sub MenuSaisieDesVentes()
    With DEVente 'comme d'hab
        ' Je verifie si il a un ticket deja saisie dans le fichier
        Select Case .rsINTicketCmd.RecordCount
            Case 0 ' il n'y a pas de ticket
                ' j'initialise le N° ticket
                TicketNumero = 1
            Case Else
                ' sinon je cherche le dernier ticket
                .rsINTicketCmd.MoveLast
                ' j'incremente le numero du ticket
                TicketNumero = .rsINTicketCmd("TicketNumero").Value + 1
        End Select
        ' j'initialise le nouveau ticket dans la table des tickets
        .rsINTicketCmd.AddNew
        ' je renseigne le N° de la caisse du ticket
        .rsINTicketCmd("CleCaisseNumero").Value = .rsSEMagasinCaisseCmd("CleCaisseNumero").Value
        ' je renseigne la date du ticket
        .rsINTicketCmd("TicketDate").Value = Date
        ' je renseigne le numero du ticket
        .rsINTicketCmd("TicketNumero").Value = TicketNumero
        ' je met a jour la base
        .rsINTicketCmd.Update
        ' je met a jour les liens des commandes avec le ticket en cours
        ChangeCleTicketNumero
        ' je cache le menu a l'user
        PVMenuPrincipaleForm.Hide
        'j'affiche la feuille de la photo
        PNMediaArticleForm.Show vbModeless, PNFondMpvForm
        'j'affiche la feuille des donnees de ticket donne
        LSTicketDonneForm.Show vbModeless, PNFondMpvForm
        'j'affiche la feuille de saisie d'une ligne du ticket
        LGTicketDonneForm.Show vbModeless, PNFondMpvForm
        'j'affiche la feuille du total du ticket
        PVTicketDonneTotalForm.Show vbModeless, PNFondMpvForm
        'j'affiche les feuilles du paver de total
        PVTotalForm.Show vbModeless, PVTicketDonneTotalForm
        'j'affiche la feuille du paver alpha numerique
        PVAlphaForm.Show vbModeless, PVTotalForm
        'j'affiche la feuille du paver clavier numerique
        PVNumForm.Show vbModeless, PNFondMpvForm
        'j'affiche la feuille du clavier
        PVClavierForm.Show vbModeless, PNFondMpvForm
        'j'affiche la feuille du paver de validation
        PVValidationForm.Show vbModeless, PNFondMpvForm
        'j'affiche la feuille du paver du type de mouvement
        PVTypeMouvementForm.Show vbModeless, PNFondMpvForm
        'j'affiche la feuille du paver de type de prix
        PVTypePrixForm.Show vbModeless, PNFondMpvForm
        ' je creer une nouvelle ligne de vente
        TicketDonneNouvelle
        'j'autorise la saisie sur le code vendeur
        ChangeTicketDonneZone (0)
        'je decharge le menu
        Unload PVMenuPrincipaleForm
    End With
End Sub
                
'pour lancer la saisie des reglements
Public Sub MenuSaisieDesReglements()
    'je decharge la feuilles du paver des types des mouvements
    Unload PVTypeMouvementForm
    'je decharge la feuilles du paver des types de prix
    Unload PVTypePrixForm
    'je decharge la feuille de saisie d'une ligne du ticket
    Unload LGTicketDonneForm
    'j'affiche la feuille des donnees de ticket reglement
    LSTicketReglementForm.Show vbModeless, LSTicketDonneForm
    'j'affiche la feuille de saisie d'une ligne de reglement
    LGTicketReglementForm.Show vbModeless, LSTicketReglementForm
    'j'affiche la feuille du paver des total des reglements
    PVTicketReglementTotalForm.Show vbModeless, LGTicketReglementForm
    'j'affiche la feuilles du paver des types de reglements
    PVTypeReglementForm.Show vbModeless, PVTicketReglementTotalForm
    'je creer une nouvelle ligne de reglement
    TicketReglementNouvelle
    'j'autorise la saisie sur la valeur du reglement
    ChangeTicketReglementZone (2)
End Sub
                
'pour lancer l'edition du ticket
Public Sub MenuEditionTicket()
    ' je decharge le paver du bouton total
    Unload PVTotalForm
    'je decharge la ligne de saisie des reglements
    Unload LGTicketReglementForm
    'je charge la feuille du menu secondaire
    PVMenuSecondaireForm.Show vbModeless, LSTicketReglementForm
    
End Sub

'pour mettre a jour les liens des commandes du ticket en cours
Public Sub ChangeCleTicketNumero()
    With DEVente
        'je ne boss que si le fichier n'est pas vide
        If .rsINTicketCmd.RecordCount <> 0 Then
            'je regarde si une cle du ticket est presente
            If IsNull(.rsINTicketCmd("CleTicketNumero").Value) = False Then
                'je n'affiche que les donnees de ticket donne concernant le ticket courant
                .rsSETicketDonneCmd.Filter = "[CLETICKETNUMERO]='" + Str(.rsINTicketCmd("CleTicketNumero").Value) + "'"
                'je n'affiche que le total du ticket concernant le ticket courant
                .rsSETicketDonneTotalCmd.Filter = "[CLETICKETNUMERO]='" + Str(.rsINTicketCmd("CleTicketNumero").Value) + "'"
                'je n'affiche que les reglements du ticket concernant le ticket courant
                .rsSETicketReglementCmd.Filter = "[CLETICKETNUMERO]='" + Str(.rsINTicketCmd("CleTicketNumero").Value) + "'"
                'je n'affiche que le total des reglements concernant le ticket courant
                .rsSETicketReglementTotalCmd.Filter = "[CLETICKETNUMERO]='" + Str(.rsINTicketCmd("CleTicketNumero").Value) + "'"
            End If
        End If
    End With
End Sub

