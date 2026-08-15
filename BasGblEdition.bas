Attribute VB_Name = "BasGblEdition"
Option Explicit

'pour lancer l'edition du ticket
Public Sub EditionTicket()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DsrDatVente
        Select Case .rsDarTicketCmd.State
            Case adStateOpen
                .rsDarTicketCmd.Requery
            Case Else
        End Select
    End With
    ' Je vais beaucoup utiliser les commandes du rapport du ticket
    With DsrDarTicketDonne
        ' J'imprimme la parti donner du ticket
        Debug.Print DsrDatVente.rsDarTicketCmd.State
        .PrintReport
        ' Je repete tant que mon travail n'est pas en attente
        While .AsyncCount = 0
            DoEvents
        Wend
        ' Je repete tant que je n'ai pas fini d'imprimmer
        While .AsyncCount <> 0
            DoEvents
        Wend
    ' Je n'ai plus besoin des commandes du rapport du ticket
    End With
    ' Je vais beaucoup utiliser les commandes du rapport du ticket
    With DsrDarTicketReglement
        ' J'imprimme la partie reglement du ticket
        .PrintReport
        ' Je repete tant que mon travail n'est pas en attente
        While .AsyncCount = 0
            DoEvents
        Wend
        ' Je repete tant que je n'ai pas fini d'imprimmer
        While .AsyncCount <> 0
            DoEvents
        Wend
    ' Je n'ai plus besoin des commandes du rapport du ticket
    End With
    ' Fin
End Sub
