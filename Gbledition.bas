Attribute VB_Name = "GBLEdition"
Option Explicit

'pour lancer l'edition du ticket
Public Sub EditionTicket()
    ' Je vais beaucoup utiliser les commandes de la base vente
    With DAEVenteDSR
        Select Case .rsDRTicketCmd.State
            Case adStateOpen
                .rsDRTicketCmd.Requery
            Case Else
        End Select
    End With
    ' Je vais beaucoup utiliser les commandes du rapport du ticket
    With DRTicketDonne
        ' J'imprimme la parti donner du ticket
        Debug.Print DAEVenteDSR.rsDRTicketCmd.State
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
    With DRTicketReglement
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
