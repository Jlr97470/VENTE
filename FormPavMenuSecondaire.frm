VERSION 5.00
Begin VB.Form FormPavMenuSecondaire 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PVMenuSecondaireCDButton 
      Caption         =   "MENU 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   2
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1500
   End
   Begin VB.CommandButton PVMenuSecondaireCDButton 
      Caption         =   "MENU 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   1
      Left            =   1590
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1500
   End
   Begin VB.CommandButton PVMenuSecondaireCDButton 
      Caption         =   "MENU 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   0
      Left            =   0
      MaskColor       =   &H80000010&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
End
Attribute VB_Name = "FormPavMenuSecondaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Fonction d'initialisation Des Objects

Private Sub FormPavMenuSecondaireInit()
    FormPavMenuSecondaire.Move Screen.Width * 2 / 20 + Screen.TwipsPerPixelX, Screen.Height * 5 / 15 + Screen.TwipsPerPixelY, Screen.Width * 15 / 20 - Screen.TwipsPerPixelX * 2, Screen.Height * 5 / 15 - Screen.TwipsPerPixelY * 2
End Sub

Private Sub PVMenuSecondaireCDButtonInit()
    Dim I As Integer
    For I = 0 To PVMenuSecondaireCDButton.Count - 1
        PVMenuSecondaireCDButton(I).Picture = RessourceIconLoad(620 + I)
        ControlInitialise PVMenuSecondaireCDButton(I), (I Mod 3) * FormPavMenuSecondaire.ScaleWidth / 3 + Screen.TwipsPerPixelX, Screen.TwipsPerPixelY, FormPavMenuSecondaire.ScaleWidth / 3 - Screen.TwipsPerPixelX * 2, FormPavMenuSecondaire.ScaleHeight - Screen.TwipsPerPixelY * 2
    Next I
End Sub

' Evenement

Private Sub Form_Load()
    FormPavMenuSecondaireInit
End Sub

Private Sub Form_Resize()
    PVMenuSecondaireCDButtonInit
End Sub

Private Sub PVMenuSecondaireCDButton_Click(Index As Integer)
' Index : N° Du Bouton Du PVMenuSecondaire
    With DsrDatVente ' comme d'hab
        Select Case Index
            Case 0 ' Bouton Edition Du Ticket
                EditionTicket
        End Select
    End With
End Sub

