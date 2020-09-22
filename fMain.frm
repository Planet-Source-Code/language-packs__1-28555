VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2280
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtText 
      Height          =   315
      Left            =   1260
      TabIndex        =   5
      Top             =   1740
      Width           =   3315
   End
   Begin VB.CommandButton CmdClickMe 
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1155
   End
   Begin VB.CommandButton CmdClickMe 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1155
   End
   Begin VB.CommandButton CmdClickMe 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1155
   End
   Begin VB.CommandButton CmdClickMe 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label LblWrite 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label LblText 
      Height          =   975
      Left            =   1380
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
   Begin VB.Menu MnuSelectLanguage 
      Caption         =   ""
      Begin VB.Menu MnuLanguage 
         Caption         =   "0"
         Index           =   0
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Const LANGPACKS_EXT = ".ept"

Dim LanguagePack As String
Dim PopupMessages As New Collection

Private Sub Form_Load()
    LanguagePack = App.Path & "\" & "english.ept"
    EnumLanguagePacks
    LoadLanguagePack
End Sub

Private Sub CmdClickMe_Click(Index As Integer)
    MsgBox PopupMessages.Item("clicked"), vbInformation
End Sub

Private Sub MnuLanguage_Click(Index As Integer)
    If Not MnuLanguage(Index).Checked Then
        Dim A As Integer
        For A = 0 To MnuLanguage.UBound
            MnuLanguage(A).Checked = False
        Next A
        MnuLanguage(Index).Checked = True
        LanguagePack = App.Path & "\" & MnuLanguage(Index).Caption ' & LANGPACKS_EXT
        LoadLanguagePack
    End If
End Sub

Private Sub LoadLanguagePack()
    LoadEPT LanguagePack, "Messages", PopupMessages
    LoadEPT LanguagePack, Name, Me
End Sub

Private Sub EnumLanguagePacks()
    Dim PackName As String
    PackName = Dir(App.Path & "\")
    Do Until PackName = ""
        If Right(PackName, Len(LANGPACKS_EXT)) = LANGPACKS_EXT Then
            MnuLanguage(0).Checked = False
            MnuLanguage(0).Caption = PackName 'StrConv(Left(PackName, Len(PackName) - Len(LANGPACKS_EXT)), vbProperCase)
            If App.Path & "\" & PackName = LanguagePack Then
                MnuLanguage(0).Checked = True
            End If
            Load MnuLanguage(MnuLanguage.Count)
        End If
        PackName = Dir$
    Loop
    Unload MnuLanguage(MnuLanguage.UBound)
End Sub
