VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text Search"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3315
      TabIndex        =   4
      Top             =   1110
      Width           =   1185
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      Height          =   360
      Left            =   2085
      TabIndex        =   3
      Top             =   1110
      Width           =   1185
   End
   Begin VB.TextBox txtStringInput 
      Height          =   330
      Left            =   705
      TabIndex        =   1
      Top             =   225
      Width           =   3780
   End
   Begin VB.Label Label1 
      Caption         =   "The text that you wish to search for?"
      Height          =   285
      Left            =   780
      TabIndex        =   2
      Top             =   645
      Visible         =   0   'False
      Width           =   3510
   End
   Begin VB.Label lblStringText 
      Caption         =   "String:"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   270
      Width           =   585
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSearch_Click()
    Call FireCallMain.SearchListBox
End Sub

