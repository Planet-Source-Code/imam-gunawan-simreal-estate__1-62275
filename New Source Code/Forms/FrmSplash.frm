VERSION 5.00
Begin VB.Form FrmSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Permainan Pembangunan Perumahan"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Mulai Program"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7020
      TabIndex        =   1
      Top             =   2280
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Oleh : Gunawan (00111290)"
      Height          =   195
      Left            =   3660
      TabIndex        =   3
      Top             =   1260
      Width           =   2040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Versi:"
      Height          =   195
      Left            =   3660
      TabIndex        =   2
      Top             =   780
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PERUMAHAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3660
      TabIndex        =   0
      Top             =   180
      Width           =   3450
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2505
      Left            =   120
      Picture         =   "FrmSplash.frx":0442
      Top             =   120
      Width           =   3330
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Call Main
End Sub

Private Sub Form_Load()
    Label2.Caption = "Versi : " & App.Major & "." & App.Minor & "." & App.Revision
End Sub


