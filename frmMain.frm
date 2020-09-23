VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crack Access 2000 Password"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPassword 
      Height          =   1095
      Left            =   45
      TabIndex        =   3
      Top             =   2430
      Width           =   4245
      Begin VB.TextBox txtPassword 
         Height          =   465
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   405
         Width           =   3885
      End
      Begin VB.Label lblResults 
         AutoSize        =   -1  'True
         Caption         =   "The password is:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   135
         Width           =   1455
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   2340
      Pattern         =   "*.mdb"
      TabIndex        =   2
      Top             =   540
      Width           =   1950
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   45
      TabIndex        =   1
      Top             =   945
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   540
      Width           =   2175
   End
   Begin VB.Label lblInstruction 
      Caption         =   "Click on the Access database you wish to recover the password for"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   4200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Dim db As String

    txtPassword.Text = ""
    db = File1.Path & "\" & File1.FileName
    db = Replace(db, "\\", "\")
    txtPassword.Text = GuessAccess2000Password(db)
End Sub


