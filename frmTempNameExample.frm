VERSION 5.00
Begin VB.Form frmTempName 
   Caption         =   "Temporary File Name Example"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6336
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   6336
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDir 
      Caption         =   "Directory"
      Height          =   732
      Left            =   4680
      TabIndex        =   7
      Top             =   0
      Width           =   1452
      Begin VB.OptionButton optCurrDir 
         Caption         =   "Current"
         Height          =   192
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1092
      End
      Begin VB.OptionButton optTempDir 
         Caption         =   "Temp "
         Height          =   192
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1212
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   972
   End
   Begin VB.CommandButton cmdKill 
      Caption         =   "Kill Temp File"
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   972
   End
   Begin VB.TextBox txtPrefix 
      Height          =   288
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   492
   End
   Begin VB.CommandButton cmdTempName 
      Caption         =   "Get a Temporary Filename"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2412
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Prefix (optional)"
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1332
   End
   Begin VB.Label Label2 
      Caption         =   "Path Returned"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1212
   End
   Begin VB.Label lblTempDirTempFile 
      Caption         =   "lblTempDirTempFile"
      Height          =   252
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   4692
   End
End
Attribute VB_Name = "frmTempName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ----------------------------------------------------------------------
Private Sub cmdTempName_Click()
' ----------------------------------------------------------------------
  Dim sDir As String
  Dim sPrefix As String
  If optCurrDir.Value = True Then
    sDir = CurDir
  Else
    sDir = ""
  End If
  sPrefix = txtPrefix.Text
  lblTempDirTempFile = fnGetTempPath(sPrefix, sDir)
End Sub
' --------------------------------------------------------------------
Private Sub cmdKill_Click()
' --------------------------------------------------------------------
' may want to add check for error 53 file not found
  Kill lblTempDirTempFile
  
End Sub
' --------------------------------------------------------------------
Private Sub cmdExit_Click()
' --------------------------------------------------------------------
  Unload Me
End Sub
