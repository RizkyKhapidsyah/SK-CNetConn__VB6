VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test CNetConn"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   2850
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCheck 
      Interval        =   5000
      Left            =   960
      Top             =   1080
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblType 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Type:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblConnected 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Connected:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Testing CNetConn..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myNetConn As CNetConn

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set myNetConn = New CNetConn
    Call tmrCheck_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set myNetConn = Nothing
End Sub

Private Sub tmrCheck_Timer()
    With myNetConn
        lblConnected.Caption = IIf(.IsConnected, "True", "False")
        lblName.Caption = .ConnName
        
        lblType.Caption = .ConnTypeDevice(.ConnType)
    End With
End Sub
