VERSION 5.00
Begin VB.Form frmRandomizer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Randomizer"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Text            =   "20"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Default         =   -1  'True
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.ListBox lstRand 
      Columns         =   1
      Height          =   5325
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmRandomizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdExit_Click()
    End
End Sub
Private Sub cmdGenerate_Click()
    If Val(txtNum) <= 0 Then Exit Sub
    Dim Num As Integer, Used() As Boolean, Count As Integer
    ReDim Used(1 To Val(txtNum)) As Boolean
    Randomize Timer
    lstRand.Clear
    lstRand.Columns = CInt(lstRand.Width / (TextWidth(txtNum) * 1.5))
    For Count = 1 To Val(txtNum)
        Do
            Num = (Rnd * (Val(txtNum) - 1)) + 1
        Loop Until Not Used(Num)
        Used(Num) = True
        lstRand.AddItem CStr(Num)
    Next Count
    Label1.Caption = CStr(lstRand.ListCount)
End Sub
