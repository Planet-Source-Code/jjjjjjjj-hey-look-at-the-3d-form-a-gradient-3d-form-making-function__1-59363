VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00A3E8FC&
   BorderStyle     =   0  'None
   Caption         =   "Sample Form"
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTest.frx":0000
   ScaleHeight     =   4005
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   120
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Public Sub Test()
    With frmMain
        ProjectForm Me, Val(.txtThkX), Val(.txtThkY), .cmdSC.BackColor, .cmdEC.BackColor, Val(.txtCurvature), Val(.txtFrames)
    End With
    lb.Caption = "3D"
End Sub

Public Sub Normal()
    Load Me: lb.Caption = "Normal"
End Sub

Private Sub Form_Load()
    Me.Show
End Sub
