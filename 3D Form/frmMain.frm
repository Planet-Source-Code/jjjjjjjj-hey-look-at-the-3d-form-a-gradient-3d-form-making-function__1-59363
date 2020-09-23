VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "3D Form By -  'Jim Jose'"
   ClientHeight    =   4140
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNormal 
      Caption         =   "&Load Normal Form"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Options"
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7695
      Begin VB.TextBox txtSC 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1800
         TabIndex        =   16
         Text            =   "255"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtEC 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   5160
         TabIndex        =   15
         Text            =   "0"
         Top             =   1320
         Width           =   855
      End
      Begin MSComDlg.CommonDialog cdlg 
         Left            =   7080
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtFrames 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   5160
         TabIndex        =   14
         Text            =   "25"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtCurvature 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1800
         TabIndex        =   13
         Text            =   "25"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdEC 
         BackColor       =   &H00000000&
         Caption         =   "..."
         Height          =   375
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton cmdSC 
         BackColor       =   &H000000FF&
         Caption         =   "..."
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txtThkY 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   5160
         TabIndex        =   6
         Text            =   "10"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtThkX 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1800
         TabIndex        =   5
         Text            =   "15"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note : The startcolor=-1 will capture the lower-right corner color of the form."
         Height          =   240
         Left            =   480
         TabIndex        =   17
         Top             =   2760
         Width           =   6495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frames"
         Height          =   240
         Left            =   3840
         TabIndex        =   12
         Top             =   2160
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Curvature"
         Height          =   240
         Left            =   480
         TabIndex        =   11
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Color"
         Height          =   240
         Left            =   3840
         TabIndex        =   8
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Color"
         Height          =   240
         Left            =   480
         TabIndex        =   7
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thickness Y"
         Height          =   240
         Left            =   3840
         TabIndex        =   4
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thickness X"
         Height          =   240
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Load Test Form"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEC_Click()
    cdlg.ShowColor
    cmdEC.BackColor = cdlg.Color
    txtEC = cdlg.Color
End Sub

Private Sub cmdNormal_Click()
    frmTest.Normal
End Sub

Private Sub cmdSC_Click()
    cdlg.ShowColor
    cmdSC.BackColor = cdlg.Color
    txtSC = cdlg.Color
End Sub

Private Sub cmdTest_Click()
    frmTest.Test
End Sub

