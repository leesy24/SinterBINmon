VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  '단일 고정
   Caption         =   "1) 1소결BIN-01 설정"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtBinAngle 
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Top             =   190
      Width           =   735
   End
   Begin VB.Label lbBinAngle 
      Caption         =   "Bin 기울기"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
