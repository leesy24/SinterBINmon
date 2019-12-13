VERSION 5.00
Begin VB.Form frmCFG 
   BorderStyle     =   1  '단일 고정
   Caption         =   "설 정"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows 기본값
   Visible         =   0   'False
   Begin VB.CommandButton cmdCFGexit 
      Caption         =   "닫 기"
      Height          =   375
      Left            =   10200
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame frTypes 
      Caption         =   "센서 종류 설정"
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      Begin VB.CommandButton cmdSetTYPE 
         Caption         =   "적 용"
         Height          =   375
         Left            =   10080
         Style           =   1  '그래픽
         TabIndex        =   3
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtCtypes 
         Height          =   270
         Index           =   0
         Left            =   1680
         TabIndex        =   2
         Top             =   310
         Width           =   615
      End
      Begin VB.Label lbBinNO2 
         Caption         =   "1) 1소결BIN-01"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Timer tmrCFG 
      Enabled         =   0   'False
      Interval        =   50000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmCFG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCFGexit_Click()
    frmSettings.Visible = False
    frmCFG.Visible = False
End Sub

Private Sub cmdSetTYPE_Click()
    Dim i
    
    For i = 0 To 19
        frmMain.ucBINdps1(i).setScanTYPE CInt(txtCtypes(i))
    Next i
    
    tmrCFG.Enabled = False
    tmrCFG.Interval = 5000
    tmrCFG.Enabled = True
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim iLeft As Long
    Dim iTop As Long

    For i = 0 To 19
        If i <> 0 Then
            Load lbBinNO2(i)
            Load txtCtypes(i)
            
            iLeft = lbBinNO2(0).Left + ((i) Mod 5) * (lbBinNO2(0).Width + txtCtypes(i).Width + 100)
            iTop = lbBinNO2(0).Top + ((i) \ 5) * 350
            
            lbBinNO2(i).Left = iLeft
            lbBinNO2(i).Top = iTop
            
            txtCtypes(i).Left = iLeft + lbBinNO2(0).Width
            txtCtypes(i).Top = iTop - 50
        End If
        
        lbBinNO2(i).Caption = frmMain.ucBINdps1(i).getBinCaption
        
        txtCtypes(i) = frmMain.ucBINdps1(i).getScanTYPE
    Next i

    For i = 0 To 19
        lbBinNO2(i).Visible = True
        txtCtypes(i).Visible = True
    Next i
    
    tmrCFG.Enabled = False
    tmrCFG.Interval = 60000 '' 60secs
    tmrCFG.Enabled = True
End Sub

Private Sub lbBinNO2_Click(Index As Integer)
'
    tmrCFG.Enabled = False
    tmrCFG.Interval = 60000 '' 60secs
    tmrCFG.Enabled = True
'
    If frmSettings.Visible = True Then
        frmSettings.Show
    End If
'
    frmSettings.Init _
        Index _
        , lbBinNO2(Index).Caption _
        , frmMain.ucBINdps1(Index).ipAddr _
        , frmMain.ucBINdps1(Index).ipPort _
        , frmMain.ucBINdps1(Index).BinAngle _
        , frmMain.ucBINdps1(Index).SensorAngle
'
    frmSettings.Visible = True
'
End Sub

Private Sub tmrCFG_Timer()

    tmrCFG.Enabled = False
    
    frmSettings.Visible = False
    frmCFG.Visible = False
    
End Sub

