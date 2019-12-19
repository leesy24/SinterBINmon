VERSION 5.00
Begin VB.Form frmCFG 
   BorderStyle     =   1  '단일 고정
   Caption         =   "설 정"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows 기본값
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "시스템 설정"
      Height          =   1335
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   11175
      Begin VB.TextBox txtAVRcnt 
         Height          =   270
         Left            =   4440
         TabIndex        =   12
         Text            =   "99"
         Top             =   310
         Width           =   615
      End
      Begin VB.TextBox txtSinterNumber2 
         Height          =   270
         Left            =   1680
         TabIndex        =   10
         Text            =   "2"
         Top             =   675
         Width           =   615
      End
      Begin VB.TextBox txtSinterNumber1 
         Height          =   270
         Left            =   1680
         TabIndex        =   7
         Text            =   "1"
         Top             =   310
         Width           =   615
      End
      Begin VB.CommandButton cmdSetSYSTEM 
         Caption         =   "적 용"
         Height          =   375
         Left            =   10080
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lbAVRcnt 
         Caption         =   "누적횟수"
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbSinterNumber2 
         Caption         =   "두번째 소결 번호"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lbSinterNumber1 
         Caption         =   "첫번째 소결 번호"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCFGexit 
      Caption         =   "닫 기"
      Height          =   375
      Left            =   10200
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "센서 종류 설정"
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   1800
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

Private Sub cmdSetSYSTEM_Click()
    Dim i
    
    If (Val(txtSinterNumber1) < 1) Or (Val(txtSinterNumber1) > 9) Then
            txtSinterNumber1 = frmMain.SinterNumber1
                MsgBox lbSinterNumber1 & "는 1 이상 9 이하 이어야 입니다.", vbOKOnly
            Exit Sub
    End If
    
    SaveSetting App.Title, "Settings", "SinterNumber1", Val(txtSinterNumber1)
    
    'frmMain.SinterNumber1 = txtSinterNumber1
    
    'frmMain.lbTitle.Caption = "[" & txtSinterNumber1 & "," & txtSinterNumber2 & "소결] BIN LEVEL MONITORING"
    
    'For i = 0 To 9
    '    frmMain.ucBINdps1(i).setBinID
    'Next i
    
    If (Val(txtSinterNumber2) < 1) Or (Val(txtSinterNumber2) > 9) Then
            txtSinterNumber2 = frmMain.SinterNumber2
                MsgBox lbSinterNumber2 & "는 1 이상 9 이하 이어야 입니다.", vbOKOnly
            Exit Sub
    End If
    
    SaveSetting App.Title, "Settings", "SinterNumber2", Val(txtSinterNumber2)
    
    'frmMain.SinterNumber2 = txtSinterNumber2
    
    'frmMain.lbTitle.Caption = "[" & txtSinterNumber1 & "," & txtSinterNumber2 & "소결] BIN LEVEL MONITORING"
    
    'For i = 10 To 19
    '    frmMain.ucBINdps1(i).setBinID
    'Next i
    
    
    If (Val(txtAVRcnt) < 10) Or (Val(txtAVRcnt) > 99) Then
            txtAVRcnt = frmMain.AOdeepMAX
                MsgBox lbAVRcnt & "는 10 이상 99 이하 이어야 입니다.", vbOKOnly
            Exit Sub
    End If
    
    SaveSetting App.Title, "Settings", "DeepMax", Val(txtAVRcnt)
    frmMain.AOdeepFull = False
    frmMain.AOdeepCNT = 0
    frmMain.AOdeepMAX = Val(txtAVRcnt)

    tmrCFG.Enabled = False
    tmrCFG.Interval = 5000
    tmrCFG.Enabled = True
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

    txtSinterNumber1 = frmMain.SinterNumber1
    txtSinterNumber2 = frmMain.SinterNumber2
    
    txtAVRcnt = frmMain.AOdeepMAX
    
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
        , frmMain.ucBINdps1(Index).SensorAngle _
        , frmMain.ucBINdps1(Index).maxHH _
        , frmMain.ucBINdps1(Index).minLH
'
    frmSettings.Visible = True
'
End Sub

Private Sub tmrCFG_Timer()

    tmrCFG.Enabled = False
    
    frmSettings.Visible = False
    frmCFG.Visible = False
    
End Sub

