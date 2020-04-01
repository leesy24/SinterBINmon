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
   Begin VB.Frame Frame2 
      Caption         =   "센서 종류 설정"
      Height          =   2415
      Left            =   240
      TabIndex        =   14
      Top             =   1800
      Width           =   11175
      Begin VB.TextBox txtCtypes 
         Height          =   270
         Index           =   0
         Left            =   1680
         TabIndex        =   16
         Top             =   310
         Width           =   615
      End
      Begin VB.CommandButton cmdSetTYPE 
         Caption         =   "적 용"
         Height          =   375
         Left            =   10080
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lbBinNO2 
         Caption         =   "1) 1소결BIN-01"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "시스템 설정"
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   11175
      Begin VB.CheckBox chkUsePLC 
         Caption         =   "PLC 이용"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkUseBeckHoof 
         Caption         =   "BeckHoff 이용"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtPLCIPPort2 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5760
         TabIndex        =   20
         Text            =   "99999"
         Top             =   1030
         Width           =   615
      End
      Begin VB.CommandButton cmdSetSYSTEM 
         Caption         =   "적 용"
         Height          =   375
         Left            =   10080
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtSinterNumber1 
         Height          =   270
         Left            =   1680
         TabIndex        =   6
         Text            =   "1"
         Top             =   310
         Width           =   615
      End
      Begin VB.TextBox txtSinterNumber2 
         Height          =   270
         Left            =   1680
         TabIndex        =   5
         Text            =   "2"
         Top             =   670
         Width           =   615
      End
      Begin VB.TextBox txtAVRcnt 
         Height          =   270
         Left            =   1680
         TabIndex        =   4
         Text            =   "99"
         Top             =   1030
         Width           =   615
      End
      Begin VB.TextBox txtPLCIPAddr 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5760
         TabIndex        =   3
         Text            =   "255.255.255.255"
         Top             =   310
         Width           =   1455
      End
      Begin VB.TextBox txtPLCIPPort1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5760
         TabIndex        =   2
         Text            =   "99999"
         Top             =   670
         Width           =   615
      End
      Begin VB.Label lbSinterNumber1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "첫번째 소결 번호"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbSinterNumber2 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "두번째 소결 번호"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lbAVRcnt 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "누적횟수"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lbPLCIPAddr 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "PLC IP addr."
         Height          =   255
         Left            =   4440
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbPLCIPPort1 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "PLC IP port 1"
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lbPLCIPPort2 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "PLC IP port 2"
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCFGexit 
      Caption         =   "닫 기"
      Height          =   375
      Left            =   10200
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
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

Const TIMEOUT = 60000 ' 60secs

Private Sub chkUseBeckHoof_Click()
'
    If (frmCFG.Visible = False) Then
        Exit Sub
    End If
'
    tmrCFG.Enabled = False
    tmrCFG.Interval = TIMEOUT
    tmrCFG.Enabled = True
'
End Sub

Private Sub chkUsePLC_Click()
'
    If (frmCFG.Visible = False) Then
        Exit Sub
    End If
'
    tmrCFG.Enabled = False
    tmrCFG.Interval = TIMEOUT
    tmrCFG.Enabled = True
'
    If (chkUsePLC.Value = 1) Then
        txtPLCIPAddr.Enabled = True
        txtPLCIPPort1.Enabled = True
        txtPLCIPPort2.Enabled = True
    Else
        txtPLCIPAddr.Enabled = False
        txtPLCIPPort1.Enabled = False
        txtPLCIPPort2.Enabled = False
    End If
'
End Sub

Private Sub cmdCFGexit_Click()
'
    tmrCFG.Enabled = False
'
    frmSettings.Visible = False
    frmCFG.Visible = False
'
End Sub

Private Sub cmdSetSYSTEM_Click()
    Dim IsValid As Boolean
    'Dim i
    
    tmrCFG.Enabled = False
    tmrCFG.Interval = TIMEOUT
    tmrCFG.Enabled = True
    
    IsValid = True
    
    If (Val(txtSinterNumber1) < 1) Or (Val(txtSinterNumber1) > 9) Then
        MsgBox lbSinterNumber1 & "는 1 이상 9 이하 이어야 합니다.", vbOKOnly
        IsValid = False
    End If
    
    If (Val(txtSinterNumber2) < 1) Or (Val(txtSinterNumber2) > 9) Then
        MsgBox lbSinterNumber2 & "는 1 이상 9 이하 이어야 합니다.", vbOKOnly
        IsValid = False
    End If
    
    If (Val(txtSinterNumber1) = Val(txtSinterNumber2)) Then
        MsgBox lbSinterNumber1 & "와 " & lbSinterNumber2 & "는 서로 다른 값이어야 합니다.", vbOKOnly
        IsValid = False
    End If
    
    If (IsValid = True) Then
        SaveSetting App.Title, "Settings", "SinterNumber1", Val(txtSinterNumber1)
        
        'frmMain.SinterNumber1 = txtSinterNumber1
        
        'frmMain.lbTitle.Caption = "[" & txtSinterNumber1 & "," & txtSinterNumber2 & "소결] BIN LEVEL MONITORING"
        
        'For i = 0 To 9
        '    frmMain.ucBINdps1(i).setBinID
        'Next i
        
        SaveSetting App.Title, "Settings", "SinterNumber2", Val(txtSinterNumber2)
    
        'frmMain.SinterNumber2 = txtSinterNumber2
        
        'frmMain.lbTitle.Caption = "[" & txtSinterNumber1 & "," & txtSinterNumber2 & "소결] BIN LEVEL MONITORING"
        
        'For i = 10 To 19
        '    frmMain.ucBINdps1(i).setBinID
        'Next i
    End If
    
    IsValid = True
    
    If (Val(txtAVRcnt) < 10) Or (Val(txtAVRcnt) > 99) Then
        MsgBox lbAVRcnt & "는 10 이상 99 이하 이어야 합니다.", vbOKOnly
        IsValid = False
    End If
    
    If (IsValid = True) Then
        SaveSetting App.Title, "Settings", "DeepMax", Val(txtAVRcnt)
        frmMain.AOdeepFull = False
        frmMain.AOdeepCNT = 0
        frmMain.AOdeepMAX = Val(txtAVRcnt)
    End If
    
    If (chkUseBeckHoof.Value <> frmMain.chkUseBeckHoof) Then
        SaveSetting App.Title, "Settings", "UseBeckHoof", chkUseBeckHoof.Value
        frmMain.chkUseBeckHoof = chkUseBeckHoof.Value
    End If
    
    If (chkUsePLC.Value <> frmMain.chkUsePLC) Then
        SaveSetting App.Title, "Settings", "UsePLC", chkUsePLC.Value
        frmMain.chkUsePLC = chkUsePLC.Value
    End If
    
    IsValid = True
    
    If IsValidIPAddress(txtPLCIPAddr) = False Then
        MsgBox lbPLCIPAddr & "는 192.168.0.1 형태의 값 이어야 합니다.", vbOKOnly
        IsValid = False
    End If
    
    If IsValidIPPort(txtPLCIPPort1) = False Then
        MsgBox lbPLCIPPort1 & "는 1024 ~ 65535 사이의 정수 값 이어야 합니다.", vbOKOnly
        IsValid = False
    End If
    
    If IsValidIPPort(txtPLCIPPort2) = False Then
        MsgBox lbPLCIPPort2 & "는 1024 ~ 65535 사이의 정수 값 이어야 합니다.", vbOKOnly
        IsValid = False
    End If
    
    If txtPLCIPPort1 = txtPLCIPPort2 Then
        MsgBox lbPLCIPPort1 & "와 " & lbPLCIPPort2 & "는 서로 다른 값 이어야 합니다.", vbOKOnly
        IsValid = False
    End If
    
    If (IsValid = True) Then
        SaveSetting App.Title, "Settings", "PLCIPAddr", txtPLCIPAddr
        SaveSetting App.Title, "Settings", "PLCIPPort1", txtPLCIPPort1
        SaveSetting App.Title, "Settings", "PLCIPPort2", txtPLCIPPort2
    
        With frmMain.wsPLC1
            .Close
            .RemoteHost = txtPLCIPAddr
            .RemotePort = txtPLCIPPort1
            .LocalPort = txtPLCIPPort1
            .Bind .LocalPort
        End With
        With frmMain.wsPLC2
            .Close
            .RemoteHost = txtPLCIPAddr
            .RemotePort = txtPLCIPPort2
            .LocalPort = txtPLCIPPort2
            .Bind .LocalPort
        End With
    End If
End Sub

Private Sub cmdSetTYPE_Click()
    Dim i
    
    tmrCFG.Enabled = False
    tmrCFG.Interval = TIMEOUT
    tmrCFG.Enabled = True
    
    For i = 0 To 19
        frmMain.ucBINdps1(i).setScanTYPE CInt(txtCtypes(i))
    Next i
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim iLeft As Long
    Dim iTop As Long

    txtSinterNumber1 = frmMain.SinterNumber1
    txtSinterNumber2 = frmMain.SinterNumber2
    
    txtAVRcnt = frmMain.AOdeepMAX
    
    chkUseBeckHoof.Value = frmMain.chkUseBeckHoof
    chkUsePLC.Value = frmMain.chkUsePLC
    If (chkUsePLC.Value = 1) Then
        txtPLCIPAddr.Enabled = True
        txtPLCIPPort1.Enabled = True
        txtPLCIPPort2.Enabled = True
    End If
    
    txtPLCIPAddr.Text = frmMain.wsPLC1.RemoteHost
    txtPLCIPPort1.Text = frmMain.wsPLC1.RemotePort
    txtPLCIPPort2.Text = frmMain.wsPLC2.RemotePort
    
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
    tmrCFG.Interval = TIMEOUT
    tmrCFG.Enabled = True
End Sub

Private Sub lbBinNO2_Click(Index As Integer)
'
    tmrCFG.Enabled = False
    tmrCFG.Interval = TIMEOUT
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

Private Sub txtAVRcnt_GotFocus()
'
    tmrCFG.Enabled = False
    tmrCFG.Interval = TIMEOUT
    tmrCFG.Enabled = True
'
End Sub

Private Sub txtCtypes_GotFocus(Index As Integer)
'
    tmrCFG.Enabled = False
    tmrCFG.Interval = TIMEOUT
    tmrCFG.Enabled = True
'
End Sub

Private Sub txtSinterNumber1_GotFocus()
'
    tmrCFG.Enabled = False
    tmrCFG.Interval = TIMEOUT
    tmrCFG.Enabled = True
'
End Sub

Private Sub txtSinterNumber2_GotFocus()
'
    tmrCFG.Enabled = False
    tmrCFG.Interval = TIMEOUT
    tmrCFG.Enabled = True
'
End Sub
