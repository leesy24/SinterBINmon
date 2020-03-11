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
   Begin VB.TextBox txtBinMinLH 
      Height          =   270
      Left            =   1320
      TabIndex        =   19
      Text            =   "500"
      Top             =   1980
      Width           =   615
   End
   Begin VB.TextBox txtBinMaxHH 
      Height          =   270
      Left            =   1320
      TabIndex        =   16
      Text            =   "2000"
      Top             =   1620
      Width           =   615
   End
   Begin VB.TextBox txtBinIPAddr 
      Height          =   270
      Left            =   1320
      TabIndex        =   1
      Text            =   "255.255.255.255"
      Top             =   195
      Width           =   1455
   End
   Begin VB.TextBox txtBinIPPort 
      Height          =   270
      Left            =   1320
      TabIndex        =   2
      Text            =   "99999"
      Top             =   555
      Width           =   615
   End
   Begin VB.CommandButton cmdSettingsExit 
      Caption         =   "닫 기"
      Height          =   495
      Left            =   3600
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdSettingsApply 
      Caption         =   "적 용"
      Height          =   375
      Left            =   3600
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtSensorAngle 
      Height          =   270
      Left            =   1320
      TabIndex        =   4
      Text            =   "-48"
      Top             =   1275
      Width           =   615
   End
   Begin VB.TextBox txtBinAngle 
      Height          =   270
      Left            =   1320
      TabIndex        =   3
      Text            =   "-10"
      Top             =   915
      Width           =   615
   End
   Begin VB.Label lbBinMinLH 
      Caption         =   "높이 최소"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "cm, 0~500"
      Height          =   255
      Left            =   2040
      TabIndex        =   17
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lbBinMaxHH 
      Caption         =   "높이 최대"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "cm, 1700~2000"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lbBinIPAddr 
      Caption         =   "Bin IP addr"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lbBinIPPort 
      Caption         =   "Bin IP port"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Serial2Net의 IP"
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Serial2Net의 port"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "°, 48~-48"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "°, 10~-10"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lbSensorAngle 
      Caption         =   "센서 기울기"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lbBinAngle 
      Caption         =   "Bin 기울기"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim Index%
Dim orgBinIPAddr$, orgBinIPPort$, orgBinAngle$, orgSensorAngle$
Dim orgBinMaxHH$, orgBinMinLH$

Public Sub Init(Index_I%, Title_I$, BinIPAddr_I$, BinIPPort_I$, BinAngle_I%, SensorAngle_I%, BinMaxHH_I%, BinMinLH_I%)
'
    Index = Index_I
'
    frmSettings.Caption = Title_I$ & " 설정"
'
    orgBinIPAddr = BinIPAddr_I
    txtBinIPAddr = BinIPAddr_I
'
    orgBinIPPort = BinIPPort_I
    txtBinIPPort = BinIPPort_I
'
    orgBinAngle = BinAngle_I
    txtBinAngle = BinAngle_I
'
    orgSensorAngle = SensorAngle_I
    txtSensorAngle = SensorAngle_I
'
    orgBinMaxHH = BinMaxHH_I
    txtBinMaxHH = BinMaxHH_I
'
    orgBinMinLH = BinMinLH_I
    txtBinMinLH = BinMinLH_I
'
End Sub

Private Sub cmdSettingsApply_Click()
'
    'Dim IsApplied As Boolean
    Dim IsValid As Boolean
'
    frmCFG.tmrCFG.Enabled = False
    frmCFG.tmrCFG.Interval = 60000 '' 60secs
    frmCFG.tmrCFG.Enabled = True
'
    'IsApplied = False
    IsValid = False
'
    If txtBinIPAddr <> orgBinIPAddr Then
        If IsValidIPAddress(txtBinIPAddr) = False Then
            MsgBox lbBinIPAddr & "는 192.168.0.1 형태의 값 이어야 합니다.", vbOKOnly
        Else
            orgBinIPAddr = txtBinIPAddr
            SaveSetting App.Title, "Settings", "BinIPAddr_" & Index, txtBinIPAddr
            IsValid = True
        End If
    End If
    If txtBinIPPort <> orgBinIPPort Then
        If IsValidIPPort(txtBinIPPort) = False Then
            MsgBox lbBinIPPort & "는 1024 ~ 65535 사이의 정수 값 이어야 합니다.", vbOKOnly
        Else
            orgBinIPPort = txtBinIPPort
            SaveSetting App.Title, "Settings", "BinIPPort_" & Index, txtBinIPPort
            IsValid = True
        End If
    End If
'
    If (IsValid = True) Then
        frmMain.ucBINdps1(Index).setIDX Index, txtBinIPAddr, txtBinIPPort
        'IsApplied = True
    End If
'
    IsValid = False
'
    If txtBinAngle <> orgBinAngle Then
        If IsNumeric(txtBinAngle) = False _
            Or CSng(CInt(Val(txtBinAngle))) <> CSng(Val(txtBinAngle)) _
            Or CInt(Val(txtBinAngle)) > 10! Or CInt(Val(txtBinAngle)) < -10! _
            Then
            MsgBox lbBinAngle & "는 10 ~ -10 사이의 정수 값 이어야 합니다.", vbOKOnly
        Else
            orgBinAngle = txtBinAngle
            SaveSetting App.Title, "Settings", "BinAngle_" & Index, txtBinAngle
            IsValid = True
        End If
    End If
    If txtSensorAngle <> orgSensorAngle Then
        If IsNumeric(txtSensorAngle) = False _
            Or CSng(CInt(Val(txtSensorAngle))) <> CSng(Val(txtSensorAngle)) _
            Or CInt(Val(txtSensorAngle)) > 48! Or CInt(Val(txtSensorAngle)) < -48! _
            Then
            MsgBox lbSensorAngle & "는 48 ~ -48 사이의 정수 값 이어야 합니다.", vbOKOnly
        Else
            orgSensorAngle = txtSensorAngle
            SaveSetting App.Title, "Settings", "SensorAngle_" & Index, txtSensorAngle
            IsValid = True
        End If
    End If
'
    If (IsValid = True) Then
        frmMain.ucBINdps1(Index).setBinSettings txtBinAngle, txtSensorAngle
        'IsApplied = True
    End If
'
    IsValid = False
'
    If txtBinMaxHH <> orgBinMaxHH Then
        If IsNumeric(txtBinMaxHH) = False _
            Or CSng(CInt(Val(txtBinMaxHH))) <> CSng(Val(txtBinMaxHH)) _
            Or CInt(Val(txtBinMaxHH)) > 2000! Or CInt(Val(txtBinMaxHH)) < 1700! _
            Then
            MsgBox lbBinMaxHH & "는 1700 ~ 2000 사이의 cm단위의 정수 값 이어야 합니다.", vbOKOnly
        Else
            orgBinMaxHH = txtBinMaxHH
            IsValid = True
        End If
    End If
    If txtBinMinLH <> orgBinMinLH Then
        If IsNumeric(txtBinMinLH) = False _
            Or CSng(CInt(Val(txtBinMinLH))) <> CSng(Val(txtBinMinLH)) _
            Or CInt(Val(txtBinMinLH)) > 500! Or CInt(Val(txtBinMinLH)) < 0! _
            Then
            MsgBox lbBinMinLH & "는 0 ~ 500 사이의 cm단위의 정수 값 이어야 합니다.", vbOKOnly
        Else
            orgBinMinLH = txtBinMinLH
            IsValid = True
        End If
    End If
'
    If (IsValid = True) Then
        frmMain.ucBINdps1(Index).set_maxHHLH txtBinMaxHH, txtBinMinLH
        'IsApplied = True
    End If
'
    'If (IsApplied = True) Then
    '    MsgBox "적용되었습니다.", vbOKOnly
    'End If
'
End Sub

Private Sub cmdSettingsExit_Click()
'
    frmCFG.tmrCFG.Enabled = False
    frmCFG.tmrCFG.Interval = 60000 '' 60secs
    frmCFG.tmrCFG.Enabled = True
'
    frmSettings.Visible = False
'
End Sub

Private Sub txtBinAngle_GotFocus()
'
    frmCFG.tmrCFG.Enabled = False
    frmCFG.tmrCFG.Interval = 60000 '' 60secs
    frmCFG.tmrCFG.Enabled = True
'
End Sub

Private Sub txtBinIPAddr_GotFocus()
'
    frmCFG.tmrCFG.Enabled = False
    frmCFG.tmrCFG.Interval = 60000 '' 60secs
    frmCFG.tmrCFG.Enabled = True
'
End Sub

Private Sub txtBinIPPort_GotFocus()
'
    frmCFG.tmrCFG.Enabled = False
    frmCFG.tmrCFG.Interval = 60000 '' 60secs
    frmCFG.tmrCFG.Enabled = True
'
End Sub

Private Sub txtBinMaxHH_GotFocus()
'
    frmCFG.tmrCFG.Enabled = False
    frmCFG.tmrCFG.Interval = 60000 '' 60secs
    frmCFG.tmrCFG.Enabled = True
'
End Sub

Private Sub txtBinMinLH_GotFocus()
'
    frmCFG.tmrCFG.Enabled = False
    frmCFG.tmrCFG.Interval = 60000 '' 60secs
    frmCFG.tmrCFG.Enabled = True
'
End Sub

Private Sub txtSensorAngle_GotFocus()
'
    frmCFG.tmrCFG.Enabled = False
    frmCFG.tmrCFG.Interval = 60000 '' 60secs
    frmCFG.tmrCFG.Enabled = True
'
End Sub
