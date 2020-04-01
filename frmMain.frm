VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{40DD8EA0-284B-11D0-A7B0-0020AFF929F4}#2.3#0"; "Adsocx.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404000&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Sinter BIN Monitor"
   ClientHeight    =   12885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13590
   FillStyle       =   0  '단색
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12885
   ScaleWidth      =   13590
   Begin ADSOCXLib.AdsOcx AdsOcx1 
      Left            =   8280
      Top             =   1200
      _Version        =   131074
      _ExtentX        =   900
      _ExtentY        =   953
      _StockProps     =   0
      AdsAmsServerNetId=   "172.16.21.20.1.1"
      AdsAmsServerPort=   800
      AdsAmsClientPort=   32807
      AdsClientType   =   ""
      AdsClientAdsState=   ""
      AdsClientAdsControl=   ""
   End
   Begin prjSinterBINmon.ucBINmon ucBINmon1 
      Height          =   10215
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   18018
   End
   Begin prjSinterBINmon.ucBINdps ucBINdps1 
      Height          =   7815
      Index           =   0
      Left            =   2040
      TabIndex        =   19
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   13785
   End
   Begin VB.TextBox txtAVRcnt 
      Alignment       =   2  '가운데 맞춤
      Enabled         =   0   'False
      Height          =   270
      Left            =   10440
      TabIndex        =   17
      Text            =   "0/0"
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdPcsRun 
      BackColor       =   &H00008000&
      Caption         =   "PcsRUN"
      Height          =   255
      Left            =   2520
      MaskColor       =   &H00E0E0E0&
      Style           =   1  '그래픽
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtPcsPort 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1800
      TabIndex        =   14
      Text            =   "8001"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtPcsIP 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   480
      TabIndex        =   13
      Text            =   "172.24.55.27"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtWSpcs 
      BackColor       =   &H00FF00FF&
      Enabled         =   0   'False
      Height          =   270
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   180
   End
   Begin VB.CommandButton cmdDmon 
      Caption         =   "dMon"
      Height          =   255
      Left            =   7320
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cboIDX 
      Height          =   300
      Left            =   6360
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSD1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   8
      Top             =   11160
      Width           =   10575
   End
   Begin VB.CommandButton cmdADSclr 
      Caption         =   "ADSclr"
      Height          =   255
      Left            =   8880
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdADS1 
      Caption         =   "ADS1"
      Height          =   255
      Left            =   9840
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picTop 
      BackColor       =   &H00808080&
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   13275
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      Begin MSWinsockLib.Winsock wsPLC2 
         Left            =   5880
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSWinsockLib.Winsock wsPLC1 
         Left            =   5760
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin VB.CommandButton cmdCFG 
         BackColor       =   &H00008000&
         Caption         =   "설 정"
         Height          =   375
         Left            =   9720
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   25
         Top             =   360
         Width           =   915
      End
      Begin VB.CommandButton adsTest1 
         Caption         =   "adsTest1"
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Timer tmrPcs 
         Interval        =   2000
         Left            =   6480
         Top             =   360
      End
      Begin MSWinsockLib.Winsock wsPcs 
         Left            =   6000
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tmrAoDo 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   6960
         Top             =   360
      End
      Begin VB.CommandButton cmdRunStop 
         BackColor       =   &H00008000&
         Caption         =   "RUN/STOP"
         Height          =   375
         Left            =   8160
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Timer tmrINIT 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   7320
         Top             =   360
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00808080&
         Caption         =   "종 료"
         Height          =   375
         Left            =   12240
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdHide 
         BackColor       =   &H00808080&
         Caption         =   "화면감추기"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10800
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lbRelDate 
         BackStyle       =   0  '투명
         Caption         =   "Release date"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   540
         Width           =   1695
      End
      Begin VB.Label lbRelVersion 
         BackStyle       =   0  '투명
         Caption         =   "Release version"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lbTitle 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "[1,2소결] BIN LEVEL MONITORING"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   705
         Left            =   4080
         TabIndex        =   4
         Top             =   0
         Width           =   9195
      End
      Begin VB.Image imgLogo1 
         BorderStyle     =   1  '단일 고정
         Height          =   495
         Left            =   120
         Picture         =   "frmMain.frx":16AC2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1605
      End
      Begin VB.Label lbTeam 
         BackColor       =   &H00808080&
         Caption         =   "DASAN-InfoTEK"
         BeginProperty Font 
            Name            =   "바탕체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Label lbTimeNow 
      BackStyle       =   0  '투명
      Caption         =   "RunTime"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "누적횟수:"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   9600
      TabIndex        =   18
      Top             =   1010
      Width           =   975
   End
   Begin VB.Label lbUpTime 
      BackStyle       =   0  '투명
      Caption         =   "Up_Time"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lbVS1 
      BackStyle       =   0  '투명
      Caption         =   "Label1"
      Height          =   255
      Left            =   7680
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===========================================================================================
'
'                       2D LEVEL Monitoring System
'                       for BIN5 with SICK LMS-211
'
'                                   BIN5mon V1.00
'
'===========================================================================================


Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" _
                    Alias "RtlMoveMemory" (hpvDest As Any, _
                                           hpvSource As Any, _
                                           ByVal cbCopy As Long)

Private Const relVersion = "v2.00.05"
Private Const relDate = "2020-03-11"

Dim d1 As Single


Public SinterNumber1 As Integer
Public SinterNumber2 As Integer

Public chkUseBeckHoof As Integer
Public chkUsePLC As Integer

Dim ipAddr(20) As String
Dim ipPort(20) As String

Dim AOdata(33) As Integer
Dim AOdata2(33) As Integer

Dim AOdeep(20, 100) As Integer
Dim Hdeep(20, 100) As Integer
Public AOdeepCNT As Integer
Public AOdeepMAX As Integer        ''<=MAX:99
Public AOdeepFull As Boolean

Dim BinWidth As Integer
''
Dim BinMaxH(33) As Integer
Dim BinMinH(33) As Integer
''
Dim BinTYPE(33) As Integer




''// WX001 :: Port: 800, IGrp: 0xF020, IOffs: 0x64, Len: 2
''// --------------------------------------------
''      pAddr1->port = 800;
''      dwData[ 0] = (short)(0);  //(0x5555);
''      dwData[ 1] = (short)(32768/64);  //(0xAAAA);
''      dwData[ 2] = (short)(32768/32);
''      dwData[ 3] = (short)(32768/16);
''      dwData[ 4] = (short)(32768/8);
''      dwData[ 5] = (short)(32768/4);
''      dwData[ 6] = (short)(32768/2);
''      dwData[ 7] = (short)(32765);
''
''      nErr = AdsSyncWriteReq( pAddr1, 0xF020,  0x64, 16, &dwData);
''
''// WX030 :: Port: 800, IGrp: 0xF020, IOffs: 0x9E, Len: 2
''      dwData[ 5] = (short)(32768/4-1);
''      dwData[ 6] = (short)(32768/2-1);
''      dwData[ 7] = (short)(32765);
''
''      nErr = AdsSyncWriteReq( pAddr1, 0xF020,  0xA0, 4, &dwData[6]);


''''    AdsSyncWriteReq
''''    Writes data of any type synchronously to an ADS device.
''''
''''    object.AdsSyncWriteReq(
''''      nIndexGroup    As Long,
''''      nIndexOffset   As Long,
''''      cbLength       As Long,
''''      pData          As YY)
''''    As Long
''''
''''    Parameter
''''
''''    nIndexGroup
''''        [in] Index group of the ADS variable
''''
''''    nIndexOffset
''''        [in] Index offset of the ADS variable
''''
''''    cbLength
''''        [in] Length of the data in bytes (see VBVarLength)
''''
''''    pData
''''        [in] Visual Basic variable from which the data is written into the ADS variable
''''

''''    Dim VBVarInteger(0) As Integer
''''    Dim VBVarLong(0) As Long
''''    Dim VBVarSingle(0) As Single
''''    Dim VBVarDouble(0) As Double
''''    Dim VBVarByte(0) As Byte
''''    Dim VBVarBoolean(0) As Boolean
''''
''''    VBVarInteger(0) = 123
''''    VBVarLong(0) = 456
''''    VBVarSingle(0) = 3,1415
''''    VBVarDouble(0) = 2,876
''''    VBVarByte(0) = 7
''''    VBVarBoolean(0) = False
''''
''''    'Write values to PLC
''''    Call AdsOcx1.AdsSyncWriteReq(&H4020&, 0&, 2&, VBVarInteger)
''''    Call AdsOcx1.AdsSyncWriteReq(&H4020&, 2&, 4&, VBVarLong)
''''    Call AdsOcx1.AdsSyncWriteReq(&H4020&, 6&, 4&, VBVarSingle)
''''    Call AdsOcx1.AdsSyncWriteReq(&H4020&, 10&, 8&, VBVarDouble)
''''    Call AdsOcx1.AdsSyncWriteReq(&H4020&, 18&, 1&, VBVarByte)
''''    Call AdsOcx1.AdsSyncWriteReq(&H4021&, 152&, 2&, VBVarBoolean)


Private Sub adsTest1_Click()

Dim ioD(33) As Integer
Dim i As Long


    ioD(0) = 10000
    
    ''Port: 301, IGrp: 0xF030, IOffs: 0x12, Len: 2
    ''Port: 301, IGrp: 0xF030, IOffs: 0x2C, Len: 2
    
    
    ''AdsOcx1.AdsAmsServerNetId = "172.16.21.20.1.1"   '''AdsOcx1.AdsAmsClientNetId
    
    AdsOcx1.AdsAmsServerNetId = "0.0.0.0.0.0"
    
    AdsOcx1.AdsAmsServerPort = 301  ''800
    AdsOcx1.EnableErrorHandling = True
    
    
    AdsOcx1.AdsSyncWriteReq &HF030&, &H12&, 2, ioD
    AdsOcx1.AdsSyncWriteReq &HF030&, &H2C&, 2, ioD
    
    
End Sub


Private Sub cmdADS1_Click()

Dim ioD(33) As Integer
Dim i As Long


    ioD(0) = 32767
    ioD(1) = 32768 / 2 - 1
    ioD(2) = 32768 / 4 - 1
    ioD(3) = 32768 / 8 - 1
    
    AdsOcx1.AdsAmsServerNetId = "172.16.21.20.1.1"   '''AdsOcx1.AdsAmsClientNetId
    AdsOcx1.AdsAmsServerPort = 800
    AdsOcx1.EnableErrorHandling = True
    
    
''    For i = 0 To 3  ''0xf020==61472
''        ''AdsOcx1.AdsSyncWriteIntegerReq &HA001, i, 1, ioD(i)
''        ''AdsOcx1.AdsSyncWriteIntegerReq CLng(61472), CLng(&H64 + (i * 2)), 2, ioD(i)
''        ''AdsOcx1.AdsSyncWriteIntegerReq &HF020&, CLng(&H64 + (i * 2)), 2, ioD(i)
''        AdsOcx1.AdsSyncWriteReq &HF020&, &H64& + (i * 2), 2, ioD(i)
''    Next i
''
''    AdsOcx1.AdsSyncWriteReq &HF020&, &H64& + (30 * 2), 4, ioD(2)  '''(30&31)
    
    
    Dim i1, i2 As Long  ''Integer
    
''    For i1 = 0 To 3
''        For i2 = 0 To 7
''            ioD((i1 * 8) + i2) = (32768 / (i2 + 1)) - 1
''        Next i2
''    Next i1
''    AdsOcx1.AdsSyncWriteReq &HF020&, &H64&, 64, ioD
    
''    For i1 = 0 To 7
''        For i2 = 0 To 3
''            ioD((i1 * 4) + i2) = (32768 / (i2 + 1)) - 1
''        Next i2
''    Next i1
''    AdsOcx1.AdsSyncWriteReq &HF020&, &H64&, 64, ioD


    ioD(0) = 32768 * 0.05       ''1
    ioD(1) = 32768 * 0.1        ''2
    ioD(2) = 32768 * 0.15       ''3
    ioD(3) = 0

    ioD(4) = 32768 * 0.2        ''4
    ioD(5) = 32768 * 0.25       ''5
    ioD(6) = 32768 * 0.3        ''6
    ioD(7) = 0

    ioD(8) = 32768 * 0.35       ''7
    ioD(9) = 32768 * 0.4        ''8
    ioD(10) = 32768 * 0.45      ''9
    ioD(11) = 0

    ioD(12) = 32768 * 0.5       ''10
    ioD(13) = 32768 * 0.55      ''11
    ''''''''''''''''''''''''''''''''''''
    ioD(14) = 32768 * 0.05      ''1
    ioD(15) = 0

    ioD(16) = 32768 * 0.1       ''2
    ioD(17) = 32768 * 0.15      ''3
    ioD(18) = 32768 * 0.2       ''4
    ioD(19) = 0

    ioD(20) = 32768 * 0.25      ''5
    ioD(21) = 32768 * 0.3       ''6
    ioD(22) = 32768 * 0.35      ''7
    ioD(23) = 0

    ioD(24) = 32768 * 0.4       ''8
    ioD(25) = 32768 * 0.45      ''9
    ioD(26) = 32768 * 0.5       ''10
    ioD(27) = 0

    ioD(28) = 32768 * 0.55      ''11
    ioD(29) = 0
    ioD(30) = 0
    ioD(31) = 0

    AdsOcx1.AdsSyncWriteReq &HF020&, &H64&, 64, ioD



    
End Sub

Private Sub cmdADSclr_Click()
Dim i As Integer
Dim d As Integer

    AdsOcx1.AdsAmsServerNetId = "172.16.21.20.1.1"   '''AdsOcx1.AdsAmsClientNetId
    AdsOcx1.AdsAmsServerPort = 800
    AdsOcx1.EnableErrorHandling = True
    
''    d = 0
''    For i = 0 To 31
''        AdsOcx1.AdsSyncWriteReq &HF020&, &H64& + (i * 2), 2, d
''    Next i
    
    Dim ioD(33) As Integer  ''(0~31)
    For i = 0 To 31
        ioD(i) = 0
    Next i
    AdsOcx1.AdsSyncWriteReq &HF020&, &H64&, 64, ioD
    
End Sub

Private Sub cmdCFG_Click()

''    frmCFG.txtMaxHH = frmMain.txtMaxHH
''    frmCFG.txtBaseHH = frmMain.txtBaseHH

    If frmCFG.Visible = True Then
        frmCFG.Show
    Else
        frmCFG.Visible = True
    End If
    
''    frmCFG.tmrCFG_update

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ret1
    
    ret1 = MsgBox("종료하면 모든 기능이 정지됩니다." & vbCrLf & "정말 종료 하시겠습니까?", vbYesNo)

    If ret1 <> vbYes Then
        Cancel = 1
        Exit Sub
    End If

    Unload frmSettings
    Unload frmCFG
End Sub

Private Sub Form_Click()
    If frmCFG.Visible = True Then
        frmCFG.tmrCFG_update
    End If
End Sub

Private Sub Form_DblClick()
    If frmCFG.Visible = True Then
        frmCFG.tmrCFG_update
    End If
End Sub

''Private Sub cmdDmon_Click()
''    ''' txtSD1 = ucBINmon1(cboIDX.ListIndex).ret_SDXY
''End Sub

Private Sub cmdExit_Click()
    Dim ret1
    
    ret1 = MsgBox("종료하면 모든 기능이 정지됩니다." & vbCrLf & "정말 종료 하시겠습니까?", vbYesNo)

    If ret1 <> vbYes Then
        Exit Sub
    End If

    End
End Sub



Private Sub cmdHide_Click()

    ''frmMain.Visible = False
    frmMain.Hide
    
    
End Sub

Private Sub cmdPcsRun_Click()
    ''&H00008000& ''G
    ''&H00000080& ''R
    ''QBColor
  Dim i As Integer
  
    If cmdPcsRun.BackColor = &H8000& Then  ''run
        wsPcs.Close
        tmrPcs.Enabled = False
        cmdPcsRun.BackColor = &H80&        ''stop
    Else  ''stop
        tmrPcs.Enabled = True
        cmdPcsRun.BackColor = &H8000&        ''run
    End If

End Sub

Private Sub cmdRunStop_Click()

    ''&H00008000& ''G
    ''&H00000080& ''R
    ''QBColor
  Dim i As Integer
  
''    If cmdRunStop.BackColor = &H8000& Then  ''run
''        For i = 0 To 10
''            ucBINmon1(i).scan_STOP
''        Next i
''        cmdRunStop.BackColor = &H80&        ''stop
''        txtMaxHH.Enabled = True
''    Else  ''stop
''        For i = 0 To 10
''            ucBINmon1(i).set_maxHH CLng(txtMaxHH)
''            ucBINmon1(i).scan_RUN
''        Next i
''        cmdRunStop.BackColor = &H8000&        ''run
''        txtMaxHH.Enabled = False
''    End If
''
        
End Sub

Private Sub Form_GotFocus()
'''
''Dim i
''    For i = 0 To 10
''        ucBINmon1(i).picCON_Cir1
''    Next i
End Sub

Private Sub Form_Load()

Dim i As Integer
Dim j As Integer

    If App.PrevInstance Then
       MsgBox "프로그램이 이미 실행되었습니다."
       Unload Me
       End
    End If
    
    lbUpTime.Caption = "Up_Time: " & Format(Now, "YYYY-MM-DD h:m:s")
    lbTimeNow.Caption = "RunTime: " & Format(Now, "YYYY-MM-DD h:m:s")
    
    frmMain.AutoRedraw = True

'    Me.Width = Screen.Width * (1280 / 1400)
'    Me.Height = Screen.Height * (1024 / 1050)

'    Me.Left = Screen.Width - Width
'    Me.Top = 0
'    frmMain.Move Screen.Width - Width, 0
    
    frmMain.Move 0, 0, Screen.Width, Screen.Height
    
    SinterNumber1 = GetSetting(App.Title, "Settings", "SinterNumber1", 1)
    SinterNumber2 = GetSetting(App.Title, "Settings", "SinterNumber2", 2)
    
    lbTitle.Caption = "[" & SinterNumber1 & "," & SinterNumber2 & "소결] BIN LEVEL MONITORING"
    
    AOdeepMAX = GetSetting(App.Title, "Settings", "DeepMax", 60)
    If AOdeepMAX < 10 Then AOdeepMAX = 10
    If AOdeepMAX > 99 Then AOdeepMAX = 99
    AOdeepFull = False
    AOdeepCNT = 0
    For i = 0 To 19
        For j = 0 To 99  ''AOdeepMAX
            AOdeep(i, j) = 0
            Hdeep(i, j) = 0
        Next j
    Next i
    
    chkUseBeckHoof = GetSetting(App.Title, "Settings", "UseBeckHoof", 1)
    If chkUseBeckHoof < 0 Or chkUseBeckHoof > 1 Then chkUseBeckHoof = 1
    
    chkUsePLC = GetSetting(App.Title, "Settings", "UsePLC", 0)
    If chkUsePLC < 0 Or chkUsePLC > 1 Then chkUsePLC = 0
    
    picTop.Left = 100
    picTop.Top = 100
    picTop.Height = 800   '''Height * 0.05 + 100
    picTop.Width = Width - 200
    ''''
        imgLogo1.Left = 100
        imgLogo1.Top = 100 ''100
        lbTitle.Left = (Width * 0.27)    ''+ 200  ''frTop.Width * 0.3
        lbTitle.Top = 50
        lbTitle.Height = 600
        lbTitle.Width = (Width * 0.5) - 500
        ''
        cmdExit.Top = 200
        cmdExit.Left = picTop.Width - 1200
        cmdHide.Top = 200
        cmdHide.Left = picTop.Width - 2600
        cmdRunStop.Top = 200
        cmdRunStop.Left = picTop.Width - 4000
        
        cmdCFG.Top = 200
        cmdCFG.Left = picTop.Width - 5000
        
        ''lbRelVersion.Top = 200
        ''lbRelVersion.Left = picTop.Width - 6050
        lbRelVersion = relVersion
        ''lbRelDate.Top = 400
        ''lbRelDate.Left = picTop.Width - 6050
        lbRelDate = relDate
        
    For i = 0 To 32
        AOdata(i) = 0
        AOdata2(i) = 0
    Next i
    
''    For i = 1 To 10
''        Load ucBINmon1(i)
''    Next i

''    For i = 0 To 10
''
''        ucBINmon1(i).Width = Width / 11 - 30
''        ucBINmon1(i).Left = (i * (Width / 11)) + 20
''        ucBINmon1(i).Height = 12200
''
''        ucBINmon1(i).Visible = True
''
''        DoEvents
''
''    Next i
'''''''''''''
    ucBINmon1(0).Width = Width / 11 - 30
    ucBINmon1(0).Left = 20  '''(i * (Width / 11)) + 20
    ucBINmon1(0).Height = 12200
    ucBINmon1(0).Visible = False '''= True
    
    
    
''    ucBINdps1(0).Width = Width / 11 - 30
''    ucBINdps1(0).Left = 2000  ''20  '''(i * (Width / 11)) + 20
''    ucBINdps1(0).Height = 6100  '''12200
''    ucBINdps1(0).Visible = True
    
    
    BinWidth = 1800  '''2000
    ''''''''
    
    For i = 1 To 9
        Load ucBINdps1(i)
        ''<--201706
    Next i
    ''''''
    For i = 0 To 9
    
        ucBINdps1(i).Top = 1400

        ucBINdps1(i).Width = BinWidth  '''Width / 11 - 50
        ucBINdps1(i).Left = (i * (Width / 11)) + 20 ''+ 1720
        ucBINdps1(i).Height = 6100  '''12200

        ucBINdps1(i).Visible = True

        DoEvents

    Next i
'''''''''''
    
    
    For i = 10 To 19
        Load ucBINdps1(i)
    Next i
    ''''''
    For i = 10 To 19
    
        ucBINdps1(i).Top = 7600

        ucBINdps1(i).Width = BinWidth  '''Width / 11 - 50
        ucBINdps1(i).Left = ((i - 10) * (Width / 11)) + 20 ''+ 1720
        ucBINdps1(i).Height = 6100  '''12200

        ucBINdps1(i).Visible = True

        DoEvents

    Next i
    
    
    txtSD1.Left = 100
    txtSD1.Top = Height - 1600
    txtSD1.Width = Width - 300
    txtSD1.Height = 1300


''''   if      ( pos == 0 )''''      m_pLmsClient[pos]->Connect( "192.168.0.21", 7001);
''''   else if ( pos == 1 )''''      m_pLmsClient[pos]->Connect( "192.168.0.21", 7002);
''''   else if ( pos == 2 )''''      m_pLmsClient[pos]->Connect( "192.168.0.21", 7003);
''''   else if ( pos == 3 )''''      m_pLmsClient[pos]->Connect( "192.168.0.21", 7004);
''''   else if ( pos == 4 )''''      m_pLmsClient[pos]->Connect( "192.168.0.22", 7001);
''''   else if ( pos == 5 )''''      m_pLmsClient[pos]->Connect( "192.168.0.22", 7002);
''''   else if ( pos == 6 )''''      m_pLmsClient[pos]->Connect( "192.168.0.22", 7003);
''''   else if ( pos == 7 )''''      m_pLmsClient[pos]->Connect( "192.168.0.22", 7004);
''''   else if ( pos == 8 )''''      m_pLmsClient[pos]->Connect( "192.168.0.22", 7005);
''''   else if ( pos == 9 )''''      m_pLmsClient[pos]->Connect( "192.168.0.22", 7006);
''''
''''//   else if ( pos == 10 )'''//  m_pLmsClient[pos]->Connect( "192.168.0.152", 7003);
''''
''''   /**
''''   else if ( pos == 11 )''''      m_pLmsClient[pos]->Connect( "192.168.0.31", 7002);
''''   else if ( pos == 12 )''''      m_pLmsClient[pos]->Connect( "192.168.0.31", 7003);
''''   else if ( pos == 13 )''''      m_pLmsClient[pos]->Connect( "192.168.0.31", 7004);
''''   else if ( pos == 14 )''''      m_pLmsClient[pos]->Connect( "192.168.0.32", 7001);
''''   else if ( pos == 15 )''''      m_pLmsClient[pos]->Connect( "192.168.0.32", 7002);
''''   else if ( pos == 16 )''''      m_pLmsClient[pos]->Connect( "192.168.0.32", 7003);
''''   else if ( pos == 17 )''''      m_pLmsClient[pos]->Connect( "192.168.0.32", 7004);
''''   else if ( pos == 18 )''''      m_pLmsClient[pos]->Connect( "192.168.0.32", 7005);
''''   else if ( pos == 19 )''''      m_pLmsClient[pos]->Connect( "192.168.0.32", 7006);
''''   **/

''    For i = 0 To 7
''        ''ipAddr(i) = "192.168.0.22"  ''151"
''        ipAddr(i) = "192.168.0.151"
''        ipPort(i) = Trim(Str(7001 + i))
''
''        ucBINmon1(i).setIDX i, ipAddr(i), ipPort(i)
''    Next i
''    ''''
''    For i = 8 To 10
''        ''ipAddr(i) = "192.168.0.21"  ''152"
''        ipAddr(i) = "192.168.0.152"
''        ipPort(i) = Trim(Str(7001 + i - 8))
''
''        ucBINmon1(i).setIDX i, ipAddr(i), ipPort(i)
''    Next i
''''''''''''''

''    ipAddr(0) = "192.168.0.21"  ''ipAddr(0) = "192.168.0.22"  ''151"
''    ipPort(0) = Trim(Str(7003))
''    ucBINdps1(0).setIDX 0, ipAddr(0), ipPort(0)
    ''
    
    If (SinterNumber1 = 3) And (SinterNumber2 = 4) Then
        txtPcsIP.Text = "172.24.55.27"
        txtPcsPort.Text = "8002"
        
        ''' Set default IP addr/port for sinter 1 and 2
        ipAddr(0) = "192.168.0.41"
        ipPort(0) = "7001"
        ipAddr(1) = "192.168.0.41"
        ipPort(1) = "7002"
        ipAddr(2) = "192.168.0.41"
        ipPort(2) = "7003"
        ipAddr(3) = "192.168.0.41"
        ipPort(3) = "7004"
        '''
        ipAddr(4) = "192.168.0.42"
        ipPort(4) = "7001"
        ipAddr(5) = "192.168.0.42"
        ipPort(5) = "7002"
        ipAddr(6) = "192.168.0.42"
        ipPort(6) = "7003"
        ipAddr(7) = "192.168.0.42"
        ipPort(7) = "7004"
        ipAddr(8) = "192.168.0.42"
        ipPort(8) = "7005"
        ipAddr(9) = "192.168.0.42"
        ipPort(9) = "7006"
        '''
        ipAddr(10) = "192.168.0.51"
        ipPort(10) = "7001"
        ipAddr(11) = "192.168.0.51"
        ipPort(11) = "7002"
        ipAddr(12) = "192.168.0.51"
        ipPort(12) = "7003"
        ipAddr(13) = "192.168.0.51"
        ipPort(13) = "7004"
        '''
        ipAddr(14) = "192.168.0.52"
        ipPort(14) = "7001"
        ipAddr(15) = "192.168.0.52"
        ipPort(15) = "7002"
        ipAddr(16) = "192.168.0.52"
        ipPort(16) = "7003"
        ipAddr(17) = "192.168.0.52"
        ipPort(17) = "7004"
        ipAddr(18) = "192.168.0.52"
        ipPort(18) = "7005"
        ipAddr(19) = "192.168.0.52"
        ipPort(19) = "7006"
    Else ' If (SinterNumber1 = 1) And (SinterNumber2 = 2) Then
        txtPcsIP.Text = "172.24.55.27"
        txtPcsPort.Text = "8001"
        ''' Set default IP addr/port for sinter 1 and 2
        ipAddr(0) = "192.168.0.21"
        ipPort(0) = "7001"
        ipAddr(1) = "192.168.0.21"
        ipPort(1) = "7002"
        ipAddr(2) = "192.168.0.21"
        ipPort(2) = "7003"
        ipAddr(3) = "192.168.0.21"
        ipPort(3) = "7004"
        '''
        ipAddr(4) = "192.168.0.22"
        ipPort(4) = "7001"
        ipAddr(5) = "192.168.0.22"
        ipPort(5) = "7002"
        ipAddr(6) = "192.168.0.22"
        ipPort(6) = "7003"
        ipAddr(7) = "192.168.0.22"
        ipPort(7) = "7004"
        ipAddr(8) = "192.168.0.22"
        ipPort(8) = "7005"
        ipAddr(9) = "192.168.0.22"
        ipPort(9) = "7006"
        '''
        ipAddr(10) = "192.168.0.31"
        ipPort(10) = "7001"
        ipAddr(11) = "192.168.0.31"
        ipPort(11) = "7002"
        ipAddr(12) = "192.168.0.31"
        ipPort(12) = "7003"
        ipAddr(13) = "192.168.0.31"
        ipPort(13) = "7004"
        '''
        ipAddr(14) = "192.168.0.32"
        ipPort(14) = "7001"
        ipAddr(15) = "192.168.0.32"
        ipPort(15) = "7002"
        ipAddr(16) = "192.168.0.32"
        ipPort(16) = "7003"
        ipAddr(17) = "192.168.0.32"
        ipPort(17) = "7004"
        ipAddr(18) = "192.168.0.32"
        ipPort(18) = "7005"
        ipAddr(19) = "192.168.0.32"
        ipPort(19) = "7006"
    End If
    
    Dim ipAddr_tmp As String
    Dim ipPort1_tmp As String
    Dim ipPort2_tmp As String
    For i = 0 To 19
        ipAddr_tmp = GetSetting(App.Title, "Settings", "BinIPAddr_" & i, "Fail")
        ipPort1_tmp = GetSetting(App.Title, "Settings", "BinIPPort_" & i, "Fail")
        If IsValidIPAddress(ipAddr_tmp) = False Then
            ipAddr_tmp = ipAddr(i)
            ''SaveSetting App.Title, "Settings", "BinIPAddr_" & i, ipAddr_tmp
        End If
        If IsValidIPPort(ipPort1_tmp) = False Then
            ipPort1_tmp = ipPort(i)
            ''SaveSetting App.Title, "Settings", "BinIPPort_" & i, ipPort1_tmp
        End If
        ucBINdps1(i).setIDX i, ipAddr_tmp, ipPort1_tmp
    Next i
    
    ipAddr_tmp = GetSetting(App.Title, "Settings", "PLCIPAddr", "Fail")
    ipPort1_tmp = GetSetting(App.Title, "Settings", "PLCIPPort1", "Fail")
    ipPort2_tmp = GetSetting(App.Title, "Settings", "PLCIPPort2", "Fail")
    If IsValidIPAddress(ipAddr_tmp) = False Then
        ipAddr_tmp = "192.168.0.2"
    End If
    If IsValidIPPort(ipPort1_tmp) = False Then
        ipPort1_tmp = "12001"
    End If
    If IsValidIPPort(ipPort2_tmp) = False Then
        ipPort2_tmp = "12002"
    End If
    
    With wsPLC1
        .Close
        .RemoteHost = ipAddr_tmp
        .RemotePort = ipPort1_tmp
        .LocalPort = ipPort1_tmp
        
        .Bind .LocalPort
    End With
    
    With wsPLC2
        .Close
        .RemoteHost = ipAddr_tmp
        .RemotePort = ipPort2_tmp
        .LocalPort = ipPort2_tmp
        
        .Bind .LocalPort
    End With
    
''    ucBINdps1(0).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(1).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(2).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(3).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(4).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(5).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(6).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(7).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(8).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(9).setOptionD "0", "0.53", "0.5"
''
''    ucBINdps1(10).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(11).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(12).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(13).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(14).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(15).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(16).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(17).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(18).setOptionD "0", "0.53", "0.5"
''    ucBINdps1(19).setOptionD "0", "0.53", "0.5"
    
    For i = 0 To 19  '''''''''''''''''''''';201705
        ucBINdps1(i).setOptionD "0", "0.6", "0.5"
    Next i
    
    ucBINdps1(6).setOptionD "0", "0.49", "0.5"
    ucBINdps1(7).setOptionD "0", "0.49", "0.5"
    ucBINdps1(16).setOptionD "0", "0.49", "0.5"
    ucBINdps1(17).setOptionD "0", "0.49", "0.5"
    
    
    
    
    
    
''    ucBINdps1(0).setOptionD "200", "0.53", "0.5"
''
''    ucBINmon1(1).setOptionD "0", "0.53", "0.5"
''    ucBINmon1(2).setOptionD "200", "0.53", "0.5"
''    ucBINmon1(3).setOptionD "0", "0.53", "0.5"
''    ucBINmon1(4).setOptionD "0", "0.53", "0.65"
''    ucBINmon1(5).setOptionD "0", "0.53", "0.65"
''    ucBINmon1(6).setOptionD "200", "0.53", "0.65"
''    ucBINmon1(7).setOptionD "200", "0.44", "0.5"
''    ucBINmon1(8).setOptionD "200", "0.44", "0.5"
''    ucBINmon1(9).setOptionD "0", "0.53", "0.5"
''    ucBINmon1(10).setOptionD "0", "0.53", "0.5"
    

    
    For i = 0 To 19  '''''''''''''''''''''';201705
        ucBINdps1(i).setBinID
        ''ucBINdps1(i).picCON_Cir1
        ''
        cboIDX.AddItem i + 1
    Next i

    
    For i = 0 To 19
        BinTYPE(i) = GetSetting(App.Title, "Settings", "BINtype_" & Trim(i), 211)
        '''''''
        '''SaveSetting App.Title, "Settings", "BINtype_" & Trim(i), BinTYPE(i)
    Next i
    '''
    For i = 0 To 19
        ucBINdps1(i).setScanTYPE BinTYPE(i) ''' 211  '''LMS-211  '''LD-LRS-3100,, DPS-2590
    Next i

    ''ucBINdps1(0).setScanTYPE 2590  '''''LD-LRS-3100,, DPS-2590
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   

    For i = 0 To 19
        BinMaxH(i) = GetSetting(App.Title, "Settings", "MaxH_" & Trim(i), 1850)
        BinMinH(i) = GetSetting(App.Title, "Settings", "MinH_" & Trim(i), 300)
    Next i
    
    For i = 0 To 19  ''12소결  '''9 '''10
        ucBINdps1(i).set_maxHHLH CLng(BinMaxH(i)), CLng(BinMinH(i))  '''CLng(txtMaxHH)
        '''''''''''''''''''''''
        ucBINdps1(i).rxMode = 0  ''7
        '''''''''''''''''''''''
        ''ucBINmon1(i).runCONN
    Next i

    Dim BinAngleTmp$, SensorAngleTmp$

    For i = 0 To 19
        BinAngleTmp = _
            GetSetting(App.Title, "Settings", "BinAngle_" & i, "Fail")
        SensorAngleTmp = _
            GetSetting(App.Title, "Settings", "SensorAngle_" & i, "Fail")
        If IsNumeric(BinAngleTmp) = False _
            Or CSng(CInt(Val(BinAngleTmp))) <> CSng(Val(BinAngleTmp)) _
            Or CInt(Val(BinAngleTmp)) > 10! Or CInt(Val(BinAngleTmp)) < -10! _
            Then
            BinAngleTmp = "0"
            SaveSetting App.Title, "Settings", "BinAngle_" & i, BinAngleTmp
        End If
        If IsNumeric(SensorAngleTmp) = False _
            Or CSng(CInt(Val(SensorAngleTmp))) <> CSng(Val(SensorAngleTmp)) _
            Or CInt(Val(SensorAngleTmp)) > 48! Or CInt(Val(SensorAngleTmp)) < -48! _
            Then
            SensorAngleTmp = "0"
            SaveSetting App.Title, "Settings", "SensorAngle_" & i, SensorAngleTmp
        End If
        ucBINdps1(i).setBinSettings CInt(BinAngleTmp), CInt(SensorAngleTmp)
    Next i
'''''''''''



    
'''    '''''''[TEST]''''LD-LRS-3100,, DPS-2590
'''    i = 20
'''    Load ucBINdps1(i)
'''    ''''''
'''        ucBINdps1(i).Top = 7600 ''4000  ''7600
'''        ''
'''        ucBINdps1(i).Width = BinWidth  '''Width / 11 - 50
'''        ucBINdps1(i).Left = ((i - 10) * (Width / 11)) + 20 ''+ 1720
'''        ucBINdps1(i).Height = 6100  '''12200
'''        ''
'''        ucBINdps1(i).Visible = True
'''        ''
'''        DoEvents
'''
'''        ucBINdps1(i).setScanTYPE 2590  '''''LD-LRS-3100,, DPS-2590
'''        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''        ucBINdps1(i).setIDX (i), "192.168.0.11", "8282"  ''"192.168.0.21", "7001"
'''
'''        ucBINdps1(i).setOptionD "0", "0.6", "0.5"
'''
'''        BinMaxH(i) = GetSetting(App.Title, "Settings", "MaxH_" & Trim(i), 1850)
'''        SaveSetting App.Title, "Settings", "MaxH_" & Trim(i), BinMaxH(i)
'''        ''''
'''        ucBINdps1(i).set_maxHH CLng(BinMaxH(i))  '''CLng(txtMaxHH)
'''        ucBINdps1(i).setBinID
'''        ''
'''        ucBINdps1(i).rxMode = 0
        
        
    
    
    cboIDX.ListIndex = 0
    cboIDX.Refresh
    

    lbVS1.Caption = Screen.Width & "x" & Screen.Height
'' _
''                    & ", " & ucBINmon1(0).Width & "x" & ucBINmon1(0).Height _
''                    & ", " & ucBINmon1(0).picGET_width & "x" & ucBINmon1(0).picGET_height


    BINLog vbCrLf & vbCrLf & Format(Now, "YYYYMMDD-hh:mm:ss") & " ====[SILO BIN-LEVEL START]===" & vbCrLf, SinterNumber1 & "소결"
    BINLog vbCrLf & vbCrLf & Format(Now, "YYYYMMDD-hh:mm:ss") & " ====[SILO BIN-LEVEL START]===" & vbCrLf, SinterNumber2 & "소결"



    tmrINIT.Interval = 5000
    tmrINIT.Enabled = True
    
    ''txtMaxHH.Enabled = False
    

''        cmdRunStop.BackColor = &H80&    ''stop
''        cmdRunStop_Click                ''<<RUN>>''


End Sub

Private Sub Form_Terminate()
    ''Return
End Sub

Private Sub tmrAoDo_Timer()

Dim i As Integer
Dim j As Integer
Dim ioD(33) As Integer
Dim str1 As String
Dim str2 As String


Dim aaD(33) As Integer

Dim avrD(20) As Integer
Dim avrDsum(20) As Long

Dim aaH(33) As Integer

Dim avrH(20) As Integer
Dim avrHsum(20) As Long

Dim UDPiV_1(29) As Integer  '''[16bit-word] to PLC : now-Use-10/30word!
Dim UDPiV_2(29) As Integer  '''[16bit-word] to PLC : now-Use-10/30word!
    
    lbTimeNow.Caption = "RunTime: " & Format(Now, "YYYY-MM-DD h:m:s")

    For i = 0 To 19
        aaD(i) = ucBINdps1(i).ret_AOd   '''' (1~32767)
        '''''''''''''''''''''''''''''
        aaH(i) = ucBINdps1(i).ret_Height
        ''''''''''''''''''''''''''''''''
    Next i
    
    ''Get--First!!
    For i = 0 To 19
        If (aaD(i) <= 0) Or (aaD(i) >= 32768) Then
            aaD(i) = GetSetting(App.Title, "Settings", "AV_" & Trim(i), 0)
        End If
    Next i

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''<AVR)
    For i = 0 To 19
        AOdeep(i, AOdeepCNT) = aaD(i)
        Hdeep(i, AOdeepCNT) = aaH(i)
    Next i
    ''
    AOdeepCNT = AOdeepCNT + 1
    ''
    If AOdeepCNT >= AOdeepMAX Then  ''99
        If AOdeepFull = False Then
            txtAVRcnt = AOdeepCNT & "/" & AOdeepMAX
            AOdeepFull = True
        End If
        AOdeepCNT = 0       ''''Loop!
    End If


    For i = 0 To 19
        avrDsum(i) = 0
        avrHsum(i) = 0
    Next i
    
    ''//??????????
    If AOdeepFull = True Then
    ''
        For i = 0 To 19
            For j = 0 To AOdeepMAX - 1
                avrDsum(i) = avrDsum(i) + AOdeep(i, j)
                avrHsum(i) = avrHsum(i) + Hdeep(i, j)
            Next j
            avrD(i) = CInt(avrDsum(i) / AOdeepMAX)
            avrH(i) = CInt(avrHsum(i) / AOdeepMAX)
        Next i
    ''
    ElseIf AOdeepCNT > 1 Then
    ''
        txtAVRcnt = AOdeepCNT & "/" & AOdeepMAX
        For i = 0 To 19
            For j = 0 To AOdeepCNT - 1
                avrDsum(i) = avrDsum(i) + AOdeep(i, j)
                avrHsum(i) = avrHsum(i) + Hdeep(i, j)
            Next j
            avrD(i) = CInt(avrDsum(i) / AOdeepCNT)
            avrH(i) = CInt(avrHsum(i) / AOdeepCNT)
        Next i
    ''
    Else
        txtAVRcnt = AOdeepCNT & "/" & AOdeepMAX
        For i = 0 To 19
            avrD(i) = aaD(i)
            avrH(i) = aaH(i)
        Next i
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''>AVR)

    ''set_avrHH for View
    For i = 0 To 19
        ucBINdps1(i).avrAOd = avrD(i)
        ucBINdps1(i).avrHeight = avrH(i)
    Next i

    
    ''Replace!!
    For i = 0 To 19
        aaD(i) = avrD(i)
    Next i

    ''Ready for PLC direct
    For i = 0 To 29
         UDPiV_1(i) = 0
         UDPiV_2(i) = 0
    Next i
    '''
    For i = 0 To 9
        UDPiV_1(i) = aaD(i)
    Next i
    ''
    For i = 10 To 19
        UDPiV_2(i - 10) = aaD(i)
    Next i
    
    ''SAVE--Replace!!
    For i = 0 To 19
        If (aaD(i) > 0) And (aaD(i) < 32768) Then
            SaveSetting App.Title, "Settings", "AV_" & Trim(i), aaD(i)
        Else
            aaD(i) = GetSetting(App.Title, "Settings", "AV_" & Trim(i), 32767)  ''0
        End If
    Next i


        ioD(0) = aaD(0)  ''ucBINdps1(0).ret_AOd
        ioD(1) = aaD(1)  ''ucBINdps1(1).ret_AOd
        ioD(2) = aaD(2)  ''ucBINdps1(2).ret_AOd
        ioD(3) = 1 ''0

        ioD(4) = aaD(3)  ''ucBINdps1(3).ret_AOd
        ioD(5) = aaD(4)  ''ucBINdps1(4).ret_AOd
        ioD(6) = aaD(5)  ''ucBINdps1(5).ret_AOd
        ioD(7) = 1 ''0

        ioD(8) = aaD(6)  ''ucBINdps1(6).ret_AOd
        ioD(9) = aaD(7)  ''ucBINdps1(7).ret_AOd
        ioD(10) = aaD(8)  ''ucBINdps1(8).ret_AOd
        ioD(11) = 1 ''0

        ioD(12) = aaD(9)  ''ucBINdps1(9).ret_AOd
        ''''''''''''''''''''''''''''''''''''
        ioD(13) = ioD(0)
        ioD(14) = ioD(1)
        ioD(15) = 1

        ioD(16) = ioD(2)
        ioD(17) = ioD(4)
        ioD(18) = ioD(5)
        ioD(19) = 1

        ioD(20) = ioD(6)
        ioD(21) = ioD(8)
        ioD(22) = ioD(9)
        ioD(23) = 1

        ioD(24) = ioD(10)
        ioD(25) = ioD(12)
        ioD(26) = 1
        ioD(27) = 1
''''

        For i = 0 To 27 ''31
    ''--------------------------------------------------------(Temp)
    ''        If (ioD(i) > 0) And (ioD(i) <= 32767) Then
    ''            AOdata(i) = ioD(i)
    ''        Else
    ''            Exit Sub
    ''            ''=========>> Cancle for Next~~ /(protect_Zero_send)
    ''        End If
    ''--------------------------------------------------------(Temp)

            AOdata(i) = ioD(i)
            ''''''''''''''''''
        Next i


        ioD(0) = aaD(10)  ''ucBINdps1(10).ret_AOd
        ioD(1) = aaD(11)  ''ucBINdps1(11).ret_AOd
        ioD(2) = aaD(12)  ''ucBINdps1(12).ret_AOd
        ioD(3) = 1 ''0

        ioD(4) = aaD(13)  ''ucBINdps1(13).ret_AOd
        ioD(5) = aaD(14)  ''ucBINdps1(14).ret_AOd
        ioD(6) = aaD(15)  ''ucBINdps1(15).ret_AOd
        ioD(7) = 1 ''0

        ioD(8) = aaD(16)  ''ucBINdps1(16).ret_AOd
        ioD(9) = aaD(17)  ''ucBINdps1(17).ret_AOd
        ioD(10) = aaD(18)  ''ucBINdps1(18).ret_AOd
        ioD(11) = 1 ''0

        ioD(12) = aaD(19)  ''ucBINdps1(19).ret_AOd
        ''''''''''''''''''''''''''''''''''''
        ioD(13) = ioD(0)
        ioD(14) = ioD(1)
        ioD(15) = 1

        ioD(16) = ioD(2)
        ioD(17) = ioD(4)
        ioD(18) = ioD(5)
        ioD(19) = 1

        ioD(20) = ioD(6)
        ioD(21) = ioD(8)
        ioD(22) = ioD(9)
        ioD(23) = 1

        ioD(24) = ioD(10)
        ioD(25) = ioD(12)
        ioD(26) = 1
        ioD(27) = 1
''''
        For i = 0 To 27 ''31
            AOdata2(i) = ioD(i)
            ''''''''''''''''''
        Next i


        If Len(txtSD1) > 6000 Then
            txtSD1 = Mid(txtSD1, 3000)
        End If
        txtSD1 = txtSD1 & vbCrLf & vbCrLf

        str1 = " <1> "
        For i = 0 To 12  ''31
            ''str1 = str1 & " [1-" & Format((i + 1), "00") & "]" & Format(AOdata(i), "00000")
            str1 = str1 & " [1-" & Format((i + 1), "00") & "]" & Format(CLng(AOdata(i)) * 100 / 32768, "00.0")
        Next i
        txtSD1 = txtSD1 & Format(Now, "YYYYMMDD-hh:mm:ss") & str1 & vbCrLf
        str2 = " <2> "
        For i = 0 To 12
            ''str2 = str2 & " [2-" & Format((i + 1), "00") & "]" & Format(AOdata2(i), "00000")
            str2 = str2 & " [2-" & Format((i + 1), "00") & "]" & Format(CLng(AOdata2(i)) * 100 / 32768, "00.0")
        Next i
        txtSD1 = txtSD1 & Format(Now, "YYYYMMDD-hh:mm:ss") & str2
        txtSD1.SelStart = Len(txtSD1)
        

        ''BINLog vbCrLf, "12소결"
        BINLog str1, SinterNumber1 & "소결"
        BINLog str2, SinterNumber2 & "소결"
        
    If (chkUsePLC = 1) Then
        Dim buffer(59) As Byte
        
        CopyMemory buffer(0), UDPiV_1(0), 30 * 2
        ''' Change little-endian to big-endian
        For i = 0 To 29
            swap buffer(i * 2), buffer(i * 2 + 1)
        Next i
        wsPLC1.SendData buffer
        
        CopyMemory buffer(0), UDPiV_2(0), 30 * 2
        ''' Change little-endian to big-endian
        For i = 0 To 29
            swap buffer(i * 2), buffer(i * 2 + 1)
        Next i
        wsPLC2.SendData buffer
    End If

    On Error GoTo wsErrADS

    If (chkUseBeckHoof = 1) Then
''''        AdsOcx1.AdsAmsServerNetId = "172.16.21.20.1.1"   '''AdsOcx1.AdsAmsClientNetId
''''        AdsOcx1.AdsAmsServerPort = 800
''''        AdsOcx1.EnableErrorHandling = True
''''        '''''''
''''        AdsOcx1.AdsSyncWriteReq &HF020&, &H64&, 64, AOdata  ''ioD
        If (SinterNumber1 = 3) And (SinterNumber2 = 4) Then
            AdsOcx1.AdsAmsServerNetId = "0.0.0.0.1.1" ''34소결!!
            AdsOcx1.AdsAmsServerPort = 301  ''800  ''3소결!!
            AdsOcx1.EnableErrorHandling = True
            ''''
            AdsOcx1.AdsSyncWriteReq &HF030&, &H0&, 56, AOdata
            
            AdsOcx1.AdsAmsServerNetId = "0.0.0.0.1.1" ''34소결!!
            AdsOcx1.AdsAmsServerPort = 302  ''800  ''4소결!!
            AdsOcx1.EnableErrorHandling = True
            ''''
            AdsOcx1.AdsSyncWriteReq &HF030&, &H0&, 56, AOdata2
        Else 'If (SinterNumber1 = 1) And (SinterNumber2 = 2) Then
            AdsOcx1.AdsAmsServerNetId = "0.0.0.0.0.0" ''12소결!!
            AdsOcx1.AdsAmsServerPort = 301  ''800  ''1소결!!
            AdsOcx1.EnableErrorHandling = True
            ''''
            AdsOcx1.AdsSyncWriteReq &HF030&, &H0&, 56, AOdata
            
            AdsOcx1.AdsAmsServerNetId = "0.0.0.0.0.0" ''12소결!!
            AdsOcx1.AdsAmsServerPort = 302  ''800  ''2소결!!
            AdsOcx1.EnableErrorHandling = True
            ''''
            AdsOcx1.AdsSyncWriteReq &HF030&, &H0&, 56, AOdata2
        End If
    End If

wsErrADS:
        '''''Just-Cancle...for next

        txtWSpcs = wsPcs.State

        If wsPcs.State = sckConnected Then
            '''''''''''''
            EditPcsData SinterNumber1
            ''DoEvents
            '''''''''''''
            EditPcsData SinterNumber2
            ''DoEvents
            '''''''''''''
            txtWSpcs.BackColor = vbGreen
        Else
            txtWSpcs.BackColor = vbRed  ''&HFF00FF
        End If
''''
End Sub


Private Sub swap(b1 As Byte, b2 As Byte)
 
  b1 = b1 Xor b2
  b2 = b1 Xor b2
  b1 = b1 Xor b2
 
End Sub


Private Sub EditPcsData(Pno As Integer)

'''Dim sendbuf(2295) As Byte  ''Variant  ''BIN5
Dim sendbuf(2087) As Byte  ''Variant  <--1,2소결 '''2088


Dim i As Integer
Dim j As Integer
Dim cnt1 As Integer
Dim ret1 As Long
Dim str1 As String
Dim L8 As Byte
Dim H8 As Byte
'''
Dim posUCidx As Integer


''struct SENDBUF
''{
''   short   head;           //0x1122 고정
''   short   size;           //Buffer 전체 Size
''   short   plant;          //1:1소결 2:2소결 3:3소결 4:4소결
''   short   spare;
''
''   short   linkstat[10];   //통신 상태  1:정상 0:이상
''   short   height[10];     //평균 높이
''   short   volume[10];     //용적       10-2 m3
''   short   data[10][101];  //BIN Level
''};
'''Connect( "172.24.55.27", 8001);

''0x0828 <= 2088 <= 2+2+2+2+20+20+20++2020 ::(1234)
'''''''''''''''''''''''''''''''''''''''''
''0x08f8 <= 2296 <= 2+2+2+2+22+22+22++2222 ::(5)
On Error GoTo wsErrPcs

'   sendbuf.head = 0x1122;
'   sendbuf.size = sizeof(sendbuf);
'   sendbuf.plant = Plant;

    sendbuf(0) = &H22:    sendbuf(1) = &H11     ''Header
    sendbuf(2) = &H28:    sendbuf(3) = &H8      ''Size <--1,2소결 '''2088==0x0828
    sendbuf(4) = &H1:     sendbuf(5) = &H0      ''Plant-No
    sendbuf(6) = &H0:     sendbuf(7) = &H0      ''spare



    If Pno = 4 Then
        sendbuf(4) = &H4  ''Plant-No
        posUCidx = 10
    ElseIf Pno = 3 Then
        sendbuf(4) = &H3  ''Plant-No
        posUCidx = 0
    ElseIf Pno = 2 Then
        sendbuf(4) = &H2  ''Plant-No
        posUCidx = 10
    ElseIf Pno = 1 Then
        sendbuf(4) = &H1  ''Plant-No
        posUCidx = 0
    Else
        Exit Sub  ''===>>>
    End If



    For i = 0 To 9
        sendbuf(8 + i * 2) = CByte(ucBINdps1(i + posUCidx).ret_Act) ''BIN_Comm_Act
        sendbuf(8 + i * 2 + 1) = &H0
        
''        If sendbuf(8 + i * 2) < 1 Then
''            Exit Sub ''=======================>>>Cancle PCS!!!
''        End If
    Next i
    For i = 0 To 9
        sendbuf(28 + i * 2) = CByte(ucBINdps1(i + posUCidx).ret_HH Mod 256)
        sendbuf(28 + i * 2 + 1) = CByte(ucBINdps1(i + posUCidx).ret_HH / 256) ''Height...AVR
    Next i
    For i = 0 To 9
        sendbuf(48 + i * 2) = CByte(ucBINdps1(i + posUCidx).ret_VV Mod 256)
        sendbuf(48 + i * 2 + 1) = CByte(ucBINdps1(i + posUCidx).ret_VV / 256) ''VVV...AVR
    Next i
    ''74''((+(11*202)==2222==>((2296))
    For j = 0 To 9
      If ucBINdps1(j + posUCidx).ret_Act > 0 Then
      ''''''''''''''''''''''''''''''''
        For i = 0 To 100
            ret1 = ucBINdps1(j + posUCidx).GETscanD(i)
            ''''''''''''''''''''''''''''''''''scan_Data''
            L8 = CByte(ret1 Mod 256)
            H8 = CByte(ret1 / 256)
            cnt1 = 68 + (j * 202) + (i * 2)
            sendbuf(cnt1) = L8
            sendbuf(cnt1 + 1) = H8
        Next i
      Else
        cnt1 = 68 + (j * 202)   '''201705~
        For i = 0 To 100
            sendbuf(cnt1 + (i * 2)) = 0:
            sendbuf(cnt1 + (i * 2) + 1) = 0
        Next i
      End If
    Next j

''On Error GoTo wsErrPcs

    txtWSpcs = wsPcs.State
    
    If wsPcs.State = sckConnected Then
        wsPcs.SendData sendbuf
        ''''''''''''''''''''''
''        str1 = "(" & Hex(Val(UBound(sendbuf))) & ") "
''        For i = 0 To UBound(sendbuf)
''            str1 = str1 & Format(Hex(sendbuf(i)), "00") & " "
''        Next i
''        txtSD1 = str1
    End If

    Exit Sub
    
wsErrPcs:
    wsPcs.Close
    DoEvents

End Sub


Private Sub tmrINIT_Timer()
    tmrINIT.Enabled = False
    
    Dim i
    
''    For i = 0 To 10
''        ucBINmon1(i).picCON_Cir1
''    Next i
'''''''''''''''
    ''ucBINmon1(0).picCON_Cir1
    

    tmrAoDo.Interval = 2000  '''3000  '''1000
    tmrAoDo.Enabled = True
    
    tmrPcs.Interval = 3000
    tmrPcs.Enabled = True
    
End Sub


Private Sub tmrPcs_Timer()

    If wsPcs.State <> sckConnected Then

        wsPcs.Close
    
        ''wsPcs.RemoteHost = "127.0.0.1"
        ''wsPcs.RemoteHost = "172.24.55.27"
        wsPcs.RemoteHost = txtPcsIP.Text
        
        ''wsPcs.RemotePort = "8003"
        wsPcs.RemotePort = txtPcsPort.Text
        
        wsPcs.Connect
    
    End If

    txtWSpcs = wsPcs.State
    
End Sub

''Private Sub ucBINmon1_upDXY(Index As Integer)
''    ''ucBINmon1(Index).ret_SDXY
''End Sub


Private Sub wsPcs_DataArrival(ByVal bytesTotal As Long)
Dim rBuf As Variant
    wsPcs.GetData rBuf  ''''null...
    
End Sub

Private Sub wsPcs_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    wsPcs.Close
    DoEvents
End Sub


