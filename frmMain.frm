VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{40DD8EA0-284B-11D0-A7B0-0020AFF929F4}#2.3#0"; "AdsOcx.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404000&
   BorderStyle     =   0  '없음
   Caption         =   "BIN5_Monitor"
   ClientHeight    =   12060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13590
   FillStyle       =   0  '단색
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12060
   ScaleWidth      =   13590
   ShowInTaskbar   =   0   'False
   Begin ADSOCXLib.AdsOcx AdsOcx1 
      Left            =   4800
      Top             =   1560
      _Version        =   131074
      _ExtentX        =   900
      _ExtentY        =   953
      _StockProps     =   0
      AdsAmsServerNetId=   "172.16.21.20.1.1"
      AdsAmsServerPort=   800
      AdsAmsClientPort=   32801
      AdsClientType   =   ""
      AdsClientAdsState=   ""
      AdsClientAdsControl=   ""
   End
   Begin VB.TextBox txtMaxHH 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   12360
      TabIndex        =   18
      Text            =   "1750"
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
      TabIndex        =   16
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtPcsPort 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1800
      TabIndex        =   15
      Text            =   "8003"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtPcsIP 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   480
      TabIndex        =   14
      Text            =   "172.24.55.27"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtWSpcs 
      Enabled         =   0   'False
      Height          =   270
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   180
   End
   Begin prjBIN5mon.ucBINmon ucBINmon1 
      Height          =   10455
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1935
      _extentx        =   3413
      _extenty        =   16748
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
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   8
      Top             =   8880
      Width           =   10575
   End
   Begin VB.CommandButton cmdADSclr 
      Caption         =   "ADSclr"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdADS1 
      Caption         =   "ADS1"
      Height          =   255
      Left            =   5040
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
         Left            =   7440
         Top             =   360
      End
      Begin VB.CommandButton cmdRunStop 
         BackColor       =   &H00008000&
         Caption         =   "RUN/STOP"
         Height          =   375
         Left            =   9360
         MaskColor       =   &H00E0E0E0&
         Style           =   1  '그래픽
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Timer tmrINIT 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   8040
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
      Begin VB.Label lbTitle 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "[5소결] BIN LEVEL MONITORING"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
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
         Caption         =   "(주)제일시스템"
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
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "MaxHigh:"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   11520
      TabIndex        =   19
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lbUpTime 
      BackStyle       =   0  '투명
      Caption         =   "UpTime"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lbVS1 
      BackStyle       =   0  '투명
      Caption         =   "Label1"
      Height          =   255
      Left            =   7680
      TabIndex        =   5
      Top             =   960
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



Dim d1 As Single



Dim ipAddr(11) As String
Dim ipPort(11) As String

Dim AOdata(33) As Integer



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


Private Sub cmdDmon_Click()
    txtSD1 = ucBINmon1(cboIDX.ListIndex).ret_SDXY
End Sub


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
  
    If cmdRunStop.BackColor = &H8000& Then  ''run
        For i = 0 To 10
            ucBINmon1(i).scan_STOP
        Next i
        cmdRunStop.BackColor = &H80&        ''stop
        txtMaxHH.Enabled = True
    Else  ''stop
        For i = 0 To 10
            ucBINmon1(i).set_maxHH CLng(txtMaxHH)
            ucBINmon1(i).scan_RUN
        Next i
        cmdRunStop.BackColor = &H8000&        ''run
        txtMaxHH.Enabled = False
    End If
    
        
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

    If App.PrevInstance Then
       MsgBox "프로그램이 이미 실행되었습니다."
       Unload Me
       End
    End If
    
    lbUpTime.Caption = "UpTime: " & Format(Now, "YYYY-MM-DD h:m:s")
    
    frmMain.AutoRedraw = True

'    Me.Width = Screen.Width * (1280 / 1400)
'    Me.Height = Screen.Height * (1024 / 1050)

'    Me.Left = Screen.Width - Width
'    Me.Top = 0
'    frmMain.Move Screen.Width - Width, 0
    
    frmMain.Move 0, 0, Screen.Width, Screen.Height
    
    picTop.Left = 100
    picTop.Top = 100
    picTop.Height = 800   '''Height * 0.05 + 100
    picTop.Width = Width - 200
    ''''
        imgLogo1.Left = 100
        imgLogo1.Top = 100 ''100
        lbTitle.Left = (Width * 0.32)    ''+ 200  ''frTop.Width * 0.3
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
        
    For i = 0 To 32
        AOdata(i) = 0
    Next i
    
    For i = 1 To 10
        Load ucBINmon1(i)
    Next i

    For i = 0 To 10

        ucBINmon1(i).Width = Width / 11 - 30
        ucBINmon1(i).Left = (i * (Width / 11)) + 20
        ucBINmon1(i).Height = 12200

        ucBINmon1(i).Visible = True
        
        DoEvents
        
    Next i

    txtSD1.Left = 100
    txtSD1.Top = Height - 1300
    txtSD1.Width = Width - 200
    txtSD1.Height = 1200


    For i = 0 To 7
        ''ipAddr(i) = "192.168.0.22"  ''151"
        ipAddr(i) = "192.168.0.151"
        ipPort(i) = Trim(Str(7001 + i))
        
        ucBINmon1(i).setIDX i, ipAddr(i), ipPort(i)
    Next i
    ''''
    For i = 8 To 10
        ''ipAddr(i) = "192.168.0.21"  ''152"
        ipAddr(i) = "192.168.0.152"
        ipPort(i) = Trim(Str(7001 + i - 8))
        
        ucBINmon1(i).setIDX i, ipAddr(i), ipPort(i)
    Next i
    
    For i = 0 To 10
        ucBINmon1(i).setBinID
        ucBINmon1(i).picCON_Cir1
        
        cboIDX.AddItem i + 1
    Next i
    cboIDX.ListIndex = 0
    cboIDX.Refresh
    
    
    ucBINmon1(0).setOptionD "200", "0.53", "0.5"
    ucBINmon1(1).setOptionD "0", "0.53", "0.5"
    ucBINmon1(2).setOptionD "200", "0.53", "0.5"
    ucBINmon1(3).setOptionD "0", "0.53", "0.5"
    ucBINmon1(4).setOptionD "0", "0.53", "0.65"
    ucBINmon1(5).setOptionD "0", "0.53", "0.65"
    ucBINmon1(6).setOptionD "200", "0.53", "0.65"
    ucBINmon1(7).setOptionD "200", "0.44", "0.5"
    ucBINmon1(8).setOptionD "200", "0.44", "0.5"
    ucBINmon1(9).setOptionD "0", "0.53", "0.5"
    ucBINmon1(10).setOptionD "0", "0.53", "0.5"
    

    lbVS1.Caption = Screen.Width & "x" & Screen.Height _
                    & ", " & ucBINmon1(0).Width & "x" & ucBINmon1(0).Height _
                    & ", " & ucBINmon1(0).picGET_width & "x" & ucBINmon1(0).picGET_height


    tmrINIT.Interval = 5000
    tmrINIT.Enabled = True
    
    txtMaxHH.Enabled = False
    
    For i = 0 To 10
        
        ucBINmon1(i).set_maxHH CLng(txtMaxHH)
    
        '''''''''''''''''''''''
        ucBINmon1(i).rxMode = 0  ''7
        '''''''''''''''''''''''
        
        ''ucBINmon1(i).runCONN
        
    Next i

''        cmdRunStop.BackColor = &H80&    ''stop
''        cmdRunStop_Click                ''<<RUN>>''


End Sub


Private Sub Form_Terminate()
    ''Return
End Sub


Private Sub tmrAoDo_Timer()

Dim i As Integer
Dim ioD(33) As Integer
Dim str1 As String

    ioD(0) = ucBINmon1(0).ret_AOd  ''32768 * 0.05       ''1
    ioD(1) = ucBINmon1(1).ret_AOd   ''32768 * 0.1        ''2
    ioD(2) = ucBINmon1(2).ret_AOd   ''32768 * 0.15       ''3
    ioD(3) = 1 ''0

    ioD(4) = ucBINmon1(3).ret_AOd   ''32768 * 0.2        ''4
    ioD(5) = ucBINmon1(4).ret_AOd   ''32768 * 0.25       ''5
    ioD(6) = ucBINmon1(5).ret_AOd   ''32768 * 0.3        ''6
    ioD(7) = 1 ''0

    ioD(8) = ucBINmon1(6).ret_AOd   ''32768 * 0.35       ''7
    ioD(9) = ucBINmon1(7).ret_AOd   ''32768 * 0.4        ''8
    ioD(10) = ucBINmon1(8).ret_AOd   ''32768 * 0.45      ''9
    ioD(11) = 1 ''0

    ioD(12) = ucBINmon1(9).ret_AOd   ''32768 * 0.5       ''10
    ioD(13) = ucBINmon1(10).ret_AOd   ''32768 * 0.55      ''11
    ''''''''''''''''''''''''''''''''''''
    ioD(14) = ioD(0)      ''1
    ioD(15) = 1 ''0

    ioD(16) = ioD(1)       ''2
    ioD(17) = ioD(2)      ''3
    ioD(18) = ioD(4)      ''4
    ioD(19) = 1 ''0

    ioD(20) = ioD(5)      ''5
    ioD(21) = ioD(6)       ''6
    ioD(22) = ioD(8)      ''7
    ioD(23) = 1 ''0

    ioD(24) = ioD(9)      ''8
    ioD(25) = ioD(10)      ''9
    ioD(26) = ioD(12)       ''10
    ioD(27) = 1 ''0

    ioD(28) = ioD(13)      ''11
    ioD(29) = 1 ''0
    ioD(30) = 1 ''0
    ioD(31) = 1 ''0
    
    
    For i = 0 To 31
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
    
    
    If Len(txtSD1) > 6000 Then
        txtSD1 = Mid(txtSD1, 3000)
    End If
    txtSD1 = txtSD1 & vbCrLf
    
    str1 = ""
    For i = 0 To 13  ''31
        str1 = str1 & " [" & Format((i + 1), "00") & "]" & Format(AOdata(i), "00000")
    Next i
    txtSD1 = txtSD1 & str1
    txtSD1.SelStart = Len(txtSD1)


On Error GoTo wsErrADS

    AdsOcx1.AdsAmsServerNetId = "172.16.21.20.1.1"   '''AdsOcx1.AdsAmsClientNetId
    AdsOcx1.AdsAmsServerPort = 800
    AdsOcx1.EnableErrorHandling = True
    
    AdsOcx1.AdsSyncWriteReq &HF020&, &H64&, 64, AOdata  ''ioD

wsErrADS:
    '''''Just-Cancle...for next

    txtWSpcs = wsPcs.State
    
    If wsPcs.State = sckConnected Then
        '''''''''''''
        EditPcsData 5
        '''''''''''''
    End If
    
End Sub


Private Sub EditPcsData(Pno As Integer)

Dim sendbuf(2295) As Byte  ''Variant
Dim i As Integer
Dim j As Integer
Dim cnt1 As Integer
Dim ret1 As Integer
Dim str1 As String
Dim L8 As Byte
Dim H8 As Byte


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


'   sendbuf.head = 0x1122;
'   sendbuf.size = sizeof(sendbuf);
'   sendbuf.plant = Plant;

    sendbuf(0) = &H22:    sendbuf(1) = &H11     ''Header
    sendbuf(2) = &HF8:    sendbuf(3) = &H8      ''Size
    sendbuf(4) = &H5:     sendbuf(5) = &H0      ''Plant-No
    sendbuf(6) = &H0:     sendbuf(7) = &H0      ''spare


    For i = 0 To 10
        sendbuf(8 + i * 2) = CByte(ucBINmon1(i).ret_Act)   ''BIN_Comm_Act
        sendbuf(8 + i * 2 + 1) = &H0
        
''        If sendbuf(8 + i * 2) < 1 Then
''            Exit Sub ''=======================>>>Cancle PCS!!!
''        End If
        
    Next i
    For i = 0 To 10
        sendbuf(30 + i * 2) = CByte(ucBINmon1(i).ret_HH Mod 256)
        sendbuf(30 + i * 2 + 1) = CByte(ucBINmon1(i).ret_HH / 256) ''Height...AVR
    Next i
    For i = 0 To 10
        sendbuf(52 + i * 2) = CByte(ucBINmon1(i).ret_VV Mod 256)
        sendbuf(52 + i * 2 + 1) = CByte(ucBINmon1(i).ret_VV / 256) ''VVV...AVR
    Next i
    ''74''((+(11*202)==2222==>((2296))
    For j = 0 To 10
      If ucBINmon1(j).ret_Act > 0 Then
      ''''''''''''''''''''''''''''''''
        For i = 0 To 100
            ret1 = ucBINmon1(j).GETscanD(i)
            ''''''''''''''''''''''''''''''''''scan_Data''
            L8 = CByte(ret1 Mod 256)
            H8 = CByte(ret1 / 256)
            cnt1 = 74 + (j * 202) + (i * 2)
            sendbuf(cnt1) = L8
            sendbuf(cnt1 + 1) = H8
        Next i
      Else
        For i = 0 To 100
            sendbuf(cnt1) = 0
            sendbuf(cnt1 + 1) = 0
        Next i
      End If
    Next j

On Error GoTo wsErrPcs

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
    For i = 0 To 10
        ucBINmon1(i).picCON_Cir1
    Next i

    tmrAoDo.Interval = 1000
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


Private Sub ucBINmon1_upDXY(Index As Integer)
    ''ucBINmon1(Index).ret_SDXY
End Sub

Private Sub ucBINmon1_Click(Index As Integer)

End Sub

Private Sub wsPcs_DataArrival(ByVal bytesTotal As Long)
Dim rBuf As Variant
    wsPcs.GetData rBuf  ''''null...
    
End Sub

Private Sub wsPcs_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    wsPcs.Close
    DoEvents
End Sub


