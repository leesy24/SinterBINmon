VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  '���� ����
   Caption         =   "1) 1�Ұ�BIN-01 ����"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdSettingsExit 
      Caption         =   "�� ��"
      Height          =   495
      Left            =   3600
      Style           =   1  '�׷���
      TabIndex        =   5
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdSettingsApply 
      Caption         =   "�� ��"
      Height          =   375
      Left            =   3600
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtSensorAngle 
      Height          =   270
      Left            =   1440
      TabIndex        =   3
      Top             =   550
      Width           =   735
   End
   Begin VB.TextBox txtBinAngle 
      Height          =   270
      Left            =   1440
      TabIndex        =   1
      Top             =   190
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "��, 48~-48"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "��, 48~-48"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lbSensorAngle 
      Caption         =   "���� ����"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lbBinAngle 
      Caption         =   "Bin ����"
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
Option Explicit
'
Dim Index%

Public Sub Init(Index_I%, Title_I$, BinAngle_I%, SeosorAngle_I%)
'
    Index = Index_I
'
    frmSettings.Caption = Title_I$ & " ����"
'
    txtBinAngle = BinAngle_I
    txtSensorAngle = SeosorAngle_I
'
End Sub

Private Sub cmdSettingsApply_Click()
'
    If IsNumeric(txtBinAngle) = False _
        Or CSng(CInt(Val(txtBinAngle))) <> CSng(Val(txtBinAngle)) _
        Or CInt(Val(txtBinAngle)) > 48! Or CInt(Val(txtBinAngle)) < -48! _
        Then
        MsgBox lbBinAngle & "�� 48 ~ -48 ������ ���� �� �̾�� �մϴ�.", vbOKOnly
        Exit Sub
    End If
    If IsNumeric(txtSensorAngle) = False _
        Or CSng(CInt(Val(txtSensorAngle))) <> CSng(Val(txtSensorAngle)) _
        Or CInt(Val(txtSensorAngle)) > 48! Or CInt(Val(txtSensorAngle)) < -48! _
        Then
        MsgBox lbSensorAngle & "�� 48 ~ -48 ������ ���� �� �̾�� �մϴ�.", vbOKOnly
        Exit Sub
    End If
'
    SaveSetting App.Title, "Settings", "BinAngle_" & Index, txtBinAngle
    SaveSetting App.Title, "Settings", "SensorAngle_" & Index, txtSensorAngle
'
    frmMain.ucBINdps1(Index).setBinSettings txtBinAngle, txtSensorAngle
'
End Sub

Private Sub cmdSettingsExit_Click()
'
    frmSettings.Visible = False
'
End Sub



