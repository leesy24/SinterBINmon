Attribute VB_Name = "modLog"
Option Explicit

''----------------''
Public Declare Function GetFileAttributes Lib "kernel32" Alias _
                        "GetFileAttributesA" (ByVal lpFileName As String) As Long

Function FileExists(ByVal strPathName As String) As Boolean
  Dim af As Long
    af = GetFileAttributes(strPathName)
    FileExists = ((af <> -1) And af <> vbDirectory)
End Function

Public Sub BINLog(data1 As String, sName As String)
    
Dim f1
Dim f2
Dim Fname     As String
Dim str1      As String
Dim str2      As String
Dim SaveDir   As String
 
''//Date-LotNo.log '' filename
    
    str1 = Format(Now, "YYYYMMDD")
    str2 = Format(Now, "YYYYMMDD-hh:mm:ss")
    
    ''SaveDir = App.Path & "\" & str1 & ".Log"
    SaveDir = "C:\BIN_LOG" ''& "\" & str1 & ".Log"
    

On Error GoTo errFile1

    If Dir(SaveDir, vbDirectory) = "" Then
        MkDir SaveDir
    End If
    
    Fname = SaveDir & "\" & sName & "_" & str1 & ".log"  ''"_DataFile.log"
    
    If Not FileExists(Fname) Then
        f1 = FreeFile
        Open Fname For Binary Access Write As #f1
            ''Put #f1, , "DAC-LOG :: " + Fname + vbCrLf + vbCrLf
            Put #f1, , str2 & " " & data1$ & vbCrLf
        Close #f1
        DoEvents
        'Sleep 10
    Else
    
         f2 = FreeFile
        Open Fname For Binary Access Write As #f2
            Seek #f2, LOF(f2) + 1
            Put #f2, , str2 & " " & data1$ & vbCrLf
            ''Put #f2, , vbCrLf & data1$
        Close #f2
        DoEvents
    
    End If

errFile1:
    SaveDir = ""
    ''''''''''''(just-cancle~)
    
End Sub

Public Function IsValidIPAddress(ByVal strIPAddress As String) As Boolean
    On Error GoTo Handler
    Dim varAddress As Variant, n As Long, lCount As Long
    
    IsValidIPAddress = False
    varAddress = Split(strIPAddress, ".", , vbTextCompare)
    '//
    If IsArray(varAddress) Then
        For n = LBound(varAddress) To UBound(varAddress)
            lCount = lCount + 1
            varAddress(n) = CByte(varAddress(n))
        Next
        '//
        IsValidIPAddress = (lCount = 4)
    End If
    '//
Handler:
End Function

Public Function IsValidIPPort(ByVal strIPPort As String) As Boolean
    On Error GoTo Handler
    
    IsValidIPPort = False
    '//
    If IsNumeric(strIPPort) = True _
        And CSng(CInt(Val(strIPPort))) = CSng(Val(strIPPort)) _
        And CInt(Val(strIPPort)) <= 65535! And CInt(Val(strIPPort)) >= 1024! _
        Then
        IsValidIPPort = True
    End If
    '//
Handler:
End Function

