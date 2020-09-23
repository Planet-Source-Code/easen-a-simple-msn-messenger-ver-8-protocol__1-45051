VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMsgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This was all done by Marc Easen, UK
' I if it wasn't my code i have said,
' even when i have edited for my use.
' Cheers Please Put my name on this code
' Marc Easen - ihatetownies666@hotmail.com

Public Event ClientConnected()
Public Event ClientDisconnected()
Public WithEvents wsMain As Winsock
Attribute wsMain.VB_VarHelpID = -1
Public WithEvents SSL As clsSSL
Attribute SSL.VB_VarHelpID = -1
Public colFLlist As Collection
Public colRLlist As Collection
Public colALlist As Collection
Public colBLlist As Collection
Public sUserName As String
Public sPassword As String
Public sFriendlyName As String
Public sStatus As String
Private iTryID As Integer
Private bReconnectToNewServer As Boolean
Private sLastProtocolCommand, iDontDoLastLine
Private sMSNMsgrServer As String, iMSNMsgrPort As Integer
Private sSSLURL, sCookie1, sCookie2, GotFirstSSL As Boolean, SSLDATAPART
Private AuthCode
Public Function TryID() As Variant
iTryID = iTryID + 1
TryID = iTryID
End Function
Public Sub Login(sUser, sPass)
sUserName = CStr(sUser)
sPassword = CStr(sPass)
Set wsMain = Nothing 'Just sets the old winsock as Nothing (deletes)
Set wsMain = New Winsock 'So every time the user logs in there is always a new winsock in use
sMSNMsgrServer = "207.46.106.186"
iMSNMsgrPort = 1863
wsMain.Connect sMSNMsgrServer, iMSNMsgrPort
End Sub

Private Sub SSL_Connected()
begin:
If GotFirstSSL = False Then
If SSL.SendCommand("GET https://loginnet.passport.com/login.srf?" & sSSLURL & vbCrLf) = False Then
    DoEvents
    GoTo begin:
End If
Else
If sCookie1 = "" Or sCookie2 = "" Then Exit Sub
sURL = "GET https://loginnet.passport.com/ppsecure/post.srf?lc=1033&id=507&tw=20&cbid=507&da=passport.com&login=" & sUserName & "&domain=hotmail.com&passwd=" & sPassword & "&sec=&mspp_shared=&padding= HTTP/1.0" & vbCrLf & _
        "Cookie: " & sCookie1 & vbCrLf & _
        "Cookie: " & sCookie2 & vbCrLf
'Debug.Print sURL
Redo:
If SSL.SendCommand(sURL) = False Then 'Second Part
    DoEvents
    GoTo Redo:
End If
    'MsgBox "SENT"
End If
End Sub

Private Sub SSL_Disconnected()
'MsgBox "CLOSE"
End Sub

Private Sub SSL_RecvData(sdata As String)
On Error Resume Next
If SSLDATAPART <> "" Then
    sdata = SSLDATAPART & sdata
    SSLDATAPART = ""
End If
If InStr(LCase(sdata), "</html>") Then
If GotFirstSSL = False Then
    Dim vLine
    vLine = Split(sdata, vbCrLf)
    For i = 1 To UBound(vLine)
        Dim vSep
        If vLine(i) <> "" Then
            vSep = Split(vLine(i), ": ")
            If vSep(0) = "Set-Cookie" Then
                If sCookie1 = "" Then
                    sCookie1 = vSep(1) 'Got first Cookie
                Else
                    sCookie2 = vSep(1) 'Got Second Cookie
                End If

            End If
        End If
    Next i
    Set SSL = New clsSSL
    SSL.ConnectSSL "loginnet.passport.com" 'Reconnect To Send Cookies
    DoEvents
    GotFirstSSL = True
Else
    Dim vTemp1, vTemp2
    vTemp1 = Split(sdata, "URL=http://messenger.msn.com/download/passportdone.asp?did=1&")
    vTemp2 = Split(vTemp1(1), "><script>function OnBack(){}</script></HEAD>")
    AuthCode = Mid(vTemp2(0), 1, Len(vTemp2(0)) - 1)
    If AuthCode = "" Then MsgBox "Wrong Password": Exit Sub
    sTemp = "USR " & TryID & " TWN S " & AuthCode & vbCrLf
    wsMain.SendData sTemp
    Debug.Print "C: " & sTemp
End If
Else
SSLDATAPART = sdata
End If
End Sub

Private Sub wsMain_Close()
If bReconnectToNewServer = False Then RaiseEvent ClientDisconnected
End Sub
Private Sub wsMain_Connect()
Dim sTemp
sTemp = "VER " & TryID & " MSNP8 CVR0" & vbCrLf
wsMain.SendData sTemp 'I have used sTemp so i add it to a log later
If bReconnectToNewServer = False Then RaiseEvent ClientConnected
If bReconnectToNewServer = True Then bReconnectToNewServer = False
Debug.Print "C: " & sTemp
End Sub
Private Sub wsMain_DataArrival(ByVal bytesTotal As Long)
Dim sdata As String, vData As Variant, iDontDoLastLine As Integer
wsMain.GetData sdata
'sData = StrConv(CStr(sData), vbUnicode)
If sLastProtocolCommand <> "" Then              'This is used because the MSN Msgr Server can only send so many characters at a time,
    sdata = sLastProtocolCommand & sdata        'and the command are half there so this part saves the first part and adds it to the
    sLastProtocolCommand = ""                   'next part.
End If
vData = Split(sdata, vbCrLf)
If Right(sdata, 2) <> vbCrLf Then               'Sees if the last 2 chars are vbcrlf, if not the data was cut into 2 parts, and puts
    sLastProtocolCommand = vData(UBound(vData)) 'them back togteher making a full command
    iDontDoLastLine = 1
End If
For i = 0 To UBound(vData) - iDontDoLastLine
    If vData(i) <> "" Then
        ClientRespond CStr(vData(i))
        Debug.Print "S: " & vData(i)
    End If
Next i
End Sub
Public Sub ClientRespond(sdata As String)
Dim FirstChars
FirstChars = Split(sdata, " ")
Select Case FirstChars(0)
Case "VER"
    sTemp = "CVR " & TryID & " 0x0409 winnt 5.1 i386 MSNMSGR 5.0.0543 MSMSGS " & sUserName & vbCrLf
Case "CVR"
    sTemp = "USR " & TryID & " TWN I " & sUserName & vbCrLf
Case "XFR"
    NewServer = Split(FirstChars(3), ":")
    bReconnectToNewServer = True
    With wsMain
        .Close
        .RemoteHost = NewServer(0)
        .RemotePort = NewServer(1)
        .Connect
    End With
Case "USR"
    If FirstChars(2) = "MD5" Then
        EncryptPass = modMD5.MSNEncryptPw(FirstChars(4) & sPassword)
        sTemp = "USR " & TryID & " MD5 S " & EncryptPass & vbCrLf
    End If
    If FirstChars(2) = "TWN" Then 'thanks to Jason K. Resch for the SSL Part :D Cheers Mate
        Set SSL = New clsSSL
        sSSLURL = Replace(DecodeNick(FirstChars(4)), ",", "&")
        sSSLURL = sSSLURL '& " HTTP/1.0"
        SSL.ConnectSSL "loginnet.passport.com"
    End If
    If FirstChars(2) = "OK" Then
        sFriendlyName = CStr(FirstChars(4)) 'Get Name
    End If
Case "MSPAuth:"
    sTemp = "SYN " & TryID & " 1" & vbCrLf
End Select
If sTemp <> "" Then wsMain.SendData sTemp: Debug.Print "C: " & sTemp
End Sub
Public Function EncodeNick(Text)
    For i = 1 To Len(Text)
        Char = Mid(Text, i, 1)
        Char = Hex(Asc(Char))
        If Len(Char) = 1 Then
            Char = "%0" & Char
        Else
            Char = "%" & Char
        End If
        WholeString = WholeString & Char
    Next
    EncodeNick = WholeString
End Function
Function DecodeNick(Text)
    For i = 1 To Len(Text)
        PartStr = Mid(Text, i, 3)
        If Left(PartStr, 1) = "%" Then
            MYHEX$ = Right(Mid(Text, i, 3), 2)
            Mydec& = Val("&H" & MYHEX$)
            HexThing = Mydec&
            PartWhole = PartWhole & Chr(HexThing)
            i = i + 2
        ElseIf Left(PartStr, 1) = "Â" Then
            i = i + 1
            PartWhole = PartWhole & Mid(Text, i, 1)
        Else
            PartWhole = PartWhole & Mid(Text, i, 1)
        End If
    Next
   DecodeNick = PartWhole
End Function
