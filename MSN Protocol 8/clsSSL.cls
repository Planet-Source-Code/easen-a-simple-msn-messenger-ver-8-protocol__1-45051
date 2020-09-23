VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is my first ever submission to Planet Source Code.  I have
'found PSC extremely useful to me over the years, and wanted to give
'back to the community.  Thanks to all the other generous coders
'out there.  This example shows how the Secure Socket Layer (SSL)
'version 2.0 protocol works.  After looking all over the Internet
'for a VB example of this I soon realized that there were no such
'examples and would have to create my own implemention of it.  It is
'compatible with any SSL server; to see an example, click the
'"Connect to Server" button, and after it says connected click
'"Send Data".  This should retrieve PayPal's HTML and display it in
'the text box.  However, it used HTTPS or the secure version of
'HTTP to get it. :-)
'
'SSL is based on public key cryptography, it works in the following manner:
'
'client-hello         C -> S: challenge, cipher_specs
'server-hello         S -> C: connection-id,server_certificate,cipher_specs
'client-master-key    C -> S: {master_key}server_public_key
'client-finish        C -> S: {connection-id}client_write_key
'server-verify        S -> C: {challenge}server_write_key
'server-finish        S -> C: {new_session_id}server_write_key
'
'First the Client sends some random data known as the CHALLENGE, along with a list of ciphers it can use, for simplicity we will only use 128-bit RC4 with MD5
'The Server responds with a random data, known as the CONNECTION-ID, and the Server's Certificate and list of cipher specs
'The Client extracts the Public Key from the Server's Certificate then uses it to Encrypt a randomly generated Master Key, this Key then sent to the Server
'The Client and Server both generate 2 keys each by hashing the Master Key with other values, and the client sends a finish message, encrypted with the client write key
'The Server Responds by returning the CHALLENGE encrypted using the Client Read Key, this proves to the Clinet that the Server is who it says its is
'The Server sends its finish message, which consists of a randomly generated value, this value can be used to re-create the session in a new connection, but that is not supported in this example
'
'For a more detailed explaination of the protocol, please see: http://colossus.net/SSL.html
'To visit my website, and check out my communications service created in Visual Basic, go to http://www.gcn.cx/
'If you need to e-mail me you can try contacting me at jason@gcn.cx, but take notice I don't check my box very often
'
'Oh, and you are free to use this code in any of your commerical or non-commercial applications, but if you redistribute this source code I ask that you preserve these comments.
'
'Thank you, and enjoy.
Public WithEvents winsock1 As Winsock
Attribute winsock1.VB_VarHelpID = -1
Public Event RecvData(sdata As String)
Public Event Connected()
Public Event Disconnected()
Public bConnected As Boolean
Public Sub RecvData(sdata)
RaiseEvent RecvData(CStr(sdata))
End Sub
Public Function SendCommand(sCommand) As Boolean
    'Send Encrypted Record if Ready
    If Layer = 3 Then
        Call SSLSend(winsock1, sCommand & vbCrLf)
        SendCommand = True
    End If
End Function
Private Sub Class_Initialize()
Set winsock1 = New Winsock
End Sub

Private Sub Class_Terminate()
Set winsock1 = Nothing
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
    Dim TheData As String
    Dim ReachLen As Long
    Do
        If SeekLen = 0 Then
            If bytesTotal >= 2 Then
                winsock1.GetData TheData, vbString, 2
                SeekLen = BytesToLen(TheData)
                bytesTotal = bytesTotal - 2
            Else
                Exit Sub
            End If
        End If
        If bytesTotal >= SeekLen Then
            winsock1.GetData TheData, vbString, SeekLen
            bytesTotal = bytesTotal - SeekLen
        Else
            Exit Sub
        End If
        Select Case Layer
            Case 0:
                ENCODED_CERT = Mid(TheData, 12, BytesToLen(Mid(TheData, 6, 2)))
                CONNECTION_ID = Right(TheData, BytesToLen(Mid(TheData, 10, 2)))
                Call IncrementRecv
                Call SendMasterKey(winsock1)
            Case 1:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If Right(TheData, Len(CHALLENGE_DATA)) = CHALLENGE_DATA Then
                    If VerifyMAC(TheData) Then Call SendClientFinish(winsock1)
                Else
                    winsock1.Close
                End If
             Case 2:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If VerifyMAC(TheData) = False Then winsock1.Close
                Layer = 3
             Case 3:
                TheData = SecureSession.RC4_Decrypt(TheData)
                If VerifyMAC(TheData) Then Call ProcessData(Mid(TheData, 17))
        End Select
        SeekLen = 0
    Loop Until bytesTotal = 0
'Debug.Print "SSL:" & TheData
End Sub
Public Sub ConnectSSL(sURL)
    'Open Socket to Remote Server
    winsock1.Close
    winsock1.Connect sURL, 443
End Sub
Public Sub Winsock1_Close()
    'Close Socket
    Layer = 0
    winsock1.Close
    Set SecureSession = Nothing
    RaiseEvent Disconnected
End Sub
Private Sub Winsock1_Connect()
    'Send Client Hello
    Set SecureSession = New clsCrypto
    Call SendClientHello(winsock1)
    If bConnected = False Then RaiseEvent Connected
    bConnected = True
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Call Close Sub
    Winsock1_Close
End Sub

