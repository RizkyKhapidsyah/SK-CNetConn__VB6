VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNetConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum NetConnTypeConstants
   INTERNET_CONNECTION_MODEM = &H1&
   INTERNET_CONNECTION_LAN = &H2&
   INTERNET_CONNECTION_PROXY = &H4&
   INTERNET_RAS_INSTALLED = &H10&
   INTERNET_CONNECTION_OFFLINE = &H20&
   INTERNET_CONNECTION_CONFIGURED = &H40&
End Enum


Private Const RAS_MAXENTRYNAME As Integer = 256
Private Const RAS_MAXDEVICETYPE As Integer = 16
Private Const RAS_MAXDEVICENAME As Integer = 128
Private Const RAS_RASCONNSIZE As Integer = 412

Private Type RasEntryName
    dwSize As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
End Type

Private Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

Private Declare Function RasEnumConnections Lib _
"rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As _
Any, lpcb As Long, lpcConnections As Long) As Long

Private Declare Function RasHangUp Lib "rasapi32.dll" Alias _
"RasHangUpA" (ByVal hRasConn As Long) As Long

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" _
(ByRef lpdwFlags As Long, _
ByVal lpszConnectionName As String, _
ByVal dwNameLen As Long, _
ByVal dwReserved As Long _
) As Long


Private Function InternetConnected( _
    Optional ByRef eConnectionInfo As NetConnTypeConstants, _
    Optional ByRef sConnectionName As String _
    ) As Boolean
   
    Dim dwFlags As Long
    Dim sNameBuf As String
    Dim lR As Long
    Dim iPos As Long
    
    sNameBuf = String$(513, 0)
    lR = InternetGetConnectedStateEx(dwFlags, sNameBuf, 512, 0&)
    eConnectionInfo = dwFlags
    iPos = InStr(sNameBuf, vbNullChar)
    
    If iPos > 0 Then
        sConnectionName = Left$(sNameBuf, iPos - 1)
    ElseIf Not sNameBuf = String$(513, 0) Then
        sConnectionName = sNameBuf
    End If
    
    InternetConnected = (lR = 1)
End Function


Public Property Get IsConnected() As Boolean
    IsConnected = InternetConnected()
End Property

Public Property Get ConnType() As Long
    Dim connInfo As NetConnTypeConstants
    InternetConnected connInfo
    ConnType = connInfo
End Property

Public Function ConnTypeDevice(nType As Long) As String
    Dim strReturn As String
    
    If nType And INTERNET_CONNECTION_LAN Then
        strReturn = "LAN"
    ElseIf nType And INTERNET_CONNECTION_MODEM Then
        strReturn = "Modem"
    ElseIf nType And INTERNET_CONNECTION_PROXY Then
        strReturn = "Proxy"
    ElseIf nType And INTERNET_CONNECTION_OFFLINE Then
        strReturn = "Offline"
    End If
    
    ConnTypeDevice = strReturn
End Function

Public Property Get ConnName() As String
    Dim strName As String
    InternetConnected , strName
    ConnName = strName
End Property

Public Sub HangUp()
    Dim i As Long
    Dim lpRasConn(255) As RasConn
    Dim lpcb As Long
    Dim lpcConnections As Long
    Dim hRasConn As Long
    Dim ReturnCode As Long
    Dim gstrISPName As String
    
    lpRasConn(0).dwSize = RAS_RASCONNSIZE
    lpcb = RAS_MAXENTRYNAME * lpRasConn(0).dwSize
    lpcConnections = 0
    ReturnCode = RasEnumConnections(lpRasConn(0), lpcb, _
    lpcConnections)

    If ReturnCode = 0 Then
        For i = 0 To lpcConnections - 1
            If Trim(ByteToString(lpRasConn(i).szEntryName)) = Trim(gstrISPName) Then
                hRasConn = lpRasConn(i).hRasConn
                ReturnCode = RasHangUp(ByVal hRasConn)
            End If
        Next i
    End If
End Sub

Private Function ByteToString(bytString() As Byte) As String
    Dim i As Integer
    
    i = 0
    While bytString(i) = 0&
        ByteToString = ByteToString & ChrB$(bytString(i))
        i = i + 1
    Wend
End Function


