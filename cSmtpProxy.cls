VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSmtpProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z
'Private Const MODULE_NAME As String = "cSmtpProxy"

'=========================================================================
' Public events
'=========================================================================

Event RecvFromClient(Data() As Byte)
Event RecvFromServer(Data() As Byte)

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_sServerAddress        As String
Private m_lServerPort           As Long
Private WithEvents m_oListen    As cTlsSocket
Attribute m_oListen.VB_VarHelpID = -1
Private WithEvents m_oClient    As cTlsSocket
Attribute m_oClient.VB_VarHelpID = -1
Private WithEvents m_oServer    As cTlsSocket
Attribute m_oServer.VB_VarHelpID = -1

'=========================================================================
' Properties
'=========================================================================

Public Property Get ServerAddress() As String
    ServerAddress = m_sServerAddress
End Property

Public Property Get ServerSocket() As cTlsSocket
    Set ServerSocket = m_oServer
End Property

Public Property Get ClientSocket() As cTlsSocket
    Set ClientSocket = m_oClient
End Property

Public Property Get ListenSocket() As cTlsSocket
    Set ListenSocket = m_oListen
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sServerAddress As String, ByVal lServerPort As Long, ByVal lLocalPort As Long, Optional sLocalAddress As String) As Boolean
    m_sServerAddress = sServerAddress
    m_lServerPort = lServerPort
    Set m_oListen = New cTlsSocket
    If Not m_oListen.Create(lLocalPort, SocketAddress:=sLocalAddress) Then
        GoTo QH
    End If
    If Not m_oListen.Listen() Then
        GoTo QH
    End If
    '--- success
    Init = True
QH:
End Function

Private Sub pvInjectStartTls(sText As String)
    If Left$(sText, 5) <> "EHLO " Then
        GoTo QH
    End If
    If Not m_oServer.SyncSendText(sText) Then
        GoTo QH
    End If
    sText = m_oServer.SyncReceiveText()
    If LenB(sText) = 0 Then
        GoTo QH
    End If
    sText = "STARTTLS" & vbCrLf
    If Not m_oServer.SyncSendText(sText) Then
        GoTo QH
    End If
    sText = m_oServer.SyncReceiveText()
    If LenB(sText) = 0 Then
        GoTo QH
    End If
    If Not m_oServer.SyncStartTls(m_sServerAddress) Then
        GoTo QH
    End If
QH:
End Sub

'=========================================================================
' Socket events
'=========================================================================

Private Sub m_oListen_OnAccept()
    Set m_oServer = New cTlsSocket
    If Not m_oServer.Connect(m_sServerAddress, m_lServerPort, UseTls:=False) Then
        Set m_oServer = Nothing
        GoTo QH
    End If
    Set m_oClient = New cTlsSocket
    m_oListen.Accept m_oClient, UseTls:=False
QH:
End Sub

Private Sub m_oServer_OnReceive()
    Dim baBuffer()      As Byte
    
    If m_oServer.ReceiveArray(baBuffer) Then
        RaiseEvent RecvFromServer(baBuffer)
        m_oClient.SendArray baBuffer
    End If
End Sub
   
Private Sub m_oClient_OnReceive()
    Dim baBuffer()      As Byte
    
    If m_oClient.ReceiveArray(baBuffer) Then
        pvInjectStartTls StrConv(baBuffer, vbUnicode)
        RaiseEvent RecvFromClient(baBuffer)
        m_oServer.SendArray baBuffer
    End If
End Sub

Private Sub m_oClient_OnClose()
    m_oServer.Close_
End Sub
