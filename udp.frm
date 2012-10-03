VERSION 5.00
Object = "{255553DB-6BA0-11D5-91E1-0050DA23487D}#1.0#0"; "Kcp.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmUdpPeerA 
   Caption         =   "frmUdpPeerA"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer broadcastTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4200
      Top             =   120
   End
   Begin VB.CommandButton kukaBtn 
      Caption         =   "initCrossCom"
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox valTxt 
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Text            =   "01"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox addrTxt 
      Height          =   495
      Left            =   3360
      TabIndex        =   14
      Text            =   "40"
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton setVarBtn 
      Caption         =   "set var"
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox ipCombo 
      Height          =   315
      Left            =   5640
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox localPortTxt 
      Height          =   285
      Left            =   5760
      TabIndex        =   11
      Text            =   "5001"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton getVarBtn 
      Caption         =   "get var"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox remotePortTxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Text            =   "5000"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton connectBtn 
      Caption         =   "connect"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox remoteIpTxt 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "192.168.1.16"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   3240
      Top             =   120
   End
   Begin VB.TextBox txtSend 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "data to send"
      Top             =   2760
      Width           =   6615
   End
   Begin VB.CommandButton sendButton 
      Caption         =   "Send"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtOutput 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "output"
      Top             =   3600
      Width           =   3135
   End
   Begin MSWinsockLib.Winsock udpPeerA 
      Left            =   6600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "192.168.0.3"
      RemotePort      =   8080
      LocalPort       =   8081
   End
   Begin MSWinsockLib.Winsock udpPeerSend 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "192.168.0.3"
      RemotePort      =   5001
      LocalPort       =   5000
   End
   Begin KCPOCXLibCtl.Kcp Kcp1 
      Left            =   8160
      OleObjectBlob   =   "udp.frx":0000
      Top             =   480
   End
   Begin VB.Label Label4 
      Caption         =   "port local"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "adresse locale"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "port distant"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "adresse distante"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmUdpPeerA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CrossCommands As cCrossComm
Dim bCrossComConnected As Boolean
Dim bSemBufferIn As Boolean
Dim bConnected As Boolean
Dim bSendVar As Boolean
Dim counter As Integer
Dim iArg As Integer
Dim bCrossCom As Boolean
Dim sendBuff As String

Const START_BYTE = 23 ''
Const END_BYTE = 32 ''
Const GET_SPEED = 0 ''
Const GET_MAG1 = 1 ''
Const GET_FAST_MODE = 2 ''
Const GET_DECHETS = 3 ''
Const GET_VIDE = 4 ''
Const GET_POUBELLE = 5 ''
Const STOP_SUBMIT_CMD = 1 ''
Const START_SUBMIT_CMD = 2 ''
Const STOP_PROG_CMD = 3 ''
Const START_PROG_CMD = 4 ''
Const UPLOAD_PROG_FROM_PC_CMD = 5 ''
Const DOWNLOAD_PROG_TO_PC_CMD = 6 ''
Const SET_DECHETS_CMD = 10 ''
Const SET_POUBELLE_CMD = 11 ''
Const SET_MAGP1_CMD = 12 ''
Const SET_SPEED_CMD = 13 ''
Const SET_FAST_MODE_CMD = 14 ''
Const SET_MODE_AUTO_CMD = 15 ''
Const SET_FIN_BATON_CMD = 16 ''
Const SET_POS_REPOS_CMD = 17 ''
Const SET_SEND_VAR_CMD = 18 ''
Const SET_REG_CMD = 19 ''
Const SET_VAR_CMD = 20 ''
Const EXEC_SHELL = 21 ''
Const PING_CMD = 22 ''
Const SET_CLIENT_CONNECTION_CMD = 23 ''

Const GET_ENV_CMD = 24 ''
Const GET_SPEED_CMD = 25 ''
Const GET_DECHETS_CMD = 26 ''
Const GET_BONNES_CMD = 27 ''
Const GET_TOTAL_A_PRODUIRE_CMD = 28 ''
Const GET_POUBELLE_CMD = 29 ''
Const GET_VIDE_CMD = 30 ''
Const GET_MAG_P1_CMD = 31 ''
Const GET_ROBOT_SPEED_CMD = 32 ''
Const GET_REG_CMD = 33 ''
Const GET_VAR_CMD = 34 ''

Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub broadcastTimer_Timer()
Dim l As Integer
Dim strData As String
    l = 0
    strData = "23" & PING_CMD & Format(l, "0000") & "32"
    udpPeerSend.SendData strData
End Sub

Private Sub connectBtn_Click()
    If connectBtn.Caption = "connect" Then
    counter = 0
    InitSocket
    bConnected = True
    connectBtn.Caption = "disconnect"
    Else
    bConnected = False
    connectBtn.Caption = "connect"
    closeSocket
    End If
    If bCrossCom Then
        initRobotCom
    End If
    
End Sub
Public Sub Test()
   Dim IpAddrs
   IpAddrs = GetIpAddrTable
   Debug.Print "Nr of IP addresses: " & UBound(IpAddrs) - LBound(IpAddrs) + 1
   Dim i As Integer
   For i = LBound(IpAddrs) To UBound(IpAddrs)
      Debug.Print IpAddrs(i)
      Next
   End Sub
   
Private Sub Form_Load()
   Dim IpAddrs
   Dim NrOfEntries As Integer
   
    bSemBufferIn = False
    
    'To get local IP address
     IpAddrs = GetIpAddrTable
     NrOfEntries = UBound(IpAddrs)
     ipCombo.Clear
     
      Dim i As Integer
   For i = 0 To NrOfEntries - 1
    ipCombo.AddItem (IpAddrs(i))
   Next
   ipCombo.ListIndex = i - 2
    
    'DEBUG MODE
    bCrossCom = True
    InitSocket
    bConnected = True
    connectBtn.Caption = "disconnect"
    If bCrossCom Then
        initRobotCom
    End If
    'start broadcast connection signal
    udpPeerSend.RemoteHost = "192.168.1.255" 'broadcast ip
    udpPeerSend.RemotePort = remotePortTxt.Text '5000
    broadcastTimer.Enabled = True
    
End Sub
' Returns an array with the local IP addresses (as strings).
' Author: Christian d'Heureuse, www.source-code.biz
Public Function GetIpAddrTable()
   Dim Buf(0 To 511) As Byte
   Dim BufSize As Long: BufSize = UBound(Buf) + 1
   Dim rc As Long
   rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
   If rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & rc
   Dim NrOfEntries As Integer: NrOfEntries = Buf(1) * 256 + Buf(0)
   If NrOfEntries = 0 Then GetIpAddrTable = Array(): Exit Function
   ReDim IpAddrs(0 To NrOfEntries - 1) As String
   Dim i As Integer
   Dim k As Integer
   k = 0
   For i = 0 To NrOfEntries - 1
      Dim j As Integer, s As String: s = ""
      For j = 0 To 3: s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j): Next
      Dim contain As String
      contain = Mid$(s, 1, 7)
      If contain = "192.168" Then
        IpAddrs(k) = s
        k = k + 1
      End If
      Next
   GetIpAddrTable = IpAddrs
   End Function
   
Public Sub Main()
    Dim bRobotDataExecuted As Boolean
    bRobotDataExecuted = False

    'While (1)
        'wait (5000)
        'If bRobotDataExecuted Then
        'udpPeerA.SendData "commande recue"
        'bRobotDataExecuted = False
        'End If
    'Wend

End Sub
Public Sub InitSocket()
    If udpPeerA.Protocol <> sckUDPProtocol Then
        udpPeerA.Protocol = sckUDPProtocol
    End If
    
    If udpPeerSend.Protocol <> sckUDPProtocol Then
        udpPeerSend.Protocol = sckUDPProtocol
    End If
    
    initTransferEnv
End Sub
Public Sub reconnect()

End Sub
Public Sub initTransferEnv()
    ''udpPeerA.Bind 5000  'local Port
    
    udpPeerA.RemoteHost = remoteIpTxt.Text
    udpPeerA.RemotePort = remotePortTxt.Text
    udpPeerA.Bind localPortTxt.Text, ipCombo.Text
End Sub
Public Sub initTransferSend()
    ''udpPeerA.Bind 5000  'local Port
    udpPeerSend.Close
    udpPeerSend.RemoteHost = remoteIpTxt.Text
    udpPeerSend.RemotePort = remotePortTxt.Text
    'udpPeerSend.Bind localPortTxt.Text, ipCombo.Text

End Sub
Public Sub closeSocket()
    udpPeerA.Close
    udpPeerSend.Close
End Sub
Private Sub getVarBtn_Click()
Dim result  As String
Dim ret As Boolean

ret = CrossCommands.ShowVar(txtSend.Text, result)
txtOutput.Text = result
If udpPeerA.State <> 1 Then
udpPeerA.Close
initTransferEnv
End If
udpPeerA.SendData "data is" & result
End Sub

Private Sub initRobotCom()
    bCrossComConnected = InitCrossComm(-1)
    bCrossComConnected = CrossCommands.CrossIsConnected
    kukaBtn.Caption = bCrossComConnected
End Sub
Private Sub kukaBtn_Click()
    initRobotCom
End Sub

Private Sub sendButton_Click()
    Dim splitter() As String
    Dim ret As Boolean
    udpPeerA.SendData txtSend.Text
    
    If bCrossCom Then
    ' Send text as soon as it's typed.
    splitter = Split(txtSend.Text, "=")
    ret = CrossCommands.SetVar(splitter(0), splitter(1))
    End If
    
End Sub

Private Sub setVarBtn_Click()
    Dim ret As Boolean
    ret = CrossCommands.SetVar(addrTxt.Text, valTxt.Text)
    MsgBox ("Connection=" & ret & addrTxt.Text & " = " & valTxt.Text)
End Sub
Private Sub Timer1_Timer()
    Dim strData As String
    If bSendVar Then
        strData = sendBuff
        udpPeerSend.SendData strData
        txtSend.Text = strData
        bSendVar = False
    End If
End Sub
Private Sub udpPeerA_Close()
    udpPeerA.Close ' has to be called
    'udpPeerA.Listen ' listen again
End Sub
Private Sub udpPeerA_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  ''MsgBox "Socket Error " & Number & ": " & Description
                                    ' show some "debug" info
  udpPeerA.Close ' close the erraneous connection
  'udpPeerA.Listen ' listen again
End Sub
Private Sub udpPeerA_ConnectionRequest(ByVal requestID As Long)
  If udpPeerA.State = sckListening Then ' if the socket is listening
    udpPeerA.Close ' reset its state to sckClosed
    udpPeerA.Accept requestID ' accept the client
  End If
End Sub

Private Sub udpPeerA_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim iData As Integer
    Dim i As Integer
    Dim k As Integer
    Dim iVal As Integer
    Dim res As String
    Dim result As String
    Dim bStart As Integer
    Dim bFunc As Integer
    Dim bLength As Integer
    Dim bEnd As Integer
    
    Dim szFunc As String
    Dim szLength As String
    Dim szArg As String
    
    
    bSemBufferIn = True
    udpPeerA.GetData strData
    'remoteIpTxt.Text = udpPeerA.RemoteHostIP
    
    i = 1
    result = ""
    strData = Replace(strData, " ", "0")
    While i < Len(strData)
        res = Mid(strData, i, 2)
        result = result & res & " "
        i = i + 2
    Wend

    txtOutput.Text = result
    
    'prepare ack
    initTransferSend
    
    If Len(strData) = 0 Then
        'bad request, send a nack
        sendBuff = "2300000032"
        bSendVar = True
    End If
    
    bStart = Left(strData, 2)
    bEnd = Right(strData, 2)
    
    If (bStart = 23) And (bEnd = 32) Then
        'extract data
        bFunc = Mid(strData, 3, 2)
        bLength = Mid(strData, 5, 4)

        szFunc = bFunc
        szLength = bLength
        If bLength > 0 Then
            szArg = Mid(strData, 9, bLength)
        Else
            szArg = ""
        End If
        ' send a ack
        sendBuff = "2301000032"
        bSendVar = True
        'If bCrossCom Then
        Call handleCommand(szFunc, szLength, szArg)
        'End If
    Else
        'bad request, send a nack
        sendBuff = "2300000032"
        bSendVar = True
    End If

    bSemBufferIn = False
    
End Sub
Private Function InitCrossComm(ByVal Mode As Integer) As Boolean
Set CrossCommands = New cCrossComm
On Error GoTo Fehler
   'Connect to KRC
   Dim RetVal As Boolean
   RetVal = False

   If Not CrossCommands.CrossIsConnected Then
      If CrossCommands.Init(Kcp1) Then
         RetVal = CrossCommands.ConnectToCross(Kcp1.name, Mode)
         
      Else
         RetVal = False
      End If
   Else
      RetVal = True
   End If
Ende:
    InitCrossComm = RetVal
    Exit Function
Fehler:
    MsgBox Err.Description + " coucou", vbCritical, Err.Number
    Resume Ende
End Function
Private Sub ExitCrossComm()
    If Not (CrossCommands Is Nothing) Then
        On Error Resume Next
        CrossCommands.ServerOff
        Set CrossCommands = Nothing
    End If
End Sub
Private Function readVar(ligne As String)
Dim ret As Boolean
Dim Etat As String
 
ret = CrossCommands.ShowVar(ligne, Etat)
udpPeerA.SendData (ligne & " = " & Etat)
End Function
 
Public Function handleCommand(cmd As String, length As String, arg As String)

Dim bSend As Boolean
Dim ret As Boolean
Dim szArg() As String
Dim arg0 As String
Dim result As String
Dim address As Integer
Dim value As Integer
Dim k As Integer
Dim szAddr As String
Dim szValue As String
Dim splitter() As String
Dim res As String
Dim l As Integer

k = 1
While k < (length)
    res = Mid(arg, k, 2)
    result = result & res & " "
    k = k + 2
Wend
szArg = Split(result, " ")
            
Select Case (cmd)
    Case SET_VAR_CMD:
        If length > 0 Then
            splitter = Split(arg, "=")
            szAddr = splitter(0)
            szValue = splitter(1)
            If bCrossCom Then
            ret = CrossCommands.SetVar(szAddr, szValue)
            End If
        End If
    Case GET_VAR_CMD:
        If length > 0 Then
            szAddr = arg
            result = szAddr & "=23"
            If bCrossCom Then
                ret = CrossCommands.ShowVar(szAddr, result)
            End If
            l = Len(result)
            sendBuff = "23" & GET_VAR_CMD & Format(l, "0000") & result & "32"
            bSendVar = True
        End If
    Case SET_CLIENT_CONNECTION_CMD:
        Dim szPort As String
        
        If length > 0 Then
            broadcastTimer.Enabled = False
            szPort = szArg(0) & szArg(1)
            remotePortTxt.Text = szPort
            remoteIpTxt.Text = udpPeerA.RemoteHostIP
            initTransferSend
            result = 1
            l = Len(Format(result, "00"))
            sendBuff = "23" & SET_CLIENT_CONNECTION_CMD & Format(l, "0000") & Format(result, "00") & "32"
            bSendVar = True
        End If
    Case SET_REG_CMD:
        If length = 8 Then
            'k = 1
            'While k < (length)
            '    res = Mid(arg, k, 2)
            '    result = result & res & " "
            '    k = k + 2
            'Wend
            'szArg = Split(result, " ")

            address = szArg(0) & szArg(1)
            szAddr = address
            value = szArg(2) & szArg(3)
           If value = 1 Then
            szValue = "TRUE"
           Else
            szValue = "FALSE"
           End If
           If bCrossCom Then
            ret = CrossCommands.SetVar("$OUT[" & szAddr & "]", szValue)
            End If
        End If
    Case GET_REG_CMD:
        If length = 4 Then
            address = szArg(0) & szArg(1)
            szAddr = address
            result = "$OUT[" & szAddr & "]=TRUE"
            If bCrossCom Then
            ret = CrossCommands.ShowVar("$OUT[" & szAddr & "]", result)
            End If
            splitter = Split(result, "=")
            If (splitter(1) = True) Then
            result = 1
            Else
            result = 0
            End If
            l = 8
            sendBuff = "23" & GET_REG_CMD & Format(l, "0000") & Format(address, "0000") & Format(result, "0000") & "32"
            bSendVar = True
        End If
    Case EXEC_SHELL:
        Dim RetVal
        If length > 0 Then
            szAddr = arg
            RetVal = Shell(szAddr, 1)
        End If
    Case SET_POS_REPOS_CMD:
        
        iArg = szArg(0)
        If iArg = 0 Then
            szArg(0) = "FALSE"
        Else
            szArg(0) = "TRUE"
        End If
        ret = CrossCommands.SetVar("POS_REPOS", szArg(0))
        
    Case SET_DECHETS_CMD:
        ret = CrossCommands.SetVar("DECHETS", szArg(0))
        
    Case SET_MAGP1_CMD:
        ret = CrossCommands.SetVar("MAG_P1", szArg(0))
        
    Case SET_FAST_MODE_CMD:
        iArg = szArg(0)
        If iArg = 0 Then
            szArg(0) = "FALSE"
        Else
            szArg(0) = "TRUE"
        End If
        ret = CrossCommands.SetVar("FAST_MODE", szArg(0))
        
    Case SET_SPEED_CMD:
        ret = CrossCommands.SetVar("$OV_PRO", szArg(0) & szArg(1))
        
    Case STOP_SUBMIT_CMD:
        ret = CrossCommands.ControlLevelStop()
    Case START_SUBMIT_CMD:
        ret = CrossCommands.RunControlLevel()
    Case STOP_PROG_CMD:
    ret = CrossCommands.RobotLevelStop()
    
    Case START_PROG_CMD:
    ret = CrossCommands.RunControlLevel()
    Case UPLOAD_PROG_FROM_PC_CMD:
    ret = CrossCommands.DownLoadDiskToRobot("\Users\lorenzzaccio\modulesKuka\moduleTest.ext")
    Case DOWNLOAD_PROG_TO_PC_CMD:
    ret = CrossCommands.UpLoadFromRobotToDisk("moduleTest", "c:\Users\lorenzzaccio\modulesKuka")
    
        
    Case GET_ENV_CMD:
        result = getVars()
        txtSend.Text = result
        l = Len(result)
        sendBuff = "23" & GET_ENV_CMD & Format(l, "0000") & result & "32" '"23" & GET_REG_CMD & "08" & Format(address, "0000") & Format(result, "0000") & "32"
        bSendVar = True
        
        Case GET_ROBOT_SPEED_CMD:
        ret = CrossCommands.ShowVar("$OV_PRO", result)
        sendBuff = result '"23" & GET_REG_CMD & "08" & Format(address, "0000") & Format(result, "0000") & "32"
        bSendVar = True
        
        Case GET_DECHETS_CMD:
        ret = CrossCommands.ShowVar("DECHETS", result)
        sendBuff = result '"23" & GET_REG_CMD & "08" & Format(address, "0000") & Format(result, "0000") & "32"
        bSendVar = True
        
        Case GET_BONNES_CMD:
        ret = CrossCommands.ShowVar("CAPS_BONNES", result)
        sendBuff = result '"23" & GET_REG_CMD & "08" & Format(address, "0000") & Format(result, "0000") & "32"
        bSendVar = True
        
        Case GET_TOTAL_A_PRODUIRE_CMD:
        ret = CrossCommands.ShowVar("TOTAL_A_PRODUIRE", result)
        sendBuff = result '"23" & GET_REG_CMD & "08" & Format(address, "0000") & Format(result, "0000") & "32"
        bSendVar = True
        
        Case GET_POUBELLE_CMD:
        ret = CrossCommands.ShowVar("POUBELLE", result)
        'udpPeerA.RemoteHost = remoteAddressTxt.Text
        udpPeerA.RemotePort = 5000
        udpPeerA.SendData (result)

        Case GET_VIDE_CMD:
        ret = CrossCommands.ShowVar("COUPS_VIDE", result)
        sendBuff = result '"23" & GET_REG_CMD & "08" & Format(address, "0000") & Format(result, "0000") & "32"
        bSendVar = True
        
        Case GET_MAG_P1_CMD:
        ret = CrossCommands.ShowVar("MAG_P1", result)
        sendBuff = result '"23" & GET_REG_CMD & "08" & Format(address, "0000") & Format(result, "0000") & "32"
        bSendVar = True
        
    Case Else
    
    End Select



End Function

Private Function getVars() As String

Dim ret As Boolean
Dim magP1 As String
Dim totalAProduire As String
Dim bonnes As String
Dim vides As String
Dim dechets As String
Dim coup_vide As String
Dim poubelle As String
Dim cadence As String
Dim robotSpeed As String
Dim result As String

magP1 = "MAG_P1=" & 0
totalAProduire = "TOTAL_A_PRODUIRE=" & 0
bonnes = "CAPS_BONNES=" & 0
dechets = "DECHETS=" & 0
vides = "COUPS_VIDE=" & 0
poubelle = "POUBELLE=" & 0
cadence = "SPEED=" & 0
robotSpeed = "$OV_PRO=" & 0
If bCrossCom Then
    ret = CrossCommands.ShowVar("MAG_P1", magP1)
    ret = CrossCommands.ShowVar("TOTAL_A_PRODUIRE", totalAProduire)
    ret = CrossCommands.ShowVar("CAPS_BONNES", bonnes)
    ret = CrossCommands.ShowVar("DECHETS", dechets)
    ret = CrossCommands.ShowVar("COUPS_VIDE", vides)
    ret = CrossCommands.ShowVar("POUBELLE", poubelle)
    ret = CrossCommands.ShowVar("SPEED", cadence)
    ret = CrossCommands.ShowVar("$OV_PRO", robotSpeed)
End If

result = cadence + "-" + totalAProduire + "-" + magP1 + "-" + bonnes + "-" + dechets + "-" + vides + "-" + poubelle + "-" + robotSpeed
getVars = result
End Function
Private Function handleCommand2(ligne As String)
Dim Longueur, Pos, i As Integer
Dim Variable As String
Dim ValVariable As String
Dim ret As Boolean

Dim ret1 As Boolean
Dim Etat As String

'bCrossComConnected = InitCrossComm(-1)
'MsgBox ("Connection=" & bCrossComConnected)
'If bCrossComConnected Then 'Connexion etablie
Longueur = Len(ligne)
Variable = ""
ValVariable = ""
  
  If Longueur > 0 Then
    Pos = InStr(ligne, "=")
    For i = 1 To Pos - 1
      Variable = Variable & Mid(ligne, i, 1)
    Next
    For i = Pos + 1 To Longueur
      ValVariable = ValVariable & Mid(ligne, i, 1)
    Next
    MsgBox ("Variable=" & Variable)
    MsgBox ("ValVariable=" & ValVariable)
    ret = CrossCommands.SetVar(Variable, ValVariable)
  End If

  'Ret = CrossCommands.ShowVar("ROBOT_RESET", Etat)
  
  'If Etat = "ROBOT_RESET=TRUE" Then
  '  ret1 = CrossCommands.SetVar("ROBOT_STOP", "TRUE")
  'Else
  '  ret1 = CrossCommands.SetVar("ROBOT_START", "TRUE")
  'End If
  
'End If

'ExitCrossComm


End Function

Private Sub varBtn_Click()
    ''handleCommand (txtSend.Text)
End Sub


   
