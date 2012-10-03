VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{255553DB-6BA0-11D5-91E1-0050DA23487D}#1.0#0"; "Kcp.ocx"
Begin VB.Form frmUdpPeerA 
   Caption         =   "frmUdpPeerA"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox remotePortTxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Text            =   "8080"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton connectBtn 
      Caption         =   "connect"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox remoteAddressTxt 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "192.168.0.3"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   120
   End
   Begin VB.CommandButton varBtn 
      Caption         =   "set variable"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "data to send"
      Top             =   2760
      Width           =   3135
   End
   Begin VB.CommandButton sendButton 
      Caption         =   "Send"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2760
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
      Left            =   5280
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "192.168.0.3"
      RemotePort      =   8080
      LocalPort       =   8081
   End
   Begin VB.Label Label2 
      Caption         =   "Remote port"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Remote IP address"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin KCPOCXLibCtl.Kcp Kcp1 
      Left            =   4200
      OleObjectBlob   =   "frmUdpPeerA.frx":0000
      Top             =   120
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




Private Sub connectBtn_Click()
    InitSocket
End Sub

Private Sub Form_Load()
    'frmUdpPeerB.Show            ' Show the second form.
    bSemBufferIn = False
    ''bCrossComConnected = InitCrossComm(-1)
    ''bCrossComConnected = CrossCommands.CrossIsConnected
    MsgBox ("Connection=" & bCrossComConnected)
    InitSocket
    
End Sub
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
udpPeerA.Protocol = sckUDPProtocol
udpPeerA.Bind 8081 'local Port
udpPeerA.RemoteHost = remoteAddressTxt.Text '"192.168.1.10"
udpPeerA.RemotePort = remotePortTxt.Text



' The control's name is udpPeerA
    With udpPeerA
        ' IMPORTANT: be sure to change the RemoteHost
        ' value to the name of your computer.
        '.RemoteHost = "192.168.0.3"
        
        '.RemotePort = 1001   ' Port to connect to.
        '.Bind= 1002,"127.0.0.1"           ' Bind to the local port.
    End With
End Sub

Private Sub sendButton_Click()
    ' Send text as soon as it's typed.
    udpPeerA.SendData txtSend.Text
End Sub

Private Sub Timer1_Timer()
If Not bSemBufferIn Then
    ''Call readVar("speed")
End If
End Sub

Private Sub udpPeerA_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    bSemBufferIn = True
    udpPeerA.GetData strData
    txtOutput.Text = strData
    ''Call handleCommand(strData)
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
Dim Ret As Boolean
Dim Etat As String
 
Ret = CrossCommands.ShowVar(ligne, Etat)
udpPeerA.SendData (ligne & " = " & Etat)
End Function
 

Private Function handleCommand(ligne As String)
Dim Longueur, Pos, i As Integer
Dim Variable As String
Dim ValVariable As String
Dim Ret As Boolean

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
    Ret = CrossCommands.SetVar(Variable, ValVariable)
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
    handleCommand (txtSend.Text)
End Sub
