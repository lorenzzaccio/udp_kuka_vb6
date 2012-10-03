VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmUdpPeerB 
   Caption         =   "frmUdpPeerB"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtOutput 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtSend 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   2655
   End
   Begin MSWinsockLib.Winsock udpPeerB 
      Left            =   480
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmUdpPeerB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' Send text as soon as it's typed.
    udpPeerB.SendData txtSend.Text
End Sub

Private Sub Form_Load()
udpPeerB.Protocol = sckUDPProtocol
 ' The control's name is udpPeerB.
    With udpPeerB
        ' IMPORTANT: be sure to change the RemoteHost
        ' value to the name of your computer.
        .RemoteHost = "127.0.0.1"
        .RemotePort = 1002    ' Port to connect to.
        .Bind 1001                ' Bind to the local port.
    End With
End Sub
Private Sub udpPeerB_DataArrival _
(ByVal bytesTotal As Long)
    Dim strData As String
    udpPeerB.GetData strData
    txtOutput.Text = strData
End Sub

