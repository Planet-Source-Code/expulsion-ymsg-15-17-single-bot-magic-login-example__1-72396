VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ymsg 15-17 Login Example"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   2775
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CboPort 
      Height          =   315
      ItemData        =   "Form1.frx":000C
      Left            =   1440
      List            =   "Form1.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox CboYmsg 
      Height          =   315
      ItemData        =   "Form1.frx":003D
      Left            =   120
      List            =   "Form1.frx":004A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox CboServers 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Form1.frx":005A
      Left            =   120
      List            =   "Form1.frx":00BB
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Login server"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Log Out"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Log In"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Pass 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Password"
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox ID 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "ID Here"
      Top             =   120
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   2160
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------
'Project Description : Single Bot Ymsg 15 - 17 Magic Login Example - I'm not gonna say don't rip it...I know you will
'Author : Expulsion
'Website : www.expulsion-creations.com
'Credits : Adam & Dubee (The ones that discovered this method)
'---------------------------------------------------
Option Explicit
Public blnconnected As Boolean
Public BotID As String
Public StrYcook As String
Public StrTcook As String

Private Sub Command1_Click()
On Error Resume Next
    If blnconnected = False Then
        BotID = ID.Text
        Winsock1.Close
        Winsock1.Connect "login.yahoo.com", "80"
    Else:
        Exit Sub
    End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Status.Caption = "Logged Out"
    Winsock2.Close
    blnconnected = False
End Sub

Private Sub Form_Load()
On Error Resume Next
    CboYmsg.Text = "15"
    CboPort.Text = "5050"
    CboServers.Text = "scs.msg.yahoo.com"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Unload Me
End Sub

Private Sub ID_DblClick()
On Error Resume Next
    ID.Text = vbNullString
End Sub

Private Sub Pass_DblClick()
On Error Resume Next
    Pass.Text = vbNullString
End Sub

Private Sub Winsock1_Connect()
On Error Resume Next
    Status.Caption = "Connecting"
    '
    Dim LoginYahoo As String
    '
    LoginYahoo = "GET http://login.yahoo.com/config/login?login=" & ID.Text & "&passwd=" & Pass.Text & " HTTP/1.1" & vbCrLf
    LoginYahoo = LoginYahoo & "Accept-Language: en-us" & vbCrLf
    LoginYahoo = LoginYahoo & "User-Agent: Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 5.1; Expulsion-Creations)" & vbCrLf
    LoginYahoo = LoginYahoo & "Accept: */*" & vbCrLf
    LoginYahoo = LoginYahoo & "Host: login.yahoo.com" & vbCrLf
    LoginYahoo = LoginYahoo & "Connection: Keep-Alive" & vbCrLf & vbCrLf
    '
    Winsock1.SendData LoginYahoo
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
    Dim Data As String
    '
    Winsock1.GetData Data
    '
    If InStr(Data, "Yahoo! - 400 Bad Request") Then
        Status.Caption = "Bad ID/Password"
        Winsock1.Close
    Exit Sub
    Else:
    If InStr(Data, "302 Found") Then
        StrYcook = Split(Data, "Y=")(1)
        StrYcook = Split(StrYcook, "np=1")(0)
        StrYcook = "Y=" & StrYcook & "np=1;"
        StrTcook = Split(Data, "T=")(1)
        StrTcook = Split(StrTcook, ";")(0)
        StrTcook = "T=" & StrTcook
        Winsock1.Close
        Winsock2.Close
        Winsock2.Connect CboServers.Text, CboPort.Text
    Else:
    Status.Caption = "Error"
    Exit Sub
    End If
    End If
End Sub


Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbCritical
End Sub

Private Sub Winsock2_Connect()
On Error Resume Next
    Winsock2.SendData Login(BotID, StrYcook, StrTcook)
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
    Dim Data As String
    '
    Winsock2.GetData Data
    '
    Select Case Asc(Mid(Data, 12, 1))
    '
    Case 85
    Status.Caption = "Logged in"
    blnconnected = True
    '
    Case 2
    If InStr(Data, "ÿÿÿÿ") Then
        Status.Caption = "Logged Out By Server"
        blnconnected = False
        Winsock2.Close
    End If
    '
    End Select
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbCritical
End Sub
