Attribute VB_Name = "PacketsModule"
Option Explicit

Private Function Header(ByVal StrPacketType As String, ByVal StrStat As String, ByVal StrSession As String, ByVal StrComm As Long) As String
    Dim Version As String
    '
    Version = Form1.CboYmsg.Text
    '
    Header = "YMSG" & Chr(Int(Version / 256)) & Chr(Int(Version Mod 256)) & String(2, Chr(0)) & Chr(Int(Len(StrPacketType) / 256)) & Chr(Int(Len(StrPacketType) Mod 256)) & Chr(Int(StrComm / 256)) & Chr(Int(StrComm Mod 256)) & Mid(StrStat, 1, 4) & Mid(StrSession, 1, 4) & StrPacketType
End Function

Public Function Login(YahooID As String, YCookie As String, TCookie As String)
    Login = Header("0��" & YahooID & "��2��" & YahooID & "��1��" & YahooID & "��244��1��6��" & YCookie & " " & TCookie & "��98��us��", String(4, Chr(0)), String(4, Chr(0)), 550)
End Function

