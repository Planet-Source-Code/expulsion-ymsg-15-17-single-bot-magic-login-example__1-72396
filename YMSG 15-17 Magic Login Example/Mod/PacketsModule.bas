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
    Login = Header("0¿Ä" & YahooID & "¿Ä2¿Ä" & YahooID & "¿Ä1¿Ä" & YahooID & "¿Ä244¿Ä1¿Ä6¿Ä" & YCookie & " " & TCookie & "¿Ä98¿Äus¿Ä", String(4, Chr(0)), String(4, Chr(0)), 550)
End Function

