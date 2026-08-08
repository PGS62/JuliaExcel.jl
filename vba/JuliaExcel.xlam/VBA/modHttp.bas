Attribute VB_Name = "modHttp"
' Copyright (c) 2021-2025 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

#If VBA7 And Win64 Then
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Private gJuliaPort As Long

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetJuliaPort
' Purpose   : Store the HTTP port used by the Julia server in the module-level variable, backed up to file
' -----------------------------------------------------------------------------------------------------------------------
Sub SetJuliaPort(Port As Long)
          Dim PortFile As String
1         On Error GoTo ErrHandler
2         gJuliaPort = Port
3         PortFile = LocalTemp() & "\Port_" & CStr(GetCurrentProcessId()) & ".txt"
4         If Port = 0 Then
5             If FileExists(PortFile) Then
6                 Kill PortFile
7             End If
8         Else
9             SaveTextFile PortFile, CStr(Port), TristateFalse
10        End If
11        Exit Sub
ErrHandler:
12        ReThrow "SetJuliaPort", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetJuliaPort
' Purpose   : Returns the HTTP port used by the Julia server, or 0 if JuliaLaunch has not
'             been called. Recovers from the port file on disk if the module-level variable
'             was reset by an unhandled VBA error. The port file is written by Julia at startup
'             and CleanLocalTemp will not remove it until it is more than 3 days old.
' -----------------------------------------------------------------------------------------------------------------------
Function GetJuliaPort() As Long
          Dim PortFile As String
          Dim PortStr As String
1         On Error GoTo ErrHandler
2         If gJuliaPort <> 0 Then
3             GetJuliaPort = gJuliaPort
4             Exit Function
5         End If
          ' Module variable was cleared - attempt recovery from the port file on disk
6         PortFile = LocalTemp() & "\Port_" & CStr(GetCurrentProcessId()) & ".txt"
7         On Error Resume Next
8         PortStr = ReadTextFile(PortFile, TristateFalse)
9         On Error GoTo ErrHandler
10        If IsNumeric(PortStr) And CLng(PortStr) > 0 Then
11            gJuliaPort = CLng(PortStr)
12            GetJuliaPort = gJuliaPort
13        End If
14        Exit Function
ErrHandler:
15        ReThrow "GetJuliaPort", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaHttpPost
' Purpose   : POST a serialised Julia expression to the JuliaExcel HTTP server and return
'             the serialised result. Uses synchronous XMLHTTP so the call blocks until Julia
'             has finished evaluating.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaHttpPost(Payload As String) As String
          Static xhr As Object
1         On Error GoTo ErrHandler
2         If xhr Is Nothing Then
3             Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
4         End If
5         xhr.Open "POST", "http://127.0.0.1:" & GetJuliaPort() & "/eval", False
6         xhr.setRequestHeader "Content-Type", "text/plain; charset=utf-8"
7         xhr.Send Payload
8         If xhr.Status <> 200 Then
9             Throw "#HTTP " & xhr.Status & ": " & xhr.statusText & "!"
10        End If
11        JuliaHttpPost = xhr.responseText
12        Exit Function
ErrHandler:
13        ReThrow "JuliaHttpPost", Err
End Function
