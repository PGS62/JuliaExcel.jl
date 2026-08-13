Attribute VB_Name = "modHttp"
' Copyright (c) 2021-2026 Philip Swannell
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
          Const ErrorString = "There is no connection between Excel and Julia. If you haven't already, call JuliaLaunch. If you have, Julia may still be starting (e.g. precompiling packages) - wait until you see ""JuliaExcel HTTP server listening"" printed in its window, then try again."
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
10        If IsNumeric(PortStr) Then
11            If CLng(PortStr) > 0 Then
12                gJuliaPort = CLng(PortStr)
13                GetJuliaPort = gJuliaPort
14                Exit Function
15            End If
16        End If
17        Throw ErrorString
18        Exit Function
ErrHandler:
19        ReThrow "GetJuliaPort", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaHttpPost
' Purpose   : POST a payload to the JuliaExcel HTTP server and return the serialised result.
'             Uses synchronous XMLHTTP so the call blocks until Julia has finished processing.
' Arguments
' Payload   : For Path = "/eval" (the default), a Julia expression to evaluate. For "/call", a
'             1D array in the JuliaExcel wire format whose first element is a function name and
'             remaining elements are its arguments.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaHttpPost(Payload As String, Optional Path As String = "/eval") As String
          Static xhr As Object
1         On Error GoTo ErrHandler
2         If xhr Is Nothing Then
3             Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
4         End If
5         xhr.Open "POST", "http://127.0.0.1:" & GetJuliaPort() & Path, False
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaIsListening
' Purpose   : Returns TRUE if the JuliaExcel HTTP server on the given port responds successfully within
'             TimeoutMs milliseconds, FALSE otherwise (including on any connection error or timeout). Uses
'             ServerXMLHTTP rather than the XMLHTTP object shared by JuliaHttpPost: XMLHTTP's setTimeouts is
'             not reliably available when late-bound, whereas ServerXMLHTTP supports it directly, and using
'             a separate object means a failed probe here cannot affect later calls that reuse that connection.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaIsListening(ByVal Port As Long, Optional ByVal TimeoutMs As Long = 2000) As Boolean
          Dim xhr As Object
1         On Error GoTo ErrHandler
2         Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
3         xhr.setTimeouts TimeoutMs, TimeoutMs, TimeoutMs, TimeoutMs
4         xhr.Open "POST", "http://127.0.0.1:" & CStr(Port) & "/eval", False
5         xhr.setRequestHeader "Content-Type", "text/plain; charset=utf-8"
6         xhr.Send "1+1"
7         JuliaIsListening = (xhr.Status = 200)
8         Exit Function
ErrHandler:
9         JuliaIsListening = False
End Function
