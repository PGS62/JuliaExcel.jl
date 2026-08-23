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
' Procedure : ReadJuliaPortFromFile
' Purpose   : Reads the HTTP port Julia most recently wrote to the port file for this Excel session,
'             bypassing the gJuliaPort module-level cache entirely. Returns 0 if the file is missing
'             or its content isn't a positive number. The port file is written by Julia at startup
'             (or by JuliaExcel.serve_xl) and CleanLocalTemp will not remove it until it is more than
'             3 days old.
' -----------------------------------------------------------------------------------------------------------------------
Function ReadJuliaPortFromFile() As Long
          Dim PortFile As String
          Dim PortStr As String
1         On Error GoTo ErrHandler
2         PortFile = LocalTemp() & "\Port_" & CStr(GetCurrentProcessId()) & ".txt"
3         On Error Resume Next
4         PortStr = ReadTextFile(PortFile, TristateFalse)
5         On Error GoTo ErrHandler
6         If IsNumeric(PortStr) Then
7             If CLng(PortStr) > 0 Then
8                 ReadJuliaPortFromFile = CLng(PortStr)
9             End If
10        End If
11        Exit Function
ErrHandler:
12        ReThrow "ReadJuliaPortFromFile", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetJuliaPort
' Purpose   : Returns the HTTP port used by the Julia server, or 0 if JuliaLaunch has not
'             been called. Recovers from the port file on disk if the module-level variable
'             was reset by an unhandled VBA error. See also ReadJuliaPortFromFile.
' -----------------------------------------------------------------------------------------------------------------------
Function GetJuliaPort() As Long
          Const ErrorString = "There is no connection between Excel and Julia. If you haven't already, call JuliaLaunch. If you have, Julia may still be starting (e.g. precompiling packages) - wait until you see ""JuliaExcel HTTP server listening"" printed in its window, then try again."
1         On Error GoTo ErrHandler
2         If gJuliaPort <> 0 Then
3             GetJuliaPort = gJuliaPort
4             Exit Function
5         End If
          ' Module variable was cleared - attempt recovery from the port file on disk
6         gJuliaPort = ReadJuliaPortFromFile()
7         If gJuliaPort > 0 Then
8             GetJuliaPort = gJuliaPort
9             Exit Function
10        End If
11        Throw ErrorString
12        Exit Function
ErrHandler:
13        ReThrow "GetJuliaPort", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TryHttpSend
' Purpose   : Attempts a synchronous POST via the given xhr object. MaxRetries is deliberately 1 (no
'             actual retry): measurement on a real machine showed a single failed connection attempt
'             to 127.0.0.1 can carry a fixed ~2 second cost, unaffected by setTimeouts, setProxy, or
'             which HTTP client makes the request (a bare PowerShell Invoke-WebRequest showed the same
'             delay) - almost certainly something outside JuliaExcel's control, e.g. security software
'             inspecting even loopback traffic. Since that cost applies identically whether it's the
'             first or a repeat attempt against the same dead port, retrying here would only multiply
'             an unavoidable delay for no benefit. (This is a different failure domain to the brief,
'             genuinely self-resolving file-locking glitches SaveTextFile in modUtils.bas retries -
'             those cost tens of milliseconds each, not seconds, so retrying them is still worthwhile.)
'             Real recovery from a dead port comes from JuliaHttpPost's separate fallback to a fresh
'             port read from disk, not from retrying here. Returns True if the attempt reaches
'             something (any HTTP status, even non-200 - that's a real error for the caller to report,
'             not a sign of a stale or flaky port), False on a connection-level failure.
' -----------------------------------------------------------------------------------------------------------------------
Private Function TryHttpSend(xhr As Object, Port As Long, Path As String, Payload As String) As Boolean
          Const DelayMs As Long = 25
          Const MaxRetries As Integer = 1
          Dim Attempts As Integer
1         For Attempts = 1 To MaxRetries
2             On Error Resume Next
3             xhr.Open "POST", "http://127.0.0.1:" & Port & Path, False
4             xhr.setRequestHeader "Content-Type", "text/plain; charset=utf-8"
5             xhr.Send Payload
6             If Err.Number = 0 Then
7                 TryHttpSend = True
8                 Exit Function
9             End If
10            Err.Clear
11            On Error GoTo 0
12            DoEvents
13            PreciseSleep DelayMs
14        Next Attempts
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaHttpPost
' Purpose   : POST a payload to the JuliaExcel HTTP server and return the serialised result.
'             Uses synchronous ServerXMLHTTP so the call blocks until Julia has finished processing.
'             ServerXMLHTTP (rather than XMLHTTP) is used specifically so setTimeouts is reliable -
'             see JuliaIsListening's comment - giving a short, known timeout for resolving/connecting
'             instead of whatever long default Windows would otherwise apply. The send/receive
'             timeouts are left unlimited (0): a real JuliaEval/JuliaCall can legitimately take a
'             long time to compute, and must never be aborted just for being slow to respond once
'             connected.
' Arguments
' Payload   : For Path = "/eval" (the default), a Julia expression to evaluate. For "/call", a
'             1D array in the JuliaExcel wire format whose first element is a function name and
'             remaining elements are its arguments.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaHttpPost(Payload As String, Optional Path As String = "/eval") As String
          Const ConnectTimeoutMs As Long = 150
          Const SXH_PROXY_SET_DIRECT As Long = 1   'MSXML2.ServerXMLHTTP's SXH_PROXY_SETTING enum
          Static xhr As Object
          Dim Port As Long
          Dim RefreshedPort As Long
1         On Error GoTo ErrHandler
2         If xhr Is Nothing Then
3             Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
              'Without this, ServerXMLHTTP's default proxy auto-detection (WPAD) can add several
              'seconds to a failed connection attempt - wholly separate from, and unaffected by,
              'setTimeouts below. There is never a legitimate proxy for a 127.0.0.1 loopback call.
4             xhr.setProxy SXH_PROXY_SET_DIRECT
5             xhr.setTimeouts ConnectTimeoutMs, ConnectTimeoutMs, 0, 0
6         End If
7         Port = GetJuliaPort()
8         If Not TryHttpSend(xhr, Port, Path, Payload) Then
              'The cached port didn't respond - it may be stale, e.g. because JuliaExcel.serve_xl was
              'called from Julia on a different port, without Excel's cached port being told. Re-read
              'the port file directly (bypassing the cache) and retry once before giving up.
9             RefreshedPort = ReadJuliaPortFromFile()
10            If RefreshedPort = 0 Or RefreshedPort = Port Then
11                Throw "There is no connection between Excel and Julia. If you haven't already, call JuliaLaunch. If you have, Julia may still be starting (e.g. precompiling packages) - wait until you see ""JuliaExcel HTTP server listening"" printed in its window, then try again."
12            End If
13            SetJuliaPort RefreshedPort
14            Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")   'the previous xhr may be left in a bad state after a failed Send
15            xhr.setProxy SXH_PROXY_SET_DIRECT
16            xhr.setTimeouts ConnectTimeoutMs, ConnectTimeoutMs, 0, 0
17            If Not TryHttpSend(xhr, RefreshedPort, Path, Payload) Then
18                Throw "There is no connection between Excel and Julia. If you haven't already, call JuliaLaunch. If you have, Julia may still be starting (e.g. precompiling packages) - wait until you see ""JuliaExcel HTTP server listening"" printed in its window, then try again."
19            End If
20        End If
21        If xhr.Status <> 200 Then
22            Throw "#HTTP " & xhr.Status & ": " & xhr.statusText & "!"
23        End If
24        JuliaHttpPost = xhr.responseText
25        Exit Function
ErrHandler:
26        ReThrow "JuliaHttpPost", Err
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
