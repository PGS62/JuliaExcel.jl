"""
    setxlpid(pid::Int64)
Set the process id of the instance of Excel that the current Julia process is serving.
"""
function setxlpid(pid::Int64)
    xlpid[] = pid
    settitle()
    println("xlpid set to $pid")
    nothing
end

"""
    getxlpid()
Returns the process id of the instance of Excel that the current Julia process is serving.
"""
function getxlpid()
    xlpid[] == 0 && throw("setxlpid has not been called in this Julia session, it must be" *
                          " called to set the process id of the active Excel session")
    xlpid[]
end

"""
    serve_xl(pid::Integer; show_results::Bool=true)
    serve_xl(; show_results::Bool=true)

Attaches this Julia session to an Excel process, so it responds to `JuliaCall`/`JuliaEval` requests
from that Excel session exactly as if it had been launched by `JuliaLaunch` - for attaching a Julia
session that's already running (e.g. one open in VS Code) instead of always launching a fresh one.

With no argument, automatically attaches to the single running Excel process (Windows only). If no
Excel process is found, or more than one is, throws an error explaining how to call `serve_xl(pid)`
instead, giving the process id explicitly - get it from Excel's `JuliaExcelPID()` worksheet function.

`show_results` is passed to `display_results` - it defaults to `true` here (unlike calling
`display_results` directly, which defaults to `false`) since attaching an interactive session like
this is exactly when seeing each call and its result echoed to the REPL is most useful. Pass
`show_results=false` to opt out.

`show_results` is a keyword rather than a positional argument to avoid ambiguity with `pid`: in
Julia, `Bool <: Integer`, so `serve_xl(true)` could otherwise be read as either `pid=true` or
`show_results=true`.

Equivalent to calling `setxlpid(pid)`, `comms_folder(...)`, `display_results(show_results)` and
`start_server()` in turn - the same steps `JuliaLaunch`'s generated startup script performs for a
session it launches itself (aside from `display_results`, which it leaves at its own default).
"""
function serve_xl(pid::Integer; show_results::Bool=true)
    setxlpid(Int64(pid))
    comms_folder(_default_commsfolder())
    display_results(show_results)
    start_server()
end

function serve_xl(; show_results::Bool=true)
    pids = _running_excel_pids()
    if isempty(pids)
        throw("No running Excel process was found. Open Excel and try again, or call serve_xl(pid)" *
              " directly, passing the process id from Excel's JuliaExcelPID() worksheet function.")
    elseif length(pids) > 1
        throw("Found $(length(pids)) running Excel processes (process ids: $(join(pids, ", "))) -" *
              " call serve_xl(pid) directly instead, passing the process id of the Excel session to" *
              " attach to, from its JuliaExcelPID() worksheet function.")
    end
    serve_xl(only(pids); show_results=show_results)
end

"""
    _running_excel_pids()::Vector{Int}
Returns the process ids of all running Excel.exe processes, found via the Windows `tasklist`
utility. Used by `serve_xl()` to find the single Excel process to attach to automatically.
"""
function _running_excel_pids()::Vector{Int}
    Sys.iswindows() || throw("Automatic Excel process detection needs Windows - call serve_xl(pid)" *
                             " directly instead, passing the process id from Excel's" *
                             " JuliaExcelPID() worksheet function.")
    pids = Int[]
    for line in readlines(`tasklist /FI "IMAGENAME eq EXCEL.EXE" /FO CSV /NH`)
        fields = split(line, "\",\"")
        length(fields) >= 2 && push!(pids, parse(Int, fields[2]))
    end
    pids
end

"""
    display_results(switch::Bool)
Switch on or off display in the REPL of both the incoming expression/function call from Excel
and the value returned to Excel, for calls via JuliaCall and JuliaEval.
"""
function display_results(switch::Bool)
    _display_results[] = switch
    "Results from JuliaCall/JuliaEval $(switch ? "will" : "will not") display in REPL"
end

"""
    display_results()
Returns whether display in the REPL of results of calls from Excel is currently switched on.
"""
display_results() = _display_results[]

"""
    comms_folder()
Returns the name of the comms folder used by JuliaExcel. See also `comms_folder(folder)`.
"""
function comms_folder()
    commsfolder[] == "" && throw("commsfolder has not been set")
    commsfolder[]
end

"""
    comms_folder(folder::String)
Sets the name of the comms folder used by JuliaExcel to `folder`. See also `comms_folder()`.
"""
comms_folder(folder::String) = (commsfolder[] = folder)

"""
    _default_commsfolder()::String
Guesses the comms folder JuliaExcel should use when none has been set explicitly - on Windows,
matches the folder VBA's LocalTemp() computes for a native (non-WSL) session. Used by `serve_xl`,
as a convenience when developing this package.
"""
function _default_commsfolder()
    if Sys.iswindows()
        joinpath(ENV["TEMP"], "@JuliaExcel")
    elseif Sys.islinux()
        trythese = ["phili", "philip", "PhilipSwannell"]
        for trythis = trythese
            f = joinpath("/mnt/c/Users", trythis, "AppData/Local/Temp/@JuliaExcel")
            isdir(f) && return f
        end
        throw("Cannot find commsfolder")
    else
        throw("operating system not supported")
    end
end

function installme()
    Sys.iswindows() || throw("JuliaExcel.installme (which installs a Microsoft Excel " *
                             "addin) can only be run from Julia on Windows")
    installscript = normpath(joinpath(@__DIR__, "..", "installer", "Install.ps1"))
    exefile = "C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"
    isfile(exefile) || throw("Cannot find PowerShell at '$exefile'")
    isfile(installscript) || throw("Cannot find install script at '$installscript'")
    run(`$exefile -ExecutionPolicy Bypass -NoProfile -File $installscript`, wait=false)
    println("Installer script has been launched, please respond to the dialogs there.")
    nothing
end

portfile() = joinpath(comms_folder(), "Port_$(getxlpid()).txt")

"""
    _encode_result_for_xl(result)::String
Encode `result` for return to Excel, shared by `srv_eval_inner` and `srv_call_inner`. If
`result` itself can't be encoded (e.g. it's a type with no `encode_for_xl` method), reports
that to the Julia console and returns an encoded error string describing the problem
instead.

Callers should invoke this via `Base.invokelatest`: if `result` is a value just defined by
an `eval` earlier in the same request (e.g. `JuliaEval("f(x)=x^2")` returns the function
`f` itself), showing its type here can require looking up a global binding that didn't
exist in the world the caller's own method was compiled in, which Julia 1.12+ reports as a
"world age" warning.
"""
function _encode_result_for_xl(result)::String
    try
        encode_for_xl(result)
    catch e
        println("")
        @error "Result of type $(typeof(result)) could not be encoded for return to Excel."
        encode_for_xl("#Expression evaluated to a variable of type $(typeof(result))," *
                      " which cannot be returned to Excel because: $(e)!")
    end
end

"""
    srv_eval_inner(expression::String)::String
Evaluate a Julia expression and return the encoded result as a string.
Called by the HTTP request handler in `start_server` for requests to `/eval`, originating
from VBA calls to JuliaEval and JuliaEvalVBA.
"""
function srv_eval_inner(expression::String)::String
    global last_question = expression
    if _display_results[]
        printstyled("question> ", color=:green)
        println(expression)
    end
    success = true
    global last_answer = try
        Main.eval(Meta.parse(expression))
    catch e
        success = false
        printstyled("Something went wrong evaluating the expression: ", color=:red)
        println(expression)
        friendly_error(e)
    end
    if _display_results[] && success
        printstyled("answer> ", color=:green)
        try
            display(last_answer)
        catch e
            printstyled("(could not display result of type $(typeof(last_answer)): $e)\n", color=:red)
        end
        println("")
    end
    Base.invokelatest(_encode_result_for_xl, last_answer)
end

"""
    friendly_error(e)::String

Reports an error to both the Julia console and back to Excel, to be called from a `catch e`
block in `srv_eval_inner`/`srv_call_inner`. Prints the full error to `stdout`, including `e`'s
stacktrace via `catch_backtrace()`, for inspection in the Julia REPL - then returns a short summary
string to send back to Excel as the encoded result.

The returned summary is deliberately short: only the first two lines of `showerror(io, e)` (the
exception's own message - `showerror` only appends a stacktrace when explicitly passed one, which
this call doesn't), further capped at 500 characters. Even without a stacktrace, a `MethodError`'s
"Closest candidates are:" section can run long for a heavily overloaded function (e.g. `+`), so
without this cap a single error could still flood the cell. "Julia REPL has more details and
stacktrace!" points the user at where the full detail actually lives.
"""
function friendly_error(e)
    print("\n")
    showerror(stdout, e, catch_backtrace())
    io = IOBuffer()
    showerror(io, e)
    print("\n\n")
    lines = split(String(take!(io)), '\n')
    error_desc = join(first(lines, 2), ' ')
    if length(error_desc) > 500
        error_desc = truncate(error_desc, 500) * "..."
    end
    return "#$error_desc Julia REPL has more details and stacktrace!"
end

"""
    srv_call_inner(payload::String)::String

Decode `payload` (in the JuliaExcel wire format - a 1D array whose first element is a
function name and remaining elements are its arguments), call the named function, and
return the encoded result as a string. Called by the HTTP request handler in `start_server`
for requests to `/call`, originating from VBA calls to functions JuliaCall and JuliaCallVBA.

Avoids the `Meta.parse` of a literal expression that `srv_eval_inner` requires, which is
slow for large arrays of arguments.

A trailing "." on the function name (e.g. "f.") requests broadcasting, matching Julia's
`f.(...)` call syntax. That syntax is a transform on the call site itself (it lowers to
`broadcast(f, ...)`) rather than a property of a standalone function reference, so "f."
can't just be handed to `Meta.parse` as-is - the dot is stripped before resolving the
function, and `broadcast` is called explicitly instead of a plain call.
"""
function srv_call_inner(payload::String)::String
    fn_name = "<unknown>"
    global args_from_xl = ["<unknown>"]
    broadcasting = false
    success = true
    global last_answer = try
        decoded = decode_from_xl(payload)
        fn_name = decoded[1]::String
        broadcasting = endswith(fn_name, ".")
        broadcasting && (fn_name = chop(fn_name))
        global last_question = broadcasting ? "$fn_name.(args_from_xl...)" : "$fn_name(args_from_xl...)"
        if _display_results[]
            printstyled("question> ", color=:green)
            println(last_question)
        end
        fn_to_call = Main.eval(Meta.parse(fn_name))             # fast: parses only the short function name
        args_from_xl = decoded[2:end]
        broadcasting ? broadcast(fn_to_call, args_from_xl...) : fn_to_call(args_from_xl...)
    catch e
        success = false
        global last_question = broadcasting ? "$fn_name.(args_from_xl...)" : "$fn_name(args_from_xl...)"
        printstyled("Something went wrong calling the Julia function $fn_name", color=:red)
        print(" from Excel against\narguments saved in args_from_xl (overwritten by the next call),")
        print(" so\nthe error should be reproducible from here with '$last_question'.\n\n")
        friendly_error(e)
    end
    if _display_results[] && success
        printstyled("answer> ", color=:green)
        try
            display(last_answer)
        catch e
            printstyled("(could not display result of type $(typeof(last_answer)): $e)\n", color=:red)
        end
        println("")
    end
    Base.invokelatest(_encode_result_for_xl, last_answer)
end

"""
    answer_again()

Evaluates `last_question` again, and you can wrap it with `@enter` or your own debugging tools.
Useful for interactively debugging a `JuliaCall`/`JuliaEval` that failed (or just misbehaved) when
called from Excel, without needing to repeat the call from there.

See also `last_question`, `last_answer`, `args_from_xl`.
"""
answer_again() = Main.eval(Meta.parse(last_question))

"""
    start_server(start::Int=2700)
Start an HTTP server on a free local port that handles evaluation requests from Excel,
trying up to 100 candidate ports starting from `start`. Tries the real bind directly rather
than probing with a separate test-listener first: a probe-then-release check has a gap
between "verified free" and the real bind moments later, during which another process (e.g.
an orphaned Julia session from a previously-closed Excel session) could still be holding
the port - or, since HTTP.jl's server binds with `reuseaddr=true`, could appear to succeed
even though something else is already listening there. Retrying on the real bind failure
avoids relying on that separate check being reliable.

Writes the chosen port to the port file so VBA can discover it during JuliaLaunch.

Closes any server already started by a previous call in this session first - otherwise, calling
this (or `serve_xl`) more than once would leave each earlier listener still running, each holding
its own port open indefinitely.
"""
function start_server(start::Int=2700)
    _server[] !== nothing && close(_server[])
    port = start
    while true
        try
            _server[] = HTTP.serve!("127.0.0.1", port) do req
                handler = req.target == "/call" ? srv_call_inner : srv_eval_inner
                HTTP.Response(200, ["Content-Type" => "text/plain; charset=utf-8"],
                    handler(String(req.body)))
            end
            break
        catch
            port += 1
            port > start + 100 && rethrow()
        end
    end
    open(portfile(), "w") do f
        write(f, string(port))
    end
    xlport[] = port
    settitle()
    println("JuliaExcel HTTP server listening on port $port")
    nothing
end

"""
    stop_server()
Stops this Julia session's HTTP server, if one is running, so it no longer responds to
`JuliaCall`/`JuliaEval` requests from Excel. No-op if no server is currently running.

Leaves the port file on disk untouched, so Excel will still try the now-dead port on its next
request - and get a clean "no connection" error rather than reaching this session again.
"""
function stop_server()
    _server[] !== nothing && close(_server[])
    _server[] = nothing
    nothing
end

"""
    server_status()
Returns a `NamedTuple` `(running, pid, port, display_results)` describing this Julia session's HTTP
server: `running` is `true` if a server is currently listening; `pid` is the Excel process id it's
set up to serve (0 if `setxlpid` has never been called), regardless of whether the server is
currently running; `port` is the port it's listening on, or 0 if not running; `display_results` is
the current setting of `display_results()`.
"""
function server_status()
    running = _server[] !== nothing && isopen(_server[])
    (running=running, pid=xlpid[], port=running ? xlport[] : 0, display_results=display_results())
end

"""
    setvar(name::String, arg)
Set a variable in global scope. Called by VBA function JuliaSetVar.
"""
function setvar(name::String, arg)

    if Base.isidentifier(name)
        Main.eval(Main.eval(Meta.parse(":(global $name = $arg)")))

        thesize = ()
        thetype = Nothing
        try
            tmp = Main.eval(Meta.parse(name))
            thesize = size(tmp)
            thetype = typeof(tmp)
        catch
        end

        numdims = length(thesize)
        if numdims == 0
            sizedesc = ""
        elseif numdims == 1
            sizedesc = "$(thesize[1])-element "
        elseif numdims > 1
            sizedesc = join(thesize, "x") * " "
        end
        "Set global variable `$name` to $sizedesc$thetype"

    else
        "#`$name` is not an allowed variable name in Julia!"
    end
end

# https://docs.microsoft.com/en-us/windows/terminal/tutorials/tab-title
function settitle()
    if Sys.islinux()
        os = "Linux"
    elseif Sys.iswindows()
        os = "Windows"
    end

    portpart = xlport[] == 0 ? "" : ", port $(xlport[])"
    print("\033]0;Julia $VERSION on $os serving Excel PID $(getxlpid())$portpart\a")
end

"""
    truncate(x::String, maxlength::Int)
Abbreviate a string to show only `maxlength` characters.
"""
function truncate(x::String, maxlength::Int)
    if (length(x)) > maxlength
        first(x, maxlength ÷ 2) * " … " * last(x, maxlength - (maxlength ÷ 2) - 1)
    else
        x
    end
end
