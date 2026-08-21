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
    getcommsfolder()
Returns the name of the comms folder used by JuliaExcel. See also `setcommsfolder`.
"""
function getcommsfolder()
    if commsfolder[] == ""
        throw("commsfolder has not been set")
    else
        commsfolder[]
    end
end

"""
    setcommsfolder(folder::String="")
Sets the name of the comms folder used by JuliaExcel. See also `getcommsfolder`.
Argument folder can be omitted as a convenience when developing this package.
"""
function setcommsfolder(folder::String="")
    if folder == ""
        if Sys.iswindows()
            folder = joinpath(ENV["TEMP"], "@JuliaExcel")
        elseif Sys.islinux()
            trythese = ["phili", "philip", "PhilipSwannell"]
            for trythis = trythese
                f = joinpath("/mnt/c/Users", trythis, "AppData/Local/Temp/@JuliaExcel")
                if isdir(f)
                    return (commsfolder[] = f)
                end
            end
            throw("Cannot find commsfolder")
        else
            throw("operating system not supported")
        end
    end
    commsfolder[] = folder
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

portfile() = joinpath(getcommsfolder(), "Port_$(getxlpid()).txt")

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
    global result = try
        Main.eval(Meta.parse(expression))
    catch e
        println("="^100)
        if length(expression) > 500
            println("Something went wrong evaluating the contents of an expression")
        else
            println("Something went wrong evaluating the expression:")
            println(expression)
        end
        friendly_error(e)
    end
    Base.invokelatest(_encode_result_for_xl, result)
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
    showerror(stdout, e, catch_backtrace())
    println("")
    println("="^100)
    io = IOBuffer()
    showerror(io, e)
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
    global result = try
        decoded = decode_from_xl(payload)
        fn_name = decoded[1]::String
        broadcasting = endswith(fn_name, ".")
        broadcasting && (fn_name = chop(fn_name))
        fn_to_call = Main.eval(Meta.parse(fn_name))             # fast: parses only the short function name
        args_from_xl = decoded[2:end]
        broadcasting ? broadcast(fn_to_call, args_from_xl...) : fn_to_call(args_from_xl...)
    catch e
        println("="^100)
        call_desc = broadcasting ? "$fn_name.(JuliaExcel.args_from_xl...)" : "$fn_name(JuliaExcel.args_from_xl...)"
        println("Something went wrong calling the Julia function $fn_name from Excel, against arguments saved in JuliaExcel.args_from_xl (until overwritten by the next call), so the error should be reproducible from here with '$call_desc'.")
        friendly_error(e)
    end
    Base.invokelatest(_encode_result_for_xl, result)
end

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
"""
function start_server(start::Int=2700)
    port = start
    while true
        try
            HTTP.serve!("127.0.0.1", port) do req
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
