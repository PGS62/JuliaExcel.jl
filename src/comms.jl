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
    installscript = normpath(joinpath(@__DIR__, "..", "installer", "install.vbs"))
    exefile = "C:/Windows/System32/wscript.exe"
    isfile(exefile) || throw("Cannot find Windows Script Host at '$exefile'")
    isfile(installscript) || throw("Cannot find install script at '$installscript'")
    run(`$exefile $installscript`, wait=false)
    println("Installer script has been launched, please respond to the dialogs there.")
    nothing
end

portfile() = joinpath(getcommsfolder(), "Port_$(getxlpid()).txt")

"""
    _encode_result_for_xl(result)::String
Encode `result` for return to Excel, shared by `srv_xl_inner` and `srv_call_inner`. If `result`
itself can't be encoded (e.g. it's a type with no `encode_for_xl` method), reports that to the
Julia console and returns an encoded error string describing the problem instead.
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
    srv_xl_inner(expression::String)::String
Evaluate a Julia expression and return the encoded result as a string.
Called by the HTTP request handler in `start_server` for requests to `/eval`.
"""
function srv_xl_inner(expression::String)::String
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
        showerror(stdout, e, catch_backtrace())
        println("")
        println("="^100)
        truncate("#($e)!", 10000)
    end
    _encode_result_for_xl(result)
end

"""
    srv_call_inner(payload::String)::String
Decode `payload` (in the JuliaExcel wire format - a 1D array whose first element is a function
name and remaining elements are its arguments), call the named function, and return the encoded
result as a string. Called by the HTTP request handler in `start_server` for requests to `/call`.
Avoids the `Meta.parse` of a literal expression that `srv_xl_inner` requires, which is slow for
large arrays of arguments.
"""
function srv_call_inner(payload::String)::String
    global result = try
        decoded = decode_from_xl(payload)
        fn_name = decoded[1]::String
        fn = Main.eval(Meta.parse(fn_name))             # fast: parses only the short function name
        fn(decoded[2:end]...)
    catch e
        println("="^100)
        println("Something went wrong calling a Julia function from Excel")
        showerror(stdout, e, catch_backtrace())
        println("")
        println("="^100)
        truncate("#($e)!", 10000)
    end
    _encode_result_for_xl(result)
end

"""
    start_server(start::Int=2700)
Start an HTTP server on a free local port that handles evaluation requests from Excel, trying
up to 100 candidate ports starting from `start`. Tries the real bind directly rather than
probing with a separate test-listener first: a probe-then-release check has a gap between
"verified free" and the real bind moments later, during which another process (e.g. an
orphaned Julia session from a previously-closed Excel session) could still be holding the
port - or, since HTTP.jl's server binds with `reuseaddr=true`, could appear to succeed even
though something else is already listening there. Retrying on the real bind failure avoids
relying on that separate check being reliable.
Writes the chosen port to the port file so VBA can discover it during JuliaLaunch.
"""
function start_server(start::Int=2700)
    port = start
    while true
        try
            HTTP.serve!("127.0.0.1", port) do req
                handler = req.target == "/call" ? srv_call_inner : srv_xl_inner
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
