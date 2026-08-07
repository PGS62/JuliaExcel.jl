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
    read_utf16(filename::String)
Returns the contents of a UTF-16 LE encoded text file, stripping the leading BOM.
The args file is written by VBA's FileSystemObject as UTF-16 LE with BOM.
See https://discourse.julialang.org/t/reading-a-utf-16-le-file/11687
"""
read_utf16(filename::String) = transcode(String, reinterpret(UInt16, read(filename)))[4:end]

"""
    srv_xl_inner(expression::String)::String
Evaluate a Julia expression and return the encoded result as a string.
Called by the HTTP request handler in `start_server`.
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

    canencode = true
    encodedresult = try
        encode_for_xl(result)
    catch e
        canencode = false
        encode_for_xl("#Expression evaluated to a variable of type $(typeof(result))," *
                      " which cannot be returned to Excel because: $(e)!")
    end

    canencode || (println("");
    @error "Result of type $(typeof(result)) could not be encoded for return to Excel.")

    return encodedresult
end

"""
    _find_free_port(start::Int=2700)
Scan for an available TCP port starting from `start`, trying up to 100 candidates.
"""
function _find_free_port(start::Int=2700)::Int
    for p in start:(start + 100)
        try
            srv = Sockets.listen(Sockets.IPv4(0), p)
            close(srv)
            return p
        catch
        end
    end
    error("no free port found in range $start to $(start + 100)")
end

"""
    start_server()
Start an HTTP server on a free local port that handles evaluation requests from Excel.
Writes the chosen port to the port file so VBA can discover it during JuliaLaunch.
"""
function start_server()
    port = _find_free_port()
    HTTP.serve!("127.0.0.1", port) do req
        HTTP.Response(200, ["Content-Type" => "text/plain; charset=utf-8"],
                      srv_xl_inner(String(req.body)))
    end
    open(portfile(), "w") do f
        write(f, string(port))
    end
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

    print("\033]0;Julia $VERSION on $os serving Excel PID $(getxlpid())\a")
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
