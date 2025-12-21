module JuliaExcel
export srv_xl, setxlpid, killflagfile, getcommsfolder, htf

using Dates, DataFrames
import StringEncodings
const global xlpid = Ref(0)
const global commsfolder = Ref("")

include("encode.jl")

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
Returns the name of the folder to which request files are written by VBA code in 
JuliaExcel.xlam and to which `srv_xl` writes results. See also `setcommsfolder`.
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
Sets the name of the folder to which request files are written by VBA code in 
JuliaExcel.xlam and to which `srv_xl` writes results. See also `getcommsfolder`.
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

flagfile() = joinpath(getcommsfolder(), "Flag_$(getxlpid()).txt")
resultfile() = joinpath(getcommsfolder(), "Result_$(getxlpid()).txt")
expressionfile() = joinpath(getcommsfolder(), "Expression_$(getxlpid()).txt")

"""
    killflagfile()
Deletes the "flag file" whose existence indicates to VBA code in JuliaExcel.xlam that 
`srv_xl()` has not yet completed its evaluation of the contents of the expression to be
evaluated. `killflagfile` can thus be used manually from the REPL if (for example) the
expression to be evaluated includes an infinite loop.
"""
function killflagfile()
    rm_retry(flagfile())
end

function rm_retry(path::AbstractString; retries::Int=10, wait::Real=0.25)
    for attempt in 1:retries
        try
            rm(path)
            attempt == 1 || @info "Successfully deleted $path on attempt $attempt"
            return true  # Success
        catch e
            @warn "Attempt $attempt to delete $path failed: $e, will retry after $wait seconds..."
            if attempt == retries
                @error "All $retries attempts to delete $path failed."
                rethrow(e)  # Final failure
            elseif isa(e, Base.IOError)
                sleep(wait)
            else
                rethrow(e)  # Unexpected error
            end
        end
    end
    return false  # Shouldn't reach here
end


"""
    read_utf16(filename::String)
Returns the contents of a UTF-16 encoded text file that has a byte option mark.
See https://discourse.julialang.org/t/reading-a-utf-16-le-file/11687
"""
read_utf16(filename::String) = transcode(String, reinterpret(UInt16, read(filename)))[4:end]

"""
    srv_xl()
Read the expression file created by JuliaExcel.xlam, evaluate it and write the result to
file, to be unserialised by JuliaExcel.xlam. Files are read from and written to the folder
given by `getcommsfolder`.
"""
function srv_xl()

    expression = read_utf16(expressionfile())
    global result = try
        Main.eval(Meta.parse(expression))
    catch e
        println("="^100)
        if length(expression) > 500
            println("Something went wrong evaluating the contents of $(expressionfile())")
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

    io = open(resultfile(), "w")
    write(io, StringEncodings.encode(encodedresult, "UTF-16"))
    close(io)

    killflagfile()
    canencode || (println("");
    @error "Result of type $(typeof(result)) could not be " *
           "encoded for return to Excel.")

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
    truncate(x::String)
Abbreviate a string to show only `maxlength` characters.
"""
function truncate(x::String, maxlength::Int)
    if (length(x)) > maxlength
        first(x, maxlength ÷ 2) * " … " * last(x, maxlength - (maxlength ÷ 2) - 1)
    else
        x
    end
end

end # module