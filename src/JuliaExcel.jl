module JuliaExcel
export srv_xl, setxlpid, killflagfile, getcommsfolder

using Dates, DataFrames
import StringEncodings
const global xlpid = Ref(0)
const global commsfolder = Ref("")

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
    #  println(truncate(expression,120))
    #  display(result)
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

#= 
encode_for_xl implements a data serialisation format that's easier and faster to
unserialise than csv.
- Singleton types are prefixed with a type indicator character.
- Dates are converted to their Excel representation - faster to unserialise in VBA.
- Arrays are written with type indicator *, then three sections separated by semi-colons:
  First section gives the number of dimensions and the dimensions themselves, comma
  delimited e.g. a 3 x 4 array would have a dimensions section "2,3,4".
  Second section gives the lengths of the encodings of each element, comma delimited with a 
  terminating comma.
  Third section gives the encodings, concatenated with no delimiter.
  - Note that arrays are written in column-major order.

When decoded (by VBA function modDecode.Decode), the type indicator characters are 
interpreted as follows:
 #   vbDouble
 £   String
 T   Boolean True
 F   Boolean False
 D   Date (D should be followed by the number that represents the date, Excel-style
           i.e. Dates.value(x) - 693594)
 E   Empty
 N   Null
 %   Integer
 &   Long
 S   Single
 C   Currency
 !   Error (! should be followed by an Excel error number, e.g. 
              2042 for the Excel error value #N/A )
 @   Decimal
 *   Array
 ^   Dictionary

  Examples:
  julia> JuliaExcel.encode_for_xl(1.0)
"#1.0"

julia> JuliaExcel.encode_for_xl(1)
"&1"

julia> JuliaExcel.encode_for_xl("Hello")
"£Hello"

julia> JuliaExcel.encode_for_xl(true)
"T"

julia> JuliaExcel.encode_for_xl(false)
"F"

julia> JuliaExcel.encode_for_xl(Date(2021,3,11))
"D44266"

julia> JuliaExcel.encode_for_xl([1 2;true π;"Hello" "World"])
"*2,3,2;2,1,6,2,18,6,;&1T£Hello&2#3.141592653589793£World" =#

# See also VBA method Decode which unserialises i.e. inverts this function
encode_for_xl(x::AbstractString) = "£" * x         # String in VBA/Excel
encode_for_xl(x::AbstractChar) = "£" * x           # String in VBA/Excel
encode_for_xl(x::Int8) = string("S", x)   # Integer in VBA
encode_for_xl(x::Int16) = string("S", x)   # Integer in VBA
encode_for_xl(x::Int32) = string("&", x)   # Long in VBA 64-bit, no native 32-bit integer
# type exists on 64 bit Excel
encode_for_xl(x::Int64) = string("&", x)   # Long in VBA 64-bit
encode_for_xl(x::Int128) = string("#", Float64(x))   # Double in VBA
encode_for_xl(x::Irrational) = string("#", Float64(x)) #Double in VBA
encode_for_xl(x::Missing) = "E"            # Empty in VBA
encode_for_xl(x::Nothing) = "E"            # Empty in VBA
encode_for_xl(x::Bool) = x ? "T" : "F"     # Boolean in VBA/Excel
encode_for_xl(x::Date) = string("D", Dates.value(x) - 693594) # Date in VBA/Excel
encode_for_xl(x::DateTime) = string("D", Dates.value(x) / 86_400_000 - 693594)  # VBA has no separate DateTime type
encode_for_xl(x::DataType) = wrapshow(x)
encode_for_xl(x::VersionNumber) = encode_for_xl("$x")
encode_for_xl(x::Tuple) = encode_for_xl([x[i] for i in eachindex(x)])
encode_for_xl(x::T) where {T<:Function} = wrapshow(x)
encode_for_xl(x::Symbol) = wrapshow(x)
encode_for_xl(x::Any) = wrapshow(x)        # Fallback

function wrapshow(x)
    io = IOBuffer()
    show(io, "text/plain", x)
    encode_for_xl(String(take!(io)))
end

function encode_for_xl(x::Float64)
    if isinf(x)
        "!2036" # #NUM! in Excel
    elseif isnan(x)
        "!2042" # #N/A in Excel
    else
        string("#", x)# Double in VBA/Excel
    end
end

function encode_for_xl(x::Union{Float16,Float32})
    if isinf(x)
        "!2036" # #NUM! in Excel
    elseif isnan(x)
        "!2042" # #N/A in Excel
    else
        string("S", x)# Single in VBA
    end
end

#=PGS 8 Dec 2021 Seeing problem:
The applicable method may be too new: running in world age 31344, while current world is 31485.
Closest candidates are:
  size(::SentinelArrays.MissingVector) at ~/.julia/packages/SentinelArrays/iHRpO/src/missingvector.jl:6 (method too new to be called from this world context.)

  Solutions tried
1)  
Used Base.invokelatest(size,x). Bit surprised this did not work...
2) write method of encode_for_xl(x::SentinelArrays.MissingVector), which
    a) requires SentinelArrays be a dependency of JuliaExcel :(
    b) solves the problem for SentinelArrays.MissingVector, but the same problem pops up
    but for NamedArrays.NamedArray
=#

#function encode_for_xl(x::SentinelArrays.MissingVector)
#    l = x.len
#    "*1,$l;$("1,"^l);$("E"^l)"
#end

function encode_for_xl(x::T) where {T<:AbstractArray}

    sx = size(x)
    dimssection = string(xl_length(sx)) * "," * join(sx, ",")
    lengths_buf = IOBuffer()
    contents_buf = IOBuffer()

    for i in eachindex(x)
        this = encode_for_xl(x[i])
        write(contents_buf, this)
        write(lengths_buf, string(xl_length(this)), ",")
    end

    "*" * dimssection * ";" * String(take!(lengths_buf)) * ";" * String(take!(contents_buf))
end

function encode_for_xl(x::DataFrame)
    nc = size(x)[2]
    data = Matrix{Any}(x)
    headers = Matrix{Any}(undef, 1, nc)
    for i in 1:nc
        headers[1, i] = names(x)[i]
    end
    encode_for_xl(vcat(headers, data))
end

function encode_for_xl(x::T) where {T<:AbstractDict}

    dimssection = string(xl_length(x))
    lengths_buf = IOBuffer()
    contents_buf = IOBuffer()

    for (k, v) in x
        thiskey = encode_for_xl(k)
        thisvalue = encode_for_xl(v)
        write(contents_buf, thiskey)
        write(contents_buf, thisvalue)
        write(lengths_buf, string(xl_length(thiskey)), ",")
        write(lengths_buf, string(xl_length(thisvalue)), ",")
    end

    "^" * dimssection * ";" * String(take!(lengths_buf)) * ";" * String(take!(contents_buf))
end

"""
    xl_length(x)
If `x` is a `Char` or `String` then `xl_length` emulates the VBA function `Len` which
returns the number of characters in a string except that characters with code point 65536
or above count as 2 instead of 1. Otherwise `xl_length` returns the (Julia) `length` of `x`.
"""
function xl_length(x::Char)
    return (codepoint(x) >= 65536 ? 2 : 1)
end
function xl_length(x::String)
    out = 0
    for char in x
        out += xl_length(char)
    end
    out
end
function xl_length(x::Any)
    length(x)
end

end # module