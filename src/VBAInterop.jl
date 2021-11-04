module VBAInterop
export z

using Dates
import StringEncodings

localtemp() = joinpath(ENV["TEMP"], "VBAInterop")
flagfile() = joinpath(localtemp(), "VBAInteropFlag_$(Main.xlpid).txt")
resultfile() = joinpath(localtemp(), "VBAInteropResult_$(Main.xlpid).txt")
expressionfile() = joinpath(localtemp(), "VBAInteropExpression_$(Main.xlpid).txt")

# read a text file with UTF-16 encoding, little endian, with byte option mark
# https://discourse.julialang.org/t/reading-a-utf-16-le-file/11687
readutf16lebom(filename::String) = transcode(String, reinterpret(UInt16, read(filename)))[4:end]

#= This function would be better named "serve_to_excel" or some such, but one-character
function name is a time saving since we have to send the function's name via PostMessage =# 
"""
    z()
Read the expression file created by Excel/VBA evaluate it and write the result to file.
"""
function z()

    expression = readutf16lebom(expressionfile())

    result = try
        Main.eval(Meta.parse(expression))
    catch e
        "#($e)!"
    end

    encodedresult = try
        encode_for_xl(result)
    catch e
        encode_for_xl("\$Expression evaluated in Julia to a variable of type $(typeof(result)) but there was a failure when encoding for return to Excel: ($e)!")
    end

    io = open(resultfile(), "w")
    write(io, StringEncodings.encode(encodedresult, "UTF-16"))
    close(io)
    
    isfile(flagfile()) && rm(flagfile())
    println(truncate(expression))
    result
end

"""
    setvar(name::String, arg)
Set a variable in global scope.    
"""
function setvar(name::String, arg)
    if Base.isidentifier(name)
        Main.eval(Main.eval(Meta.parse(":(global $name = $arg)")))
        "Set global variable `$name` to a value with type $(typeof(Main.eval(Meta.parse(name))))"
    else
        "#`$name` is not an allowed variable name in Julia!"
    end
end

#= Overriding base include method to avoid serializing issue
Issue is `include` returns the last thing that it encounters in the file. Which may be 
something that is not serializable. To avoid the error, we add a `nothing` at the end. =#
function include(x::String)
    if isfile(x)
        Base.MainInclude.include(x)
        "File `$(normpath(abspath(x)))` was included"
    else
        "#Cannot find file `$(normpath(abspath(x)))`!"
    end
end

# https://docs.microsoft.com/en-us/windows/terminal/tutorials/tab-title
function settitle()
    print("\033]0;Julia $VERSION PID $(getpid()) serving Excel PID $(Main.xlpid)\a")
end

"""
    truncate(x::String)
Abbreviate a string to show only 120 characters, the usual width of the REPL.
"""
function truncate(x::String)
    if (length(x)) > 120
        x[1:58] * " … " * x[end - 58:end]
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

  Examples:
  julia> VBAInterop.encode_for_xl(1.0)
"#1.0"

julia> VBAInterop.encode_for_xl(1)
"&1"

julia> VBAInterop.encode_for_xl("Hello")
"£Hello"

julia> VBAInterop.encode_for_xl(true)
"T"

julia> VBAInterop.encode_for_xl(false)
"F"

julia> VBAInterop.encode_for_xl(Date(2021,3,11))
"D44266"

julia> VBAInterop.encode_for_xl([1 2;true π;"Hello" "World"])
"*2,3,2;2,1,6,2,18,6,;&1T£Hello&2#3.141592653589793£World" =#

# See also VBA method Decode which unserialises i.e. inverts this function
encode_for_xl(x::String) = "£" * x        # String in VBA/Excel
encode_for_xl(x::Int64) = string("&", x)   # Long in VBA 64-bit
encode_for_xl(x::Int32) = string("&", x)   # Long in VBA 64-bit, no native 32-bit integer type exists on 64 bit Excel
encode_for_xl(x::Int16) = string("S", x)   # Integer in VBA
encode_for_xl(x::Irrational) = string("#", Float64(x))
encode_for_xl(x::Missing) = "E"            # Empty in VBA
encode_for_xl(x::Nothing) = "E"            # Empty in VBA
encode_for_xl(x::Bool) = x ? "T" : "F"     # Boolean in VBA/Excel
encode_for_xl(x::Date) = string("D", Dates.value(x) - 693594)      # Date in VBA/Excel
encode_for_xl(x::DateTime) = string("D", Dates.value(x) - 693594)  # VBA has no separate DateTime type
encode_for_xl(x::DataType) = encode_for_xl("$x")
encode_for_xl(x::VersionNumber) = encode_for_xl("$x")
encode_for_xl(x::Tuple) = encode_for_xl([x[i] for i in eachindex(x)])
encode_for_xl(x::Any) = throw("No method exists to encode a variable of type $(typeof(x)) for return to Excel")

function encode_for_xl(x::T) where T <: Function
    io = IOBuffer()
    show(io,"text/plain",x)
    encode_for_xl(String(take!(io)))
end

function encode_for_xl(x::Float64)
    if isinf(x)
        "#$(prevfloat(x, x > 0 ? 1 : -1))"
    elseif isnan(x)
        "!2042" # #N/A in Excel
    else
        string("#", x)# Double in VBA/Excel
    end
end
    
function encode_for_xl(x::Float32)
    if isinf(x)
        "#$(prevfloat(x, x > 0 ? 1 : -1))"
    elseif isnan(x)
        "!2042" # #N/A in Excel
    else
        string("S", x)# Single in VBA
    end
end

function encode_for_xl(x::T) where T <: AbstractArray

    dimssection = string(length(size(x))) * "," * join(size(x), ",")

    lengths_buf = IOBuffer()
    contents_buf = IOBuffer()

    for i in eachindex(x)
        this = encode_for_xl(x[i])
        write(contents_buf, this)
        write(lengths_buf, string(length(this)), ",")
    end   

    "*" * dimssection * ";" * String(take!(lengths_buf)) * ";" * String(take!(contents_buf))
end

end # module