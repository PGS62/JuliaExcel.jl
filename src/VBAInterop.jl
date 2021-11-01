module VBAInterop
export z

using Dates
using StringEncodings

localtemp() = joinpath(ENV["TEMP"], "VBAInterop")
flagfile() = joinpath(localtemp(), "VBAInteropFlag_$(Main.xlpid).txt")
resultfile() = joinpath(localtemp(), "VBAInteropResult_$(Main.xlpid).csv")
expressionfile() = joinpath(localtemp(), "VBAInteropExpression_$(Main.xlpid).txt")
encode_for_csv(x::String) = "\"" * replace(x, "\"" => "\"\"") * "\""
encode_for_csv(x::Irrational) = "$(Float64(x))"
encode_for_csv(x::Missing) = ""#will end up in Excel as the value of the ShowMissingsAs argument to CSVRead
encode_for_csv(x::Any) =  "$x"

# read a text file with UTF-16 encoding, little endian, with byte option mark
# https://discourse.julialang.org/t/reading-a-utf-16-le-file/11687
readutf16lebom(filename::String) = transcode(String, reinterpret(UInt16, read(filename)))[4:end]

#=This function would be better named "serve_to_excel" or some such, but one-character
function name is a time saving since we have to send the function's name via PostMessage 
=#    
function z()

    expression = readutf16lebom(expressionfile())

    success = true
    result =
    try
        Main.eval(Meta.parse(expression))
    catch e
        success = false
        "$e"
    end

    if success
        serializeresult(result, resultfile())
    else
        reporterror(result, resultfile()) 
    end
    
    isfile(flagfile()) && rm(flagfile())
    println(truncate(expression))
    result
end

function setvar(name::String,arg)
    if Base.isidentifier(name)
        Main.eval(Main.eval(Meta.parse(":(global $name = $arg)")))
        "Set global variable `$name` to a value with type $(typeof(Main.eval(Meta.parse(name))))"
    else
        "#`$name` is not an allowed variable name in Julia!"
    end
end

# Overriding base include method to avoid serializing issue
# Issue is `include` returns the last thing that it encounters in the file. Which may be something that is not serializable. To avoid the error, we add a `nothing` at the end
function include(x::String)
    if isfile(x)
        Base.MainInclude.include(x)
        "File `$(normpath(abspath(x)))` was included"
    else
        "#Cannot find file `$(normpath(abspath(x)))`!"
    end
end

function serializeresult(x::Union{Number,String,Bool,Date,DateTime,Nothing,DataType,Missing,VersionNumber}, 
                         filename::String, success::Bool=true)
    io = open(filename, "w")
    if success
        write(io, encode_for_csv("NumDims=0|Type=$(typeof(x))") * "\n")
    else
        write(io, "NumDims=0|Type=Exception\n")
    end
    write(io, encode_for_csv(x))
    write(io, "\n")
    close(io)
end

function reporterror(error::String, filename)
    io = open(filename, "w")
    write(io, "NumDims=0|Type=ErrorException\n")
    write(io, encode_for_csv("#$(error)!"))
    close(io)
end

function serializeresult(x::Any, filename::String)

    success = true
    xasarray = try
        Array(x)
    catch
        success = false
    end

    if success
        success = length(size(x)) <= 2
    end

    if success
        serializeresult(xasarray, filename, typeof(x))
    else
        io = open(filename, "w")
        write(io, encode_for_csv("NumDims=?|Type=$(typeof(x))") * "\n")
        write(io, encode_for_csv("#Expression evaluates to a variable of type $(typeof(x)), and no method exists to return variables of that type to Excel!"))
        close(io)
    end
end

function serializeresult(x::Vector{T}, filename::String, thetype::DataType=typeof(x)) where T
    io = open(filename, "w")
    write(io, encode_for_csv("NumDims=1|Type=$(thetype)") * "\n")
    for i in eachindex(x)
        write(io, "$(encode_for_csv(x[i]))\n")
    end    
    close(io)
end

function serializeresult(x::Matrix{T}, filename::String, thetype::DataType=typeof(x)) where T
    nr, nc = size(x)
    io = open(filename, "w")
    write(io, encode_for_csv("NumDims=2|Type=$(thetype)") * "\n")
    for i in 1:nr
        for j in 1:nc
            write(io, encode_for_csv(x[i,j]) * (j == nc ? "\n" : ","))
        end
    end    
    close(io)
end

#https://docs.microsoft.com/en-us/windows/terminal/tutorials/tab-title
function settitle()
    print("\033]0;Julia $VERSION PID $(getpid()) serving Excel PID $(Main.xlpid)\a")
end

function truncate(x::String)
    if (length(x)) > 120
        x[1:58] * " … " * x[end - 58:end]
    else
        x
    end
end


end # module