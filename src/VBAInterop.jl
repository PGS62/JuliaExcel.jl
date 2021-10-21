module VBAInterop
export z

using Dates

localtemp() = joinpath(ENV["TEMP"], "VBAInterop")
flagfile() = joinpath(localtemp(), "VBAInteropFlag_$(Main.xlpid).txt")
resultfile() = joinpath(localtemp(), "VBAInteropResult_$(Main.xlpid).csv")
expressionfile() = joinpath(localtemp(), "VBAInteropExpression_$(Main.xlpid).txt")
encode(x::String) = "\"" * replace(x, "\"" => "\"\"") * "\""
encode(x::Irrational) = Float64(x)
encode(x::DataType) = "$x"
encode(x::VersionNumber) = "$x"
encode(x::Missing) = ""#will end up in Excel as the value of the ShowMissingsAs argument to CSVRead
encode(x::Any) =  x

# read a text file with UTF-16 encoding, little endian, with byte option mark
# https://discourse.julialang.org/t/reading-a-utf-16-le-file/11687
readutf16lebom(filename::String) = transcode(String, reinterpret(UInt16, read(filename)))[4:end]

#=This function would be better named "serve_to_excel" or some such, but one-character
function name is a time saving since we have to send the function's name via SendMessage 
=#    
function z()

    expression = readutf16lebom(expressionfile())

    #special case!
    if expression == "exit()"
        result = "Julia has shut down"
        serializeresult(result, resultfile())
        isfile(flagfile()) && rm(flagfile())
        exit()
    end

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

function serializeresult(x::Union{Number,String,Bool,Date,DateTime,Nothing,DataType,Missing,VersionNumber}, 
                         filename::String, success::Bool=true)
    io = open(filename, "w")
    if success
        write(io, encode("NumDims=0|Type=$(typeof(x))") * "\n")
    else
        write(io, "NumDims=0|Type=Exception\n")
    end
    write(io, "$(encode(x))")
    write(io, "\n")
    close(io)
end

function reporterror(error::String, filename)
    io = open(filename, "w")
    write(io, "NumDims=0|Type=ErrorException\n")
    write(io, encode("#$(error)!"))
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
        write(io, encode("NumDims=?|Type=$(typeof(x))") * "\n")
        write(io, encode("#Expression evaluated in Julia, but returned a variable of type $(typeof(x)), which the function serializeresult does not (currently) handle!"))
        close(io)
    end
end

function serializeresult(x::Vector{T}, filename::String, thetype::DataType=typeof(x)) where T
    io = open(filename, "w")
    write(io, encode("NumDims=1|Type=$(thetype)") * "\n")
    for i in eachindex(x)
        write(io, "$(encode(x[i]))\n")
    end    
    close(io)
end

function serializeresult(x::Matrix{T}, filename::String, thetype::DataType=typeof(x)) where T
    nr, nc = size(x)
    io = open(filename, "w")
    write(io, encode("NumDims=2|Type=$(thetype)") * "\n")
    for i in 1:nr
        for j in 1:nc
            write(io, "$(encode(x[i,j]))" * (j == nc ? "\n" : ","))
        end
    end    
    close(io)
end

#https://docs.microsoft.com/en-us/windows/terminal/tutorials/tab-title
function settitle()
    print("\033]0;Julia $VERSION serving Excel PID $(Main.xlpid)\a")
end

function truncate(x::String)
    if (length(x)) > 120
        x[1:58] * " … " * x[end - 58:end]
    else
        x
    end
end

end # module