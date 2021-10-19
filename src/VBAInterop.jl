module VBAInterop
export z

using Dates

localtemp() = joinpath(ENV["TEMP"], "VBAInterop")
flagfile() = joinpath(localtemp(), "VBAInteropFlag.txt")
resultfile() = joinpath(localtemp(), "VBAInteropResult.csv")
expressionfile() = joinpath(localtemp(), "VBAInteropExpression.txt")
encode(x::String) = "\"" * replace(x, "\"" => "\"\"") * "\""
encode(x::Irrational) = Float64(x)
encode(x::DataType) = "$x"
encode(x::Any) =  x

# read a text file with UTF-16 encoding, little endian, with byte option mark
# https://discourse.julialang.org/t/reading-a-utf-16-le-file/11687
readutf16lebom(filename::String) = transcode(String, reinterpret(UInt16, read(filename)))[4:end]

function z()#One-character function name thanks to extreme slowness of SendKeys...

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
    println(expression)
    result
end

function serializeresult(x::Union{Number,String,Bool,Date,DateTime,Nothing,DataType}, 
                         filename::String, success::Bool=true)
    io = open(filename, "w")
    if success
        write(io, encode("NumDims=0|Type=$(typeof(x))") * "\n")
    else
        write(io, "NumDims=0|Type=Exception\n")
    end
    write(io, "$(encode(x))")
    close(io)
end

function reporterror(error::String, filename)
    io = open(filename, "w")
    write(io, "NumDims=0|Type=ErrorException\n")
    write(io, encode("#$(error)!"))
    close(io)
end

function serializeresult(x::Any, filename::String)
    io = open(filename, "w")
    write(io, encode("NumDims=?|Type=$(typeof(x))") * "\n")
    write(io, encode("#Expression evaluated in Julia, but returned a variable of type $(typeof(x)), which the function serializeresult cannot (yet) handle!"))
    close(io)
end

function serializeresult(x::Vector{T}, filename::String) where T
    io = open(filename, "w")
    write(io, encode("NumDims=2|Type=$(typeof(x))") * "\n")
    for i in eachindex(x)
        write(io, "$(encode(x[i]))\n")
    end    
    close(io)
end

function serializeresult(x::Matrix{T}, filename::String) where T
    nr, nc = size(x)
    io = open(filename, "w")
    write(io, encode("NumDims=2|Type=$(typeof(x))") * "\n")
    for i in 1:nr
        for j in 1:nc
            write(io, "$(encode(x[i,j]))" * (j == nc ? "\n" : ","))
    end
    end    
    close(io)
end

end # module