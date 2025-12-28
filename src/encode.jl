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

When decoded (by VBA function modSerialise.Unserialise), the type indicator characters are 
interpreted as follows:
 #   Double (followed by hex represention of the value, see float64_to_hex)
 £   String (followed by the string)
 T   Boolean True
 F   Boolean False
 D   Date (followed by the number that represents the date, Excel-style
           i.e. Dates.value(x) - 693594)
 G   Date (with time, no separate type exists in VBA. Followed by hex representation of the 
     Double that is equivalent in Excel)
 E   Empty
 N   Null
 %   Integer (followed by decimal representation of the value)
 &   Long (followed by decimal representation of the value)
 ^   LongLong (followed by decimal representation of the value)
 S   Single (followed by hex represention of the value, see float32_to_hex)
 !   Error (followed by an Excel error number, e.g. 
              2042 for the Excel error value #N/A )
 *   Array
 H   Dictionary

  Examples:
  julia> JuliaExcel.encode_for_xl(1.0)
"#3FF0000000000000"

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

# See also VBA method Unserialise which unserialises i.e. inverts this function
encode_for_xl(x::AbstractString) = "£" * x         # String in VBA/Excel
encode_for_xl(x::AbstractChar) = "£" * x           # String in VBA/Excel
encode_for_xl(x::Int8) = string("%", x)   # Integer in VBA
encode_for_xl(x::Int16) = string("%", x)   # Integer in VBA
encode_for_xl(x::Int32) = string("&", x)   # Long in VBA 64-bit, no native 32-bit integer
# type exists on 64 bit Excel
encode_for_xl(x::Int64) = string("^", x)   # LongLong in VBA 64-bit
encode_for_xl(x::Int128) = encode_for_xl(Float64(x))   # Double in VBA
encode_for_xl(x::Irrational) = encode_for_xl(Float64(x)) #Double in VBA
encode_for_xl(x::Missing) = "E"            # Empty in VBA
encode_for_xl(x::Nothing) = "E"            # Empty in VBA
encode_for_xl(x::Bool) = x ? "T" : "F"     # Boolean in VBA/Excel
encode_for_xl(x::Date) = string("D", Dates.value(x) - 693594) # Date in VBA/Excel
encode_for_xl(x::DateTime) = "G" * float64_to_hex(Dates.value(x) / 86_400_000 - 693594)  # VBA has no separate DateTime type
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
        "#" * float64_to_hex(x)
    end
end

function encode_for_xl(x::Float32)
    if isinf(x)
        "!2036" # #NUM! in Excel
    elseif isnan(x)
        "!2042" # #N/A in Excel
    else
        "S" * float32_to_hex(x)# Single in VBA
    end
end

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

    "H" * dimssection * ";" * String(take!(lengths_buf)) * ";" * String(take!(contents_buf))
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

"""
    float64_to_hex(x::Float64)::String

Return a 16-character uppercase hexadecimal string representing the IEEE-754
bit pattern of `x` (Float64). Does not special-case NaN or +0.0 and -0.0.
"""
function float64_to_hex(x::Float64)::String
    bits = reinterpret(UInt64, x)
    s = uppercase(string(bits, base=16))
    return lpad(s, 16, '0')
end

"""
    hex_to_float64(hex::AbstractString)::Float64

Parse a 16-character hex string (uppercase or lowercase) as the IEEE-754
bit pattern of a `Float64` and return the corresponding `Float64` value.
"""
function hex_to_float64(hex::AbstractString)::Float64

    length(hex) == 16 || throw(ArgumentError("input must be 16 characters, but got $(length(hex))"))

    bits = parse(UInt64, hex; base=16)
    return reinterpret(Float64, bits)
end

"""
    float32_to_hex(x::Float32)::String

Return an 8-character uppercase hexadecimal string representing the IEEE-754
bit pattern of `x` (Float32). Does not special-case NaN or +0.0 and -0.0.
"""
function float32_to_hex(x::Float32)::String
    bits = reinterpret(UInt32, x)
    s = uppercase(string(bits, base=16))
    return lpad(s, 8, '0')
end

"""
    hex_to_float32(hex::AbstractString)::Float32

Parse an 8-character hex string (uppercase or lowercase) as the IEEE-754
bit pattern of a `Float32` and return the corresponding `Float32` value.
"""
function hex_to_float32(hex::AbstractString)::Float32
    length(hex) == 8 || throw(ArgumentError("input must be 8 characters, but got $(length(hex))"))
    bits = parse(UInt32, hex; base=16)
    return reinterpret(Float32, bits)
end

# For brevity in the output of the VBA function MakeJuliaLiteral
htd = hex_to_float64
hts = hex_to_float32