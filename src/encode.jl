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
- Dictionaries are written with type indicator H, then three sections separated by semi-colons:
  First section gives the number of key-value pairs.
  Second section gives the lengths of the encodings of each key and value in alternating order
  (key1_len, val1_len, key2_len, val2_len, ...), comma delimited with a terminating comma.
  Third section gives the encodings of each key and value concatenated with no delimiter.

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
 V   Array of Float64 only, rank 1-9, no per-element type indicator or length - see
     encode_for_xl(x::Array{Float64,N}) below. Decoded by the Case 86 'V' branch of the
     production VBA decoder (modUnserialise.bas). Also used in the opposite direction (VBA's
     TrySerialiseArrayAsV, modSerialise.bas), decoded by decode_xl_array_v (src/decode.jl).
 R   Range (UnitRange/StepRange/StepRangeLen/LinRange etc.) - encodes only first/step/length,
     not every element, so wire size doesn't scale with the range's length at all. Julia -> Excel
     only. See encode_for_xl(x::AbstractRange{Float64})/(x::AbstractRange{<:Integer}) below.
     Decoded by the 'R' branch of the production VBA decoder (modUnserialise.bas).

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

# An ExcelError decoded from VBA (see decode_from_xl's '!' branch, decode.jl) re-emits the same
# wire form it was decoded from, so an Excel error genuinely round-trips through Julia - e.g.
# JuliaCall("identity", SomeErrorCell) returns the same error, unlike the Inf/Float64 case above,
# which only ever produces #NUM!/#N/A from Julia's own numeric NaN/Inf, not a full 14-error set.
encode_for_xl(x::ExcelError) = "!" * string(x.code)

function encode_array_general(x::AbstractArray)
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

"""
    encode_for_xl(x::AbstractArray)

Fallback for any array whose *static* element type isn't already known to be Float64 (unlike the
Array{Float64,N} method below, which gets the fast "V" encoding for free via dispatch, with no
runtime check needed). Julia allows a Vector{Any} (or similar) to hold nothing but Float64 values
at runtime - e.g. from `Any[1.0, 2.0]`, `push!` onto an untyped `[]`, or values collected from
mixed sources - so this checks at runtime whether that's actually the case, and if so still takes
the fast "V" path rather than leaving that data on the slow general path just because of how it
happened to be typed. The check (a single pass testing `isa(Float64)`, `isnan`, `isinf`) is cheap
relative to the general-format encoding it would otherwise fall through to (per-element
encode_for_xl calls plus string-building for every element), so this only adds cost to paths that
were already slow - never to the already-fast concrete Array{Float64,N} case, which dispatches
straight past this method.

Deliberately uses the `Array{Float64,N}(x)` *constructor*, not `Float64.(x)` broadcasting: for a
plain Vector{Any}/Matrix{Any}/etc. they'd likely agree, but for some other AbstractArray subtype,
broadcasting can return that same container type with eltype now Float64 rather than a genuine
Array{Float64,N} - which would dispatch straight back into this same fallback method (now
trivially passing the check) and recurse without ever reaching a concrete-type base case. The
explicit constructor always produces a genuine dense Array, so the recursive call below is
guaranteed to land on the fast method and never re-enter this one.

Rank is restricted to 1-9, matching what the VBA decoder's Case 86 'V' branch
(modUnserialise.bas) supports; anything else falls through to encode_array_general unchanged.
Ranks 3-9 can only be returned to a VBA variable (JuliaEvalVBA/JuliaCallVBA), not a worksheet -
same restriction the general "*" format already has for those ranks.
"""
function encode_for_xl(x::T) where {T<:AbstractArray}
    if ndims(x) in 1:9 && !isempty(x) && all(v -> v isa Float64 && !isnan(v) && !isinf(v), x)
        return encode_for_xl(Array{Float64,ndims(x)}(x))
    end
    encode_array_general(x)
end

"""
    encode_for_xl(x::Array{Float64,N}) where {N}

Compact encoding for arrays whose element type is statically known to be Float64 - dispatch on
the concrete type means Julia has already done the work of confirming every element is a plain
Float64, so (unlike encode_array_general) neither a per-element type indicator nor a
per-element length is needed: every element is always exactly 16 hex characters. Format:
"V<rank>,<dims>;<raw hex, no delimiters>", e.g. "V1,5;..." for a 5-element vector,
"V2,3,4;..." for a 3x4 matrix, or "V3,2,3,4;..." for a 2x3x4 array - decoded by the Case 86 'V'
branch of the VBA function Unserialise (modUnserialise.bas), which relies on HexToDouble there
(the same function used for scalar "#" values). One method handles every rank since
Vector{Float64}/Matrix{Float64} are just Array{Float64,1}/Array{Float64,2}, and the encoding
itself (a single bulk byte-reinterpret over the whole linear buffer) doesn't care how many
dimensions the buffer is being viewed as.

Falls back to encode_array_general for rank > 9 (matching VBA's own ReDimVariantArray MAX_RANK),
an empty array, or one containing any NaN/Inf - the fast path assumes every element decodes as an
ordinary hex-encoded Double, which isn't true for those (see the NaN/Inf special-casing in the
scalar encode_for_xl(::Float64) above).

The hex uses the same big-endian-style convention as `float64_to_hex` (scalars), achieved by
`bswap`-ing every element (a cheap, fully vectorized bulk operation) before reinterpreting as
bytes and calling `bytes2hex`. An earlier version of this format hex-encoded the raw
little-endian memory bytes directly (skipping the bswap here, in favour of a dedicated
little-endian decode function in VBA) - but a direct measurement (VFormatDecodeSpeedTest,
modPerformance.bas) showed that VBA-side little-endian decode was enough slower, per element,
than the existing big-endian HexToDouble that the whole "V" format decoded slower than the
general "*" format it was meant to replace. Doing the bswap here instead keeps the decode side
simple and fast.
"""
function encode_for_xl(x::Array{Float64,N}) where {N}
    N > 9 && return encode_array_general(x)
    dims = size(x)
    (any(==(0), dims) || any(v -> isnan(v) || isinf(v), x)) && return encode_array_general(x)
    # reinterpret over the raw linear buffer, which Julia already stores column-major -
    # matching the wire format's column-major convention with no reshaping needed, regardless of N.
    "V" * string(N) * "," * join(dims, ",") * ";" * bytes2hex(reinterpret(UInt8, bswap.(reinterpret(UInt64, x))))
end

"""
    encode_for_xl(x::AbstractRange{Float64})
    encode_for_xl(x::AbstractRange{<:Integer})

Compact encoding for a Julia Range (`UnitRange`, `StepRange`, `StepRangeLen`, `LinRange`, etc.) -
encodes only `first(x)`, `step(x)` and `length(x)`, rather than materializing and encoding every
element as "V"/"*" would. One method per element kind covers every concrete range type, since
`first`/`step`/`length` are defined generically for all of them regardless of internal
representation - dispatch on the *element type* (`Float64` vs `<:Integer`), not the concrete range
type. This is the only array format in this file whose wire size doesn't scale with the number of
elements at all: `1:1_000_000` encodes to a few dozen bytes instead of the ~17 MB the general
format (or ~16 MB even with "V") would need.

VBA reconstructs each element by plain arithmetic (`first + (i-1)*step`) rather than decoding any
per-element wire data - verified (informally, not proven for every possible start/step/length) to
exactly reproduce Julia's own range materialization, including for `StepRangeLen`'s
twice-precision internal representation, over a 1,000,000-element case.

Integer ranges encode `first`/`step` as plain decimal (exact, matching the "^" LongLong
convention) - format "RI,<n>,<first>,<step>;". Float64 ranges encode them as 16-character
big-endian hex (matching scalar "#"/`float64_to_hex`) - format "RF,<n>;<hex first><hex step>".
Both are decoded by the 'R' branch of the VBA function Unserialise (modUnserialise.bas).

Julia -> Excel only: VBA arrays are always fully materialized, so there's no equivalent lazy
"range" concept on the encode (VBA -> Julia) side to compress this way - unlike "V", this format
isn't a candidate for the reverse direction.

Falls back to encode_array_general for an empty range, non-finite first/step (Float64 only), or -
defensively - if `last(x)` doesn't come out to exactly `first(x) + (length(x)-1)*step(x)`. Julia's
abstract type hierarchy doesn't compiler-enforce that a type subtyping AbstractRange actually
behaves as an arithmetic progression (nothing stops a third-party AbstractRange subtype from
implementing `step` to mean something else entirely) - this O(1) check verifies the assumption
this whole format rests on actually holds for the specific range in hand, rather than trusting
`AbstractRange` as a semantic promise it doesn't compiler-enforce. Base itself takes the same care
in the other direction: `Base.LogRange` (representing a geometric, not arithmetic, progression)
deliberately does *not* subtype `AbstractRange`, and doesn't even implement `step`.
"""
function encode_for_xl(x::AbstractRange{Float64})
    n = length(x)
    f, s = Float64(first(x)), Float64(step(x))
    (n == 0 || !isfinite(f) || !isfinite(s) || last(x) !== f + (n - 1) * s) && return encode_array_general(x)
    "RF," * string(n) * ";" * float64_to_hex(f) * float64_to_hex(s)
end

function encode_for_xl(x::AbstractRange{<:Integer})
    n = length(x)
    f, s = Int64(first(x)), Int64(step(x))
    (n == 0 || Int64(last(x)) != f + (n - 1) * s) && return encode_array_general(x)
    "RI," * string(n) * "," * string(f) * "," * string(s) * ";"
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