#=
decode_from_xl is the inverse of encode_for_xl (in encode.jl). It decodes a string in the
JuliaExcel wire format back to a Julia value. The same format is used by VBA's
SerialiseElement function (in modSerialise.bas) to encode arguments before POSTing them to
the /call HTTP endpoint, handled by srv_call_inner (in comms.jl).

Type indicators:
 #   Float64  (followed by 16 hex chars, IEEE-754 bit pattern)
 S   Float32  (followed by 8 hex chars)
 £   String   (followed by the string content)
 T   Bool true
 F   Bool false
 E   Missing  (VBA Empty)
 N   Nothing  (VBA Null)
 %   Int16    (followed by decimal)
 &   Int32    (followed by decimal)
 ^   Int64    (followed by decimal)
 D   Date     (followed by Excel serial integer: days since 1899-12-30)
 G   DateTime (followed by 16 hex chars representing Excel serial as Float64)
 !   Error    (followed by Excel error number, e.g. 2042 for #N/A)
 *   Array    (*<rank>,<d1>[,<d2>];<len1>,<len2>,...,;<elements> column-major)
 H   Dict     (H<count>;<key1_len>,<val1_len>,...,;<key1><val1>... pairs column-order)
 V   Array of Float64 only (VBA -> Julia direction; see decode_xl_array_v below and
     TrySerialiseArrayAsV in modSerialise.bas, which produces this format)

See also VBA function SerialiseElement (modSerialise.bas) which serialises i.e. inverts this.
=#

"""
    decode_from_xl(s::String)

Decode a string in the JuliaExcel wire format and return the corresponding Julia value.
Inverse of `encode_for_xl`.
"""
function decode_from_xl(s::String)
    isempty(s) && return missing
    c = s[1]
    if c == '#'
        hex_to_float64(s[2:end])
    elseif c == '£'
        s[nextind(s, 1):end]                                   # skip the 2-byte £ prefix
    elseif c == 'T'
        true
    elseif c == 'F'
        false
    elseif c == 'E'
        missing
    elseif c == 'N'
        nothing
    elseif c == '&'
        parse(Int32, s[2:end])
    elseif c == '%'
        parse(Int16, s[2:end])
    elseif c == '^'
        parse(Int64, s[2:end])
    elseif c == 'S'
        hex_to_float32(s[2:end])
    elseif c == 'D'
        Dates.Date(1899, 12, 30) + Dates.Day(parse(Int, s[2:end]))
    elseif c == 'G'
        Dates.DateTime(1899, 12, 30) +
            Dates.Millisecond(round(Int64, hex_to_float64(s[2:end]) * 86_400_000))
    elseif c == '!'
        "#ExcelError$(s[2:end])!"
    elseif c == '*'
        decode_xl_array(s)
    elseif c == 'H'
        decode_xl_dict(s)
    elseif c == 'V'
        decode_xl_array_v(s)
    else
        error("decode_from_xl: unknown type indicator $(repr(c)) in '$(first(s, 50))'")
    end
end

"""
    xl_advance(s, start, xl_len) -> Int

Starting at byte position `start` in `s`, advance `xl_len` xl-length units and return the
byte index of the first character past the advanced region. An xl-length unit is one per
BMP character (< U+10000) and two per supplementary character (≥ U+10000), matching VBA's
`Len()` and Julia's `xl_length()`.
"""
function xl_advance(s::String, start::Int, xl_len::Int)::Int
    pos = start
    remaining = xl_len
    while remaining > 0
        b = codeunit(s, pos)
        if b < 0x80
            pos += 1; remaining -= 1            # ASCII: 1 byte, 1 xl-unit
        elseif b < 0xE0
            pos += 2; remaining -= 1            # 2-byte UTF-8: BMP char, 1 xl-unit
        elseif b < 0xF0
            pos += 3; remaining -= 1            # 3-byte UTF-8: BMP char, 1 xl-unit
        else
            pos += 4; remaining -= 2            # 4-byte UTF-8: supplementary, 2 xl-units
        end
    end
    pos
end

"""
    decode_xl_array(s::String)

Decode an array-encoded string (starting with `*`) from the JuliaExcel wire format.
Returns a typed array when all elements share the same type, otherwise `Array{Any}`.
"""
function decode_xl_array(s::String)
    # Format: *<rank>,<d1>[,<d2>];<len1>,<len2>,...,;<elements>
    p1 = findfirst(isequal(';'), s)::Int
    p2 = findnext(isequal(';'), s, p1 + 1)::Int

    # Parse rank and dims (all ASCII, safe to byte-index)
    parts = split(s[2:p1-1], ',')
    rank  = parse(Int, parts[1])
    dims  = tuple(parse.(Int, parts[2:end])...)
    n     = prod(dims; init=1)

    if n == 0
        return rank == 1 ? Any[] : Array{Any}(undef, dims...)
    end

    # Parse lengths section (comma-separated with trailing comma, all ASCII)
    lengths = Vector{Int}(undef, n)
    pos = p1 + 1
    for i in 1:n
        comma = findnext(isequal(','), s, pos)::Int
        lengths[i] = parse(Int, s[pos:comma-1])
        pos = comma + 1
    end

    # Decode elements from contents section
    elements = Vector{Any}(undef, n)
    pos = p2 + 1
    for i in 1:n
        next_pos = xl_advance(s, pos, lengths[i])
        elements[i] = decode_from_xl(s[pos:next_pos-1])
        pos = next_pos
    end

    # Reshape for 2D; try to return a typed array
    arr = rank == 1 ? elements : reshape(elements, dims)
    _maybe_typed(arr)
end

"""
    decode_xl_array_v(s::String)

Decode a "V"-format array-of-Float64 string (starting with `V`) - the compact encoding VBA's
`TrySerialiseArrayAsV` (modSerialise.bas) produces for a 1-D or 2-D array whose elements are all
finite Doubles: no per-element type indicator or length, since every element is always exactly 16
hex characters. Unlike `decode_xl_array`, this never needs `_maybe_typed` - `reinterpret`/`reshape`
already produce a genuinely typed `Vector{Float64}`/`Matrix{Float64}` directly.

The hex is big-endian (matching `hex_to_float64`'s scalar convention), produced on the VBA side by
plain `DoubleToHex` (no byte reordering needed there, since VBA already reads/writes hex MSB-first)
- decoded here by `hex2bytes` (giving the bytes in wire/big-endian order) followed by `bswap` to
get the native (little-endian) `UInt64` bit pattern before reinterpreting as `Float64`. This is the
same convention `encode_for_xl(::Vector{Float64})` (encode.jl) uses for the Julia -> Excel
direction, just decoded instead of encoded.
"""
function decode_xl_array_v(s::String)
    # Format: V<rank>,<d1>[,<d2>];<hex, no delimiters, 16 hex chars per Float64>
    p1 = findfirst(isequal(';'), s)::Int
    parts = split(s[2:p1-1], ',')
    rank = parse(Int, parts[1])
    dims = tuple(parse.(Int, parts[2:end])...)

    raw = hex2bytes(s[p1+1:end])
    vals = reinterpret(Float64, bswap.(reinterpret(UInt64, raw)))
    rank == 1 ? collect(vals) : reshape(collect(vals), dims)
end

"""
    decode_xl_dict(s::String)

Decode a dict-encoded string (starting with `H`) from the JuliaExcel wire format.
Returns a `Dict{Any,Any}` with keys and values decoded by `decode_from_xl`.
"""
function decode_xl_dict(s::String)
    # Format: H<count>;<key1_len>,<val1_len>,...,;<key1><val1>...
    p1 = findfirst(isequal(';'), s)::Int
    p2 = findnext(isequal(';'), s, p1 + 1)::Int

    n = parse(Int, s[2:p1-1])   # number of key-value pairs
    n == 0 && return Dict{Any,Any}()

    # Parse 2n lengths (alternating: key_len, val_len, ...)
    lengths = Vector{Int}(undef, 2n)
    pos = p1 + 1
    for i in 1:2n
        comma = findnext(isequal(','), s, pos)::Int
        lengths[i] = parse(Int, s[pos:comma-1])
        pos = comma + 1
    end

    # Decode key-value pairs from contents section
    result = Dict{Any,Any}()
    pos = p2 + 1
    for i in 1:n
        key_end = xl_advance(s, pos, lengths[2i-1])
        key     = decode_from_xl(s[pos:key_end-1])
        val_end = xl_advance(s, key_end, lengths[2i])
        val     = decode_from_xl(s[key_end:val_end-1])
        result[key] = val
        pos = val_end
    end
    result
end

# Convert Array{Any} to a typed array when all elements share the same concrete type.
function _maybe_typed(a::Array{Any})
    isempty(a) && return a
    T = typeof(a[1])
    all(x -> typeof(x) === T, a) || return a
    try
        convert(Array{T}, a)
    catch
        a
    end
end
