module JuliaExcel
export start_server, setxlpid, getcommsfolder, ExcelError

using DataFrames: DataFrames, DataFrame, Missing
using Dates: Dates, Date, DateTime
using HTTP: HTTP
using Sockets: Sockets

const global xlpid = Ref(0)
const global commsfolder = Ref("")
const global xlport = Ref(0)

"""
    ExcelError(code::Int)

Represents an Excel error value (e.g. `#DIV/0!`, `#N/A`) received from VBA. Decoded from the wire
format's `!` type indicator (see `decode_from_xl` in decode.jl) and re-encoded the same way by
`encode_for_xl` (encode.jl), so an Excel error genuinely round-trips through Julia unchanged - e.g.
`JuliaCall("identity", SomeErrorCell)` returns the same error.

Deliberately a distinct type rather than a `String`: a Julia function that isn't written to expect
an Excel error, and receives one where it expects an ordinary value, fails immediately with a
`MethodError` - rather than silently treating error text as ordinary string data. Functions that
want to handle Excel errors explicitly can dispatch on this type, e.g. `myfunc(::ExcelError) =
missing` to swallow them, or `myfunc(e::ExcelError) = e` to propagate them onward unchanged.

The numeric code is stored as-is, with no validation against the current list of known Excel error
codes - Excel could add new ones in future versions, and any code decodes/re-encodes identically
regardless of whether it's recognised (see `EXCEL_ERROR_NAMES` below, used only for display).
"""
struct ExcelError
    code::Int
end

# Display names for known Excel error codes, purely for a friendlier `show` - a code not in this
# table (e.g. a future Excel error not yet known to this package) still round-trips correctly,
# just without a friendly name.
const EXCEL_ERROR_NAMES = Dict(
    2000 => "#NULL!",
    2007 => "#DIV/0!",
    2015 => "#VALUE!",
    2023 => "#REF!",
    2029 => "#NAME?",
    2036 => "#NUM!",
    2042 => "#N/A",           # no trailing "!"
    2043 => "#GETTING_DATA",  # no trailing "!"
    2045 => "#SPILL!",
    2046 => "#CONNECT!",
    2047 => "#BLOCKED!",
    2048 => "#UNKNOWN!",
    2049 => "#FIELD!",
    2050 => "#CALC!",
)

Base.show(io::IO, e::ExcelError) = print(io, get(EXCEL_ERROR_NAMES, e.code, "#ERROR($(e.code))!"))

include("comms.jl")
include("encode.jl")
include("decode.jl")

end