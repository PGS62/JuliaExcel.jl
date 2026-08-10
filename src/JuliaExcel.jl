module JuliaExcel
export start_server, setxlpid, getcommsfolder

using DataFrames: DataFrames, DataFrame, Missing
using Dates: Dates, Date, DateTime
using HTTP: HTTP
using Sockets: Sockets

const global xlpid = Ref(0)
const global commsfolder = Ref("")
const global xlport = Ref(0)

include("comms.jl")
include("encode.jl")
include("decode.jl")

end