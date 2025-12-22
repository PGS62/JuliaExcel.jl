module JuliaExcel
export srv_xl, setxlpid, killflagfile, getcommsfolder, htd, hts

using DataFrames: DataFrames, DataFrame, Missing
using Dates: Dates, Date, DateTime
import StringEncodings

const global xlpid = Ref(0)
const global commsfolder = Ref("")

include("comms.jl")
include("encode.jl")

end