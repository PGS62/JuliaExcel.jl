#This is the code of JuliaInXL, which I have here for reference.

#Copyright (c) 2015: Julia Computing Inc. All rights reserved.
module JuliaInXL

export xldate, parse_and_eval, jlsetvar
#import Base.include_from_node1; export include_from_node1

using Reexport
@reexport using JuliaWebAPI
using Logging
using Dates
using ZMQ



global const SECONDS_PER_MINUTE = 60
global const MINUTES_PER_HOUR = 60
global const HOURS_PER_DAY = 24
global const SECONDS_PER_DAY = (HOURS_PER_DAY * MINUTES_PER_HOUR * SECONDS_PER_MINUTE)
global const DAY_MILLISECONDS = SECONDS_PER_DAY * 1000
#global const threadid = Vector{Int}(128)
global const threadid = Vector(undef,128)

#The xldata function is adapted from the code in the Apache POI project
function xldate(date::Number; use1904windowing=false, roundtoSeconds=false)
  wholeDays = floor(Int, date)
  millisInDay = round(Int, (date-wholeDays)*DAY_MILLISECONDS)
  startYear = 1900
  dayAdjust = -1 #Excel thinks 2/29/1900 is a valid date, which it isn't
  if (use1904windowing)
    startYear = 1904
    dayAdjust = 1 #// 1904 date windowing uses 1/2/1904 as the first day
  elseif (wholeDays < 61)
    dayAdjust = 0
  end

  d = DateTime(startYear, 1, 1)
  d = d+Day(wholeDays + dayAdjust - 1)
  if roundtoSeconds
    millisInDay = round(Int, millisInDay/1000)*1000
  end
  d = d + Millisecond(millisInDay)
end

parse_and_eval(arg::String) = Main.eval(Meta.parse(arg))

function jlsetvar(name::String, arg)
    y = Any[arg]
    JuliaWebAPI.narrow_args!(y)
    Main.eval(Main.eval(Meta.parse(":(global $name = $(y[1]))")))
    "Set global variable $name"
end

# Overriding base include method to avoid serializing issue
# Issue is `include` returns the last thing that it encounters in the file. Which may be something that is not serializable. To avoid the error, we add a `nothing` at the end

function include(x)
    Base.MainInclude.include(x)
    nothing
end

# entry point for new thread
function heartbeat_thread(sock::Ptr{Nothing})
    ccall((:zmq_device,ZMQ.zmq), Cint, (Cint, Ptr{Nothing}, Ptr{Nothing}),
          ZMQ.QUEUE, sock, sock)
    nothing
end

function start_heartbeat(port=9998)
	ctx=Context()
	sock=Socket(ctx, REP)
	ZMQ.bind(sock, "tcp://*:$port")
    heartbeat_c = cfunction(heartbeat_thread, Nothing, (Ptr{Nothing},))
    ccall(:uv_thread_create, Cint, (Ptr{Int}, Ptr{Nothing}, Ptr{Nothing}),
          threadid, heartbeat_c, sock.data)
end

function start_async_server(port=9999)
  global_logger(SimpleLogger(stdout, Logging.Warn))
  transport = ZMQTransport("tcp://127.0.0.1:$port", REP, true)
  msgformat = JSONMsgFormat()
  api = APIResponder(transport, msgformat, nothing, true)
  map(fn->register(api, fn), [+, -, *, /, include, parse_and_eval, jlsetvar])
  conn = process(api; async=true)

  Base.eval(Main, :(conn=$conn))
  #start_heartbeat(port-1)
  println("JuliaInXL server bound to variable 'conn' at tcp://127.0.0.1:$port")
  println("JuliaInXL server Connection accessible from tcp://localhost:$port inside of Excel")
end

end # module
