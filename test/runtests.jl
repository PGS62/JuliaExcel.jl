using JuliaExcel
using Dates
using Test

# isequal alone is type-blind (isequal(2, 2.0) is true), so it would miss a type morph introduced
# on the wire (e.g. decode returning Int32 where the input was Int64). === is too strict the other
# way: it's identity comparison for mutable types, so it fails for every array/Dict round trip even
# when content and type both match, since decode_from_xl always builds a fresh object. Checking
# typeof equality alongside isequal gets both right - exact type, and isequal's correct handling of
# NaN/missing/nested structures for the value itself.
function round_trip(x)
    y = JuliaExcel.decode_from_xl(JuliaExcel.encode_for_xl(x))
    typeof(y) == typeof(x) && isequal(y, x)
end

@testset "JuliaExcel.jl" begin
    #For compatibility with VBA function Unserialise (which should be exact equivalent of decode_from_excel)
    @test JuliaExcel.encode_for_xl(1) == "^1"
    @test JuliaExcel.encode_for_xl(1.0) == "#3FF0000000000000"
    @test JuliaExcel.encode_for_xl(Int8(1)) == "%1"
    @test JuliaExcel.encode_for_xl(Int16(1)) == "%1"
    @test JuliaExcel.encode_for_xl(Int32(1)) == "&1"
    @test JuliaExcel.encode_for_xl(Int64(1)) == "^1"
    @test JuliaExcel.encode_for_xl(UInt8(200)) == "B200"
    @test JuliaExcel.decode_from_xl("B200") == UInt8(200)
    @test JuliaExcel.encode_for_xl(true) == "T"
    @test JuliaExcel.encode_for_xl(false) == "F"
    @test JuliaExcel.encode_for_xl("foo") == "£foo"
    @test JuliaExcel.encode_for_xl('x') == "£x"
    @test JuliaExcel.encode_for_xl(:x) == "£:x"
    @test JuliaExcel.encode_for_xl(nothing) == "E"
    @test JuliaExcel.encode_for_xl(missing) == "E"
    @test JuliaExcel.encode_for_xl(Inf) == "!2036"
    @test JuliaExcel.encode_for_xl(-Inf) == "!2036"
    @test JuliaExcel.encode_for_xl(NaN) == "!2042"
    @test JuliaExcel.encode_for_xl(JuliaExcel.ExcelError(2007)) == "!2007"
    @test JuliaExcel.decode_from_xl("!2007") == JuliaExcel.ExcelError(2007)
    @test JuliaExcel.encode_for_xl(Date("2021-11-8")) == "D44508"
    @test JuliaExcel.encode_for_xl(DateTime("2021-11-8T12:00:00")) == "G40E5BB9000000000"
    @test JuliaExcel.encode_for_xl(Int64) == "£Int64"
    @test JuliaExcel.encode_for_xl(v"1.2.3") == "£1.2.3"
    @test JuliaExcel.encode_for_xl((1, 2)) == "*1,2;2,2,;^1^2"
    @test JuliaExcel.encode_for_xl([1, 2, 3]) == "*1,3;2,2,2,;^1^2^3"
    @test JuliaExcel.encode_for_xl(Any[1, 2, 3.0, π]) == "*1,4;2,2,17,17,;^1^2#4008000000000000#400921FB54442D18"
    @test JuliaExcel.encode_for_xl([1, true, "x"]) == "*1,3;2,1,2,;^1T£x"
    @test JuliaExcel.encode_for_xl([1, [2, 3]]) == "*1,2;2,14,;^1*1,2;2,2,;^2^3"
    # Dict key order isn't guaranteed by the language, and can differ between Julia versions/hash
    # seeds - accept either encoding order rather than hardcoding one.
    @test JuliaExcel.encode_for_xl(Dict("a" => 1, "b" => 2)) in
          ("H2;2,2,2,2,;£a^1£b^2", "H2;2,2,2,2,;£b^2£a^1")

    @test round_trip(1)
    @test round_trip(1.0)
  #  @test round_trip(Int8(1)) # No dedicated Int8 wire type - morphs to Int16, the smallest one
    @test round_trip(Int16(1))
    @test round_trip(Int32(1))
    @test round_trip(Int64(1))
    @test round_trip(UInt8(200))
    @test round_trip(true)
    @test round_trip(false)
    @test round_trip("foo")
  #  @test round_trip('x') # Characters morph to String
  #  @test round_trip(:x)  # Symbols morph to String
  #  @test round_trip(nothing) # Nothing morphs to missing (Empty in Excel)
     @test round_trip(missing)
  #  @test round_trip(Inf)  # Morphs to ExcelError(2036) (#NUM! in Excel)
  #  @test round_trip(-Inf) # Morphs to ExcelError(2036) (#NUM! in Excel)
  #  @test round_trip(NaN)  # Morphs to ExcelError(2042) (#N/A in Excel)
    @test round_trip(Date("2021-11-8"))
    @test round_trip(DateTime("2021-11-8T12:00:00"))
  #  @test round_trip(Int64) # Types morph to String
  #  @test round_trip(v"1.2.3") # VersionNumber morph to String
  #  @test round_trip((1, 2)) # Tuples morph to Vector
    @test round_trip([1, 2, 3])
  #  @test round_trip(Any[1, 2, 3.0, π]) #Irrational morphs to nearest Float64
    @test round_trip([1, true, "x"])
    @test round_trip([1, [2, 3]])
    @test round_trip(Dict("a"=>1, "b"=>2))
    @test typeof(JuliaExcel.decode_from_xl(JuliaExcel.encode_for_xl(Dict("a"=>1, "b"=>2)))) == Dict{String,Int64}

    # Mixed-type dict correctly declines the typed conversion, staying Dict{Any,Any} - matching
    # decode_xl_array's own dynamic fallback for a Vector{Any} with mixed element types.
    @test JuliaExcel.decode_from_xl(JuliaExcel.encode_for_xl(Dict("a"=>1, "b"=>"x"))) isa Dict{Any,Any}

    # ExcelError - unlike Inf/NaN above (which only ever produce 2036/2042 from Julia's own
    # numeric values), every documented Excel error code round-trips exactly through Julia when it
    # originates as a genuine VBA error - see ExcelError's docstring (JuliaExcel.jl) for why this is
    # a distinct type rather than a String.
    for code in keys(JuliaExcel.EXCEL_ERROR_NAMES)
        @test JuliaExcel.decode_from_xl("!$code") == JuliaExcel.ExcelError(code)
        @test round_trip(JuliaExcel.ExcelError(code))
    end

    # "V" format - compact encoding for arrays of Float64 (see encode_for_xl(::Vector{Float64})/
    # (::Matrix{Float64}) and the dynamic AbstractArray fallback in src/encode.jl).
    @test JuliaExcel.encode_for_xl([1.0, 2.0, 3.0]) == "V1,3;3ff000000000000040000000000000004008000000000000"
    @test JuliaExcel.encode_for_xl([1.0 2.0; 3.0 4.0]) == "V2,2,2;3ff0000000000000400800000000000040000000000000004010000000000000"

    # Empty array and NaN/Inf-containing arrays must fall back to the general "*" format, not "V" -
    # every element of "V" is assumed to be an ordinary finite Double, which isn't true for these.
    @test JuliaExcel.encode_for_xl(Float64[]) == JuliaExcel.encode_array_general(Float64[])
    @test JuliaExcel.encode_for_xl([1.0, NaN, 3.0]) == JuliaExcel.encode_array_general([1.0, NaN, 3.0])
    @test JuliaExcel.encode_for_xl([1.0, Inf, 3.0]) == JuliaExcel.encode_array_general([1.0, Inf, 3.0])
    @test !startswith(JuliaExcel.encode_for_xl([1.0, NaN, 3.0]), "V")
    @test !startswith(JuliaExcel.encode_for_xl([1.0, Inf, 3.0]), "V")

    # Dynamic fallback: a Vector{Any}/Matrix{Any} that's actually all Float64 at runtime should
    # still get the fast "V" encoding, byte-for-byte identical to the concrete Vector{Float64}/
    # Matrix{Float64} case - dispatch alone (the type being literally Vector{Float64}) isn't the
    # only way to reach it.
    @test JuliaExcel.encode_for_xl(Any[1.0, 2.0, 3.0]) == JuliaExcel.encode_for_xl([1.0, 2.0, 3.0])
    @test JuliaExcel.encode_for_xl(Any[1.0 2.0; 3.0 4.0]) == JuliaExcel.encode_for_xl([1.0 2.0; 3.0 4.0])
    @test startswith(JuliaExcel.encode_for_xl(Any[1.0, 2.0, 3.0]), "V")

    # ...but the dynamic check must still correctly decline "V" (falling back to the general
    # format) for a Vector{Any} containing a NaN, a non-Float64 element, or with rank > 9 (matching
    # VBA's own ReDimVariantArray MAX_RANK, modUnserialise.bas).
    @test JuliaExcel.encode_for_xl(Any[1.0, 2.0, NaN]) == JuliaExcel.encode_for_xl([1.0, 2.0, NaN])
    @test !startswith(JuliaExcel.encode_for_xl(Any[1.0, 2.0, NaN]), "V")
    @test JuliaExcel.encode_for_xl(Any[1.0, 2, 3.0]) == JuliaExcel.encode_array_general(Any[1.0, 2, 3.0])
    @test !startswith(JuliaExcel.encode_for_xl(Any[1.0, 2, 3.0]), "V")
    @test JuliaExcel.encode_for_xl(Any[]) == JuliaExcel.encode_array_general(Any[])

    # Rank 3-9: one method (encode_for_xl(x::Array{Float64,N})) handles every rank, since
    # Vector{Float64}/Matrix{Float64} are just Array{Float64,1}/Array{Float64,2} and the bulk
    # byte-reinterpret doesn't care how many dimensions the buffer is viewed as. Decoded by the
    # Case 86 'V' branch's Case Else (rank 3-9) in Unserialise (modUnserialise.bas), which reuses
    # ParseDims/ReDimVariantArray/AssignByRank - the same helpers the general "*" format's own
    # >=3-D handling already used.
    let a3d = reshape(collect(1.0:24.0), 2, 3, 4)
        @test startswith(JuliaExcel.encode_for_xl(a3d), "V3,2,3,4;")
        @test round_trip(a3d)
        @test JuliaExcel.decode_from_xl(JuliaExcel.encode_for_xl(a3d)) isa Array{Float64,3}
        # dynamic AbstractArray fallback also generalises to rank 3-9, not just 1-2
        @test JuliaExcel.encode_for_xl(Array{Any}(a3d)) == JuliaExcel.encode_for_xl(a3d)
    end

    # Rank > 9 must fall back to the general "*" format - matches VBA's own ReDimVariantArray
    # MAX_RANK, which Case 86's Case Else would otherwise be asked to exceed.
    let a10d = reshape(collect(1.0:2.0^10), fill(2, 10)...)
        @test !startswith(JuliaExcel.encode_for_xl(a10d), "V")
        @test round_trip(a10d)
    end

    let x3d = fill(1.0, 2, 2, 2) |> a -> Array{Any}(a)
        @test startswith(JuliaExcel.encode_for_xl(x3d), "V3,")  # now gets "V", not the general format
        @test JuliaExcel.encode_for_xl(x3d) == JuliaExcel.encode_for_xl(fill(1.0, 2, 2, 2))
    end

    # decode_xl_array_v (src/decode.jl) - the VBA -> Julia direction, added alongside
    # TrySerialiseArrayAsV (modSerialise.bas). Real round trips are possible now that
    # decode_from_xl understands "V", unlike the encode-only assertions above.
    @test round_trip([1.0, 2.0, 3.0])
    @test round_trip([1.0 2.0 3.0; 4.0 5.0 6.0])  # non-square, to catch a row/column transposition bug
    @test JuliaExcel.decode_from_xl(JuliaExcel.encode_for_xl([1.0, 2.0, 3.0])) isa Vector{Float64}
    @test JuliaExcel.decode_from_xl(JuliaExcel.encode_for_xl([1.0 2.0; 3.0 4.0])) isa Matrix{Float64}

    # Hand-built, VBA-style big-endian hex (plain MSB-first hex per element, matching what
    # DoubleToHex in modSerialise.bas actually produces) - confirms decode_xl_array_v works against
    # genuinely VBA-encoded-style input, not only against Julia's own encoder output.
    let be_hex(v) = uppercase(string(reinterpret(UInt64, v), base=16, pad=16))
        manual = "V1,3;" * be_hex(1.0) * be_hex(-2.5) * be_hex(3.14159265358979)
        @test JuliaExcel.decode_from_xl(manual) == [1.0, -2.5, 3.14159265358979]

        # Rank 3, exercising decode_xl_array_v's general reshape path against hand-built VBA-style
        # hex - this is what TrySerialiseArrayAsV's new Case Else (modSerialise.bas) now produces
        # for the Excel -> Julia direction of a 3D+ array.
        vals3d = collect(1.0:24.0)
        manual3d = "V3,2,3,4;" * join(be_hex.(vals3d))
        @test JuliaExcel.decode_from_xl(manual3d) == reshape(vals3d, 2, 3, 4)
    end

    # "R" format - compact encoding for Ranges (UnitRange/StepRange/StepRangeLen/LinRange etc.):
    # encodes only first/step/length, not every element, so wire size is O(1) regardless of the
    # range's length - see encode_for_xl(::AbstractRange{Float64})/(::AbstractRange{<:Integer}) in
    # src/encode.jl. Julia -> Excel only: VBA arrays are always fully materialized, so there's no
    # lazy range concept on the VBA side to send this format back - unlike "V", decode_from_xl has
    # no 'R' case and doesn't need one, so these are encode-only checks (no round_trip).
    @test JuliaExcel.encode_for_xl(1:5) == "RI,5,1,1;"
    @test JuliaExcel.encode_for_xl(5:3:47) == "RI,15,5,3;"  # non-1 step
    @test JuliaExcel.encode_for_xl((1:3) .* pi) == "RF,3;" * JuliaExcel.float64_to_hex(Float64(pi)) * JuliaExcel.float64_to_hex(Float64(pi))

    # Wire size doesn't scale with length at all - the whole point of this format.
    @test length(JuliaExcel.encode_for_xl(1:1_000_000)) < 20
    @test length(JuliaExcel.encode_for_xl((1:1_000_000) .* pi)) < 60

    # Empty ranges and non-finite first/step (Float64 only - integer ranges can't hold NaN/Inf)
    # must fall back to the general "*" format, not "R".
    @test JuliaExcel.encode_for_xl(1:0) == JuliaExcel.encode_array_general(1:0)
    @test JuliaExcel.encode_for_xl(1.0:1.0:0.0) == JuliaExcel.encode_array_general(1.0:1.0:0.0)
    @test !startswith(JuliaExcel.encode_for_xl(1:0), "R")
    let weird = range(1.0, step=NaN, length=5)
        @test !startswith(JuliaExcel.encode_for_xl(weird), "R")
    end

    # Exactness: VBA reconstructs each element via plain "first + (i-1)*step" arithmetic, not by
    # decoding any per-element wire data - confirms that matches Julia's own range materialization
    # exactly, including StepRangeLen's twice-precision internal representation, over a large case.
    let r = (1:1_000_000) .* pi
        f, s = first(r), step(r)
        @test all(i -> (f + (i - 1) * s) === r[i], 1:length(r))
    end

    # Defensive guard: nothing compiler-enforces that a type subtyping AbstractRange actually
    # behaves as an arithmetic progression (Julia's abstract types carry no semantic guarantees,
    # only nominal ones - a third-party AbstractRange subtype could implement `step` to mean
    # anything). This checks that encode_for_xl verifies last(x) == first(x) + (n-1)*step(x) before
    # trusting the fast path, using a deliberately "lying" range as a regression test - a real
    # geometric progression whose step() claims (wrongly) that it's arithmetic with step 1.0.
    # Must fall back to the general format and still round-trip correctly, not go through "R" and
    # silently produce a linear approximation of non-arithmetic data.
    struct LyingRange <: AbstractRange{Float64}
        data::Vector{Float64}
    end
    Base.length(r::LyingRange) = length(r.data)
    Base.first(r::LyingRange) = r.data[1]
    Base.last(r::LyingRange) = r.data[end]
    Base.step(r::LyingRange) = 1.0
    Base.getindex(r::LyingRange, i::Int) = r.data[i]
    Base.size(r::LyingRange) = (length(r.data),)

    let lr = LyingRange([1.0, 2.0, 100.0])  # not actually evenly spaced by 1.0
        @test !startswith(JuliaExcel.encode_for_xl(lr), "R")
        @test JuliaExcel.decode_from_xl(JuliaExcel.encode_for_xl(lr)) == [1.0, 2.0, 100.0]
    end

    # friendly_error (comms.jl) - srv_eval_inner's error path for a JuliaEval expression that
    # itself errors (as opposed to returning a value that fails to encode - see
    # _encode_result_for_xl above). Checks the exact text so a regression (a stray backtrace frame,
    # a changed cap, a wording change) is caught, not just "it's some string starting with #".
    let msg = JuliaExcel.decode_from_xl(JuliaExcel.srv_eval_inner("1+\"1\""))
        @test msg == "#MethodError: no method matching +(::Int64, ::String) The function `+` " *
                     "exists, but no method is defined for this combination of argument types. " *
                     "Julia REPL has more details and stacktrace!"
    end

    # display_results (comms.jl) - toggles REPL echoing of calls from Excel; check both the
    # setter's return-value contract and the getter round-trips it correctly. Ends on false (the
    # default) so it doesn't leak into any test that follows.
    @test JuliaExcel.display_results(true) == "Results from JuliaCall/JuliaEval will display in REPL"
    @test JuliaExcel.display_results() == true
    @test JuliaExcel.display_results(false) == "Results from JuliaCall/JuliaEval will not display in REPL"
    @test JuliaExcel.display_results() == false

    # last_question/last_answer/answer_again (comms.jl) - srv_eval_inner records the expression
    # text itself as last_question...
    JuliaExcel.decode_from_xl(JuliaExcel.srv_eval_inner("1+1"))
    @test JuliaExcel.last_question == "1+1"
    @test JuliaExcel.last_answer == 2
    @test JuliaExcel.answer_again() == 2

    # ...while srv_call_inner reconstructs a "fn_name(args_from_xl...)" call expression, with a
    # ".(...)" broadcasting form when the function name arrives with a trailing ".".
    let payload = JuliaExcel.encode_for_xl(Any["identity", 5])
        @test JuliaExcel.decode_from_xl(JuliaExcel.srv_call_inner(payload)) == 5
        @test JuliaExcel.last_question == "identity(args_from_xl...)"
        @test JuliaExcel.last_answer == 5
        @test JuliaExcel.answer_again() == 5
    end
    let payload = JuliaExcel.encode_for_xl(Any["identity.", [1, 2, 3]])
        @test JuliaExcel.decode_from_xl(JuliaExcel.srv_call_inner(payload)) == [1, 2, 3]
        @test JuliaExcel.last_question == "identity.(args_from_xl...)"
        @test JuliaExcel.answer_again() == [1, 2, 3]
    end

end

# The VBA-side test suite (modTest.RunTests) needs a live Excel/VBA session (via COM automation),
# which GitHub Actions doesn't have - so this is skipped in CI (detected via the "CI" environment
# variable GitHub Actions sets, not the OS: the CI matrix in .github/workflows/Runtests.yml
# currently runs on windows-latest, so checking Sys.iswindows() alone wouldn't distinguish CI from
# a local run). Locally, run scripts\Run-VbaTests.ps1, which expects workbooks\JuliaExcel.xlam to
# already be open in Excel and fails clearly (rather than skipping silently) if it isn't.
if get(ENV, "CI", "false") != "true" && Sys.iswindows()
    @testset "VBA test suite (modTest.RunTests, requires Excel open locally)" begin
        script = joinpath(@__DIR__, "..", "scripts", "Run-VbaTests.ps1")
        proc = run(ignorestatus(`powershell -ExecutionPolicy Bypass -File $script`))
        @test proc.exitcode == 0
    end
end
