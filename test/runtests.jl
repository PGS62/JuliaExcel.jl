using JuliaExcel
using Dates
using Test

round_trip(x) = isequal(JuliaExcel.decode_from_xl(JuliaExcel.encode_for_xl(x)), x)

@testset "JuliaExcel.jl" begin
    #For compatibility with VBA function Unserialise (which should be exact equivalent of decode_from_excel)
    @test JuliaExcel.encode_for_xl(1) == "^1"
    @test JuliaExcel.encode_for_xl(1.0) == "#3FF0000000000000"
    @test JuliaExcel.encode_for_xl(Int8(1)) == "%1"
    @test JuliaExcel.encode_for_xl(Int16(1)) == "%1"
    @test JuliaExcel.encode_for_xl(Int32(1)) == "&1"
    @test JuliaExcel.encode_for_xl(Int64(1)) == "^1"
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
    @test JuliaExcel.encode_for_xl(Date("2021-11-8")) == "D44508"
    @test JuliaExcel.encode_for_xl(DateTime("2021-11-8T12:00:00")) == "G40E5BB9000000000"
    @test JuliaExcel.encode_for_xl(Int64) == "£Int64"
    @test JuliaExcel.encode_for_xl(v"1.2.3") == "£1.2.3"
    @test JuliaExcel.encode_for_xl((1, 2)) == "*1,2;2,2,;^1^2"
    @test JuliaExcel.encode_for_xl([1, 2, 3]) == "*1,3;2,2,2,;^1^2^3"
    @test JuliaExcel.encode_for_xl(Any[1, 2, 3.0, π]) == "*1,4;2,2,17,17,;^1^2#4008000000000000#400921FB54442D18"
    @test JuliaExcel.encode_for_xl([1, true, "x"]) == "*1,3;2,1,2,;^1T£x"
    @test JuliaExcel.encode_for_xl([1, [2, 3]]) == "*1,2;2,14,;^1*1,2;2,2,;^2^3"
    @test JuliaExcel.encode_for_xl(Dict("a"=>1, "b"=>2)) == "H2;2,2,2,2,;£b^2£a^1"

    @test round_trip(1)
    @test round_trip(1.0)
    @test round_trip(Int8(1))
    @test round_trip(Int16(1))
    @test round_trip(Int32(1))
    @test round_trip(Int64(1))
    @test round_trip(true)
    @test round_trip(false)
    @test round_trip("foo")
  #  @test round_trip('x') # Characters morph to String
  #  @test round_trip(:x)  # Symbols morph to String
  #  @test round_trip(nothing) # Nothing morphs to missing (Empty in Excel)
     @test round_trip(missing)
  #  @test round_trip(Inf)
  #  @test round_trip(-Inf)
  #  @test round_trip(NaN)
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

    # "V" format - compact encoding for arrays of Float64 (see encode_for_xl(::Vector{Float64})/
    # (::Matrix{Float64}) and the dynamic AbstractArray fallback in src/encode.jl). No round_trip
    # tests here: decode_from_xl has no 'V' case (nothing currently sends "V" strings back to
    # Julia - VBA is the only consumer), so these check encode_for_xl's output directly, the same
    # way the plain type-indicator tests above do.
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
    # format) for a Vector{Any} containing a NaN, a non-Float64 element, or with rank > 2 (Case 86
    # in the VBA decoder, modUnserialise.bas, only supports rank 1 and 2).
    @test JuliaExcel.encode_for_xl(Any[1.0, 2.0, NaN]) == JuliaExcel.encode_for_xl([1.0, 2.0, NaN])
    @test !startswith(JuliaExcel.encode_for_xl(Any[1.0, 2.0, NaN]), "V")
    @test JuliaExcel.encode_for_xl(Any[1.0, 2, 3.0]) == JuliaExcel.encode_array_general(Any[1.0, 2, 3.0])
    @test !startswith(JuliaExcel.encode_for_xl(Any[1.0, 2, 3.0]), "V")
    @test JuliaExcel.encode_for_xl(Any[]) == JuliaExcel.encode_array_general(Any[])
    let x3d = fill(1.0, 2, 2, 2) |> a -> Array{Any}(a)
        @test !startswith(JuliaExcel.encode_for_xl(x3d), "V")
        @test JuliaExcel.encode_for_xl(x3d) == JuliaExcel.encode_array_general(x3d)
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
