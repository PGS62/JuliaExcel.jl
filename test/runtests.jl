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

end
