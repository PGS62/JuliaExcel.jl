using JuliaExcel
using Dates
using Test

@testset "JuliaExcel.jl" begin
   @test JuliaExcel.encode_for_xl(1) == "&1"
   @test JuliaExcel.encode_for_xl(1.0) == "#1.0"
   @test JuliaExcel.encode_for_xl(Int8(1)) == "S1" 
   @test JuliaExcel.encode_for_xl(Int16(1)) == "S1" 
   @test JuliaExcel.encode_for_xl(Int32(1)) == "&1"
   @test JuliaExcel.encode_for_xl(Int64(1)) == "&1"
   @test JuliaExcel.encode_for_xl(true) == "T"
   @test JuliaExcel.encode_for_xl(false) == "F"
   @test JuliaExcel.encode_for_xl("foo") == "£foo"
   @test JuliaExcel.encode_for_xl('x') == "£x"
   @test JuliaExcel.encode_for_xl(:x) == "£:x"
   @test JuliaExcel.encode_for_xl(nothing) == "E"
   @test JuliaExcel.encode_for_xl(missing) == "E"
   @test JuliaExcel.encode_for_xl(Inf) =="!2036"
   @test JuliaExcel.encode_for_xl(-Inf) =="!2036"
   @test JuliaExcel.encode_for_xl(NaN) =="!2042"
   @test JuliaExcel.encode_for_xl(Date("2021-11-8")) == "D44508"
   @test JuliaExcel.encode_for_xl(DateTime("2021-11-8T12:00:00")) == "D44508.5"
   @test JuliaExcel.encode_for_xl(Int64) == "£Int64"
   @test JuliaExcel.encode_for_xl(v"1.2.3") == "£1.2.3"
   @test JuliaExcel.encode_for_xl((1,2)) == "*1,2;2,2,;&1&2"
   @test JuliaExcel.encode_for_xl([1,2,3]) == "*1,3;2,2,2,;&1&2&3"
   @test JuliaExcel.encode_for_xl(Any[1,2,3.0]) == "*1,3;2,2,4,;&1&2#3.0"
   @test JuliaExcel.encode_for_xl([1,true,"x"]) == "*1,3;2,1,2,;&1T£x"
   @test JuliaExcel.encode_for_xl([1,[2,3]]) == "*1,2;2,14,;&1*1,2;2,2,;&2&3"
end
