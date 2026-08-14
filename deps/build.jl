# Launches installer/Install.vbs to register the JuliaExcel.xlam Excel add-in, so that
# `Pkg.add("JuliaExcel")` alone is enough to install it - no separate manual call to
# JuliaExcel.installme() needed. Pkg runs this script automatically right after add/build.
#
# Doesn't `using JuliaExcel` here - the package may not be safely loadable during its own
# build, so the handful of lines from installme() (src/comms.jl) are duplicated rather than
# shared; installme() stays as-is for manual re-install/repair, where throwing on failure is
# the right behaviour, unlike here.
#
# Skipped entirely (does nothing) when:
#   - not running on Windows - the installer registers a Windows Excel add-in via COM/wscript.exe
#   - running under CI - the installer is interactive (message-box dialogs), which would hang a
#     non-interactive build waiting for clicks that will never come
if Sys.iswindows() && get(ENV, "CI", "false") != "true"
    try
        installscript = normpath(joinpath(@__DIR__, "..", "installer", "Install.vbs"))
        exefile = "C:/Windows/System32/wscript.exe"
        if isfile(exefile) && isfile(installscript)
            run(`$exefile $installscript`, wait=false)
            println("JuliaExcel: launched the JuliaExcel add-in installer - please respond to its dialogs.")
        else
            @warn "JuliaExcel: could not find wscript.exe or the install script; run `using JuliaExcel; JuliaExcel.installme()` manually instead."
        end
    catch e
        @warn "JuliaExcel: automatic installer failed to launch ($e); run `using JuliaExcel; JuliaExcel.installme()` manually instead."
    end
end
