if "%1" == "a" mingw32-make -fmakefile.mingw BUILDDIR=Build64 clean
mingw32-make -fmakefile.mingw PREFIX=x86_64-w64-mingw32- BUILDDIR=Build64 SUFFIX=64 PNGUtils 2> PNGUtils.log
mingw32-make -fmakefile.mingw PREFIX=x86_64-w64-mingw32- BUILDDIR=Build64 SUFFIX=64 Iconizer 2> Iconizer.log
