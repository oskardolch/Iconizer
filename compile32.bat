if "%1" == "a" mingw32-make -fmakefile.mingw BUILDDIR=Build32 clean
mingw32-make -fmakefile.mingw PREFIX=i686-w64-mingw32- TOOLPFX=i686-w64-mingw32- BUILDDIR=Build32 SUFFIX=32 PNGUtils 2> PNGUtils.log
mingw32-make -fmakefile.mingw PREFIX=i686-w64-mingw32- TOOLPFX=i686-w64-mingw32- BUILDDIR=Build32 SUFFIX=32 Iconizer 2> Iconizer.log
