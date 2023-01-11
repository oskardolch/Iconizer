/*
    Copyright (C) 2009  Pavel Krejcir

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

#ifndef _PNGUTILS_HPP_
#define _PNGUTILS_HPP_

#include <windows.h>

typedef CALLBACK long BackgroundFunction(int i, int j);

extern "C" BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved);
extern "C" long PNGImageCreate(char* buf, long bufsize);
extern "C" void PNGImageDestroy(long img);
extern "C" BOOL PNGImageIsPNG(long img);
extern "C" void PNGImageGetDimen(long img, long* width, long* height);
extern "C" int PNGImageGetBitDepth(long img);
extern "C" int PNGImageRender(long img, int vw, int vh, HDC ADC,
    int destw, int desth, BackgroundFunction backfnc);
extern "C" BOOL PNGImageGetBitmaps(long img, HDC ADC,
    HBITMAP hBmpColor, HBITMAP hBmpMask);
extern "C" BOOL IsPNGFile(char* buf, long bufsize);

#endif
