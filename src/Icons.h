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

#ifndef _ICONS_H_
#define _ICONS_H_

#include <windows.h>

//#pragma pack(1)

typedef struct
{
  BYTE        bWidth;          // Width, in pixels, of the image
  BYTE        bHeight;         // Height, in pixels, of the image
  BYTE        bColorCount;     // Number of colors in image (0 if >=8bpp)
  BYTE        bReserved;       // Reserved ( must be 0)
  WORD        wPlanes;         // Color Planes
  WORD        wBitCount;       // Bits per pixel
  DWORD       dwBytesInRes;    // How many bytes in this resource?
  DWORD       dwImageOffset;   // Where in the file is this image?
} ICONDIRENTRY, *LPICONDIRENTRY;

typedef struct
{
  WORD           idReserved;   // Reserved (must be 0)
  WORD           idType;       // Resource Type (1 for icons)
  WORD           idCount;      // How many images?
  ICONDIRENTRY   idEntries[1]; // An entry for each image (idCount of 'em)
} ICONDIR, *LPICONDIR;

typedef struct
{
  BITMAPINFOHEADER   icHeader;      // DIB header
  RGBQUAD         icColors[1];   // Color table
  BYTE            icXOR[1];      // DIB bits for XOR mask
  BYTE            icAND[1];      // DIB bits for AND mask
} ICONIMAGE, *LPICONIMAGE;

typedef struct
{
  HBITMAP hPreview;
  ICONDIRENTRY icdi;
  ICONIMAGE icoimg;
} ICONDATA, *LPICONDATA;

#endif
