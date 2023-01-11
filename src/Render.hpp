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

#ifndef _RENDER_HPP_
#define _RENDER_HPP_

#include <stdio.h>
#include "Icons.h"

HBITMAP RenderIcon(HWND hWndStatic, LPICONIMAGE pIcoImg, int width, int height);
HBITMAP RenderPNG(HWND hWndStatic, LPICONIMAGE pIcoImg, int iSizeint, int *piWidth, int *piHeight);
HBITMAP RenderBitmap(HWND hWndStatic, BYTE *pData, DWORD dwTranspCol, int *piWidth, int *piHeight);
void SaveIconBitmap(LPICONDATA pIcoData, FILE *fp);
LPICONDATA PNGtoIconCopy(char *pData, int iDataSize);
LPICONDATA PNGtoIconData(char *pData, int iDataSize, int iBitDepth,
  int iWidth, int iHeight, BOOL bDither, BOOL bOptimize);
LPICONDATA BMPtoIconData(char *pData, int iDataSize, int iBitDepth,
  int iWidth, int iHeight, DWORD dwTranspColor, BOOL bDither, BOOL bOptimize);
void GetEditImage(LPICONIMAGE pIcoImg, LPBITMAPINFO pBitmap, COLORREF clTransp);
void SaveEditImage(LPICONIMAGE pIcoImg, LPBITMAPINFO pBitmap);

#endif
