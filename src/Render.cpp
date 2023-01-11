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

#include "Render.hpp"
#include <math.h>
#include "../PNGUtils/PNGUtils.hpp"
#include "../Common/Quantize.h"

HBITMAP RenderIcon(HWND hWndStatic, LPICONIMAGE pIcoImg, int width, int height)
{
  int clmask, clback;
  int ri, gi, bi;

  int mnwidth = width/16;
  if(mnwidth % 2) mnwidth++;
  mnwidth *= 16;

  int bitoffs = 0;
  int clidx;
  int srcbitcnt = pIcoImg->icHeader.biBitCount;

  if(srcbitcnt < 16)
  {
    if(pIcoImg->icHeader.biClrUsed > 0)
    bitoffs = pIcoImg->icHeader.biClrUsed;
    else
    {
      bitoffs = 1;
      for(int i = 0; i < srcbitcnt; i++) bitoffs *= 2;
    }
  }

  int imgrowsize = srcbitcnt*width/8;
  if(imgrowsize % 2) imgrowsize++;
  int maskrowsize = mnwidth/8;
  if(maskrowsize % 2) maskrowsize++;
  int resrowsize = 4*width;

  BYTE *pbimg = (BYTE*)&pIcoImg->icColors[bitoffs];
  BYTE *pbmask = NULL;
  if(pIcoImg->icHeader.biBitCount < 32) pbmask = &pbimg[imgrowsize*height];
  BYTE *pbimgrow = pbimg;
  BYTE *pbmaskrow = pbmask;

  BITMAPINFO bif;
  bif.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
  bif.bmiHeader.biWidth = width;
  bif.bmiHeader.biHeight = height;
  bif.bmiHeader.biPlanes = 1;
  bif.bmiHeader.biBitCount = 32;
  bif.bmiHeader.biCompression = BI_RGB;
  bif.bmiHeader.biSizeImage = resrowsize*height;
  bif.bmiHeader.biXPelsPerMeter = 0;
  bif.bmiHeader.biYPelsPerMeter = 0;
  bif.bmiHeader.biClrUsed = 0;
  bif.bmiHeader.biClrImportant = 0;

  BYTE *pbres = (BYTE*)malloc(bif.bmiHeader.biSizeImage*sizeof(BYTE));
  BYTE *pbresrow = pbres;
  WORD wcol;

  for(int j = 0; j < height; j++)
  {
    for(int i = 0; i < width; i++)
    {
      if((i / 4 + (height - j - 1) / 4) % 2) clback = 224;
      else clback = 192;
      if(pbmaskrow)
      {
        clmask = (pbmaskrow[i/8] >> (7 - (i % 8))) & 0x01;
        if(srcbitcnt < 16)
        {
          clidx = (pbimgrow[srcbitcnt*i/8] >>
            (8 - srcbitcnt - srcbitcnt*(i % (8/srcbitcnt)))) &
            (0xFF >> (8 - srcbitcnt));
          ri = pIcoImg->icColors[clidx].rgbRed;
          gi = pIcoImg->icColors[clidx].rgbGreen;
          bi = pIcoImg->icColors[clidx].rgbBlue;
        }
        else if(srcbitcnt == 16)
        {
          wcol = (pbimgrow[2*i] << 8) | pbimgrow[2*i + 1];
          bi = 8*((wcol >> 11) & 0x001F);
          bi = 8*((wcol >> 6) & 0x001F);
          bi = 8*((wcol >> 1) & 0x001F);
        }
        else if(srcbitcnt == 24)
        {
          bi = pbimgrow[3*i] & 0xFF;
          gi = pbimgrow[3*i + 1] & 0xFF;
          ri = pbimgrow[3*i + 2] & 0xFF;
        }
        pbresrow[4*i] = clmask*clback + (1 - clmask)*bi;
        pbresrow[4*i + 1] = clmask*clback + (1 - clmask)*gi;
        pbresrow[4*i + 2] = clmask*clback + (1 - clmask)*ri;
        pbresrow[4*i + 3] = 0;
      }
      else
      {
        bi = pbimgrow[4*i] & 0xFF;
        gi = pbimgrow[4*i + 1] & 0xFF;
        ri = pbimgrow[4*i + 2] & 0xFF;
        clmask = pbimgrow[4*i + 3] & 0xFF;
        pbresrow[4*i] = (clmask*bi + (255 - clmask)*clback)/255;
        pbresrow[4*i + 1] = (clmask*gi + (255 - clmask)*clback)/255;
        pbresrow[4*i + 2] = (clmask*ri + (255 - clmask)*clback)/255;
        pbresrow[4*i + 3] = 0;
      }
    }
    pbimgrow += imgrowsize;
    if(pbmaskrow) pbmaskrow += maskrowsize;
    pbresrow += resrowsize;
  }

  HDC hdc = GetDC(hWndStatic);
  HBITMAP hbmpRes = CreateDIBitmap(hdc, &bif.bmiHeader, CBM_INIT, pbres, &bif, DIB_RGB_COLORS);
  ReleaseDC(hWndStatic, hdc);

  free(pbres);

  SendMessage(hWndStatic, STM_SETIMAGE, IMAGE_BITMAP, (LPARAM)hbmpRes);
  return hbmpRes;
}

CALLBACK long BgrFunc(int i, int j)
{
  int clback = 192;
  if((i / 4 + j / 4) % 2) clback = 224;
  return RGB(clback, clback, clback);
}

HBITMAP RenderPNG(HWND hWndStatic, LPICONIMAGE pIcoImg, int iSize, int *piWidth, int *piHeight)
{
/*wchar_t buf[64];
swprintf(buf, L"%d", iSize);
MessageBox(0, buf, L"Debug", MB_OK);*/
  long himg = PNGImageCreate((char*)pIcoImg, iSize);
  if(!PNGImageIsPNG(himg))
  {
    PNGImageDestroy(himg);
    return NULL;
  }

  long srcwidth, srcheight, dstwidth, dstheight;
  PNGImageGetDimen(himg, &srcwidth, &srcheight);
  dstwidth = srcwidth;
  dstheight = srcheight;
  if(dstwidth > 256) dstwidth = 256;
  if(dstheight > 256) dstheight = 256;

  HDC hdc = GetDC(hWndStatic);
  HDC dc = CreateCompatibleDC(hdc);
  HBITMAP hbmpRes = CreateCompatibleBitmap(hdc, dstwidth, dstheight);
  SelectObject(dc, hbmpRes);
  PNGImageRender(himg, srcwidth, srcheight, dc, dstwidth, dstheight, BgrFunc);
  DeleteDC(dc);
  ReleaseDC(hWndStatic, hdc);

  PNGImageDestroy(himg);

  if(piWidth) *piWidth = dstwidth;
  if(piHeight) *piHeight = dstheight;

  SendMessage(hWndStatic, STM_SETIMAGE, IMAGE_BITMAP, (LPARAM)hbmpRes);
  return hbmpRes;
}

HBITMAP RenderBitmap(HWND hWndStatic, BYTE *pData, DWORD dwTranspCol, int *piWidth, int *piHeight)
{
  HBITMAP hbmpRes = NULL;
  BITMAPFILEHEADER *pbfh = (BITMAPFILEHEADER*)pData;
  BITMAPCOREHEADER *pbci = (BITMAPCOREHEADER*)&pData[sizeof(BITMAPFILEHEADER)];

  HDC hdc = GetDC(hWndStatic);
  if(pbci->bcSize == sizeof(BITMAPCOREHEADER))
  {
    //MessageBox(0, L"Core", L"Debug", MB_OK);
  }
  else if(pbci->bcSize == sizeof(BITMAPINFOHEADER))
  {
    BITMAPINFO *pbi = (BITMAPINFO*)pbci;
    BYTE *pbits = (BYTE*)&pData[pbfh->bfOffBits];

    int srcw = pbi->bmiHeader.biWidth;
    int srch = pbi->bmiHeader.biHeight;

    // if some application saves 32bit bmp with alpha channel, we can
    // uncomment this block:
    /*if(pbi->bmiHeader.biBitCount == 32)
    {
      BYTE br, bg, bb, ba, bbg;
      for(int j = 0; j < srch; j++)
      {
        for(int i = 0; i < srcw; i++)
        {
          br = pbits[4*srcw*j + 4*i];
          bg = pbits[4*srcw*j + 4*i + 1];
          bb = pbits[4*srcw*j + 4*i + 2];
          ba = pbits[4*srcw*j + 4*i + 3];
          bbg = 192;
          if((i / 4 + j / 4) % 2) bbg = 224;
          pbits[4*srcw*j + 4*i] = ((255 - ba)*br + ba*bbg)/256;
          pbits[4*srcw*j + 4*i + 1] = ((255 - bg)*br + ba*bbg)/256;
          pbits[4*srcw*j + 4*i + 2] = ((255 - bb)*br + ba*bbg)/256;
        }
      }
    }*/
    int dstw = srcw;
    int dsth = srch;
    if(dstw > 256) dstw = 256;
    if(dsth > 256) dsth = 256;
    if(srcw == dstw && srch == dsth)
      hbmpRes = CreateDIBitmap(hdc, &pbi->bmiHeader, CBM_INIT, pbits, pbi, DIB_RGB_COLORS);
    else
    {
      HDC lhdc = CreateCompatibleDC(hdc);
      hbmpRes = CreateCompatibleBitmap(hdc, dstw, dsth);
      SelectObject(lhdc, hbmpRes);
      SetStretchBltMode(lhdc, HALFTONE);
      StretchDIBits(lhdc, 0, 0, dstw, dsth, 0, 0, srcw, srch, pbits,
      pbi, DIB_RGB_COLORS, SRCCOPY);
      DeleteDC(lhdc);
    }

    if(dwTranspCol != 0xFF000000)
    {
      int imgrowsize = 4*dstw;
      BYTE *plocalbits = (BYTE*)malloc(imgrowsize*dsth*sizeof(BYTE));

      BITMAPINFO bif;
      bif.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
      bif.bmiHeader.biWidth = dstw;
      bif.bmiHeader.biHeight = dsth;
      bif.bmiHeader.biPlanes = 1;
      bif.bmiHeader.biBitCount = 32;
      bif.bmiHeader.biCompression = BI_RGB;
      bif.bmiHeader.biSizeImage = imgrowsize*dsth;
      bif.bmiHeader.biXPelsPerMeter = 0;
      bif.bmiHeader.biYPelsPerMeter = 0;
      bif.bmiHeader.biClrUsed = 0;
      bif.bmiHeader.biClrImportant = 0;

      GetDIBits(hdc, hbmpRes, 0, dsth, plocalbits, &bif, DIB_RGB_COLORS);
      DWORD col;
      BYTE *plrow = plocalbits;
      int clback;
      for(int j = 0; j < dsth; j++)
      {
        for(int i = 0; i < dstw; i++)
        {
          col = plrow[4*i + 2] | (plrow[4*i + 1] << 8) | (plrow[4*i] << 16);
          if(col == dwTranspCol)
          {
            if((i / 4 + j / 4) % 2) clback = 224;
            else clback = 192;
            plrow[4*i] = clback;
            plrow[4*i + 1] = clback;
            plrow[4*i + 2] = clback;
          }
        }
        plrow += imgrowsize;
      }
      DeleteObject(hbmpRes);
      hbmpRes = CreateDIBitmap(hdc, &bif.bmiHeader, CBM_INIT,
      plocalbits, &bif, DIB_RGB_COLORS);
      free(plocalbits);
    }

    if(piWidth) *piWidth = dstw;
    if(piHeight) *piHeight = dsth;
  }
  ReleaseDC(hWndStatic, hdc);

  SendMessage(hWndStatic, STM_SETIMAGE, IMAGE_BITMAP, (LPARAM)hbmpRes);
  return hbmpRes;
}

void SaveIconBitmap(LPICONDATA pIcoData, FILE *fp)
{
  BITMAPINFOHEADER bih;
  int bifsize = sizeof(BITMAPINFO);
  memcpy(&bih, &pIcoData->icoimg.icHeader, sizeof(BITMAPINFOHEADER));
  bih.biHeight = pIcoData->icoimg.icHeader.biHeight / 2;

  int bitoffs = 0;
  int srcbitcnt = pIcoData->icoimg.icHeader.biBitCount;

  if(srcbitcnt < 16)
  {
    if(pIcoData->icoimg.icHeader.biClrUsed > 0)
      bitoffs = pIcoData->icoimg.icHeader.biClrUsed;
    else
    {
      bitoffs = 1;
      for(int i = 0; i < srcbitcnt; i++) bitoffs *= 2;
    }
  }

  int imgrowsize = srcbitcnt*bih.biWidth/8;
  if(imgrowsize % 2) imgrowsize++;

  BITMAPFILEHEADER bfh;
  bfh.bfType = 0x4D42;
  bfh.bfSize = sizeof(BITMAPINFOHEADER) + bitoffs*sizeof(RGBQUAD) +
    imgrowsize*bih.biHeight;
  bfh.bfReserved1 = 0;
  bfh.bfReserved2 = 0;
  bfh.bfOffBits = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER) +
    bitoffs*sizeof(RGBQUAD);

  fwrite(&bfh, 1, sizeof(BITMAPFILEHEADER), fp);
  fwrite(&bih, 1, sizeof(BITMAPINFOHEADER), fp);
  fwrite(pIcoData->icoimg.icColors, 1,
    bitoffs*sizeof(RGBQUAD) + imgrowsize*bih.biHeight, fp);
  return;
}

LPICONDATA PNGtoIconCopy(char *pData, int iDataSize)
{
  LPICONDATA icdRes = NULL;
  long lpng = PNGImageCreate(pData, iDataSize);
  if(PNGImageIsPNG(lpng))
  {
    long width, height;
    PNGImageGetDimen(lpng, &width, &height);
    icdRes = (LPICONDATA)malloc(sizeof(HBITMAP) + sizeof(ICONDIRENTRY) + iDataSize);
    icdRes->hPreview = NULL;
    icdRes->icdi.bWidth = width;
    icdRes->icdi.bHeight = height;
    icdRes->icdi.bColorCount = 0;
    icdRes->icdi.bReserved = 0;
    icdRes->icdi.wPlanes = 1;
    icdRes->icdi.wBitCount = PNGImageGetBitDepth(lpng);
    icdRes->icdi.dwBytesInRes = iDataSize;
    icdRes->icdi.dwImageOffset = 0;

    memcpy(&icdRes->icoimg, pData, iDataSize);
  }
  PNGImageDestroy(lpng);
  return icdRes;
}

int DistFunc(int i1, int i2, int i3)
{
  return(abs(i1) + abs(i2) + abs(i3));
  //return(i1*i1 + i2*i2 + i3*i3);
}

int GetPalIndex(RGBQUAD *pPal, int iPalSize, BYTE r, BYTE g, BYTE b)
{
  int mindist = DistFunc(255, 255, 255);
  int minidx = 0;
  int curdist;
  for(int i = 0; i < iPalSize; i++)
  {
    curdist = DistFunc(r - pPal[i].rgbRed, g - pPal[i].rgbGreen, b - pPal[i].rgbBlue);
    if(curdist < mindist)
    {
      mindist = curdist;
      minidx = i;
    }
  }
  return minidx;
}

int Round(double d)
{
  int id = floor(d);
  if(d - id > 0.5) id = ceil(d);
  return id;
}

BYTE AddByte(BYTE b, int i)
{
  if(i + b < 0) return 0;
  if(i + b > 255) return 255;
  return i + b;
}

void GetAlphaImage(HDC hdc, HBITMAP hbmpColor, HBITMAP hbmpMask, int iWidth,
  int iHeight, BYTE *pimg)
{
  BITMAPINFO bi;
  bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
  bi.bmiHeader.biWidth = iWidth;
  bi.bmiHeader.biHeight = iHeight;
  bi.bmiHeader.biPlanes = 1;
  bi.bmiHeader.biBitCount = 32;
  bi.bmiHeader.biCompression = BI_RGB;
  bi.bmiHeader.biSizeImage = 4*iWidth*iHeight;
  bi.bmiHeader.biXPelsPerMeter = 0;
  bi.bmiHeader.biYPelsPerMeter = 0;
  bi.bmiHeader.biClrUsed = 0;
  bi.bmiHeader.biClrImportant = 0;

  GetDIBits(hdc, hbmpColor, 0, iHeight, pimg, &bi, DIB_RGB_COLORS);
  BYTE *pmask = (BYTE*)malloc(4*iWidth*iHeight);
  GetDIBits(hdc, hbmpMask, 0, iHeight, pmask, &bi, DIB_RGB_COLORS);

  BYTE *pbimgrow = pimg;
  BYTE *pbmaskrow = pmask;
  int rowsize = 4*iWidth;
  for(int j = 0; j < iHeight; j++)
  {
    for(int i = 0; i < iWidth; i++)
    {
      pbimgrow[4*i + 3] = 255 - pbmaskrow[4*i];
      if(pbimgrow[4*i + 3] < 5)
      {
        pbimgrow[4*i] = 0;
        pbimgrow[4*i + 1] = 0;
        pbimgrow[4*i + 2] = 0;
      }
    }
    pbimgrow += rowsize;
    pbmaskrow += rowsize;
  }

  free(pmask);
  return;
}

int GetStdColors(RGBQUAD *cols)
{
  DWORD clrefs[16] = {0x00000000, 0x00000080, 0x00008000, 0x00008080,
    0x00800000, 0x00800080, 0x00808000, 0x00808080, 0x00C0C0C0, 0x000000FF,
    0x0000FF00, 0x0000FFFF, 0x00FF0000, 0x00FF00FF, 0x00FFFF00, 0x00FFFFFF};
  for(int i = 0; i < 16; i++)
  {
    cols[i].rgbRed = clrefs[i] & 0xFF;
    cols[i].rgbGreen = (clrefs[i] >> 8) & 0xFF;
    cols[i].rgbBlue = (clrefs[i] >> 16) & 0xFF;
    cols[i].rgbReserved = 0;
  }
  return(16);
}

int GetStdColors256(RGBQUAD *cols)
{
  DWORD clrefs[256] = {0x00000000, 0x00000080, 0x00008000, 0x00008080,
    0x00800000, 0x00800080, 0x00808000, 0x00C0C0C0, 0x00C0DCC0, 0x00F0CAA6,
    0x00002040, 0x00002060, 0x00002080, 0x000020A0, 0x000020C0, 0x000020E0,
    0x00004000, 0x00004020, 0x00004040, 0x00004060, 0x00004080, 0x000040A0,
    0x000040C0, 0x000040E0, 0x00006000, 0x00006020, 0x00006040, 0x00006060,
    0x00006080, 0x000060A0, 0x000060C0, 0x000060E0, 0x00008000, 0x00008020,
    0x00008040, 0x00008060, 0x00008080, 0x000080A0, 0x000080C0, 0x000080E0,
    0x0000A000, 0x0000A020, 0x0000A040, 0x0000A060, 0x0000A080, 0x0000A0A0,
    0x0000A0C0, 0x0000A0E0, 0x0000C000, 0x0000C020, 0x0000C040, 0x0000C060,
    0x0000C080, 0x0000C0A0, 0x0000C0C0, 0x0000C0E0, 0x0000E000, 0x0000E020,
    0x0000E040, 0x0000E060, 0x0000E080, 0x0000E0A0, 0x0000E0C0, 0x0000E0E0,
    0x00400000, 0x00400020, 0x00400040, 0x00400060, 0x00400080, 0x004000A0,
    0x004000C0, 0x004000E0, 0x00402000, 0x00402020, 0x00402040, 0x00402060,
    0x00402080, 0x004020A0, 0x004020C0, 0x004020E0, 0x00404000, 0x00404020,
    0x00404040, 0x00404060, 0x00404080, 0x004040A0, 0x004040C0, 0x004040E0,
    0x00406000, 0x00406020, 0x00406040, 0x00406060, 0x00406080, 0x004060A0,
    0x004060C0, 0x004060E0, 0x00408000, 0x00408020, 0x00408040, 0x00408060,
    0x00408080, 0x004080A0, 0x004080C0, 0x004080E0, 0x0040A000, 0x0040A020,
    0x0040A040, 0x0040A060, 0x0040A080, 0x0040A0A0, 0x0040A0C0, 0x0040A0E0,
    0x0040C000, 0x0040C020, 0x0040C040, 0x0040C060, 0x0040C080, 0x0040C0A0,
    0x0040C0C0, 0x0040C0E0, 0x0040E000, 0x0040E020, 0x0040E040, 0x0040E060,
    0x0040E080, 0x0040E0A0, 0x0040E0C0, 0x0040E0E0, 0x00800000, 0x00800020,
    0x00800040, 0x00800060, 0x00800080, 0x008000A0, 0x008000C0, 0x008000E0,
    0x00802000, 0x00802020, 0x00802040, 0x00802060, 0x00802080, 0x008020A0,
    0x008020C0, 0x008020E0, 0x00804000, 0x00804020, 0x00804040, 0x00804060,
    0x00804080, 0x008040A0, 0x008040C0, 0x008040E0, 0x00806000, 0x00806020,
    0x00806040, 0x00806060, 0x00806080, 0x008060A0, 0x008060C0, 0x008060E0,
    0x00808000, 0x00808020, 0x00808040, 0x00808060, 0x00808080, 0x008080A0,
    0x008080C0, 0x008080E0, 0x0080A000, 0x0080A020, 0x0080A040, 0x0080A060,
    0x0080A080, 0x0080A0A0, 0x0080A0C0, 0x0080A0E0, 0x0080C000, 0x0080C020,
    0x0080C040, 0x0080C060, 0x0080C080, 0x0080C0A0, 0x0080C0C0, 0x0080C0E0,
    0x0080E000, 0x0080E020, 0x0080E040, 0x0080E060, 0x0080E080, 0x0080E0A0,
    0x0080E0C0, 0x0080E0E0, 0x00C00000, 0x00C00020, 0x00C00040, 0x00C00060,
    0x00C00080, 0x00C000A0, 0x00C000C0, 0x00C000E0, 0x00C02000, 0x00C02020,
    0x00C02040, 0x00C02060, 0x00C02080, 0x00C020A0, 0x00C020C0, 0x00C020E0,
    0x00C04000, 0x00C04020, 0x00C04040, 0x00C04060, 0x00C04080, 0x00C040A0,
    0x00C040C0, 0x00C040E0, 0x00C06000, 0x00C06020, 0x00C06040, 0x00C06060,
    0x00C06080, 0x00C060A0, 0x00C060C0, 0x00C060E0, 0x00C08000, 0x00C08020,
    0x00C08040, 0x00C08060, 0x00C08080, 0x00C080A0, 0x00C080C0, 0x00C080E0,
    0x00C0A000, 0x00C0A020, 0x00C0A040, 0x00C0A060, 0x00C0A080, 0x00C0A0A0,
    0x00C0A0C0, 0x00C0A0E0, 0x00C0C000, 0x00C0C020, 0x00C0C040, 0x00C0C060,
    0x00C0C080, 0x00C0C0A0, 0x00F0FBFF, 0x00A4A0A0, 0x00808080, 0x000000FF,
    0x0000FF00, 0x0000FFFF, 0x00FF0000, 0x00FF00FF, 0x00FFFF00, 0x00FFFFFF};
  for(int i = 0; i < 256; i++)
  {
    cols[i].rgbRed = clrefs[i] & 0xFF;
    cols[i].rgbGreen = (clrefs[i] >> 8) & 0xFF;
    cols[i].rgbBlue = (clrefs[i] >> 16) & 0xFF;
    cols[i].rgbReserved = 0;
  }
  return 256;
}

void SimplifyAlphaImg(BITMAPINFO *pbi, BYTE *pAlphaImg, int iWidth,
  int iHeight, int iImgRowSize, int iMaskRowSize, BOOL bDither, BOOL bOptimal)
{
  int iBitDepth = pbi->bmiHeader.biBitCount;
  int iPalSize = pbi->bmiHeader.biClrUsed;

  BYTE *imgbufrow = pAlphaImg;
  BYTE *imgbufnextrow;
  BYTE *pbimgrow = (BYTE*)&pbi->bmiColors[iPalSize];
  BYTE *pbmaskrow = &pbimgrow[iImgRowSize*iHeight];

  if(iBitDepth < 8)
  {
    pbi->bmiHeader.biClrImportant = GetStdColors(pbi->bmiColors);
  }
  else
  {
    if(bOptimal)
    {
      CQuantizer *quant = new CQuantizer(iPalSize, 6);
      quant->ProcessAlphaImg(pAlphaImg, iWidth, iHeight);
      pbi->bmiHeader.biClrImportant = quant->GetColorCount();
      quant->GetColorTable(pbi->bmiColors);
      delete quant;
    }
    else pbi->bmiHeader.biClrImportant = GetStdColors256(pbi->bmiColors);
  }

  DWORD col;
  int palidx;
  double dr, dg, db;
  for(int j = 0; j < iHeight; j++)
  {
    if(j < iHeight - 1) imgbufnextrow = imgbufrow + 4*iWidth;
    else imgbufnextrow = NULL;

    memset(pbimgrow, 0, iImgRowSize);
    memset(pbmaskrow, 0, iMaskRowSize);
    for(int i = 0; i < iWidth; i++)
    {
      if(imgbufrow[4*i + 3] < 128) palidx = 0;
      else palidx = GetPalIndex(pbi->bmiColors,
        pbi->bmiHeader.biClrImportant, imgbufrow[4*i + 2],
        imgbufrow[4*i + 1], imgbufrow[4*i]);
      switch(iBitDepth)
      {
      case 4:
        pbimgrow[i/2] |= palidx << 4*(1 - i % 2);
        break;
      case 8:
        pbimgrow[i] = palidx;
        break;
      }
      if(imgbufrow[4*i + 3] < 128)
      {
        pbmaskrow[i/8] |= 1 << 7 - i % 8;
      }
      // Floyd-Steinberg ditherring
      else if(bDither)
      {
        dr = (imgbufrow[4*i + 2] - pbi->bmiColors[palidx].rgbRed)/16;
        dg = (imgbufrow[4*i + 1] - pbi->bmiColors[palidx].rgbGreen)/16;
        db = (imgbufrow[4*i] - pbi->bmiColors[palidx].rgbBlue)/16;
        if(i < iWidth - 1)
        {
          imgbufrow[4*(i + 1)] = AddByte(imgbufrow[4*(i + 1)], Round(7*db));
          imgbufrow[4*(i + 1) + 1] = AddByte(imgbufrow[4*(i + 1) + 1], Round(7*dg));
          imgbufrow[4*(i + 1) + 2] = AddByte(imgbufrow[4*(i + 1) + 2], Round(7*dr));
        }
        if(imgbufnextrow)
        {
          if(i > 0)
          {
            imgbufnextrow[4*(i - 1)] = AddByte(imgbufnextrow[4*(i - 1)], Round(3*db));
            imgbufnextrow[4*(i - 1) + 1] = AddByte(imgbufnextrow[4*(i - 1) + 1], Round(3*dg));
            imgbufnextrow[4*(i - 1) + 2] = AddByte(imgbufnextrow[4*(i - 1) + 2], Round(3*dr));
          }
          imgbufnextrow[4*i] = AddByte(imgbufnextrow[4*i], Round(5*db));
          imgbufnextrow[4*i + 1] = AddByte(imgbufnextrow[4*i + 1], Round(5*dg));
          imgbufnextrow[4*i + 2] = AddByte(imgbufnextrow[4*i + 2], Round(5*dr));
          if(i < iWidth - 1)
          {
            imgbufnextrow[4*(i + 1)] = AddByte(imgbufnextrow[4*(i + 1)], Round(db));
            imgbufnextrow[4*(i + 1) + 1] = AddByte(imgbufnextrow[4*(i + 1) + 1], Round(dg));
            imgbufnextrow[4*(i + 1) + 2] = AddByte(imgbufnextrow[4*(i + 1) + 2], Round(dr));
          }
        }
      }
      // end of FSD
    }
    imgbufrow += 4*iWidth;
    pbimgrow += iImgRowSize;
    pbmaskrow += iMaskRowSize;
  }

  return;
}

LPICONDATA PNGtoIconData(char *pData, int iDataSize, int iBitDepth,
  int iWidth, int iHeight, BOOL bDither, BOOL bOptimize)
{
  LPICONDATA icdRes = NULL;

  long lpng = PNGImageCreate(pData, iDataSize);
  if(PNGImageIsPNG(lpng))
  {
    HDC dc = GetDC(0);
    HDC srcdc = CreateCompatibleDC(dc);
    HDC dstdc = CreateCompatibleDC(dc);

    long srcw = 0;
    long srch = 0;
    PNGImageGetDimen(lpng, &srcw, &srch);

    HBITMAP bmpSrcColor = CreateCompatibleBitmap(dc, srcw, srch);
    HBITMAP bmpSrcMask = CreateCompatibleBitmap(dc, srcw, srch);

    PNGImageGetBitmaps(lpng, srcdc, bmpSrcColor, bmpSrcMask);

    HBITMAP bmpDstColor = CreateCompatibleBitmap(dc, iWidth, iHeight);
    HBITMAP bmpDstMask = CreateCompatibleBitmap(dc, iWidth, iHeight);

    SetStretchBltMode(dstdc, HALFTONE);

    SelectObject(srcdc, bmpSrcColor);
    SelectObject(dstdc, bmpDstColor);
    StretchBlt(dstdc, 0, 0, iWidth, iHeight, srcdc, 0, 0, srcw, srch, SRCCOPY);

    SelectObject(srcdc, bmpSrcMask);
    SelectObject(dstdc, bmpDstMask);
    StretchBlt(dstdc, 0, 0, iWidth, iHeight, srcdc, 0, 0, srcw, srch, SRCCOPY);
    DeleteObject(bmpSrcColor);
    DeleteObject(bmpSrcMask);
    DeleteDC(srcdc);

    BYTE *pAlphaImg = (BYTE*)malloc(4*iWidth*iHeight);
    GetAlphaImage(dstdc, bmpDstColor, bmpDstMask, iWidth, iHeight, pAlphaImg);

    int mnwidth = iWidth/16;
    if(mnwidth % 2) mnwidth++;
    mnwidth *= 16;

    int bitoffs = 0;
    if(iBitDepth < 16)
    {
      bitoffs = 1;
      for(int i = 0; i < iBitDepth; i++) bitoffs *= 2;
    }

    int imgrowsize = iBitDepth*iWidth/8;
    if(imgrowsize % 2) imgrowsize++;
    int maskrowsize = mnwidth/8;
    if(maskrowsize % 2) maskrowsize++;

    int iressize = sizeof(BITMAPINFOHEADER) + bitoffs*sizeof(RGBQUAD) +
      (imgrowsize + maskrowsize)*iHeight;
    icdRes = (LPICONDATA)malloc(sizeof(HBITMAP) + sizeof(ICONDIRENTRY) + iressize);
    BITMAPINFO *pbi = (BITMAPINFO*)&icdRes->icoimg;

    pbi->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    pbi->bmiHeader.biWidth = iWidth;
    pbi->bmiHeader.biHeight = iHeight;
    pbi->bmiHeader.biPlanes = 1;
    pbi->bmiHeader.biBitCount = iBitDepth;
    pbi->bmiHeader.biCompression = BI_RGB;
    pbi->bmiHeader.biSizeImage = imgrowsize*iHeight;
    pbi->bmiHeader.biXPelsPerMeter = 0;
    pbi->bmiHeader.biYPelsPerMeter = 0;
    pbi->bmiHeader.biClrUsed = bitoffs;
    pbi->bmiHeader.biClrImportant = 0;

    if(bitoffs > 0)
    {
      /*if(iBitDepth == 8 && !bOptimize)
      {
        BYTE *pbimg = (BYTE*)&pbi->bmiColors[bitoffs];
        pbimg += imgrowsize*iHeight;

        pbi->bmiHeader.biBitCount = 1;
        pbi->bmiHeader.biSizeImage = maskrowsize*iHeight;
        pbi->bmiHeader.biClrUsed = 2;
        GetDIBits(dc, bmpDstMask, 0, iHeight, pbimg, pbi, DIB_RGB_COLORS);

        pbi->bmiHeader.biBitCount = iBitDepth;
        pbi->bmiHeader.biSizeImage = imgrowsize*iHeight;
        pbi->bmiHeader.biClrUsed = bitoffs;

        pbimg = (BYTE*)&pbi->bmiColors[bitoffs];
        GetDIBits(dc, bmpDstColor, 0, iHeight, pbimg, pbi, DIB_RGB_COLORS);
      }
      else*/ SimplifyAlphaImg(pbi, pAlphaImg, iWidth, iHeight, imgrowsize,
        maskrowsize, bDither, bOptimize);
    }
    else
    {
      BYTE *pbimg = (BYTE*)&pbi->bmiColors[bitoffs];
      pbimg += imgrowsize*iHeight;

      pbi->bmiHeader.biBitCount = 1;
      pbi->bmiHeader.biSizeImage = maskrowsize*iHeight;
      pbi->bmiHeader.biClrUsed = 2;
      GetDIBits(dc, bmpDstMask, 0, iHeight, pbimg, pbi, DIB_RGB_COLORS);

      pbi->bmiHeader.biBitCount = iBitDepth;
      pbi->bmiHeader.biSizeImage = imgrowsize*iHeight;
      pbi->bmiHeader.biClrUsed = bitoffs;

      pbimg = (BYTE*)&pbi->bmiColors[bitoffs];
      memcpy(pbimg, pAlphaImg, 4*iWidth*iHeight);
    }
    free(pAlphaImg);

    DeleteObject(bmpDstColor);
    DeleteObject(bmpDstMask);
    DeleteDC(dstdc);
    ReleaseDC(0, dc);

    pbi->bmiHeader.biHeight = 2*iHeight;
    //pbi->bmiHeader.biClrUsed = 0;

    icdRes->hPreview = NULL;
    icdRes->icdi.bWidth = iWidth;
    icdRes->icdi.bHeight = iHeight;
    if(iBitDepth < 8) icdRes->icdi.bColorCount = bitoffs;
    else icdRes->icdi.bColorCount = 0;
    icdRes->icdi.bReserved = 0;
    icdRes->icdi.wPlanes = 1;
    icdRes->icdi.wBitCount = iBitDepth;
    icdRes->icdi.dwBytesInRes = iressize;
    icdRes->icdi.dwImageOffset = 0;
  }
  PNGImageDestroy(lpng);

  return icdRes;
}

HBITMAP GetMaskFromBitmap(HDC hdc, char *pData, int iDataSize, int iBitDepth,
  int iWidth, int iHeight, DWORD dwTranspColor)
{
  HBITMAP hbmpRes = NULL;

  BITMAPFILEHEADER *pbfh = (BITMAPFILEHEADER*)pData;
  BITMAPCOREHEADER *pbci = (BITMAPCOREHEADER*)&pData[sizeof(BITMAPFILEHEADER)];

  if(pbci->bcSize == sizeof(BITMAPCOREHEADER))
  {
    //I haven't seen such an image yet
  }
  else if(pbci->bcSize == sizeof(BITMAPINFOHEADER))
  {
    BITMAPINFO *pbi = (BITMAPINFO*)pbci;
    BYTE *pbits = (BYTE*)&pData[pbfh->bfOffBits];

    int srcw = pbi->bmiHeader.biWidth;
    int srch = pbi->bmiHeader.biHeight;

    HDC maskdc = CreateCompatibleDC(hdc);
    HBITMAP bmpMask = CreateDIBitmap(hdc, &pbi->bmiHeader, CBM_INIT,
    pbits, pbi, DIB_RGB_COLORS);
    SelectObject(maskdc, bmpMask);

    int iRowSize = 4*srcw;

    pbi = (BITMAPINFO*)malloc(sizeof(BITMAPINFOHEADER) + iRowSize*srch);
    pbi->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    pbi->bmiHeader.biWidth = srcw;
    pbi->bmiHeader.biHeight = srch;
    pbi->bmiHeader.biPlanes = 1;
    pbi->bmiHeader.biBitCount = 32;
    pbi->bmiHeader.biCompression = BI_RGB;
    pbi->bmiHeader.biSizeImage = iRowSize*srch;
    pbi->bmiHeader.biXPelsPerMeter = 0;
    pbi->bmiHeader.biYPelsPerMeter = 0;
    pbi->bmiHeader.biClrUsed = 0;
    pbi->bmiHeader.biClrImportant = 0;

    BYTE *maskBits = (BYTE*)&pbi->bmiColors[0];
    GetDIBits(maskdc, bmpMask, 0, srch, maskBits, pbi, DIB_RGB_COLORS);
    DeleteObject(bmpMask);

    if(dwTranspColor == 0xFF000000)
    {
      memset(maskBits, 0x00, iRowSize*srch);
    }
    else
    {
      BYTE *maskBitsRow = maskBits;
      DWORD col;
      for(int j = 0; j < srch; j++)
      {
        for(int i = 0; i < srch; i++)
        {
          col = maskBitsRow[4*i + 2] | (maskBitsRow[4*i + 1] << 8) | (maskBitsRow[4*i] << 16);
          if(col == dwTranspColor)
          {
            maskBitsRow[4*i] = 0xFF;
            maskBitsRow[4*i + 1] = 0xFF;
            maskBitsRow[4*i + 2] = 0xFF;
            maskBitsRow[4*i + 3] = 0x00;
          }
          else
          {
            maskBitsRow[4*i] = 0x00;
            maskBitsRow[4*i + 1] = 0x00;
            maskBitsRow[4*i + 2] = 0x00;
            maskBitsRow[4*i + 3] = 0x00;
          }
        }
        maskBitsRow += iRowSize;
      }
    }

    hbmpRes = CreateCompatibleBitmap(hdc, iWidth, iHeight);
    SelectObject(maskdc, hbmpRes);
    SetStretchBltMode(maskdc, HALFTONE);
    StretchDIBits(maskdc, 0, 0, iWidth, iHeight, 0, 0, srcw, srch, maskBits,
    pbi, DIB_RGB_COLORS, SRCCOPY);

    free(pbi);

    DeleteDC(maskdc);
  }
  return(hbmpRes);
}

LPICONDATA BMPtoIconData(char *pData, int iDataSize, int iBitDepth,
  int iWidth, int iHeight, DWORD dwTranspColor, BOOL bDither, BOOL bOptimize)
{
  LPICONDATA icdRes = NULL;

  BITMAPFILEHEADER *pbfh = (BITMAPFILEHEADER*)pData;
  BITMAPCOREHEADER *pbci = (BITMAPCOREHEADER*)&pData[sizeof(BITMAPFILEHEADER)];

  if(pbci->bcSize == sizeof(BITMAPCOREHEADER))
  {
    //I haven't seen such an image yet
  }
  else if(pbci->bcSize == sizeof(BITMAPINFOHEADER))
  {
    BITMAPINFO *pbi = (BITMAPINFO*)pbci;
    BYTE *pbits = (BYTE*)&pData[pbfh->bfOffBits];

    int srcw = pbi->bmiHeader.biWidth;
    int srch = pbi->bmiHeader.biHeight;

    HDC dc = GetDC(0);
    HDC cdc = CreateCompatibleDC(dc);
    HBITMAP bmpColors = CreateCompatibleBitmap(dc, iWidth, iHeight);
    HBITMAP bmpMask = GetMaskFromBitmap(dc, pData, iDataSize, iBitDepth,
    iWidth, iHeight, dwTranspColor);
    SelectObject(cdc, bmpColors);

    SetStretchBltMode(cdc, HALFTONE);
    StretchDIBits(cdc, 0, 0, iWidth, iHeight, 0, 0, srcw, srch, pbits,
    pbi, DIB_RGB_COLORS, SRCCOPY);

    BYTE *pAlphaImg = (BYTE*)malloc(4*iWidth*iHeight);
    GetAlphaImage(cdc, bmpColors, bmpMask, iWidth, iHeight, pAlphaImg);

    int mnwidth = iWidth/16;
    if(mnwidth % 2) mnwidth++;
    mnwidth *= 16;

    int bitoffs = 0;
    if(iBitDepth < 16)
    {
      bitoffs = 1;
      for(int i = 0; i < iBitDepth; i++) bitoffs *= 2;
    }

    int imgrowsize = iBitDepth*iWidth/8;
    if(imgrowsize % 2) imgrowsize++;
    int maskrowsize = 0;
    maskrowsize = mnwidth/8;
    if(maskrowsize % 2) maskrowsize++;

    int iressize = sizeof(BITMAPINFOHEADER) + bitoffs*sizeof(RGBQUAD) + (imgrowsize + maskrowsize)*iHeight;
    icdRes = (LPICONDATA)malloc(sizeof(HBITMAP) + sizeof(ICONDIRENTRY) + iressize);
    pbi = (BITMAPINFO*)&icdRes->icoimg;

    pbi->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    pbi->bmiHeader.biWidth = iWidth;
    pbi->bmiHeader.biHeight = iHeight;
    pbi->bmiHeader.biPlanes = 1;
    pbi->bmiHeader.biBitCount = iBitDepth;
    pbi->bmiHeader.biCompression = BI_RGB;
    pbi->bmiHeader.biSizeImage = imgrowsize*iHeight;
    pbi->bmiHeader.biXPelsPerMeter = 0;
    pbi->bmiHeader.biYPelsPerMeter = 0;
    pbi->bmiHeader.biClrUsed = bitoffs;
    pbi->bmiHeader.biClrImportant = 0;

    if(bitoffs > 0)
    {
      /*if(iBitDepth == 8 && !bOptimize)
      {
        BYTE *pbimg = (BYTE*)&pbi->bmiColors[bitoffs];
        pbimg += imgrowsize*iHeight;

        pbi->bmiHeader.biBitCount = 1;
        pbi->bmiHeader.biSizeImage = maskrowsize*iHeight;
        pbi->bmiHeader.biClrUsed = 2;
        GetDIBits(dc, bmpMask, 0, iHeight, pbimg, pbi, DIB_RGB_COLORS);

        pbi->bmiHeader.biBitCount = iBitDepth;
        pbi->bmiHeader.biSizeImage = imgrowsize*iHeight;
        pbi->bmiHeader.biClrUsed = bitoffs;

        pbimg = (BYTE*)&pbi->bmiColors[bitoffs];
        GetDIBits(dc, bmpColors, 0, iHeight, pbimg, pbi, DIB_RGB_COLORS);
      }
      else*/ SimplifyAlphaImg(pbi, pAlphaImg, iWidth, iHeight, imgrowsize,
        maskrowsize, bDither, bOptimize);
    }
    else
    {
      BYTE *pbimg = (BYTE*)&pbi->bmiColors[bitoffs];
      pbimg += imgrowsize*iHeight;

      pbi->bmiHeader.biBitCount = 1;
      pbi->bmiHeader.biSizeImage = maskrowsize*iHeight;
      pbi->bmiHeader.biClrUsed = 2;
      GetDIBits(dc, bmpMask, 0, iHeight, pbimg, pbi, DIB_RGB_COLORS);

      pbi->bmiHeader.biBitCount = iBitDepth;
      pbi->bmiHeader.biSizeImage = imgrowsize*iHeight;
      pbi->bmiHeader.biClrUsed = bitoffs;

      pbimg = (BYTE*)&pbi->bmiColors[bitoffs];
      memcpy(pbimg, pAlphaImg, 4*iWidth*iHeight);
    }
    free(pAlphaImg);

    DeleteObject(bmpColors);
    DeleteObject(bmpMask);
    DeleteDC(cdc);
    ReleaseDC(0, dc);

    pbi->bmiHeader.biHeight = 2*iHeight;

    icdRes->hPreview = NULL;
    icdRes->icdi.bWidth = iWidth;
    icdRes->icdi.bHeight = iHeight;
    icdRes->icdi.bColorCount = bitoffs;
    icdRes->icdi.bReserved = 0;
    icdRes->icdi.wPlanes = 1;
    icdRes->icdi.wBitCount = iBitDepth;
    icdRes->icdi.dwBytesInRes = iressize;
    icdRes->icdi.dwImageOffset = 0;
  }

  return icdRes;
}

void GetEditImage(LPICONIMAGE pIcoImg, LPBITMAPINFO pBitmap, COLORREF clTransp)
{
  int clmask;
  BYTE br, bg, bb, ri, gi, bi;
  br = clTransp & 0xFF;
  bg = (clTransp >> 8) & 0xFF;
  bb = (clTransp >> 16) & 0xFF;

  int width = pIcoImg->icHeader.biWidth;
  int height = pIcoImg->icHeader.biHeight / 2;
  int mnwidth = width/16;
  if(mnwidth % 2) mnwidth++;
  mnwidth *= 16;

  int bitoffs = 0;
  int clidx;
  int srcbitcnt = pIcoImg->icHeader.biBitCount;

  if(srcbitcnt < 16)
  {
    if(pIcoImg->icHeader.biClrUsed > 0)
      bitoffs = pIcoImg->icHeader.biClrUsed;
    else
    {
      bitoffs = 1;
      for(int i = 0; i < srcbitcnt; i++) bitoffs *= 2;
    }
  }

  int imgrowsize = srcbitcnt*width/8;
  if(imgrowsize % 2) imgrowsize++;
  int maskrowsize = mnwidth/8;
  if(maskrowsize % 2) maskrowsize++;
  int resrowsize = 4*width;

  BYTE *pbimg = (BYTE*)&pIcoImg->icColors[bitoffs];
  BYTE *pbmask = NULL;
  if(pIcoImg->icHeader.biBitCount < 32) pbmask = &pbimg[imgrowsize*height];
  BYTE *pbimgrow = pbimg;
  BYTE *pbmaskrow = pbmask;

  pBitmap->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
  pBitmap->bmiHeader.biWidth = width;
  pBitmap->bmiHeader.biHeight = height;
  pBitmap->bmiHeader.biPlanes = 1;
  pBitmap->bmiHeader.biBitCount = 32;
  pBitmap->bmiHeader.biCompression = BI_RGB;
  pBitmap->bmiHeader.biSizeImage = resrowsize*height;
  pBitmap->bmiHeader.biXPelsPerMeter = 0;
  pBitmap->bmiHeader.biYPelsPerMeter = 0;
  pBitmap->bmiHeader.biClrUsed = 0;
  pBitmap->bmiHeader.biClrImportant = 0;

  BYTE *pbres = (BYTE*)pBitmap->bmiColors;
  BYTE *pbresrow = pbres;
  WORD wcol;

  for(int j = 0; j < height; j++)
  {
    for(int i = 0; i < width; i++)
    {
      if(pbmaskrow)
      {
        clmask = (pbmaskrow[i/8] >> (7 - (i % 8))) & 0x01;
        if(srcbitcnt < 16)
        {
          clidx = (pbimgrow[srcbitcnt*i/8] >>
            (8 - srcbitcnt - srcbitcnt*(i % (8/srcbitcnt)))) &
            (0xFF >> (8 - srcbitcnt));
          ri = pIcoImg->icColors[clidx].rgbRed;
          gi = pIcoImg->icColors[clidx].rgbGreen;
          bi = pIcoImg->icColors[clidx].rgbBlue;
        }
        else if(srcbitcnt == 16)
        {
          wcol = (pbimgrow[2*i] << 8) | pbimgrow[2*i + 1];
          bi = 8*((wcol >> 11) & 0x001F);
          bi = 8*((wcol >> 6) & 0x001F);
          bi = 8*((wcol >> 1) & 0x001F);
        }
        else if(srcbitcnt == 24)
        {
          bi = pbimgrow[3*i] & 0xFF;
          gi = pbimgrow[3*i + 1] & 0xFF;
          ri = pbimgrow[3*i + 2] & 0xFF;
        }
        if(clmask)
        {
          pbresrow[4*i] = bb;
          pbresrow[4*i + 1] = bg;
          pbresrow[4*i + 2] = br;
          pbresrow[4*i + 3] = 1; // we may use the fourth byte as
          // a buffer for various information. For example flag
          // 2 means selected, flag 1 transparent
        }
        else
        {
          pbresrow[4*i] = bi;
          pbresrow[4*i + 1] = gi;
          pbresrow[4*i + 2] = ri;
          pbresrow[4*i + 3] = 0;
        }
      }
    }
    pbimgrow += imgrowsize;
    if(pbmaskrow) pbmaskrow += maskrowsize;
    pbresrow += resrowsize;
  }

  return;
}

void SaveEditImage(LPICONIMAGE pIcoImg, LPBITMAPINFO pBitmap)
{
  int clmask;
  BYTE ri, gi, bi;

  int width = pIcoImg->icHeader.biWidth;
  int height = pIcoImg->icHeader.biHeight / 2;
  int mnwidth = width/16;
  if(mnwidth % 2) mnwidth++;
  mnwidth *= 16;

  int bitoffs = 0;
  int clidx;
  int srcbitcnt = pIcoImg->icHeader.biBitCount;

  if(srcbitcnt < 16)
  {
    if(pIcoImg->icHeader.biClrUsed > 0)
      bitoffs = pIcoImg->icHeader.biClrUsed;
    else
    {
      bitoffs = 1;
      for(int i = 0; i < srcbitcnt; i++) bitoffs *= 2;
    }
  }

  int imgrowsize = srcbitcnt*width/8;
  if(imgrowsize % 2) imgrowsize++;
  int maskrowsize = mnwidth/8;
  if(maskrowsize % 2) maskrowsize++;
  int resrowsize = 4*width;

  BYTE *pbimg = (BYTE*)&pIcoImg->icColors[bitoffs];
  BYTE *pbmask = NULL;
  if(pIcoImg->icHeader.biBitCount < 32) pbmask = &pbimg[imgrowsize*height];
  BYTE *pbimgrow = pbimg;
  BYTE *pbmaskrow = pbmask;

  BYTE *pbres = (BYTE*)pBitmap->bmiColors;
  BYTE *pbresrow = pbres;
  WORD wcol;

  for(int j = 0; j < height; j++)
  {
    if(pbmaskrow) memset(pbmaskrow, 0, maskrowsize);
    memset(pbimgrow, 0, imgrowsize);
    for(int i = 0; i < width; i++)
    {
      if(pbmaskrow)
      {
        if(pbresrow[4*i + 3] & 1)
        {
          ri = 0;
          gi = 0;
          bi = 0;
          clmask = 1;
        }
        else
        {
          bi = pbresrow[4*i];
          gi = pbresrow[4*i + 1];
          ri = pbresrow[4*i + 2];
          clmask = 0;
        }
        pbmaskrow[i/8] |= clmask << (7 - (i % 8));
        if(srcbitcnt < 16)
        {
          clidx = GetPalIndex(pIcoImg->icColors, bitoffs, ri, gi, bi);
          pbimgrow[srcbitcnt*i/8] |= clidx << (8 - srcbitcnt - srcbitcnt*(i % (8/srcbitcnt)));
        }
        else if(srcbitcnt == 16)
        {
          wcol = ((bi/8) << 11) | ((gi/8) << 6) | ((ri/8) << 1);
          pbimgrow[2*i] = (wcol >> 8) & 0xFF;
          pbimgrow[2*i + 1] = wcol & 0xFF;
        }
        else if(srcbitcnt == 24)
        {
          pbimgrow[3*i] = bi;
          pbimgrow[3*i + 1] = gi;
          pbimgrow[3*i + 2] = ri;
        }
      }
    }
    pbimgrow += imgrowsize;
    if(pbmaskrow) pbmaskrow += maskrowsize;
    pbresrow += resrowsize;
  }
  return;
}
