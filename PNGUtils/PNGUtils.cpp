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

#include "PNGUtils.hpp"
#include "png.h"

struct PNGImageRec
{
    long imgHandle;
    char *imgData;
    int imgDataLen;
    int iCurPos;
    bool bImageRead;
    bool bDeleted;
    png_structp png_ptr;
    png_infop info_ptr;
    png_infop end_info;
    png_uint_32 width, height;
    int bit_depth, color_type, interlace_type, compression_type, filter_method;
};

int imgcount;
int imgsize;
PNGImageRec *imgsbuf;

void pu_read_data(png_structp png_ptr, png_bytep data, png_size_t length)
{
    PNGImageRec *pir = (PNGImageRec*)png_get_io_ptr(png_ptr);
    int bytes_to_read = pir->imgDataLen - pir->iCurPos;
    if(length < bytes_to_read) bytes_to_read = length;
    memcpy(data, &pir->imgData[pir->iCurPos], bytes_to_read);
    pir->iCurPos += bytes_to_read;
    return;
}

void pu_error_fn(png_structp png_ptr, png_const_charp error_msg)
{
    return;
}

void pu_warning_fn(png_structp png_ptr, png_const_charp warning_msg)
{
    return;
}


extern "C" BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	switch(fdwReason)
	{
	case DLL_PROCESS_ATTACH:
	    imgcount = 0;
	    imgsize = 8;
	    imgsbuf = (PNGImageRec*)malloc(imgsize*sizeof(PNGImageRec));
	    break;
    case DLL_PROCESS_DETACH:
        free(imgsbuf);
        break;
	}
	return TRUE;
}

int FindImage(long img)
{
    int i = 0;
    bool bFound = false;
    while(!bFound && (i < imgcount))
    {
        bFound = imgsbuf[i++].imgHandle == img;
    }
    if(bFound)
    {
        if(imgsbuf[--i].bDeleted) return(-1);
        return(i);
    }
    return(-1);
}

long PNGImageCreate(char* buf, long bufsize)
{
    long lret = imgcount + 1;
    if(imgcount >= imgsize)
    {
	    imgsize += 8;
	    imgsbuf = (PNGImageRec*)realloc(imgsbuf, imgsize*sizeof(PNGImageRec));
    }
    imgsbuf[imgcount].imgHandle = lret;
    imgsbuf[imgcount].imgData = buf;
    imgsbuf[imgcount].imgDataLen = bufsize;
    imgsbuf[imgcount].iCurPos = 0;
    imgsbuf[imgcount].bImageRead = false;
    imgsbuf[imgcount].bDeleted = false;

    imgsbuf[imgcount].png_ptr = png_create_read_struct(PNG_LIBPNG_VER_STRING,
        (png_voidp)NULL, pu_error_fn, pu_warning_fn);
    if(!imgsbuf[imgcount].png_ptr) return(-1);

    imgsbuf[imgcount].info_ptr =
        png_create_info_struct(imgsbuf[imgcount].png_ptr);
    if(!imgsbuf[imgcount].info_ptr)
    {
        png_destroy_read_struct(&imgsbuf[imgcount].png_ptr, (png_infopp)NULL,
            (png_infopp)NULL);
        return(-1);
    }

    imgsbuf[imgcount].end_info =
        png_create_info_struct(imgsbuf[imgcount].png_ptr);
    if(!imgsbuf[imgcount].end_info)
    {
        png_destroy_read_struct(&imgsbuf[imgcount].png_ptr,
            &imgsbuf[imgcount].info_ptr, (png_infopp)NULL);
        return(-1);
    }

    /*png_set_read_fn(imgsbuf[imgcount].png_ptr,
        png_get_io_ptr(imgsbuf[imgcount].png_ptr), pu_read_data);*/
    png_set_read_fn(imgsbuf[imgcount].png_ptr, &imgsbuf[imgcount],
        pu_read_data);
    png_read_info(imgsbuf[imgcount].png_ptr, imgsbuf[imgcount].info_ptr);

    png_get_IHDR(imgsbuf[imgcount].png_ptr, imgsbuf[imgcount].info_ptr,
        &imgsbuf[imgcount].width, &imgsbuf[imgcount].height,
        &imgsbuf[imgcount].bit_depth, &imgsbuf[imgcount].color_type,
        &imgsbuf[imgcount].interlace_type, &imgsbuf[imgcount].compression_type,
        &imgsbuf[imgcount].filter_method);

    imgcount++;
    return(lret);
}

void PNGImageDestroy(long img)
{
    int i = FindImage(img);
    if(i > -1)
    {
        if(imgsbuf[i].bImageRead)
            png_read_end(imgsbuf[i].png_ptr, imgsbuf[i].end_info);
        png_destroy_read_struct(&imgsbuf[i].png_ptr, &imgsbuf[i].info_ptr,
            &imgsbuf[i].end_info);
        imgsbuf[i].bDeleted = true;
    }
    if(i == imgcount - 1) imgcount--;
    return;
}

BOOL PNGImageIsPNG(long img)
{
    int i = FindImage(img);
    if(i > -1)
    {
        if(imgsbuf[i].imgDataLen < 8) return(FALSE);
        return(!png_sig_cmp((png_byte*)imgsbuf[i].imgData, 0, 8));
    }
    return(FALSE);
}

void PNGImageGetDimen(long img, long* width, long* height)
{
    *width = 0;
    *height = 0;
    int i = FindImage(img);
    if(i > -1)
    {
        *width = imgsbuf[i].width;
        *height = imgsbuf[i].height;
    }
    return;
}

int PNGImageGetBitDepth(long img)
{
    int iRes = 0;
    int i = FindImage(img);
    if(i > -1) iRes = imgsbuf[i].bit_depth;
    if(imgsbuf[i].color_type == PNG_COLOR_TYPE_RGB) iRes = 24;
    else if(imgsbuf[i].color_type == PNG_COLOR_TYPE_RGB_ALPHA) iRes = 32;
    return(iRes);
}

int CopyImage(HDC hdc, png_byte **rows, int width, int height,
    BackgroundFunction backfnc)
{
    int resrowsize = 4*width;

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

    DWORD bkcol = 0;
    png_byte *row;
    png_byte alpha;
    for(int j = 0; j < height; j++)
    {
        row = rows[height - j - 1];
        for(int i = 0; i < width; i++)
        {
            if(backfnc) bkcol = backfnc(i, j);
            alpha = row[4*i + 3];
            pbresrow[4*i + 2] = ((255 - alpha)*((bkcol >> 16) & 0xFF) +
                alpha*row[4*i])/255;
            pbresrow[4*i + 1] = ((255 - alpha)*((bkcol >> 8) & 0xFF) +
                alpha*row[4*i + 1])/255;
            pbresrow[4*i] = ((255 - alpha)*(bkcol & 0xFF) +
                alpha*row[4*i + 2])/255;
            pbresrow[4*i + 3] = 0;
        }
        pbresrow += resrowsize;
    }

    SetDIBitsToDevice(hdc, 0, 0, width, height, 0, 0, 0, height, pbres, &bif,
        DIB_RGB_COLORS);
    free(pbres);
    return(0);
}

int StretchImage(HDC hdc, png_byte **rows, int srcwidth, int srcheight,
    int dstwidth, int dstheight, BackgroundFunction backfnc)
{
    HDC ldc = CreateCompatibleDC(hdc);
    HBITMAP bmp  = CreateCompatibleBitmap(hdc, dstwidth, dstheight);
    SelectObject(ldc, bmp);
    SetStretchBltMode(ldc, HALFTONE);

    int inrowsize = 4*srcwidth;
    int outrowsize = 4*dstwidth;

    BITMAPINFO bifin, bifout;
    bifin.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bifin.bmiHeader.biWidth = srcwidth;
    bifin.bmiHeader.biHeight = srcheight;
    bifin.bmiHeader.biPlanes = 1;
    bifin.bmiHeader.biBitCount = 32;
    bifin.bmiHeader.biCompression = BI_RGB;
    bifin.bmiHeader.biSizeImage = inrowsize*srcheight;
    bifin.bmiHeader.biXPelsPerMeter = 0;
    bifin.bmiHeader.biYPelsPerMeter = 0;
    bifin.bmiHeader.biClrUsed = 0;
    bifin.bmiHeader.biClrImportant = 0;

    bifout.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bifout.bmiHeader.biWidth = dstwidth;
    bifout.bmiHeader.biHeight = dstheight;
    bifout.bmiHeader.biPlanes = 1;
    bifout.bmiHeader.biBitCount = 32;
    bifout.bmiHeader.biCompression = BI_RGB;
    bifout.bmiHeader.biSizeImage = outrowsize*dstheight;
    bifout.bmiHeader.biXPelsPerMeter = 0;
    bifout.bmiHeader.biYPelsPerMeter = 0;
    bifout.bmiHeader.biClrUsed = 0;
    bifout.bmiHeader.biClrImportant = 0;

    BYTE *pbin = (BYTE*)malloc(bifin.bmiHeader.biSizeImage*sizeof(BYTE));
    BYTE *pbcol = (BYTE*)malloc(bifout.bmiHeader.biSizeImage*sizeof(BYTE));
    BYTE *pbmask = (BYTE*)malloc(bifout.bmiHeader.biSizeImage*sizeof(BYTE));
    BYTE *pbinrow = pbin;

    DWORD bkcol = 0;
    png_byte *row;
    png_byte alpha;

    // first fill the color data
    for(int j = 0; j < srcheight; j++)
    {
        row = rows[srcheight - j - 1];
        for(int i = 0; i < srcwidth; i++)
        {
            pbinrow[4*i + 2] = row[4*i];
            pbinrow[4*i + 1] = row[4*i + 1];
            pbinrow[4*i] = row[4*i + 2];
            pbinrow[4*i + 3] = 0;
        }
        pbinrow += inrowsize;
    }

    StretchDIBits(ldc, 0, 0, dstwidth, dstheight, 0, 0, srcwidth, srcheight,
        pbin, &bifin, DIB_RGB_COLORS, SRCCOPY);
    // and get the stretched color bits
    GetDIBits(ldc, bmp, 0, dstheight, pbcol, &bifout, DIB_RGB_COLORS);

    pbinrow = pbin;
    // then fill the mask data
    for(int j = 0; j < srcheight; j++)
    {
        row = rows[srcheight - j - 1];
        for(int i = 0; i < srcwidth; i++)
        {
            pbinrow[4*i + 2] = row[4*i + 3];
            pbinrow[4*i + 1] = row[4*i + 3];
            pbinrow[4*i] = row[4*i + 3];
            pbinrow[4*i + 3] = 0;
        }
        pbinrow += inrowsize;
    }

    StretchDIBits(ldc, 0, 0, dstwidth, dstheight, 0, 0, srcwidth, srcheight,
        pbin, &bifin, DIB_RGB_COLORS, SRCCOPY);
    // and get the stretched mask bits
    GetDIBits(ldc, bmp, 0, dstheight, pbmask, &bifout, DIB_RGB_COLORS);

    free(pbin);

    // finally combine the color and mask bits
    BYTE *pbcolrow = pbcol;
    BYTE *pbmaskrow = pbmask;
    for(int j = 0; j < dstheight; j++)
    {
        for(int i = 0; i < dstwidth; i++)
        {
            if(backfnc) bkcol = backfnc(i, j);
            alpha = pbmaskrow[4*i];
            pbcolrow[4*i + 2] = ((255 - alpha)*((bkcol >> 16) & 0xFF) +
                alpha*pbcolrow[4*i + 2])/255;
            pbcolrow[4*i + 1] = ((255 - alpha)*((bkcol >> 8) & 0xFF) +
                alpha*pbcolrow[4*i + 1])/255;
            pbcolrow[4*i] = ((255 - alpha)*(bkcol & 0xFF) +
                alpha*pbcolrow[4*i])/255;
            pbcolrow[4*i + 3] = 0;
        }
        pbcolrow += outrowsize;
        pbmaskrow += outrowsize;
    }

    DeleteObject(bmp);
    DeleteDC(ldc);

    SetDIBitsToDevice(hdc, 0, 0, dstwidth, dstheight, 0, 0, 0, dstheight,
        pbcol, &bifout, DIB_RGB_COLORS);
    free(pbcol);
    free(pbmask);
    return(0);
}

int PNGImageRender(long img, int vw, int vh, HDC ADC, int destw,
    int desth, BackgroundFunction backfnc)
{
    int i = FindImage(img);
    if(i > -1)
    {
//MessageBox(0, L"Dobry 1", L"Debug", MB_OK);
        if(imgsbuf[i].bit_depth < 8) png_set_packing(imgsbuf[i].png_ptr);

        png_color_8p sig_bit;
        if(png_get_sBIT(imgsbuf[i].png_ptr, imgsbuf[i].info_ptr, &sig_bit))
            png_set_shift(imgsbuf[i].png_ptr, sig_bit);

        if(imgsbuf[i].color_type == PNG_COLOR_TYPE_PALETTE)
            png_set_palette_to_rgb(imgsbuf[i].png_ptr);
        if(imgsbuf[i].color_type == PNG_COLOR_TYPE_GRAY &&
            imgsbuf[i].bit_depth < 8)
                png_set_expand_gray_1_2_4_to_8(imgsbuf[i].png_ptr);
        if(imgsbuf[i].color_type == PNG_COLOR_TYPE_RGB ||
            imgsbuf[i].color_type == PNG_COLOR_TYPE_GRAY)
                png_set_add_alpha(imgsbuf[i].png_ptr, 255, PNG_FILLER_AFTER);
        if(png_get_valid(imgsbuf[i].png_ptr, imgsbuf[i].info_ptr,
            PNG_INFO_tRNS))
                png_set_tRNS_to_alpha(imgsbuf[i].png_ptr);
        png_read_update_info(imgsbuf[i].png_ptr, imgsbuf[i].info_ptr);

        png_byte **rows = (png_byte**)malloc(vh*sizeof(png_byte*));
        for(int j = 0; j < vh; j++)
        {
            rows[j] = (png_byte*)malloc(vw*4*sizeof(png_byte));
        }
        png_read_image(imgsbuf[i].png_ptr, rows);
        imgsbuf[i].bImageRead = true;

        if(vw == destw && vh == desth) CopyImage(ADC, rows, vw, vh, backfnc);
        else StretchImage(ADC, rows, vw, vh, destw, desth, backfnc);

        for(int j = 0; j < vh; j++) free(rows[j]);
        free(rows);
    }
    return(0);
}

void GetColorBmp(HDC hdc, png_byte **rows, int width, int height)
{
    int resrowsize = 4*width;

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

    png_byte *row;
    for(int j = 0; j < height; j++)
    {
        row = rows[height - j - 1];
        for(int i = 0; i < width; i++)
        {
            pbresrow[4*i + 2] = row[4*i];
            pbresrow[4*i + 1] = row[4*i + 1];
            pbresrow[4*i] = row[4*i + 2];
            pbresrow[4*i + 3] = 0;
        }
        pbresrow += resrowsize;
    }

    SetDIBitsToDevice(hdc, 0, 0, width, height, 0, 0, 0, height, pbres, &bif,
        DIB_RGB_COLORS);
    free(pbres);
    return;
}

void GetMaskBmp(HDC hdc, png_byte **rows, int width, int height)
{
    int resrowsize = 4*width;

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

    png_byte *row;
    png_byte alpha;
    for(int j = 0; j < height; j++)
    {
        row = rows[height - j - 1];
        for(int i = 0; i < width; i++)
        {
            alpha = 255 - row[4*i + 3];
            pbresrow[4*i + 2] = alpha;
            pbresrow[4*i + 1] = alpha;
            pbresrow[4*i] = alpha;
            pbresrow[4*i + 3] = 0;
        }
        pbresrow += resrowsize;
    }

    SetDIBitsToDevice(hdc, 0, 0, width, height, 0, 0, 0, height, pbres, &bif,
        DIB_RGB_COLORS);
    free(pbres);
    return;
}

BOOL PNGImageGetBitmaps(long img, HDC ADC, HBITMAP hBmpColor,
    HBITMAP hBmpMask)
{
    int i = FindImage(img);
    if(i > -1)
    {
        if(imgsbuf[i].bit_depth < 8) png_set_packing(imgsbuf[i].png_ptr);

        png_color_8p sig_bit;
        if(png_get_sBIT(imgsbuf[i].png_ptr, imgsbuf[i].info_ptr, &sig_bit))
            png_set_shift(imgsbuf[i].png_ptr, sig_bit);

        if(imgsbuf[i].color_type == PNG_COLOR_TYPE_PALETTE)
            png_set_palette_to_rgb(imgsbuf[i].png_ptr);
        if(imgsbuf[i].color_type == PNG_COLOR_TYPE_GRAY &&
            imgsbuf[i].bit_depth < 8)
                png_set_expand_gray_1_2_4_to_8(imgsbuf[i].png_ptr);
        if(imgsbuf[i].color_type == PNG_COLOR_TYPE_RGB ||
            imgsbuf[i].color_type == PNG_COLOR_TYPE_GRAY)
                png_set_add_alpha(imgsbuf[i].png_ptr, 255, PNG_FILLER_AFTER);
        if(png_get_valid(imgsbuf[i].png_ptr, imgsbuf[i].info_ptr,
            PNG_INFO_tRNS))
                png_set_tRNS_to_alpha(imgsbuf[i].png_ptr);
        png_read_update_info(imgsbuf[i].png_ptr, imgsbuf[i].info_ptr);

        png_byte **rows = (png_byte**)malloc(imgsbuf[i].height*sizeof(png_byte*));
        for(int j = 0; j < imgsbuf[i].height; j++)
        {
            rows[j] = (png_byte*)malloc(imgsbuf[i].width*4*sizeof(png_byte));
        }
        png_read_image(imgsbuf[i].png_ptr, rows);
        imgsbuf[i].bImageRead = true;

        SelectObject(ADC, hBmpColor);
        GetColorBmp(ADC, rows, imgsbuf[i].width, imgsbuf[i].height);
        SelectObject(ADC, hBmpMask);
        GetMaskBmp(ADC, rows, imgsbuf[i].width, imgsbuf[i].height);

        for(int j = 0; j < imgsbuf[i].height; j++) free(rows[j]);
        free(rows);

        return(TRUE);
    }
    return(FALSE);
}

BOOL IsPNGFile(char* buf, long bufsize)
{
    if(bufsize < 8) return(FALSE);
    return(!png_sig_cmp((png_byte*)buf, 0, 8));
}
