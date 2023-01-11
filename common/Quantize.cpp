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

#include "Quantize.h"

CQuantizer::CQuantizer (UINT nMaxColors, UINT nColorBits)
{
    //ASSERT (nColorBits <= 8);

    m_pTree = NULL;
    m_nLeafCount = 0;
    for (int i=0; i<=(int) nColorBits; i++)
        m_pReducibleNodes[i] = NULL;
    m_nMaxColors = nMaxColors;
    m_nColorBits = nColorBits;
}

CQuantizer::~CQuantizer ()
{
    if (m_pTree != NULL)
        DeleteTree (&m_pTree);
}

BOOL CQuantizer::ProcessImage (HANDLE hImage)
{
    DWORD rmask, gmask, bmask;
    int rright, gright, bright;
    int rleft, gleft, bleft;
    BYTE* pbBits;
    WORD* pwBits;
    DWORD* pdwBits;
    BYTE r, g, b;
    WORD wColor;
    DWORD dwColor;
    int i, j;
    HDC hdc;
    BYTE* pBuffer;
    BITMAPINFO bmi;

    DIBSECTION ds;
    ::GetObject (hImage, sizeof (ds), &ds);
    int nPad = ds.dsBm.bmWidthBytes - (((ds.dsBmih.biWidth *
        ds.dsBmih.biBitCount) + 7) / 8);

    switch (ds.dsBmih.biBitCount) {

    case 1: // 1-bit DIB
    case 4: // 4-bit DIB
    case 8: // 8-bit DIB
        //
        // The strategy here is to use ::GetDIBits to convert the
        // image into a 24-bit DIB one scan line at a time. A pleasant
        // side effect of using ::GetDIBits in this manner is that RLE-
        // encoded 4-bit and 8-bit DIBs will be uncompressed.
        //
        hdc = ::GetDC (NULL);
        pBuffer = new BYTE[ds.dsBmih.biWidth * 3];

        ::ZeroMemory (&bmi, sizeof (bmi));
        bmi.bmiHeader.biSize = sizeof (BITMAPINFOHEADER);
        bmi.bmiHeader.biWidth = ds.dsBmih.biWidth;
        bmi.bmiHeader.biHeight = ds.dsBmih.biHeight;
        bmi.bmiHeader.biPlanes = 1;
        bmi.bmiHeader.biBitCount = 24;
        bmi.bmiHeader.biCompression = BI_RGB;

        for (i=0; i<ds.dsBmih.biHeight; i++) {
            ::GetDIBits (hdc, (HBITMAP) hImage, i, 1, pBuffer, &bmi,
                DIB_RGB_COLORS);
            pbBits = pBuffer;
            for (j=0; j<ds.dsBmih.biWidth; j++) {
                b = *pbBits++;
                g = *pbBits++;
                r = *pbBits++;
                AddColor (&m_pTree, r, g, b, m_nColorBits, 0, &m_nLeafCount,
                    m_pReducibleNodes);
                while (m_nLeafCount > m_nMaxColors)
                    ReduceTree (m_nColorBits, &m_nLeafCount,
                        m_pReducibleNodes);
            }
        }

        delete pBuffer;
        ::ReleaseDC (NULL, hdc);
        break;

    case 16: // 16-bit DIB
        if (ds.dsBmih.biCompression == BI_BITFIELDS) {
            rmask = ds.dsBitfields[0];
            gmask = ds.dsBitfields[1];
            bmask = ds.dsBitfields[2];
        }
        else {
            rmask = 0x7C00;
            gmask = 0x03E0;
            bmask = 0x001F;
        }

        rright = GetRightShiftCount (rmask);
        gright = GetRightShiftCount (gmask);
        bright = GetRightShiftCount (bmask);

        rleft = GetLeftShiftCount (rmask);
        gleft = GetLeftShiftCount (gmask);
        bleft = GetLeftShiftCount (bmask);

        pwBits = (WORD*) ds.dsBm.bmBits;
        for (i=0; i<ds.dsBmih.biHeight; i++) {
            for (j=0; j<ds.dsBmih.biWidth; j++) {
                wColor = *pwBits++;
                b = (BYTE) (((wColor & (WORD) bmask) >> bright) << bleft);
                g = (BYTE) (((wColor & (WORD) gmask) >> gright) << gleft);
                r = (BYTE) (((wColor & (WORD) rmask) >> rright) << rleft);
                AddColor (&m_pTree, r, g, b, m_nColorBits, 0, &m_nLeafCount,
                    m_pReducibleNodes);
                while (m_nLeafCount > m_nMaxColors)
                    ReduceTree (m_nColorBits, &m_nLeafCount, m_pReducibleNodes);
            }
            pwBits = (WORD*) (((BYTE*) pwBits) + nPad);
        }
        break;

    case 24: // 24-bit DIB
        pbBits = (BYTE*) ds.dsBm.bmBits;
        for (i=0; i<ds.dsBmih.biHeight; i++) {
            for (j=0; j<ds.dsBmih.biWidth; j++) {
                b = *pbBits++;
                g = *pbBits++;
                r = *pbBits++;
                AddColor (&m_pTree, r, g, b, m_nColorBits, 0, &m_nLeafCount,
                    m_pReducibleNodes);
                while (m_nLeafCount > m_nMaxColors)
                    ReduceTree (m_nColorBits, &m_nLeafCount, m_pReducibleNodes);
            }
            pbBits += nPad;
        }
        break;

    case 32: // 32-bit DIB
        if (ds.dsBmih.biCompression == BI_BITFIELDS) {
            rmask = ds.dsBitfields[0];
            gmask = ds.dsBitfields[1];
            bmask = ds.dsBitfields[2];
        }
        else {
            rmask = 0x00FF0000;
            gmask = 0x0000FF00;
            bmask = 0x000000FF;
        }

        rright = GetRightShiftCount (rmask);
        gright = GetRightShiftCount (gmask);
        bright = GetRightShiftCount (bmask);

        pdwBits = (DWORD*) ds.dsBm.bmBits;
        for (i=0; i<ds.dsBmih.biHeight; i++) {
            for (j=0; j<ds.dsBmih.biWidth; j++) {
                dwColor = *pdwBits++;
                b = (BYTE) ((dwColor & bmask) >> bright);
                g = (BYTE) ((dwColor & gmask) >> gright);
                r = (BYTE) ((dwColor & rmask) >> rright);
                AddColor (&m_pTree, r, g, b, m_nColorBits, 0, &m_nLeafCount,
                    m_pReducibleNodes);
                while (m_nLeafCount > m_nMaxColors)
                    ReduceTree (m_nColorBits, &m_nLeafCount, m_pReducibleNodes);
            }
            pdwBits = (DWORD*) (((BYTE*) pdwBits) + nPad);
        }
        break;

    default: // Unrecognized color format
        return FALSE;
    }
    return TRUE;
}

BOOL CQuantizer::ProcessBitmap(HBITMAP hBmp, int iWidth, int iHeight)
{
    DWORD rmask, gmask, bmask;
    int rright, gright, bright;
    int rleft, gleft, bleft;
    BYTE* pbBits;
    WORD* pwBits;
    DWORD* pdwBits;
    BYTE r, g, b;
    WORD wColor;
    DWORD dwColor;
    int i, j;
    HDC hdc;
    BYTE* pBuffer;
    BITMAPINFO bmi;

    /*DIBSECTION ds;
    ::GetObject (hImage, sizeof (ds), &ds);
    int nPad = ds.dsBm.bmWidthBytes - (((ds.dsBmih.biWidth *
        ds.dsBmih.biBitCount) + 7) / 8);*/

    //
    // The strategy here is to use ::GetDIBits to convert the
    // image into a 24-bit DIB one scan line at a time. A pleasant
    // side effect of using ::GetDIBits in this manner is that RLE-
    // encoded 4-bit and 8-bit DIBs will be uncompressed.
    //
    hdc = ::GetDC (NULL);
    pBuffer = new BYTE[iWidth * 3];

    ::ZeroMemory (&bmi, sizeof (bmi));
    bmi.bmiHeader.biSize = sizeof (BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = iWidth;
    bmi.bmiHeader.biHeight = iHeight;
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 24;
    bmi.bmiHeader.biCompression = BI_RGB;

    for (i=0; i<iHeight; i++) {
        ::GetDIBits (hdc, hBmp, i, 1, pBuffer, &bmi, DIB_RGB_COLORS);
        pbBits = pBuffer;
        for (j=0; j<iWidth; j++) {
            b = *pbBits++;
            g = *pbBits++;
            r = *pbBits++;
            AddColor (&m_pTree, r, g, b, m_nColorBits, 0, &m_nLeafCount,
                m_pReducibleNodes);
            while (m_nLeafCount > m_nMaxColors)
                ReduceTree (m_nColorBits, &m_nLeafCount,
                    m_pReducibleNodes);
        }
    }

    delete pBuffer;
    ::ReleaseDC (NULL, hdc);
    return TRUE;
}

BOOL CQuantizer::ProcessAlphaImg(BYTE *pImg, int iWidth, int iHeight)
{
    BYTE* pbBits;
    BYTE r, g, b, a;
    BYTE* pBuffer = pImg;
    int irowsize = 4*iWidth;

    for(int i = 0; i < iHeight; i++)
    {
        pbBits = pBuffer;
        for(int j = 0; j < iWidth; j++)
        {
            b = *pbBits++;
            g = *pbBits++;
            r = *pbBits++;
            a = *pbBits++;
            AddColor(&m_pTree, r, g, b, m_nColorBits, 0, &m_nLeafCount,
                m_pReducibleNodes);
            while(m_nLeafCount > m_nMaxColors)
                ReduceTree(m_nColorBits, &m_nLeafCount, m_pReducibleNodes);
        }
        pBuffer += irowsize;
    }

    return TRUE;
}

int CQuantizer::GetLeftShiftCount (DWORD dwVal)
{
    int nCount = 0;
    for (int i=0; i<sizeof (DWORD) * 8; i++) {
        if (dwVal & 1)
            nCount++;
        dwVal >>= 1;
    }
    return (8 - nCount);
}

int CQuantizer::GetRightShiftCount (DWORD dwVal)
{
    for (int i=0; i<sizeof (DWORD) * 8; i++) {
        if (dwVal & 1)
            return i;
        dwVal >>= 1;
    }
    return -1;
}

void CQuantizer::AddColor (NODE** ppNode, BYTE r, BYTE g, BYTE b,
    UINT nColorBits, UINT nLevel, UINT* pLeafCount, NODE** pReducibleNodes)
{
    static BYTE mask[8] = { 0x80, 0x40, 0x20, 0x10, 0x08, 0x04, 0x02, 0x01 };

    //
    // If the node doesn't exist, create it.
    //
    if (*ppNode == NULL)
        *ppNode = CreateNode (nLevel, nColorBits, pLeafCount,
            pReducibleNodes);

    //
    // Update color information if it's a leaf node.
    //
    if ((*ppNode)->bIsLeaf) {
        (*ppNode)->nPixelCount++;
        (*ppNode)->nRedSum += r;
        (*ppNode)->nGreenSum += g;
        (*ppNode)->nBlueSum += b;
    }

    //
    // Recurse a level deeper if the node is not a leaf.
    //
    else {
        int shift = 7 - nLevel;
        int nIndex = (((r & mask[nLevel]) >> shift) << 2) |
            (((g & mask[nLevel]) >> shift) << 1) |
            ((b & mask[nLevel]) >> shift);
        AddColor (&((*ppNode)->pChild[nIndex]), r, g, b, nColorBits,
            nLevel + 1, pLeafCount, pReducibleNodes);
    }
}

NODE* CQuantizer::CreateNode (UINT nLevel, UINT nColorBits, UINT* pLeafCount,
    NODE** pReducibleNodes)
{
    NODE* pNode;

    if ((pNode = (NODE*) HeapAlloc (GetProcessHeap (), HEAP_ZERO_MEMORY,
        sizeof (NODE))) == NULL)
        return NULL;

    pNode->bIsLeaf = (nLevel == nColorBits) ? TRUE : FALSE;
    if (pNode->bIsLeaf)
        (*pLeafCount)++;
    else {
        pNode->pNext = pReducibleNodes[nLevel];
        pReducibleNodes[nLevel] = pNode;
    }
    return pNode;
}

void CQuantizer::ReduceTree (UINT nColorBits, UINT* pLeafCount,
    NODE** pReducibleNodes)
{
    //
    // Find the deepest level containing at least one reducible node.
    //
    //for (i=nColorBits - 1; (i>0) && (pReducibleNodes[i] == NULL); i--);

    //Pavel Krejcir: some strange notation, I rewrite it as:
    int i = nColorBits - 1;
    while((i > 0) && (pReducibleNodes[i] == NULL)) i--;

    //
    // Reduce the node most recently added to the list at level i.
    //
    NODE* pNode = pReducibleNodes[i];
    pReducibleNodes[i] = pNode->pNext;

    UINT nRedSum = 0;
    UINT nGreenSum = 0;
    UINT nBlueSum = 0;
    UINT nChildren = 0;

    for (i=0; i<8; i++) {
        if (pNode->pChild[i] != NULL) {
            nRedSum += pNode->pChild[i]->nRedSum;
            nGreenSum += pNode->pChild[i]->nGreenSum;
            nBlueSum += pNode->pChild[i]->nBlueSum;
            pNode->nPixelCount += pNode->pChild[i]->nPixelCount;
            HeapFree (GetProcessHeap (), 0, pNode->pChild[i]);
            pNode->pChild[i] = NULL;
            nChildren++;
        }
    }

    pNode->bIsLeaf = TRUE;
    pNode->nRedSum = nRedSum;
    pNode->nGreenSum = nGreenSum;
    pNode->nBlueSum = nBlueSum;
    *pLeafCount -= (nChildren - 1);
}

void CQuantizer::DeleteTree (NODE** ppNode)
{
    for (int i=0; i<8; i++) {
        if ((*ppNode)->pChild[i] != NULL)
            DeleteTree (&((*ppNode)->pChild[i]));
    }
    HeapFree (GetProcessHeap (), 0, *ppNode);
    *ppNode = NULL;
}

void CQuantizer::GetPaletteColors (NODE* pTree, RGBQUAD* prgb, UINT* pIndex)
{
    if (pTree->bIsLeaf) {
        prgb[*pIndex].rgbRed =
            (BYTE) ((pTree->nRedSum) / (pTree->nPixelCount));
        prgb[*pIndex].rgbGreen =
            (BYTE) ((pTree->nGreenSum) / (pTree->nPixelCount));
        prgb[*pIndex].rgbBlue =
            (BYTE) ((pTree->nBlueSum) / (pTree->nPixelCount));
        prgb[*pIndex].rgbReserved = 0;
        (*pIndex)++;
    }
    else {
        for (int i=0; i<8; i++) {
            if (pTree->pChild[i] != NULL)
                GetPaletteColors (pTree->pChild[i], prgb, pIndex);
        }
    }
}

UINT CQuantizer::GetColorCount ()
{
    return m_nLeafCount;
}

void CQuantizer::GetColorTable (RGBQUAD* prgb)
{
    UINT nIndex = 0;
    GetPaletteColors (m_pTree, prgb, &nIndex);
}
