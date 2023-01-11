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

//****** NOTE ********
// this quantization code has been taken from a PalGen Microsoft's sample
// however, it is not used as the defautl option for image simplification
// The images are converted into less bits using standard web-safe palettes
// and dithering

#ifndef _QUANTIZE_H_
#define _QUANTIZE_H_

#include <windows.h>

typedef struct _NODE {
    BOOL bIsLeaf;               // TRUE if node has no children
    UINT nPixelCount;           // Number of pixels represented by this leaf
    UINT nRedSum;               // Sum of red components
    UINT nGreenSum;             // Sum of green components
    UINT nBlueSum;              // Sum of blue components
    struct _NODE* pChild[8];    // Pointers to child nodes
    struct _NODE* pNext;        // Pointer to next reducible node
} NODE;

class CQuantizer
{
protected:
    NODE* m_pTree;
    UINT m_nLeafCount;
    NODE* m_pReducibleNodes[9];
    UINT m_nMaxColors;
    UINT m_nColorBits;

public:
    CQuantizer (UINT nMaxColors, UINT nColorBits);
    virtual ~CQuantizer ();
    BOOL ProcessImage (HANDLE hImage);
    BOOL ProcessBitmap(HBITMAP hBmp, int iWidth, int iHeight);
    BOOL ProcessAlphaImg(BYTE *pImg, int iWidth, int iHeight);
    UINT GetColorCount ();
    void GetColorTable (RGBQUAD* prgb);

protected:
    int GetLeftShiftCount (DWORD dwVal);
    int GetRightShiftCount (DWORD dwVal);
    void AddColor (NODE** ppNode, BYTE r, BYTE g, BYTE b, UINT nColorBits,
        UINT nLevel, UINT* pLeafCount, NODE** pReducibleNodes);
    NODE* CreateNode (UINT nLevel, UINT nColorBits, UINT* pLeafCount,
        NODE** pReducibleNodes);
    void ReduceTree (UINT nColorBits, UINT* pLeafCount,
        NODE** pReducibleNodes);
    void DeleteTree (NODE** ppNode);
    void GetPaletteColors (NODE* pTree, RGBQUAD* prgb, UINT* pIndex);
};

#endif
