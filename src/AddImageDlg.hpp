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

#ifndef _ADDIMGDLG_HPP_
#define _ADDIMGDLG_HPP_

#include <windows.h>

#define AID_SIZE_48 1
#define AID_SIZE_32 2
#define AID_SIZE_24 4
#define AID_SIZE_16 8
#define AID_SIZE_256 16
#define AID_SIZE_CUSTOM 32
#define AID_DEPTH_4 1
#define AID_DEPTH_8 2
#define AID_DEPTH_RGBA 4
#define AID_DEPTH_PNG 8

typedef struct
{
    BYTE bSize; // mask containing one or more AID_SIZE_* flags
    BYTE bColorDepth; // mask containing one or more AID_DEPTH_* flags
    DWORD dwTranspColor;
    int iCustWidth; // width and height applicable if AID_SIZE_CUSTOM
    int iCustHeight; // flag is set
    BYTE bImgType; // 0 - unknown, 1 - PNG, 2 - BMP
    BOOL bDither;
    BOOL bOptimalPal;
    long lRawDataSize;
    BYTE *pbRawData;
} ADDIMGDATA, *PADDIMGDATA;

class CAddImageDlg
{
private:
    HINSTANCE m_hInstance;
    LPTSTR m_sFileName;
    BYTE *m_bData;
    HBITMAP m_Bmp;
    PADDIMGDATA m_AddImgData;
    BOOL m_bSelectCol, m_bHasMoved;
    POINT m_ptImageStartClick, m_ptImageMove;
    HDC m_ImageDC;
    int m_iImageROP2;

    INT_PTR CustomBtnClick(HWND hWnd, HWND hwndCtl);
    INT_PTR VistaBtnClick(HWND hWnd);
    INT_PTR XPBtnClick(HWND hWnd);
    void SetTransparentColor(HWND hWnd, HWND hwndImg);
    void LoadImage(HWND hWnd, HWND hwndCtl);

    INT_PTR OKBtnClick(HWND hWnd);
public:
    CAddImageDlg(HINSTANCE hInstance);
    ~CAddImageDlg();
    BOOL ShowModal(HWND hWnd, LPTSTR sFileName, PADDIMGDATA pAddImgData);

    // windows messages sections
    INT_PTR WMInitDialog(HWND hWnd, HWND hwndFocus, LPARAM lInitParam);
    INT_PTR WMCommand(HWND hWnd, WORD wNotifyCode, WORD wID, HWND hwndCtl);

    // image control messages
    LONG_PTR ImageWMSetCursor(HWND hWnd, HWND hwnd, WORD nHittest,
        WORD wMouseMsg);
    LONG_PTR ImageWMLButtonDown(HWND hWnd, WPARAM fwKeys, WORD xPos, WORD yPos);
    LONG_PTR ImageWMLButtonUp(HWND hWnd, WPARAM fwKeys, WORD xPos, WORD yPos);
    LONG_PTR ImageWMMouseMove(HWND hWnd, WPARAM fwKeys, WORD xPos, WORD yPos);
};

#endif
