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

#ifndef _EDITOR_HPP_
#define _EDITOR_HPP_

#include <windows.h>
#include <commctrl.h>
#include "Icons.h"

class CEditImageDlg
{
private:
  HINSTANCE m_hInstance;
  HIMAGELIST m_hImgs;
  int m_iZoom;
  LPICONIMAGE m_IcoImg;
  LPBITMAPINFO m_pImage;
  int m_iSelColor; // non-negative number means item selected in the palette
    // -1 means transparent color selected, -2 means no palette item selected
  COLORREF m_cTransp;
  POINT m_ptDown, m_ptPalOrig, m_ptImgOrig;
  int m_iPalCount;
  BOOL m_bMouseDrag;
  BOOL m_bHasChanged;

  INT_PTR NFToolBarNeedText(HWND hWnd, LPTOOLTIPTEXT pttt);
  void SetupPalette(HWND hWnd);
  void SetupImage(HWND hWnd);
  void DrawPalette(HWND hWnd, HDC hdc);
  void DrawImage(HWND hWnd, HDC hdc);
  INT_PTR ZoomIn(HWND hWnd, HWND hwndCtl);
  INT_PTR ZoomOut(HWND hWnd, HWND hwndCtl);

  void LButtonClick(HWND hWnd);
  void SelImageClick(HWND hWnd);
  void SelPaletteClick(HWND hWnd);
  int IsPalettePoint(HWND hWnd, LPPOINT ppt); // returns the index of palette
    // entry, or -2 if outside of palette entries
  int IsImagePoint(HWND hWnd, LPPOINT ppt); // returns pixel position
    // of the image, or -1 if oustide of the image
public:
  CEditImageDlg(HINSTANCE hInstance);
  ~CEditImageDlg();
  BOOL ShowModal(HWND hWnd, LPICONIMAGE pIcoImg, int iSize);
  void SaveImage();

  // windows messages sections
  INT_PTR WMInitDialog(HWND hWnd, HWND hwndFocus, LPARAM lInitParam);
  INT_PTR WMCommand(HWND hWnd, WORD wNotifyCode, WORD wID, HWND hwndCtl);
  INT_PTR WMNotify(HWND hWnd, int idCtrl, LPNMHDR pnmh);
  INT_PTR WMPaint(HWND hWnd, HDC hdc);
  INT_PTR WMLButtonDown(HWND hWnd, WPARAM fwKeys, WORD xPos, WORD yPos);
  INT_PTR WMLButtonUp(HWND hWnd, WPARAM fwKeys, WORD xPos, WORD yPos);
  INT_PTR WMMouseMove(HWND hWnd, WPARAM fwKeys, WORD xPos, WORD yPos);
};

#endif
