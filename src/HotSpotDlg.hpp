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

#ifndef _HOTSPOTDLG_HPP_
#define _HOTSPOTDLG_HPP_

#include <windows.h>

class CHotSpotDlg
{
private:
  HINSTANCE m_hInstance;
  INT_PTR OKBtnClick(HWND hWnd);
  double *m_pdOffsetX;
  double *m_pdOffsetY;
public:
  CHotSpotDlg(HINSTANCE hInstance);
  ~CHotSpotDlg();
  BOOL ShowModal(HWND hWnd, double *pdOffsetX, double *pdOffsetY);

  // windows messages sections
  INT_PTR WMInitDialog(HWND hWnd, HWND hwndFocus, LPARAM lInitParam);
  INT_PTR WMCommand(HWND hWnd, WORD wNotifyCode, WORD wID, HWND hwndCtl);
};

#endif
