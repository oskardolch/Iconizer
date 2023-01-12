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

#include "HotSpotDlg.hpp"
#include <tchar.h>
#include <stdio.h>
#include "Iconizer.rh"

INT_PTR CALLBACK HotSpotDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  CHotSpotDlg *dlg = NULL;
  if(uMsg == WM_INITDIALOG) dlg = (CHotSpotDlg*)lParam;
  else dlg = (CHotSpotDlg*)GetWindowLongPtr(hwndDlg, DWLP_USER);

  switch(uMsg)
  {
  case WM_INITDIALOG:
    return dlg->WMInitDialog(hwndDlg, (HWND)wParam, lParam);
  case WM_COMMAND:
    return dlg->WMCommand(hwndDlg, HIWORD(wParam), LOWORD(wParam), (HWND)lParam);
  default:
    return 0;
  }
}

CHotSpotDlg::CHotSpotDlg(HINSTANCE hInstance)
{
  m_hInstance = hInstance;
}

CHotSpotDlg::~CHotSpotDlg()
{
}

BOOL CHotSpotDlg::ShowModal(HWND hWnd, double *pdOffsetX, double *pdOffsetY)
{
  m_pdOffsetX = pdOffsetX;
  m_pdOffsetY = pdOffsetY;
  return DialogBoxParam(m_hInstance, _T("HOTSPOTDLG"), hWnd, HotSpotDlgProc, (LPARAM)this) == IDOK;
}

INT_PTR CHotSpotDlg::WMInitDialog(HWND hWnd, HWND hwndFocus, LPARAM lInitParam)
{
  SetWindowLongPtr(hWnd, DWLP_USER, lInitParam);
  TCHAR buf[64];

  _stprintf(buf, _T("%f"), *m_pdOffsetX);
  HWND wnd = GetDlgItem(hWnd, HOTSPOTDLG_CTLR_EDTOFFSETX);
  SendMessage(wnd, WM_SETTEXT, 0, (LPARAM)buf);

  _stprintf(buf, _T("%f"), *m_pdOffsetY);
  wnd = GetDlgItem(hWnd, HOTSPOTDLG_CTLR_EDTOFFSETY);
  SendMessage(wnd, WM_SETTEXT, 0, (LPARAM)buf);
  return 1;
}

INT_PTR CHotSpotDlg::WMCommand(HWND hWnd, WORD wNotifyCode, WORD wID, HWND hwndCtl)
{
  switch(wID)
  {
  case IDOK:
    return OKBtnClick(hWnd);
  case IDCANCEL:
    EndDialog(hWnd, 0);
    return 1;
  default:
    return 0;
  }
}

INT_PTR CHotSpotDlg::OKBtnClick(HWND hWnd)
{
  float fOffsetX, fOffsetY;
  HWND wnd = GetDlgItem(hWnd, HOTSPOTDLG_CTLR_EDTOFFSETX);
  TCHAR buf[64];
  TCHAR errmsg[128];
  SendMessage(wnd, WM_GETTEXT, 63, (LPARAM)buf);
  if(_stscanf(buf, _T("%f"), &fOffsetX) != 1 || fOffsetX < 0.0 || fOffsetX > 100.0)
  {
    LoadString(m_hInstance, IDS_DLG_ERR, buf, 16);
    LoadString(m_hInstance, IDS_MSG_INVALIDUSERINPUT, errmsg, 128);
    MessageBox(hWnd, errmsg, buf, MB_OK | MB_ICONERROR);
    SendMessage(wnd, EM_SETSEL, 0, (LPARAM)-1);
    SetFocus(wnd);
    return 1;
  }
  wnd = GetDlgItem(hWnd, HOTSPOTDLG_CTLR_EDTOFFSETY);
  SendMessage(wnd, WM_GETTEXT, 63, (LPARAM)buf);
  if(_stscanf(buf, _T("%f"), &fOffsetY) != 1 || fOffsetY < 0.0 || fOffsetY > 100.0)
  {
    LoadString(m_hInstance, IDS_DLG_ERR, buf, 16);
    LoadString(m_hInstance, IDS_MSG_INVALIDUSERINPUT, errmsg, 128);
    MessageBox(hWnd, errmsg, buf, MB_OK | MB_ICONERROR);
    SendMessage(wnd, EM_SETSEL, 0, (LPARAM)-1);
    SetFocus(wnd);
    return 1;
  }
  *m_pdOffsetX = fOffsetX;
  *m_pdOffsetY = fOffsetY;
  EndDialog(hWnd, 1);
  return 1;
}
