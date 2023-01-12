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

#ifndef MAINFORM_HPP_
#define MAINFORM_HPP_

#include <windows.h>
#include <commctrl.h>
#include "Icons.h"
#include "AddImageDlg.hpp"
#include "Editor.hpp"
#include "HotSpotDlg.hpp"

class CMainForm
{
private:
  LPTSTR m_sAppName;
  HINSTANCE m_hInstance;
  BOOL m_bHasChanged;
  TCHAR m_sLastOpenPath[MAX_PATH], m_sCurFileName[MAX_PATH];
  UINT m_uiLBDragMsg;
  int m_iDragItem;
  double m_dOffsetX;
  double m_dOffsetY;
  CAddImageDlg *m_AddImgDlg;
  CEditImageDlg *m_EditImgDlg;
  CHotSpotDlg *m_HotSpotDlg;

  void RestoreSettings(HWND hWnd);
  void SaveSettings(HWND hWnd);
  BOOL ExitApp(HWND hWnd);
  INT_PTR OpenFile(HWND hWnd);
  void EnableMenuItems(HWND hWnd);
  void OpenIconFile(HWND hWnd, LPTSTR sFileName);
  BOOL CheckChangedFile(HWND hWnd);
  INT_PTR SaveFile(HWND hWnd);
  BOOL SaveFileAs(HWND hWnd);
  BOOL ReadIconFile(HWND hWnd, BOOL bClearLB, LPTSTR sFileName);
  void SaveIconFile(HWND hWnd);
  INT_PTR ListBoxChange(HWND hWnd, HWND hwndCtl);
  void ClearListBox(HWND hWnd);
  INT_PTR ListBoxDragMsg(HWND hWnd, int idCtl, LPDRAGLISTINFO pDragInfo);
  INT_PTR ListBoxBeginDrag(HWND hWnd, int idCtl, HWND hwndCtl, POINT ptCursor);
  INT_PTR ListBoxCancelDrag(HWND hWnd, int idCtl, HWND hwndCtl, POINT ptCursor);
  INT_PTR ListBoxDragging(HWND hWnd, int idCtl, HWND hwndCtl, POINT ptCursor);
  INT_PTR ListBoxDropped(HWND hWnd, int idCtl, HWND hwndCtl, POINT ptCursor);
  INT_PTR AddImage(HWND hWnd);
  INT_PTR NewFile(HWND hWnd);
  INT GetIcoFileType(LPTSTR sFileName); // 1 - ICO, 2 - CUR, 3 - Unknown
  INT_PTR ImageDelete(HWND hWnd);
  INT_PTR ImageEdit(HWND hWnd);
  INT_PTR ImageHotSpot(HWND hWnd);
  INT_PTR ImageSave(HWND hWnd);
  LPICONDATA IconDataFromAddImgData(PADDIMGDATA pAid, int iBitDepth, int iWidth, int iHeight);
  void InsertIconImages(HWND hWnd, PADDIMGDATA pAid);

public:
  CMainForm(HINSTANCE hInstance, DLGPROC lpDialogFunc, HWND *phWnd, HACCEL *phAccTable);
  ~CMainForm();

  // windows messages sections
  INT_PTR WMInitDialog(HWND hWnd, HWND hwndFocus, LPARAM lInitParam);
  INT_PTR WMCommand(HWND hWnd, WORD wNotifyCode, WORD wID, HWND hwndCtl);
  INT_PTR WMSysCommand(HWND hWnd, WPARAM uCmdType, WORD xPos, WORD yPos);
  INT_PTR WMDestroy(HWND hWnd);
  INT_PTR WMClose(HWND hWnd);
  INT_PTR WMUnknownMessage(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
};

#endif
