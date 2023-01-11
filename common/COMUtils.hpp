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

#ifndef _COMUTILS_HPP_
#define _COMUTILS_HPP_

#include <basetyps.h>
#include <wtypes.h>

LPWSTR StrToWideChar(LPCSTR AStr);
LPSTR WideCharToStr(LPCWSTR AStr);
BSTR StrToBSTR(LPCSTR AStr);
int VarToInt(VARIANT V);
BOOL VarToBool(VARIANT V);
int VarStrLen(VARIANT V);
BOOL VarToChar(VARIANT V, char *buf, int bufsize);
BOOL VarToChar(VARIANT V, wchar_t *buf, int bufsize);
LPTSTR LoadGResString(HINSTANCE hInstance, UINT uID);
void FreeTString(LPTSTR *ABuf);

#endif
