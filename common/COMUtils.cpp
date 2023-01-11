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

#include "COMUtils.hpp"
#include <stdio.h>

LPWSTR StrToWideChar(LPCSTR AStr)
{
    int slen = strlen(AStr);
    wchar_t *res = (wchar_t*)malloc((slen + 1)*sizeof(wchar_t));
    MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, AStr, -1, res, slen + 1);
    return(res);
}

LPSTR WideCharToStr(LPCWSTR AStr)
{
    int slen = wcslen(AStr);
    char *res = (char*)malloc((slen + 1)*sizeof(char));
    WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, AStr, -1, res, slen + 1,
        NULL, NULL);
    return(res);
}

BSTR StrToBSTR(LPCSTR AStr)
{
    int slen = strlen(AStr);
    BSTR res = SysAllocStringLen(NULL, slen);
    MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, AStr, -1, res, slen + 1);
    return(res);
}

int VarToInt(VARIANT V)
{
    switch(V.vt)
    {
    case VT_I4: return(V.lVal);
    case VT_UI1: return(V.bVal);
    case VT_I2: return(V.iVal);
    case VT_BOOL: return(V.boolVal);
    case VT_I1: return(V.cVal);
    case VT_UI2: return(V.uiVal);
    case VT_UI4: return(V.ulVal);
    case VT_INT: return(V.intVal);
    case VT_UINT: return(V.uintVal);
    default: return(-1);
    }
}

BOOL VarToBool(VARIANT V)
{
    switch(V.vt)
    {
    case VT_I4: return(V.lVal);
    case VT_UI1: return(V.bVal);
    case VT_I2: return(V.iVal);
    case VT_BOOL: return(V.boolVal);
    case VT_I1: return(V.cVal);
    case VT_UI2: return(V.uiVal);
    case VT_UI4: return(V.ulVal);
    case VT_INT: return(V.intVal);
    case VT_UINT: return(V.uintVal);
    default: return(FALSE);
    }
}

int VarStrLen(VARIANT V)
{
    switch(V.vt)
    {
    case VT_BSTR:
        return(wcslen(V.bstrVal));
    default:
        return(32);
    }
}

BOOL VarToChar(VARIANT V, char *buf, int bufsize)
{
    switch(V.vt)
    {
    case VT_I4:
        sprintf(buf, "%d", V.lVal);
        return(TRUE);
    case VT_UI1:
        sprintf(buf, "%d", V.bVal);
        return(TRUE);
    case VT_I2:
        sprintf(buf, "%d", V.iVal);
        return(TRUE);
    case VT_BOOL:
        sprintf(buf, "%d", V.boolVal);
        return(TRUE);
    case VT_I1:
        sprintf(buf, "%d", V.cVal);
        return(TRUE);
    case VT_UI2:
        sprintf(buf, "%d", V.uiVal);
        return(TRUE);
    case VT_UI4:
        sprintf(buf, "%d", V.ulVal);
        return(TRUE);
    case VT_INT:
        sprintf(buf, "%d", V.intVal);
        return(TRUE);
    case VT_UINT:
        sprintf(buf, "%d", V.uintVal);
        return(TRUE);
    case VT_R4:
        sprintf(buf, "%.4f", V.fltVal);
        return(TRUE);
    case VT_R8:
        sprintf(buf, "%.4f", V.dblVal);
        return(TRUE);
    case VT_BSTR:
        return(WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, V.bstrVal, -1,
            buf, bufsize, NULL, NULL));
    case VT_ARRAY | VT_UI1:
        strcpy(buf, "(Blob)");
        return(TRUE);
    default:
        buf[0] = 0;
        return(FALSE);
    }
}

BOOL VarToChar(VARIANT V, wchar_t *buf, int bufsize)
{
    int i;
    switch(V.vt)
    {
    case VT_I4:
        swprintf(buf, L"%d", V.lVal);
        return(TRUE);
    case VT_UI1:
        swprintf(buf, L"%d", V.bVal);
        return(TRUE);
    case VT_I2:
        swprintf(buf, L"%d", V.iVal);
        return(TRUE);
    case VT_BOOL:
        swprintf(buf, L"%d", V.boolVal);
        return(TRUE);
    case VT_I1:
        swprintf(buf, L"%d", V.cVal);
        return(TRUE);
    case VT_UI2:
        swprintf(buf, L"%d", V.uiVal);
        return(TRUE);
    case VT_UI4:
        swprintf(buf, L"%d", V.ulVal);
        return(TRUE);
    case VT_INT:
        swprintf(buf, L"%d", V.intVal);
        return(TRUE);
    case VT_UINT:
        swprintf(buf, L"%d", V.uintVal);
        return(TRUE);
    case VT_R4:
        swprintf(buf, L"%.4f", V.fltVal);
        return(TRUE);
    case VT_R8:
        swprintf(buf, L"%.4f", V.dblVal);
        return(TRUE);
    case VT_BSTR:
        i = wcslen(V.bstrVal);
        if(i < bufsize)
        {
            wcscpy(buf, V.bstrVal);
            return(TRUE);
        }
        else return(FALSE);
    case VT_ARRAY | VT_UI1:
        wcscpy(buf, L"(Blob)");
        return(TRUE);
    default:
        buf[0] = 0;
        return(FALSE);
    }
}

LPTSTR LoadGResString(HINSTANCE hInstance, UINT uID)
{
    long nm = uID / 16;
    HRSRC rh = FindResource(hInstance, MAKEINTRESOURCE(nm + 1), RT_STRING);
    HGLOBAL hg = LoadResource(hInstance, rh);
    wchar_t *P = (wchar_t*)LockResource(hg);
    nm = uID % 16;
    int j = 0;
    int i;
    LPTSTR res = NULL;

    while(nm > 0)
    {
        i = (WORD)P[j];
        j += (i + 1);
        nm--;
    }

    i = (WORD)P[j];
    if(i > 0)
    {
        res = (LPTSTR)malloc((i + 1)*sizeof(TCHAR));
#ifdef UNICODE
        wcsncpy(res, &P[j + 1], i);
#else
        WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, &P[j + 1], i, res,
            i + 1, NULL, NULL);
#endif
        res[i] = 0;
    }
    UnlockResource(hg);
    return(res);
};

void FreeTString(LPTSTR *ABuf)
{
#ifdef UNICODE
    SysFreeString(*ABuf);
#else
    free(*ABuf);
#endif
    *ABuf = NULL;
    return;
}
