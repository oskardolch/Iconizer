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

#include "msxml_DLL.hpp"

static DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
static DISPID DispIdPropPut = DISPID_PROPERTYPUT;

BOOL IXMLDOMImplementationDisp::hasFeature(LPTSTR feature, LPTSTR version)
{
  VARIANT pVarResult;
  int i;
  BSTR Buf[2];
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf[0] = SysAllocString(version);
#else
  Buf[0] = SysAllocStringLen(NULL, strlen(version));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, version, -1, Buf[0], strlen(version) + 1);
  params[0].bstrVal = Buf[0];
#endif
  params[1].vt = VT_BSTR;
#ifdef UNICODE
  Buf[1] = SysAllocString(feature);
#else
  Buf[1] = SysAllocStringLen(NULL, strlen(feature));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, feature, -1, Buf[1], strlen(feature) + 1);
  params[1].bstrVal = Buf[1];
#endif
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000091, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  for(i = 0; i < 2; i++) SysFreeString(Buf[i]);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMNodeDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMNodeDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMNodeDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMNodeDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMNodeDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMNodeDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNodeDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNodeDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNodeDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNodeDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMNodeDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNodeDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNodeDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNodeDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNodeDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMNodeDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMNodeDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNodeDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMNodeDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMNodeDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMNodeDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMNodeDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMNodeDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMNodeDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMNodeDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMNodeDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMNodeDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMNodeDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMNodeDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMNodeDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNodeDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMNodeDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMNodeDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMNodeDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMNodeDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMNodeDisp::transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

IXMLDOMNode* IXMLDOMNodeListDisp::get_item(long index)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_I4;
  params.lVal = index;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000000, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

long IXMLDOMNodeListDisp::get_length()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000004A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMNodeListDisp::nextNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000004C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

void IXMLDOMNodeListDisp::reset()
{
  Invoke((DISPID)0x0000004D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, NULL, NULL, NULL);
};

IUnknown * IXMLDOMNodeListDisp::get__newEnum()
{
  VARIANT pVarResult;
  Invoke((DISPID)0xFFFFFFFC, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.punkVal);
};

IXMLDOMNode* IXMLDOMNamedNodeMapDisp::getNamedItem(LPTSTR name)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(name);
#else
  Buf = SysAllocStringLen(NULL, strlen(name));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, name, -1, Buf, strlen(name) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000053, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNamedNodeMapDisp::setNamedItem(IXMLDOMNode* newItem)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newItem;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000054, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNamedNodeMapDisp::removeNamedItem(LPTSTR name)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(name);
#else
  Buf = SysAllocStringLen(NULL, strlen(name));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, name, -1, Buf, strlen(name) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000055, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNamedNodeMapDisp::get_item(long index)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_I4;
  params.lVal = index;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000000, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

long IXMLDOMNamedNodeMapDisp::get_length()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000004A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMNamedNodeMapDisp::getQualifiedItem(LPTSTR baseName, LPTSTR namespaceURI)
{
  VARIANT pVarResult;
  int i;
  BSTR Buf[2];
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf[0] = SysAllocString(namespaceURI);
#else
  Buf[0] = SysAllocStringLen(NULL, strlen(namespaceURI));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, namespaceURI, -1, Buf[0], strlen(namespaceURI) + 1);
  params[0].bstrVal = Buf[0];
#endif
  params[1].vt = VT_BSTR;
#ifdef UNICODE
  Buf[1] = SysAllocString(baseName);
#else
  Buf[1] = SysAllocStringLen(NULL, strlen(baseName));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, baseName, -1, Buf[1], strlen(baseName) + 1);
  params[1].bstrVal = Buf[1];
#endif
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000057, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  for(i = 0; i < 2; i++) SysFreeString(Buf[i]);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNamedNodeMapDisp::removeQualifiedItem(LPTSTR baseName, LPTSTR namespaceURI)
{
  VARIANT pVarResult;
  int i;
  BSTR Buf[2];
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf[0] = SysAllocString(namespaceURI);
#else
  Buf[0] = SysAllocStringLen(NULL, strlen(namespaceURI));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, namespaceURI, -1, Buf[0], strlen(namespaceURI) + 1);
  params[0].bstrVal = Buf[0];
#endif
  params[1].vt = VT_BSTR;
#ifdef UNICODE
  Buf[1] = SysAllocString(baseName);
#else
  Buf[1] = SysAllocStringLen(NULL, strlen(baseName));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, baseName, -1, Buf[1], strlen(baseName) + 1);
  params[1].bstrVal = Buf[1];
#endif
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000058, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  for(i = 0; i < 2; i++) SysFreeString(Buf[i]);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNamedNodeMapDisp::nextNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000059, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

void IXMLDOMNamedNodeMapDisp::reset()
{
  Invoke((DISPID)0x0000005A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, NULL, NULL, NULL);
};

IUnknown * IXMLDOMNamedNodeMapDisp::get__newEnum()
{
  VARIANT pVarResult;
  Invoke((DISPID)0xFFFFFFFC, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.punkVal);
};

LPTSTR IXMLDOMDocumentDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMDocumentDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMDocumentDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMDocumentDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMDocumentDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMDocumentDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMDocumentDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMDocumentDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMDocumentDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMDocumentDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMDocumentDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMDocumentDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMDocumentDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMDocumentDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMDocumentDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMDocumentDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMDocumentDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMDocumentDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMDocumentDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMDocumentDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMDocumentDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMDocumentDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMDocumentDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMDocumentDisp::transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

IXMLDOMDocumentType* IXMLDOMDocumentDisp::get_doctype()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000026, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocumentType*)pVarResult.pdispVal);
};

IXMLDOMImplementation* IXMLDOMDocumentDisp::get_implementation()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000027, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMImplementation*)pVarResult.pdispVal);
};

IXMLDOMElement* IXMLDOMDocumentDisp::get_documentElement()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000028, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMElement*)pVarResult.pdispVal);
};

void IXMLDOMDocumentDisp::set_documentElement(IXMLDOMElement* Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)Value;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000028, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

IXMLDOMElement* IXMLDOMDocumentDisp::createElement(LPTSTR tagName)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(tagName);
#else
  Buf = SysAllocStringLen(NULL, strlen(tagName));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, tagName, -1, Buf, strlen(tagName) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000029, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMElement*)pVarResult.pdispVal);
};

IXMLDOMDocumentFragment* IXMLDOMDocumentDisp::createDocumentFragment()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000002A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocumentFragment*)pVarResult.pdispVal);
};

IXMLDOMText* IXMLDOMDocumentDisp::createTextNode(LPTSTR data)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000002B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMText*)pVarResult.pdispVal);
};

IXMLDOMComment* IXMLDOMDocumentDisp::createComment(LPTSTR data)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000002C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMComment*)pVarResult.pdispVal);
};

IXMLDOMCDATASection* IXMLDOMDocumentDisp::createCDATASection(LPTSTR data)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000002D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMCDATASection*)pVarResult.pdispVal);
};

IXMLDOMProcessingInstruction* IXMLDOMDocumentDisp::createProcessingInstruction(LPTSTR target,
  LPTSTR data)
{
  VARIANT pVarResult;
  int i;
  BSTR Buf[2];
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf[0] = SysAllocString(data);
#else
  Buf[0] = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf[0], strlen(data) + 1);
  params[0].bstrVal = Buf[0];
#endif
  params[1].vt = VT_BSTR;
#ifdef UNICODE
  Buf[1] = SysAllocString(target);
#else
  Buf[1] = SysAllocStringLen(NULL, strlen(target));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, target, -1, Buf[1], strlen(target) + 1);
  params[1].bstrVal = Buf[1];
#endif
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000002E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  for(i = 0; i < 2; i++) SysFreeString(Buf[i]);
  return((IXMLDOMProcessingInstruction*)pVarResult.pdispVal);
};

IXMLDOMAttribute* IXMLDOMDocumentDisp::createAttribute(LPTSTR name)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(name);
#else
  Buf = SysAllocStringLen(NULL, strlen(name));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, name, -1, Buf, strlen(name) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000002F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMAttribute*)pVarResult.pdispVal);
};

IXMLDOMEntityReference* IXMLDOMDocumentDisp::createEntityReference(LPTSTR name)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(name);
#else
  Buf = SysAllocStringLen(NULL, strlen(name));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, name, -1, Buf, strlen(name) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000031, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMEntityReference*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMDocumentDisp::getElementsByTagName(LPTSTR tagName)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(tagName);
#else
  Buf = SysAllocStringLen(NULL, strlen(tagName));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, tagName, -1, Buf, strlen(tagName) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000032, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::createNode(VARIANT type, LPTSTR name, LPTSTR namespaceURI)
{
  VARIANT pVarResult;
  int i;
  BSTR Buf[2];
  VARIANT params[3];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf[0] = SysAllocString(namespaceURI);
#else
  Buf[0] = SysAllocStringLen(NULL, strlen(namespaceURI));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, namespaceURI, -1, Buf[0], strlen(namespaceURI) + 1);
  params[0].bstrVal = Buf[0];
#endif
  params[1].vt = VT_BSTR;
#ifdef UNICODE
  Buf[1] = SysAllocString(name);
#else
  Buf[1] = SysAllocStringLen(NULL, strlen(name));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, name, -1, Buf[1], strlen(name) + 1);
  params[1].bstrVal = Buf[1];
#endif
  params[2].vt = type.vt;
  if(params[2].vt == VT_ERROR) params[2].scode = DISP_E_PARAMNOTFOUND;
  else params[2].dblVal = type.dblVal;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 3;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000036, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  for(i = 0; i < 2; i++) SysFreeString(Buf[i]);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentDisp::nodeFromID(LPTSTR idString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(idString);
#else
  Buf = SysAllocStringLen(NULL, strlen(idString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, idString, -1, Buf, strlen(idString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000038, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMDocumentDisp::load(VARIANT xmlSource)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = xmlSource.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = xmlSource.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000003A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

long IXMLDOMDocumentDisp::get_readyState()
{
  VARIANT pVarResult;
  Invoke((DISPID)0xFFFFFDF3, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMParseError* IXMLDOMDocumentDisp::get_parseError()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000003B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMParseError*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMDocumentDisp::get_url()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000003C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

BOOL IXMLDOMDocumentDisp::get_async()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000003D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

void IXMLDOMDocumentDisp::set_async(BOOL Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = Value;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000003D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

void IXMLDOMDocumentDisp::abort()
{
  Invoke((DISPID)0x0000003E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, NULL, NULL, NULL);
};

BOOL IXMLDOMDocumentDisp::loadXML(LPTSTR bstrXML)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(bstrXML);
#else
  Buf = SysAllocStringLen(NULL, strlen(bstrXML));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, bstrXML, -1, Buf, strlen(bstrXML) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000003F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return(pVarResult.boolVal);
};

void IXMLDOMDocumentDisp::save(VARIANT destination)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = destination.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = destination.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000040, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

BOOL IXMLDOMDocumentDisp::get_validateOnParse()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000041, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

void IXMLDOMDocumentDisp::set_validateOnParse(BOOL Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = Value;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000041, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

BOOL IXMLDOMDocumentDisp::get_resolveExternals()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000042, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

void IXMLDOMDocumentDisp::set_resolveExternals(BOOL Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = Value;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000042, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

BOOL IXMLDOMDocumentDisp::get_preserveWhiteSpace()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000043, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

void IXMLDOMDocumentDisp::set_preserveWhiteSpace(BOOL Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = Value;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000043, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

void IXMLDOMDocumentDisp::set_onreadystatechange(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000044, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

void IXMLDOMDocumentDisp::set_ondataavailable(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000045, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

void IXMLDOMDocumentDisp::set_ontransformnode(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000046, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMDocumentTypeDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMDocumentTypeDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMDocumentTypeDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMDocumentTypeDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMDocumentTypeDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMDocumentTypeDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentTypeDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentTypeDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentTypeDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentTypeDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMDocumentTypeDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentTypeDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentTypeDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentTypeDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentTypeDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMDocumentTypeDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMDocumentTypeDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentTypeDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMDocumentTypeDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMDocumentTypeDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMDocumentTypeDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMDocumentTypeDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMDocumentTypeDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMDocumentTypeDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMDocumentTypeDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMDocumentTypeDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMDocumentTypeDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMDocumentTypeDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMDocumentTypeDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMDocumentTypeDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentTypeDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMDocumentTypeDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMDocumentTypeDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMDocumentTypeDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMDocumentTypeDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMDocumentTypeDisp::transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMDocumentTypeDisp::get_name()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000083, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNamedNodeMap* IXMLDOMDocumentTypeDisp::get_entities()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000084, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMDocumentTypeDisp::get_notations()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000085, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMElementDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMElementDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMElementDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMElementDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMElementDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMElementDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMElementDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMElementDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMElementDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMElementDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMElementDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMElementDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMElementDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMElementDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMElementDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMElementDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMElementDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMElementDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMElementDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMElementDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMElementDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMElementDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMElementDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMElementDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMElementDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMElementDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMElementDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMElementDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMElementDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMElementDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMElementDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMElementDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMElementDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMElementDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMElementDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMElementDisp::transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMElementDisp::get_tagName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000061, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMElementDisp::getAttribute(LPTSTR name)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(name);
#else
  Buf = SysAllocStringLen(NULL, strlen(name));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, name, -1, Buf, strlen(name) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000063, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return(pVarResult);
};

void IXMLDOMElementDisp::setAttribute(LPTSTR name, VARIANT value)
{
  BSTR Buf;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = value.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = value.dblVal;
  params[1].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(name);
#else
  Buf = SysAllocStringLen(NULL, strlen(name));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, name, -1, Buf, strlen(name) + 1);
  params[1].bstrVal = Buf;
#endif
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000064, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

void IXMLDOMElementDisp::removeAttribute(LPTSTR name)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(name);
#else
  Buf = SysAllocStringLen(NULL, strlen(name));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, name, -1, Buf, strlen(name) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000065, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

IXMLDOMAttribute* IXMLDOMElementDisp::getAttributeNode(LPTSTR name)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(name);
#else
  Buf = SysAllocStringLen(NULL, strlen(name));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, name, -1, Buf, strlen(name) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000066, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMAttribute*)pVarResult.pdispVal);
};

IXMLDOMAttribute* IXMLDOMElementDisp::setAttributeNode(IXMLDOMAttribute* DOMAttribute)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)DOMAttribute;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000067, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMAttribute*)pVarResult.pdispVal);
};

IXMLDOMAttribute* IXMLDOMElementDisp::removeAttributeNode(IXMLDOMAttribute* DOMAttribute)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)DOMAttribute;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000068, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMAttribute*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMElementDisp::getElementsByTagName(LPTSTR tagName)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(tagName);
#else
  Buf = SysAllocStringLen(NULL, strlen(tagName));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, tagName, -1, Buf, strlen(tagName) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000069, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

void IXMLDOMElementDisp::normalize()
{
  Invoke((DISPID)0x0000006A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, NULL, NULL, NULL);
};

LPTSTR IXMLDOMAttributeDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMAttributeDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMAttributeDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMAttributeDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMAttributeDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMAttributeDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMAttributeDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMAttributeDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMAttributeDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMAttributeDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMAttributeDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMAttributeDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMAttributeDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMAttributeDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMAttributeDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMAttributeDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMAttributeDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMAttributeDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMAttributeDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMAttributeDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMAttributeDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMAttributeDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMAttributeDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMAttributeDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMAttributeDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMAttributeDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMAttributeDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMAttributeDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMAttributeDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMAttributeDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMAttributeDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMAttributeDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMAttributeDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMAttributeDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMAttributeDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMAttributeDisp::transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMAttributeDisp::get_name()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000076, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMAttributeDisp::get_value()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000078, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMAttributeDisp::set_value(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000078, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMDocumentFragmentDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMDocumentFragmentDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMDocumentFragmentDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMDocumentFragmentDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMDocumentFragmentDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMDocumentFragmentDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentFragmentDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentFragmentDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentFragmentDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentFragmentDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMDocumentFragmentDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentFragmentDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentFragmentDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentFragmentDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentFragmentDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMDocumentFragmentDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMDocumentFragmentDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentFragmentDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMDocumentFragmentDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMDocumentFragmentDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMDocumentFragmentDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMDocumentFragmentDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMDocumentFragmentDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMDocumentFragmentDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMDocumentFragmentDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMDocumentFragmentDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMDocumentFragmentDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMDocumentFragmentDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMDocumentFragmentDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMDocumentFragmentDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMDocumentFragmentDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMDocumentFragmentDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMDocumentFragmentDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMDocumentFragmentDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMDocumentFragmentDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMDocumentFragmentDisp::transformNodeToObject(IXMLDOMNode* stylesheet,
  VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMCharacterDataDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMCharacterDataDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMCharacterDataDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMCharacterDataDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMCharacterDataDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMCharacterDataDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCharacterDataDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCharacterDataDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCharacterDataDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCharacterDataDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMCharacterDataDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCharacterDataDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCharacterDataDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCharacterDataDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCharacterDataDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMCharacterDataDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMCharacterDataDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCharacterDataDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMCharacterDataDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMCharacterDataDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMCharacterDataDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMCharacterDataDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMCharacterDataDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMCharacterDataDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMCharacterDataDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMCharacterDataDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMCharacterDataDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMCharacterDataDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMCharacterDataDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMCharacterDataDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCharacterDataDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMCharacterDataDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMCharacterDataDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMCharacterDataDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMCharacterDataDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMCharacterDataDisp::transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMCharacterDataDisp::get_data()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000006D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMCharacterDataDisp::set_data(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000006D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

long IXMLDOMCharacterDataDisp::get_length()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000006E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

LPTSTR IXMLDOMCharacterDataDisp::substringData(long offset, long count)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_I4;
  params[0].lVal = count;
  params[1].vt = VT_I4;
  params[1].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000006F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMCharacterDataDisp::appendData(LPTSTR data)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000070, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

void IXMLDOMCharacterDataDisp::insertData(long offset, LPTSTR data)
{
  BSTR Buf;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params[0].bstrVal = Buf;
#endif
  params[1].vt = VT_I4;
  params[1].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000071, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

void IXMLDOMCharacterDataDisp::deleteData(long offset, long count)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_I4;
  params[0].lVal = count;
  params[1].vt = VT_I4;
  params[1].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000072, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

void IXMLDOMCharacterDataDisp::replaceData(long offset, long count, LPTSTR data)
{
  BSTR Buf;
  VARIANT params[3];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params[0].bstrVal = Buf;
#endif
  params[1].vt = VT_I4;
  params[1].lVal = count;
  params[2].vt = VT_I4;
  params[2].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 3;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000073, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

LPTSTR IXMLDOMTextDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMTextDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMTextDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMTextDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMTextDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMTextDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMTextDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMTextDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMTextDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMTextDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMTextDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMTextDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMTextDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMTextDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMTextDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMTextDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMTextDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMTextDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMTextDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMTextDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMTextDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMTextDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMTextDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMTextDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMTextDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMTextDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMTextDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMTextDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMTextDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMTextDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMTextDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMTextDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMTextDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMTextDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMTextDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMTextDisp::transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMTextDisp::get_data()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000006D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMTextDisp::set_data(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000006D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

long IXMLDOMTextDisp::get_length()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000006E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

LPTSTR IXMLDOMTextDisp::substringData(long offset, long count)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_I4;
  params[0].lVal = count;
  params[1].vt = VT_I4;
  params[1].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000006F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMTextDisp::appendData(LPTSTR data)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000070, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

void IXMLDOMTextDisp::insertData(long offset, LPTSTR data)
{
  BSTR Buf;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params[0].bstrVal = Buf;
#endif
  params[1].vt = VT_I4;
  params[1].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000071, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

void IXMLDOMTextDisp::deleteData(long offset, long count)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_I4;
  params[0].lVal = count;
  params[1].vt = VT_I4;
  params[1].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000072, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

void IXMLDOMTextDisp::replaceData(long offset, long count, LPTSTR data)
{
  BSTR Buf;
  VARIANT params[3];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params[0].bstrVal = Buf;
#endif
  params[1].vt = VT_I4;
  params[1].lVal = count;
  params[2].vt = VT_I4;
  params[2].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 3;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000073, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

IXMLDOMText* IXMLDOMTextDisp::splitText(long offset)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_I4;
  params.lVal = offset;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000007B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMText*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMCommentDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMCommentDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMCommentDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMCommentDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMCommentDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMCommentDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCommentDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCommentDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCommentDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCommentDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMCommentDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCommentDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCommentDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCommentDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCommentDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMCommentDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMCommentDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCommentDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMCommentDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMCommentDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMCommentDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMCommentDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMCommentDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMCommentDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMCommentDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMCommentDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMCommentDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMCommentDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMCommentDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMCommentDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCommentDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMCommentDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMCommentDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMCommentDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMCommentDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMCommentDisp::transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMCommentDisp::get_data()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000006D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMCommentDisp::set_data(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000006D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

long IXMLDOMCommentDisp::get_length()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000006E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

LPTSTR IXMLDOMCommentDisp::substringData(long offset, long count)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_I4;
  params[0].lVal = count;
  params[1].vt = VT_I4;
  params[1].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000006F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMCommentDisp::appendData(LPTSTR data)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000070, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

void IXMLDOMCommentDisp::insertData(long offset, LPTSTR data)
{
  BSTR Buf;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params[0].bstrVal = Buf;
#endif
  params[1].vt = VT_I4;
  params[1].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000071, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

void IXMLDOMCommentDisp::deleteData(long offset, long count)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_I4;
  params[0].lVal = count;
  params[1].vt = VT_I4;
  params[1].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000072, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

void IXMLDOMCommentDisp::replaceData(long offset, long count, LPTSTR data)
{
  BSTR Buf;
  VARIANT params[3];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params[0].bstrVal = Buf;
#endif
  params[1].vt = VT_I4;
  params[1].lVal = count;
  params[2].vt = VT_I4;
  params[2].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 3;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000073, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

LPTSTR IXMLDOMCDATASectionDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMCDATASectionDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMCDATASectionDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMCDATASectionDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMCDATASectionDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMCDATASectionDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCDATASectionDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCDATASectionDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCDATASectionDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCDATASectionDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMCDATASectionDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCDATASectionDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCDATASectionDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCDATASectionDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCDATASectionDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMCDATASectionDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMCDATASectionDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCDATASectionDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMCDATASectionDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMCDATASectionDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMCDATASectionDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMCDATASectionDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMCDATASectionDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMCDATASectionDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMCDATASectionDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMCDATASectionDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMCDATASectionDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMCDATASectionDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMCDATASectionDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMCDATASectionDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMCDATASectionDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMCDATASectionDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMCDATASectionDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMCDATASectionDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMCDATASectionDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMCDATASectionDisp::transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMCDATASectionDisp::get_data()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000006D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMCDATASectionDisp::set_data(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000006D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

long IXMLDOMCDATASectionDisp::get_length()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000006E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

LPTSTR IXMLDOMCDATASectionDisp::substringData(long offset, long count)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_I4;
  params[0].lVal = count;
  params[1].vt = VT_I4;
  params[1].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000006F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMCDATASectionDisp::appendData(LPTSTR data)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000070, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

void IXMLDOMCDATASectionDisp::insertData(long offset, LPTSTR data)
{
  BSTR Buf;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params[0].bstrVal = Buf;
#endif
  params[1].vt = VT_I4;
  params[1].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000071, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

void IXMLDOMCDATASectionDisp::deleteData(long offset, long count)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_I4;
  params[0].lVal = count;
  params[1].vt = VT_I4;
  params[1].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000072, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

void IXMLDOMCDATASectionDisp::replaceData(long offset, long count, LPTSTR data)
{
  BSTR Buf;
  VARIANT params[3];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(data);
#else
  Buf = SysAllocStringLen(NULL, strlen(data));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, data, -1, Buf, strlen(data) + 1);
  params[0].bstrVal = Buf;
#endif
  params[1].vt = VT_I4;
  params[1].lVal = count;
  params[2].vt = VT_I4;
  params[2].lVal = offset;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 3;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000073, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

IXMLDOMText* IXMLDOMCDATASectionDisp::splitText(long offset)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_I4;
  params.lVal = offset;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000007B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMText*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMProcessingInstructionDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMProcessingInstructionDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMProcessingInstructionDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMProcessingInstructionDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMProcessingInstructionDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMProcessingInstructionDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMProcessingInstructionDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMProcessingInstructionDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMProcessingInstructionDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMProcessingInstructionDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMProcessingInstructionDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMProcessingInstructionDisp::insertBefore(IXMLDOMNode* newChild,
  VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMProcessingInstructionDisp::replaceChild(IXMLDOMNode* newChild,
  IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMProcessingInstructionDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMProcessingInstructionDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMProcessingInstructionDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMProcessingInstructionDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMProcessingInstructionDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMProcessingInstructionDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMProcessingInstructionDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMProcessingInstructionDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMProcessingInstructionDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMProcessingInstructionDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMProcessingInstructionDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMProcessingInstructionDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMProcessingInstructionDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMProcessingInstructionDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMProcessingInstructionDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMProcessingInstructionDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMProcessingInstructionDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMProcessingInstructionDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMProcessingInstructionDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMProcessingInstructionDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMProcessingInstructionDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMProcessingInstructionDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMProcessingInstructionDisp::transformNodeToObject(IXMLDOMNode* stylesheet,
  VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMProcessingInstructionDisp::get_target()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000007F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMProcessingInstructionDisp::get_data()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000080, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMProcessingInstructionDisp::set_data(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000080, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

LPTSTR IXMLDOMEntityReferenceDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMEntityReferenceDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMEntityReferenceDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMEntityReferenceDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMEntityReferenceDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMEntityReferenceDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityReferenceDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityReferenceDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityReferenceDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityReferenceDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMEntityReferenceDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityReferenceDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityReferenceDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityReferenceDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityReferenceDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMEntityReferenceDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMEntityReferenceDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityReferenceDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMEntityReferenceDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMEntityReferenceDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMEntityReferenceDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMEntityReferenceDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMEntityReferenceDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMEntityReferenceDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMEntityReferenceDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMEntityReferenceDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMEntityReferenceDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMEntityReferenceDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMEntityReferenceDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMEntityReferenceDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityReferenceDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMEntityReferenceDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMEntityReferenceDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMEntityReferenceDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMEntityReferenceDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMEntityReferenceDisp::transformNodeToObject(IXMLDOMNode* stylesheet,
  VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMParseErrorDisp::get_errorCode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000000, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

LPTSTR IXMLDOMParseErrorDisp::get_url()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x000000B3, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMParseErrorDisp::get_reason()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x000000B4, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMParseErrorDisp::get_srcText()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x000000B5, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

long IXMLDOMParseErrorDisp::get_line()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x000000B6, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

long IXMLDOMParseErrorDisp::get_linepos()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x000000B7, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

long IXMLDOMParseErrorDisp::get_filepos()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x000000B8, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

LPTSTR IXMLDOMNotationDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMNotationDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMNotationDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMNotationDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMNotationDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMNotationDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNotationDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNotationDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNotationDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNotationDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMNotationDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNotationDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNotationDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNotationDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNotationDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMNotationDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMNotationDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNotationDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMNotationDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMNotationDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMNotationDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMNotationDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMNotationDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMNotationDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMNotationDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMNotationDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMNotationDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMNotationDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMNotationDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMNotationDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMNotationDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMNotationDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMNotationDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMNotationDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMNotationDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMNotationDisp::transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMNotationDisp::get_publicId()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000088, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

VARIANT IXMLDOMNotationDisp::get_systemId()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000089, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

LPTSTR IXMLDOMEntityDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLDOMEntityDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMEntityDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDOMEntityDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXMLDOMEntityDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXMLDOMEntityDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXMLDOMEntityDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMEntityDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXMLDOMEntityDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXMLDOMEntityDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMEntityDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMEntityDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXMLDOMEntityDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXMLDOMEntityDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXMLDOMEntityDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMEntityDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMEntityDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXMLDOMEntityDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLDOMEntityDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMEntityDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXMLDOMEntityDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXMLDOMEntityDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXMLDOMEntityDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXMLDOMEntityDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMEntityDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDOMEntityDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDOMEntityDisp::transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

VARIANT IXMLDOMEntityDisp::get_publicId()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000008C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

VARIANT IXMLDOMEntityDisp::get_systemId()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000008D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

LPTSTR IXMLDOMEntityDisp::get_notationName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000008E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXTLRuntimeDisp::get_nodeName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXTLRuntimeDisp::get_nodeValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXTLRuntimeDisp::set_nodeValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXTLRuntimeDisp::get_nodeType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDOMNode* IXTLRuntimeDisp::get_parentNode()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNodeList* IXTLRuntimeDisp::get_childNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXTLRuntimeDisp::get_firstChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXTLRuntimeDisp::get_lastChild()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXTLRuntimeDisp::get_previousSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXTLRuntimeDisp::get_nextSibling()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNamedNodeMap* IXTLRuntimeDisp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNamedNodeMap*)pVarResult.pdispVal);
};

IXMLDOMNode* IXTLRuntimeDisp::insertBefore(IXMLDOMNode* newChild, VARIANT refChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = refChild.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = refChild.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXTLRuntimeDisp::replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)oldChild;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXTLRuntimeDisp::removeChild(IXMLDOMNode* childNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)childNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000000F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

IXMLDOMNode* IXTLRuntimeDisp::appendChild(IXMLDOMNode* newChild)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)newChild;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000010, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXTLRuntimeDisp::hasChildNodes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000011, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMDocument* IXTLRuntimeDisp::get_ownerDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000012, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

IXMLDOMNode* IXTLRuntimeDisp::cloneNode(BOOL deep)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BOOL;
  params.boolVal = deep;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000013, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

LPTSTR IXTLRuntimeDisp::get_nodeTypeString()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000015, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXTLRuntimeDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXTLRuntimeDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000018, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

BOOL IXTLRuntimeDisp::get_specified()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000016, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

IXMLDOMNode* IXTLRuntimeDisp::get_definition()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000017, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

VARIANT IXTLRuntimeDisp::get_nodeTypedValue()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXTLRuntimeDisp::set_nodeTypedValue(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00000019, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

VARIANT IXTLRuntimeDisp::get_dataType()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

void IXTLRuntimeDisp::set_dataType(VARIANT Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = Value.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = Value.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000001A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXTLRuntimeDisp::get_xml()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000001B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXTLRuntimeDisp::transformNode(IXMLDOMNode* stylesheet)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDOMNodeList* IXTLRuntimeDisp::selectNodes(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNodeList*)pVarResult.pdispVal);
};

IXMLDOMNode* IXTLRuntimeDisp::selectSingleNode(LPTSTR queryString)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(queryString);
#else
  Buf = SysAllocStringLen(NULL, strlen(queryString));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, queryString, -1, Buf, strlen(queryString) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0000001E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return((IXMLDOMNode*)pVarResult.pdispVal);
};

BOOL IXTLRuntimeDisp::get_parsed()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000001F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.boolVal);
};

LPTSTR IXTLRuntimeDisp::get_namespaceURI()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000020, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXTLRuntimeDisp::get_prefix()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000021, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXTLRuntimeDisp::get_baseName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000022, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXTLRuntimeDisp::transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject)
{
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = outputObject.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = outputObject.dblVal;
  params[1].vt = VT_DISPATCH;
  params[1].pdispVal = (IDispatch*)stylesheet;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000023, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

long IXTLRuntimeDisp::uniqueID(IXMLDOMNode* pNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)pNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000000BB, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

long IXTLRuntimeDisp::depth(IXMLDOMNode* pNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)pNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000000BC, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

long IXTLRuntimeDisp::childNumber(IXMLDOMNode* pNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)pNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000000BD, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

long IXTLRuntimeDisp::ancestorChildNumber(LPTSTR bstrNodeName, IXMLDOMNode* pNode)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_DISPATCH;
  params[0].pdispVal = (IDispatch*)pNode;
  params[1].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(bstrNodeName);
#else
  Buf = SysAllocStringLen(NULL, strlen(bstrNodeName));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, bstrNodeName, -1, Buf, strlen(bstrNodeName) + 1);
  params[1].bstrVal = Buf;
#endif
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000000BE, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return(pVarResult.lVal);
};

long IXTLRuntimeDisp::absoluteChildNumber(IXMLDOMNode* pNode)
{
  VARIANT pVarResult;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)pNode;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000000BF, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

LPTSTR IXTLRuntimeDisp::formatIndex(long lIndex, LPTSTR bstrFormat)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  BSTR Buf;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(bstrFormat);
#else
  Buf = SysAllocStringLen(NULL, strlen(bstrFormat));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, bstrFormat, -1, Buf, strlen(bstrFormat) + 1);
  params[0].bstrVal = Buf;
#endif
  params[1].vt = VT_I4;
  params[1].lVal = lIndex;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000000C0, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXTLRuntimeDisp::formatNumber(double dblNumber, LPTSTR bstrFormat)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  BSTR Buf;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(bstrFormat);
#else
  Buf = SysAllocStringLen(NULL, strlen(bstrFormat));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, bstrFormat, -1, Buf, strlen(bstrFormat) + 1);
  params[0].bstrVal = Buf;
#endif
  params[1].vt = VT_R8;
  params[1].dblVal = dblNumber;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000000C1, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXTLRuntimeDisp::formatDate(VARIANT varDate, LPTSTR bstrFormat, VARIANT varDestLocale)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  BSTR Buf;
  VARIANT params[3];
  DISPPARAMS dispparams;
  params[0].vt = varDestLocale.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = varDestLocale.dblVal;
  params[1].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(bstrFormat);
#else
  Buf = SysAllocStringLen(NULL, strlen(bstrFormat));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, bstrFormat, -1, Buf, strlen(bstrFormat) + 1);
  params[1].bstrVal = Buf;
#endif
  params[2].vt = varDate.vt;
  if(params[2].vt == VT_ERROR) params[2].scode = DISP_E_PARAMNOTFOUND;
  else params[2].dblVal = varDate.dblVal;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 3;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000000C2, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXTLRuntimeDisp::formatTime(VARIANT varTime, LPTSTR bstrFormat, VARIANT varDestLocale)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  BSTR Buf;
  VARIANT params[3];
  DISPPARAMS dispparams;
  params[0].vt = varDestLocale.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = varDestLocale.dblVal;
  params[1].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(bstrFormat);
#else
  Buf = SysAllocStringLen(NULL, strlen(bstrFormat));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, bstrFormat, -1, Buf, strlen(bstrFormat) + 1);
  params[1].bstrVal = Buf;
#endif
  params[2].vt = varTime.vt;
  if(params[2].vt == VT_ERROR) params[2].scode = DISP_E_PARAMNOTFOUND;
  else params[2].dblVal = varTime.dblVal;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 3;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000000C3, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void XMLDOMDocumentEvents::ondataavailable()
{
  Invoke((DISPID)0x000000C6, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, NULL, NULL, NULL);
};

void XMLDOMDocumentEvents::onreadystatechange()
{
  Invoke((DISPID)0xFFFFFD9F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, NULL, NULL, NULL);
};

IXMLDOMDocument* CreateDOMDocument()
{
  IXMLDOMDocument *res;
  CoCreateInstance(CLASS_DOMDocument, NULL,
    CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, DIID_IXMLDOMDocument, (void**)&res);
  return(res);
};

IXMLDOMDocument* CreateDOMFreeThreadedDocument()
{
  IXMLDOMDocument *res;
  CoCreateInstance(CLASS_DOMFreeThreadedDocument, NULL,
    CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, DIID_IXMLDOMDocument, (void**)&res);
  return(res);
};

void IXMLHttpRequestDisp::open(LPTSTR bstrMethod, LPTSTR bstrUrl, VARIANT varAsync,
  VARIANT bstrUser, VARIANT bstrPassword)
{
  int i;
  BSTR Buf[2];
  VARIANT params[5];
  DISPPARAMS dispparams;
  params[0].vt = bstrPassword.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = bstrPassword.dblVal;
  params[1].vt = bstrUser.vt;
  if(params[1].vt == VT_ERROR) params[1].scode = DISP_E_PARAMNOTFOUND;
  else params[1].dblVal = bstrUser.dblVal;
  params[2].vt = varAsync.vt;
  if(params[2].vt == VT_ERROR) params[2].scode = DISP_E_PARAMNOTFOUND;
  else params[2].dblVal = varAsync.dblVal;
  params[3].vt = VT_BSTR;
#ifdef UNICODE
  Buf[0] = SysAllocString(bstrUrl);
#else
  Buf[0] = SysAllocStringLen(NULL, strlen(bstrUrl));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, bstrUrl, -1, Buf[0], strlen(bstrUrl) + 1);
  params[3].bstrVal = Buf[0];
#endif
  params[4].vt = VT_BSTR;
#ifdef UNICODE
  Buf[1] = SysAllocString(bstrMethod);
#else
  Buf[1] = SysAllocStringLen(NULL, strlen(bstrMethod));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, bstrMethod, -1, Buf[1], strlen(bstrMethod) + 1);
  params[4].bstrVal = Buf[1];
#endif
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 5;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000001, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  for(i = 0; i < 2; i++) SysFreeString(Buf[i]);
};

void IXMLHttpRequestDisp::setRequestHeader(LPTSTR bstrHeader, LPTSTR bstrValue)
{
  int i;
  BSTR Buf[2];
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = VT_BSTR;
#ifdef UNICODE
  Buf[0] = SysAllocString(bstrValue);
#else
  Buf[0] = SysAllocStringLen(NULL, strlen(bstrValue));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, bstrValue, -1, Buf[0], strlen(bstrValue) + 1);
  params[0].bstrVal = Buf[0];
#endif
  params[1].vt = VT_BSTR;
#ifdef UNICODE
  Buf[1] = SysAllocString(bstrHeader);
#else
  Buf[1] = SysAllocStringLen(NULL, strlen(bstrHeader));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, bstrHeader, -1, Buf[1], strlen(bstrHeader) + 1);
  params[1].bstrVal = Buf[1];
#endif
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  for(i = 0; i < 2; i++) SysFreeString(Buf[i]);
};

LPTSTR IXMLHttpRequestDisp::getResponseHeader(LPTSTR bstrHeader)
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(bstrHeader);
#else
  Buf = SysAllocStringLen(NULL, strlen(bstrHeader));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, bstrHeader, -1, Buf, strlen(bstrHeader) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLHttpRequestDisp::getAllResponseHeaders()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000004, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLHttpRequestDisp::send(VARIANT varBody)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = varBody.vt;
  if(params.vt == VT_ERROR) params.scode = DISP_E_PARAMNOTFOUND;
  else params.dblVal = varBody.dblVal;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00000005, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

void IXMLHttpRequestDisp::abort()
{
  Invoke((DISPID)0x00000006, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparamsNoArgs, NULL, NULL, NULL);
};

long IXMLHttpRequestDisp::get_status()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000007, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

LPTSTR IXMLHttpRequestDisp::get_statusText()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00000008, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IDispatch * IXMLHttpRequestDisp::get_responseXML()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00000009, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.pdispVal);
};

LPTSTR IXMLHttpRequestDisp::get_responseText()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0000000A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

VARIANT IXMLHttpRequestDisp::get_responseBody()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

VARIANT IXMLHttpRequestDisp::get_responseStream()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult);
};

long IXMLHttpRequestDisp::get_readyState()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0000000D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

void IXMLHttpRequestDisp::set_onreadystatechange(IDispatch * Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = Value;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0000000E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

IXMLHttpRequest* CreateXMLHTTPRequest()
{
  IXMLHttpRequest *res;
  CoCreateInstance(CLASS_XMLHTTPRequest, NULL,
    CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, DIID_IXMLHttpRequest, (void**)&res);
  return(res);
};

IXMLDOMDocument* IXMLDSOControlDisp::get_XMLDocument()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00010001, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLDOMDocument*)pVarResult.pdispVal);
};

void IXMLDSOControlDisp::set_XMLDocument(IXMLDOMDocument* Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)Value;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00010001, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDSOControlDisp::get_JavaDSOCompatible()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00010002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

void IXMLDSOControlDisp::set_JavaDSOCompatible(long Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_I4;
  params.lVal = Value;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00010002, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLDSOControlDisp::get_readyState()
{
  VARIANT pVarResult;
  Invoke((DISPID)0xFFFFFDF3, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IXMLDSOControl* CreateXMLDSOControl()
{
  IXMLDSOControl *res;
  CoCreateInstance(CLASS_XMLDSOControl, NULL,
    CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, DIID_IXMLDSOControl, (void**)&res);
  return(res);
};

void IXMLElementCollectionDisp::set_length(long Value)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_I4;
  params.lVal = Value;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00010001, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
};

long IXMLElementCollectionDisp::get_length()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00010001, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

IUnknown * IXMLElementCollectionDisp::get__newEnum()
{
  VARIANT pVarResult;
  Invoke((DISPID)0xFFFFFFFC, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.punkVal);
};

IDispatch * IXMLElementCollectionDisp::item(VARIANT var1, VARIANT var2)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = var2.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = var2.dblVal;
  params[1].vt = var1.vt;
  if(params[1].vt == VT_ERROR) params[1].scode = DISP_E_PARAMNOTFOUND;
  else params[1].dblVal = var1.dblVal;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x00010003, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return(pVarResult.pdispVal);
};

IXMLElement* IXMLDocumentDisp::get_root()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x00010065, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLElement*)pVarResult.pdispVal);
};

LPTSTR IXMLDocumentDisp::get_fileSize()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00010066, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDocumentDisp::get_fileModifiedDate()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00010067, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDocumentDisp::get_fileUpdatedDate()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00010068, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDocumentDisp::get_url()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00010069, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDocumentDisp::set_url(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x00010069, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

LPTSTR IXMLDocumentDisp::get_mimeType()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0001006A, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

long IXMLDocumentDisp::get_readyState()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x0001006B, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

LPTSTR IXMLDocumentDisp::get_charset()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0001006D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLDocumentDisp::set_charset(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x0001006D, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

LPTSTR IXMLDocumentDisp::get_version()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0001006E, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDocumentDisp::get_doctype()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x0001006F, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLDocumentDisp::get_dtdURL()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00010070, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLElement* IXMLDocumentDisp::createElement(VARIANT vType, VARIANT var1)
{
  VARIANT pVarResult;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = var1.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = var1.dblVal;
  params[1].vt = vType.vt;
  if(params[1].vt == VT_ERROR) params[1].scode = DISP_E_PARAMNOTFOUND;
  else params[1].dblVal = vType.dblVal;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x0001006C, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  return((IXMLElement*)pVarResult.pdispVal);
};

LPTSTR IXMLElementDisp::get_tagName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x000100C9, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLElementDisp::set_tagName(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x000100C9, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

IXMLElement* IXMLElementDisp::get_parent()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x000100CA, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLElement*)pVarResult.pdispVal);
};

void IXMLElementDisp::setAttribute(LPTSTR strPropertyName, VARIANT PropertyValue)
{
  BSTR Buf;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = PropertyValue.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = PropertyValue.dblVal;
  params[1].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(strPropertyName);
#else
  Buf = SysAllocStringLen(NULL, strlen(strPropertyName));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, strPropertyName, -1, Buf, strlen(strPropertyName) + 1);
  params[1].bstrVal = Buf;
#endif
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000100CB, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

VARIANT IXMLElementDisp::getAttribute(LPTSTR strPropertyName)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(strPropertyName);
#else
  Buf = SysAllocStringLen(NULL, strlen(strPropertyName));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, strPropertyName, -1, Buf, strlen(strPropertyName) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000100CC, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return(pVarResult);
};

void IXMLElementDisp::removeAttribute(LPTSTR strPropertyName)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(strPropertyName);
#else
  Buf = SysAllocStringLen(NULL, strlen(strPropertyName));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, strPropertyName, -1, Buf, strlen(strPropertyName) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000100CD, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

IXMLElementCollection* IXMLElementDisp::get_children()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x000100CE, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLElementCollection*)pVarResult.pdispVal);
};

long IXMLElementDisp::get_type()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x000100CF, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

LPTSTR IXMLElementDisp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x000100D0, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLElementDisp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x000100D0, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

void IXMLElementDisp::addChild(IXMLElement* pChildElem, long lIndex, long lReserved)
{
  VARIANT params[3];
  DISPPARAMS dispparams;
  params[0].vt = VT_I4;
  params[0].lVal = lReserved;
  params[1].vt = VT_I4;
  params[1].lVal = lIndex;
  params[2].vt = VT_DISPATCH;
  params[2].pdispVal = (IDispatch*)pChildElem;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 3;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000100D1, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

void IXMLElementDisp::removeChild(IXMLElement* pChildElem)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)pChildElem;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000100D2, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

LPTSTR IXMLElement2Disp::get_tagName()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x000100C9, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLElement2Disp::set_tagName(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x000100C9, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

IXMLElement2* IXMLElement2Disp::get_parent()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x000100CA, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLElement2*)pVarResult.pdispVal);
};

void IXMLElement2Disp::setAttribute(LPTSTR strPropertyName, VARIANT PropertyValue)
{
  BSTR Buf;
  VARIANT params[2];
  DISPPARAMS dispparams;
  params[0].vt = PropertyValue.vt;
  if(params[0].vt == VT_ERROR) params[0].scode = DISP_E_PARAMNOTFOUND;
  else params[0].dblVal = PropertyValue.dblVal;
  params[1].vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(strPropertyName);
#else
  Buf = SysAllocStringLen(NULL, strlen(strPropertyName));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, strPropertyName, -1, Buf, strlen(strPropertyName) + 1);
  params[1].bstrVal = Buf;
#endif
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000100CB, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

VARIANT IXMLElement2Disp::getAttribute(LPTSTR strPropertyName)
{
  VARIANT pVarResult;
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(strPropertyName);
#else
  Buf = SysAllocStringLen(NULL, strlen(strPropertyName));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, strPropertyName, -1, Buf, strlen(strPropertyName) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000100CC, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, &pVarResult, NULL, NULL);
  SysFreeString(Buf);
  return(pVarResult);
};

void IXMLElement2Disp::removeAttribute(LPTSTR strPropertyName)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(strPropertyName);
#else
  Buf = SysAllocStringLen(NULL, strlen(strPropertyName));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, strPropertyName, -1, Buf, strlen(strPropertyName) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000100CD, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

IXMLElementCollection* IXMLElement2Disp::get_children()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x000100CE, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLElementCollection*)pVarResult.pdispVal);
};

long IXMLElement2Disp::get_type()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x000100CF, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return(pVarResult.lVal);
};

LPTSTR IXMLElement2Disp::get_text()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x000100D0, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

void IXMLElement2Disp::set_text(LPTSTR Value)
{
  BSTR Buf;
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_BSTR;
#ifdef UNICODE
  Buf = SysAllocString(Value);
#else
  Buf = SysAllocStringLen(NULL, strlen(Value));
  MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Value, -1, Buf, strlen(Value) + 1);
  params.bstrVal = Buf;
#endif
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = &DispIdPropPut;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 1;
  Invoke((DISPID)0x000100D0, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);
  SysFreeString(Buf);
};

void IXMLElement2Disp::addChild(IXMLElement2* pChildElem, long lIndex, long lReserved)
{
  VARIANT params[3];
  DISPPARAMS dispparams;
  params[0].vt = VT_I4;
  params[0].lVal = lReserved;
  params[1].vt = VT_I4;
  params[1].lVal = lIndex;
  params[2].vt = VT_DISPATCH;
  params[2].pdispVal = (IDispatch*)pChildElem;
  dispparams.rgvarg = params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 3;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000100D1, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

void IXMLElement2Disp::removeChild(IXMLElement2* pChildElem)
{
  VARIANT params;
  DISPPARAMS dispparams;
  params.vt = VT_DISPATCH;
  params.pdispVal = (IDispatch*)pChildElem;
  dispparams.rgvarg = &params;
  dispparams.rgdispidNamedArgs = NULL;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  Invoke((DISPID)0x000100D2, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_METHOD, &dispparams, NULL, NULL, NULL);
};

IXMLElementCollection* IXMLElement2Disp::get_attributes()
{
  VARIANT pVarResult;
  Invoke((DISPID)0x000100D3, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
  return((IXMLElementCollection*)pVarResult.pdispVal);
};

LPTSTR IXMLAttributeDisp::get_name()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00010191, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

LPTSTR IXMLAttributeDisp::get_value()
{
  VARIANT pVarResult;
  char *ResBuf;
  int ilen;
  Invoke((DISPID)0x00010192, IID_NULL, (LCID)LOCALE_USER_DEFAULT,
    (WORD)DISPATCH_PROPERTYGET, &dispparamsNoArgs, &pVarResult, NULL, NULL);
#ifdef UNICODE
  return(pVarResult.bstrVal);
#else
  ilen = WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, NULL, 0, NULL, NULL);
  ResBuf = (char*)LocalAlloc(LMEM_FIXED, ilen + 1);
  WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, pVarResult.bstrVal, -1, ResBuf, ilen, NULL, NULL);
  VariantClear(&pVarResult);
  return(ResBuf);
#endif
};

IXMLDocument2* CreateXMLDocument()
{
  IXMLDocument2 *res;
  CoCreateInstance(CLASS_XMLDocument, NULL,
    CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_IXMLDocument2, (void**)&res);
  return(res);
};

