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

#ifndef _XMLUtils_HPP
#define _XMLUtils_HPP

#ifdef _MSXML_H_
#include <msxml.h>
#else
#include "msxml_DLL.hpp"
#endif

class CXMLReader
{
private:
    char* fName;
    IXMLDOMDocument* fDoc;
    IXMLDOMElement* fRoot;
    BOOL fOpened;
public:
    CXMLReader(LPSTR AFileName);
    ~CXMLReader();
    IXMLDOMElement* OpenSection(LPCSTR AName);
    IXMLDOMElement* OpenSubSection(IXMLDOMElement* AParent, LPCSTR AName);
    IXMLDOMElement* OpenSubSection(IXMLDOMElement* AParent, BSTR AName);
    IXMLDOMElement* GetFirstNode(IXMLDOMElement* AParent, LPSTR *ANodeName);
    IXMLDOMElement* GetFirstNode(IXMLDOMElement* AParent, BSTR *ANodeName);
    IXMLDOMElement* GetNextNode(IXMLDOMElement* APrevNode, LPSTR *ANodeName);
    IXMLDOMElement* GetNextNode(IXMLDOMElement* APrevNode, BSTR *ANodeName);
    BOOL GetByteValue(IXMLDOMElement* ASec, LPCSTR AName, BYTE *Value);
    BOOL GetIntValue(IXMLDOMElement* ASec, LPCSTR AName, int *Value);
    BOOL GetLongValue(IXMLDOMElement* ASec, LPCSTR AName, long *Value);
    BOOL GetDoubleValue(IXMLDOMElement* ASec, LPCSTR AName, double *Value);
    BOOL GetStringValue(IXMLDOMElement* ASec, LPCSTR AName, LPSTR *Value);
    BOOL GetStringValue(IXMLDOMElement* ASec, LPCSTR AName, BSTR *Value);
    BOOL GetStringValueBuf(IXMLDOMElement* ASec, LPCSTR AName, LPSTR Value,
        int ABufSize);
    BOOL GetStringValueBuf(IXMLDOMElement* ASec, LPCSTR AName, BSTR Value,
        int ABufSize);
};

class CXMLWritter
{
private:
    char* fName;
    IXMLDOMDocument* fDoc;
    IXMLDOMElement* fRoot;
    void InsertNode(IDispatch* AParent, IDispatch* ANode);
public:
    CXMLWritter(LPSTR AFileName);
    ~CXMLWritter();
    void Save();
    void WriteComment(LPCSTR Value);
    void CreateRoot(LPCSTR AName);
    IXMLDOMElement* CreateSection(LPCSTR AName);
    IXMLDOMElement* CreateSubSection(IXMLDOMElement* AParent, LPCSTR AName);
    IXMLDOMElement* CreateSubSection(IXMLDOMElement* AParent, BSTR AName);
    void AddByteValue(IXMLDOMElement* ASec, LPCSTR AName, BYTE Value);
    void AddIntValue(IXMLDOMElement* ASec, LPCSTR AName, int Value);
    void AddLongValue(IXMLDOMElement* ASec, LPCSTR AName, long Value);
    void AddDoubleValue(IXMLDOMElement* ASec, LPCSTR AName, double Value);
    void AddStringValue(IXMLDOMElement* ASec, LPCSTR AName, char* Value);
    void AddStringValue(IXMLDOMElement* ASec, LPCSTR AName, BSTR Value);
};

#endif
