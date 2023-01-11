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
#include "XMLUtils.hpp"

#ifdef _MSXML_H_
#define DIID_IXMLDOMElement IID_IXMLDOMElement
#define DIID_IXMLDOMNode IID_IXMLDOMNode
#define DIID_IXMLDOMDocument IID_IXMLDOMDocument

const CLSID CLASS_DOMDocument = {0x2933BF90,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};

IXMLDOMDocument* CreateDOMDocument()
{
  IXMLDOMDocument *res;
  CoCreateInstance(CLASS_DOMDocument, NULL,
    CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, DIID_IXMLDOMDocument, (void**)&res);
  return(res);
}

#endif

CXMLReader::CXMLReader(LPSTR AFileName)
{
    fName = (char*)malloc(strlen(AFileName) + 1);
    strcpy(fName, AFileName);
    fDoc = CreateDOMDocument();

    VARIANT V;
    V.vt = VT_BSTR;
    V.bstrVal = StrToBSTR(AFileName);

    VARIANT_BOOL ob;
    fDoc->load(V, &ob);
    VariantClear(&V);
    fOpened = ob;

    fRoot = NULL;
    fDoc->get_documentElement(&fRoot);
    return;
}

CXMLReader::~CXMLReader()
{
    if(fRoot) fRoot->Release();
    fDoc->Release();
    free(fName);
    return;
}

IXMLDOMElement* CXMLReader::OpenSection(LPCSTR AName)
{
    BSTR WS = StrToBSTR(AName);
    IXMLDOMNode* N = NULL;
    IXMLDOMElement* res = NULL;

    if(fOpened)
    {
        fRoot->selectSingleNode(WS, &N);
        SysFreeString(WS);
    }
    if(N)
    {
        N->QueryInterface(DIID_IXMLDOMElement, (void**)&res);
        N->Release();
    };
    return(res);
}

IXMLDOMElement* CXMLReader::OpenSubSection(IXMLDOMElement* AParent, LPCSTR AName)
{
    BSTR WS = StrToBSTR(AName);
    IXMLDOMNode* N = NULL;
    IXMLDOMElement* res = NULL;

    if(fOpened)
    {
        AParent->selectSingleNode(WS, &N);
        SysFreeString(WS);
    }
    if(N)
    {
        N->QueryInterface(DIID_IXMLDOMElement, (void**)&res);
        N->Release();
    };
    return(res);
}

IXMLDOMElement* CXMLReader::OpenSubSection(IXMLDOMElement* AParent, BSTR AName)
{
    IXMLDOMNode* N = NULL;
    IXMLDOMElement* res = NULL;

    if(fOpened) AParent->selectSingleNode(AName, &N);
    if(N)
    {
        N->QueryInterface(DIID_IXMLDOMElement, (void**)&res);
        N->Release();
    };
    return(res);
}

IXMLDOMElement* CXMLReader::GetFirstNode(IXMLDOMElement* AParent, LPSTR *ANodeName)
{
    *ANodeName = NULL;
    IXMLDOMNode *N1 = NULL;
    IXMLDOMElement *res = NULL;
    BSTR wnodename;

    HRESULT hres = AParent->get_firstChild(&N1);
    if(N1 && (hres == S_OK))
    {
        hres = N1->QueryInterface(DIID_IXMLDOMElement, (void**)&res);
        if(hres != S_OK) res = NULL;
        else
        {
            N1->get_nodeName(&wnodename);
            *ANodeName = WideCharToStr(wnodename);
            SysFreeString(wnodename);
        }
        N1->Release();
    }
    return(res);
}

IXMLDOMElement* CXMLReader::GetFirstNode(IXMLDOMElement* AParent, BSTR *ANodeName)
{
    *ANodeName = NULL;
    IXMLDOMNode *N1 = NULL;
    IXMLDOMElement *res = NULL;
    HRESULT hres = AParent->get_firstChild(&N1);
    if(N1 && (hres == S_OK))
    {
        hres = N1->QueryInterface(DIID_IXMLDOMElement, (void**)&res);
        if(hres != S_OK) res = NULL;
        else N1->get_nodeName(ANodeName);
        N1->Release();
    }
    return(res);
}

IXMLDOMElement* CXMLReader::GetNextNode(IXMLDOMElement* APrevNode, LPSTR *ANodeName)
{
    *ANodeName = NULL;
    IXMLDOMNode *N1 = NULL, *N2 = NULL;
    IXMLDOMElement *res = NULL;
    BSTR wnodename;

    HRESULT hres = APrevNode->QueryInterface(DIID_IXMLDOMNode, (void**)&N1);
    if(N1 && (hres == S_OK))
    {
        hres = N1->get_nextSibling(&N2);
        if(N2 && (hres == S_OK))
        {
            hres = N2->QueryInterface(DIID_IXMLDOMElement, (void**)&res);
            if(hres != S_OK) res = NULL;
            else
            {
                N2->get_nodeName(&wnodename);
                *ANodeName = WideCharToStr(wnodename);
                SysFreeString(wnodename);
            }
            N2->Release();
        }
        N1->Release();
    }
    // Caution! We clear the previous node so that it is possible to reuse
    // the same Element variable in a "while" loop
    APrevNode->Release();
    return(res);
}

IXMLDOMElement* CXMLReader::GetNextNode(IXMLDOMElement* APrevNode, BSTR *ANodeName)
{
    *ANodeName = NULL;
    IXMLDOMNode *N1 = NULL, *N2 = NULL;
    IXMLDOMElement *res = NULL;
    HRESULT hres = APrevNode->QueryInterface(DIID_IXMLDOMNode, (void**)&N1);
    if(N1 && (hres == S_OK))
    {
        hres = N1->get_nextSibling(&N2);
        if(N2 && (hres == S_OK))
        {
            hres = N2->QueryInterface(DIID_IXMLDOMElement, (void**)&res);
            if(hres != S_OK) res = NULL;
            else N2->get_nodeName(ANodeName);
            N2->Release();
        }
        N1->Release();
    }
    // Caution! We clear the previous node so that it is possible to reuse
    // the same Element variable in a "while" loop
    APrevNode->Release();
    return(res);
}

BOOL CXMLReader::GetByteValue(IXMLDOMElement* ASec, LPCSTR AName, BYTE *Value)
{
    BSTR WS = StrToBSTR(AName);
    VARIANT V;

    VariantInit(&V);
    ASec->getAttribute(WS, &V);
    SysFreeString(WS);

    BOOL res = V.vt == VT_BSTR;
    if(res)
    {
        BYTE i;
        VarUI1FromStr(V.bstrVal, 0, 0, &i);
        *Value = i;
    };
    VariantClear(&V);
    return(res);
}

BOOL CXMLReader::GetIntValue(IXMLDOMElement* ASec, LPCSTR AName, int *Value)
{
    BSTR WS = StrToBSTR(AName);
    VARIANT V;

    VariantInit(&V);
    ASec->getAttribute(WS, &V);
    SysFreeString(WS);

    BOOL res = V.vt == VT_BSTR;
    if(res)
    {
        short i;
        VarI2FromStr(V.bstrVal, 0, 0, &i);
        *Value = i;
    };
    VariantClear(&V);
    return(res);
}

BOOL CXMLReader::GetLongValue(IXMLDOMElement* ASec, LPCSTR AName, long *Value)
{
    BSTR WS = StrToBSTR(AName);
    VARIANT V;

    VariantInit(&V);
    ASec->getAttribute(WS, &V);
    SysFreeString(WS);

    BOOL res = V.vt == VT_BSTR;
    if(res) VarI4FromStr(V.bstrVal, 0, 0, Value);
    VariantClear(&V);
    return(res);
}

BOOL CXMLReader::GetDoubleValue(IXMLDOMElement* ASec, LPCSTR AName, double *Value)
{
    //const LOCALE_NOUSEROVERRIDE = $8000000; // invalid but magic ?

    BSTR WS = StrToBSTR(AName);
    VARIANT V;

    VariantInit(&V);
    ASec->getAttribute(WS, &V);
    SysFreeString(WS);
    BOOL res = V.vt == VT_BSTR;
    if(res) VarR8FromStr(V.bstrVal, 0, LOCALE_NOUSEROVERRIDE, Value);
    VariantClear(&V);
    return(res);
}

BOOL CXMLReader::GetStringValue(IXMLDOMElement* ASec, LPCSTR AName, LPSTR *Value)
{
    *Value = NULL;

    BSTR WS = StrToBSTR(AName);
    VARIANT V;

    VariantInit(&V);
    ASec->getAttribute(WS, &V);
    SysFreeString(WS);
    BOOL res = V.vt == VT_BSTR;

    if(res) *Value = WideCharToStr(V.bstrVal);
    VariantClear(&V);
    return(res);
}

BOOL CXMLReader::GetStringValue(IXMLDOMElement* ASec, LPCSTR AName, BSTR *Value)
{
    *Value = NULL;

    BSTR WS = StrToBSTR(AName);
    VARIANT V;

    VariantInit(&V);
    ASec->getAttribute(WS, &V);
    SysFreeString(WS);
    BOOL res = V.vt == VT_BSTR;
    if(res) *Value = V.bstrVal;
    return(res);
}

BOOL CXMLReader::GetStringValueBuf(IXMLDOMElement* ASec, LPCSTR AName,
    LPSTR Value, int ABufSize)
{
    Value[0] = 0;

    BSTR WS = StrToBSTR(AName);
    VARIANT V;

    VariantInit(&V);
    ASec->getAttribute(WS, &V);
    SysFreeString(WS);
    BOOL res = V.vt == VT_BSTR;
    if(res)
    {
        WideCharToMultiByte(CP_ACP, WC_COMPOSITECHECK, V.bstrVal, -1,
            Value, ABufSize, NULL, NULL);
    }
    VariantClear(&V);
    return(res);
}

BOOL CXMLReader::GetStringValueBuf(IXMLDOMElement* ASec, LPCSTR AName,
    BSTR Value, int ABufSize)
{
    Value[0] = 0;

    BSTR WS = StrToBSTR(AName);
    VARIANT V;

    VariantInit(&V);
    ASec->getAttribute(WS, &V);
    SysFreeString(WS);
    BOOL res = V.vt == VT_BSTR;
    if(res) wcscpy(Value, V.bstrVal);
    VariantClear(&V);
    return(res);
}


CXMLWritter::CXMLWritter(LPSTR AFileName)
{
    fName = (char*)malloc(strlen(AFileName) + 1);
    strcpy(fName, AFileName);
    BSTR WS = StrToBSTR("xml");
    //BSTR WName = StrToBSTR("version=\"1.0\" encoding=\"windows-1250\"");
    BSTR WName = StrToBSTR("version=\"1.0\"");

    fDoc = CreateDOMDocument();

    IXMLDOMProcessingInstruction* PI = NULL;
    fDoc->createProcessingInstruction(WS, WName, &PI);
    SysFreeString(WName);
    SysFreeString(WS);
    InsertNode(fDoc, PI);
    PI->Release();

    fRoot = NULL;
    return;
}

CXMLWritter::~CXMLWritter()
{
    if(fRoot) fRoot->Release();
    fDoc->Release();
    free(fName);
    return;
}

void CXMLWritter::InsertNode(IDispatch* AParent, IDispatch* ANode)
{
    IXMLDOMNode* ParN = NULL;
    IXMLDOMNode* ChildN = NULL;
    IXMLDOMNode* DummyN = NULL;

    AParent->QueryInterface(DIID_IXMLDOMNode, (void**)&ParN);
    ANode->QueryInterface(DIID_IXMLDOMNode, (void**)&ChildN);
    ParN->appendChild(ChildN, &DummyN);
    if(DummyN) DummyN->Release();
    ChildN->Release();
    ParN->Release();
    return;
}

void CXMLWritter::Save()
{
    VARIANT V;
    V.vt = VT_BSTR;
    V.bstrVal = StrToBSTR(fName);
    fDoc->save(V);
    VariantClear(&V);
    return;
}

void CXMLWritter::WriteComment(LPCSTR Value)
{
    BSTR WS = StrToBSTR(Value);
    IXMLDOMComment* Comment;
    fDoc->createComment(WS, &Comment);
    SysFreeString(WS);
    InsertNode(fDoc, Comment);
    Comment->Release();
    return;
}

void CXMLWritter::CreateRoot(LPCSTR AName)
{
    BSTR WS = StrToBSTR(AName);
    fDoc->createElement(WS, &fRoot);
    SysFreeString(WS);
    InsertNode(fDoc, fRoot);
    return;
}

IXMLDOMElement* CXMLWritter::CreateSection(LPCSTR AName)
{
    BSTR WS = NULL;
    IXMLDOMText* TN = NULL;
    VARIANT_BOOL hch;

    fRoot->hasChildNodes(&hch);
    if(!hch)
    {
        WS = SysAllocString(L"\r\n");
        fDoc->createTextNode(WS, &TN);
        SysFreeString(WS);
        InsertNode(fRoot, TN);
        TN->Release();
    }

    WS = SysAllocString(L"\t");
    fDoc->createTextNode(WS, &TN);
    SysFreeString(WS);
    InsertNode(fRoot, TN);
    TN->Release();
    TN = NULL;

    IXMLDOMElement* res = NULL;
    WS = StrToBSTR(AName);
    fDoc->createElement(WS, &res);
    SysFreeString(WS);
    InsertNode(fRoot, res);

    WS = SysAllocString(L"\r\n");
    fDoc->createTextNode(WS, &TN);
    SysFreeString(WS);
    InsertNode(fRoot, TN);
    TN->Release();
    return(res);
}

IXMLDOMElement* CXMLWritter::CreateSubSection(IXMLDOMElement* AParent,
    LPCSTR AName)
{
    BSTR WS = NULL;
    IXMLDOMText* TN = NULL;
    IXMLDOMNode *N1, *N2;
    IXMLDOMNode *NPar = NULL;
    VARIANT_BOOL hch;

    AParent->hasChildNodes(&hch);
    if(!hch)
    {
        WS = SysAllocString(L"\r\n");
        fDoc->createTextNode(WS, &TN);
        SysFreeString(WS);
        InsertNode(AParent, TN);
        TN->Release();
        TN = NULL;
        N1 = NULL;
        AParent->get_parentNode(&NPar);
        if(NPar)
        {
            NPar->get_parentNode(&N1);
        }
        WS = SysAllocString(L"\t");
        while(N1)
        {
            fDoc->createTextNode(WS, &TN);
            InsertNode(AParent, TN);
            TN->Release();
            TN = NULL;
            N2 = N1;
            N1 = NULL;
            N2->get_parentNode(&N1);
            N2->Release();
        }
        SysFreeString(WS);
    }
    WS = SysAllocString(L"\t");
    fDoc->createTextNode(WS, &TN);
    SysFreeString(WS);
    InsertNode(AParent, TN);
    TN->Release();
    TN = NULL;

    IXMLDOMElement* res = NULL;
    WS = StrToBSTR(AName);
    fDoc->createElement(WS, &res);
    SysFreeString(WS);
    InsertNode(AParent, res);

    WS = SysAllocString(L"\r\n");
    fDoc->createTextNode(WS, &TN);
    SysFreeString(WS);
    InsertNode(AParent, TN);
    TN->Release();
    TN = NULL;

    N1 = NULL;
    AParent->get_parentNode(&NPar);
    if(NPar)
    {
        NPar->get_parentNode(&N1);
    }

    WS = SysAllocString(L"\t");
    while(N1)
    {
        fDoc->createTextNode(WS, &TN);
        InsertNode(AParent, TN);
        TN->Release();
        TN = NULL;
        N2 = N1;
        N1 = NULL;
        N2->get_parentNode(&N1);
        N2->Release();
    }
    SysFreeString(WS);

    return(res);
}

IXMLDOMElement* CXMLWritter::CreateSubSection(IXMLDOMElement* AParent, BSTR AName)
{
    BSTR WS = NULL;
    IXMLDOMText* TN = NULL;
    IXMLDOMNode *N1, *N2;
    IXMLDOMNode *NPar = NULL;
    VARIANT_BOOL hch;

    AParent->hasChildNodes(&hch);
    if(!hch)
    {
        WS = SysAllocString(L"\r\n");
        fDoc->createTextNode(WS, &TN);
        SysFreeString(WS);
        InsertNode(AParent, TN);
        TN->Release();
        TN = NULL;
        N1 = NULL;
        AParent->get_parentNode(&NPar);
        if(NPar)
        {
            NPar->get_parentNode(&N1);
        }
        WS = SysAllocString(L"\t");
        while(N1)
        {
            fDoc->createTextNode(WS, &TN);
            InsertNode(AParent, TN);
            TN->Release();
            TN = NULL;
            N2 = N1;
            N1 = NULL;
            N2->get_parentNode(&N1);
            N2->Release();
        }
        SysFreeString(WS);
    }
    WS = SysAllocString(L"\t");
    fDoc->createTextNode(WS, &TN);
    SysFreeString(WS);
    InsertNode(AParent, TN);
    TN->Release();
    TN = NULL;

    IXMLDOMElement* res = NULL;
    fDoc->createElement(AName, &res);
    InsertNode(AParent, res);

    WS = SysAllocString(L"\r\n");
    fDoc->createTextNode(WS, &TN);
    SysFreeString(WS);
    InsertNode(AParent, TN);
    TN->Release();
    TN = NULL;

    N1 = NULL;
    AParent->get_parentNode(&NPar);
    if(NPar)
    {
        NPar->get_parentNode(&N1);
    }

    WS = SysAllocString(L"\t");
    while(N1)
    {
        fDoc->createTextNode(WS, &TN);
        InsertNode(AParent, TN);
        TN->Release();
        TN = NULL;
        N2 = N1;
        N1 = NULL;
        N2->get_parentNode(&N1);
        N2->Release();
    }
    SysFreeString(WS);

    return(res);
}

void CXMLWritter::AddByteValue(IXMLDOMElement* ASec, LPCSTR AName, BYTE Value)
{
    BSTR WS = StrToBSTR(AName);
    VARIANT V;
    IXMLDOMAttribute* Attr = NULL;

    fDoc->createAttribute(WS, &Attr);
    SysFreeString(WS);
    V.vt = VT_UI1;
    V.bVal = Value;
#ifdef _MSXML_H_
	Attr->put_value(V);
#else
    Attr->set_value(V);
#endif
    ASec->setAttributeNode(Attr, NULL);
    Attr->Release();
    return;
}

void CXMLWritter::AddIntValue(IXMLDOMElement* ASec, LPCSTR AName, int Value)
{
    BSTR WS = StrToBSTR(AName);
    VARIANT V;
    IXMLDOMAttribute* Attr = NULL;

    fDoc->createAttribute(WS, &Attr);
    SysFreeString(WS);
    V.vt = VT_I2;
    V.iVal = Value;
#ifdef _MSXML_H_
	Attr->put_value(V);
#else
    Attr->set_value(V);
#endif
    ASec->setAttributeNode(Attr, NULL);
    Attr->Release();
    return;
}

void CXMLWritter::AddLongValue(IXMLDOMElement* ASec, LPCSTR AName, long Value)
{
    BSTR WS = StrToBSTR(AName);
    VARIANT V;
    IXMLDOMAttribute* Attr = NULL;

    fDoc->createAttribute(WS, &Attr);
    SysFreeString(WS);
    V.vt = VT_I4;
    V.lVal = Value;
#ifdef _MSXML_H_
	Attr->put_value(V);
#else
    Attr->set_value(V);
#endif
    ASec->setAttributeNode(Attr, NULL);
    Attr->Release();
    return;
}

void CXMLWritter::AddDoubleValue(IXMLDOMElement* ASec, LPCSTR AName, double Value)
{
    BSTR WS = StrToBSTR(AName);
    VARIANT V;
    IXMLDOMAttribute* Attr = NULL;

    fDoc->createAttribute(WS, &Attr);
    SysFreeString(WS);
    V.vt = VT_R8;
    V.dblVal = Value;
#ifdef _MSXML_H_
	Attr->put_value(V);
#else
    Attr->set_value(V);
#endif
    ASec->setAttributeNode(Attr, NULL);
    Attr->Release();
    return;
}

void CXMLWritter::AddStringValue(IXMLDOMElement* ASec, LPCSTR AName, char* Value)
{
    BSTR WS = StrToBSTR(AName);
    IXMLDOMAttribute* Attr = NULL;

    fDoc->createAttribute(WS, &Attr);
    SysFreeString(WS);
    WS = StrToBSTR(Value);
#ifdef _MSXML_H_
	Attr->put_text(WS);
#else
    Attr->set_text(WS);
#endif
    ASec->setAttributeNode(Attr, NULL);
    Attr->Release();
    SysFreeString(WS);
    return;
}

void CXMLWritter::AddStringValue(IXMLDOMElement* ASec, LPCSTR AName, BSTR Value)
{
    BSTR WS = StrToBSTR(AName);
    IXMLDOMAttribute* Attr = NULL;

    fDoc->createAttribute(WS, &Attr);
    SysFreeString(WS);
#ifdef _MSXML_H_
	Attr->put_text(Value);
#else
    Attr->set_text(Value);
#endif
    ASec->setAttributeNode(Attr, NULL);
    Attr->Release();
    return;
}
