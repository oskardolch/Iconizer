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

#ifndef _msxml_DLL_H
#define _msxml_DLL_H
#if __GNUC__ >= 3
#pragma GCC system_header
#endif

#ifdef __cplusplus
extern "C" {
#endif

#include <basetyps.h>
#include <wtypes.h>

#include <ocidl.h>

//
// GUID constant declarations
//

const IID DIID_IXMLDOMImplementation = {0x2933BF8F,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMNode = {0x2933BF80,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMNodeList = {0x2933BF82,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMNamedNodeMap = {0x2933BF83,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMDocument = {0x2933BF81,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMDocumentType = {0x2933BF8B,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMElement = {0x2933BF86,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMAttribute = {0x2933BF85,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMDocumentFragment = {0x3EFAA413,0x272F,0x11D2,{0x83,0x6F,0x00,0x00,0xF8,0x7A,0x77,0x82}};
const IID DIID_IXMLDOMCharacterData = {0x2933BF84,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMText = {0x2933BF87,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMComment = {0x2933BF88,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMCDATASection = {0x2933BF8A,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMProcessingInstruction = {0x2933BF89,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMEntityReference = {0x2933BF8E,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMParseError = {0x3EFAA426,0x272F,0x11D2,{0x83,0x6F,0x00,0x00,0xF8,0x7A,0x77,0x82}};
const IID DIID_IXMLDOMNotation = {0x2933BF8C,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLDOMEntity = {0x2933BF8D,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXTLRuntime = {0x3EFAA425,0x272F,0x11D2,{0x83,0x6F,0x00,0x00,0xF8,0x7A,0x77,0x82}};
const IID DIID_XMLDOMDocumentEvents = {0x3EFAA427,0x272F,0x11D2,{0x83,0x6F,0x00,0x00,0xF8,0x7A,0x77,0x82}};
const CLSID CLASS_DOMDocument = {0x2933BF90,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const CLSID CLASS_DOMFreeThreadedDocument = {0x2933BF91,0x7B36,0x11D2,{0xB2,0x0E,0x00,0xC0,0x4F,0x98,0x3E,0x60}};
const IID DIID_IXMLHttpRequest = {0xED8C108D,0x4349,0x11D2,{0x91,0xA4,0x00,0xC0,0x4F,0x79,0x69,0xE8}};
const CLSID CLASS_XMLHTTPRequest = {0xED8C108E,0x4349,0x11D2,{0x91,0xA4,0x00,0xC0,0x4F,0x79,0x69,0xE8}};
const IID DIID_IXMLDSOControl = {0x310AFA62,0x0575,0x11D2,{0x9C,0xA9,0x00,0x60,0xB0,0xEC,0x3D,0x39}};
const CLSID CLASS_XMLDSOControl = {0x550DDA30,0x0541,0x11D2,{0x9C,0xA9,0x00,0x60,0xB0,0xEC,0x3D,0x39}};
const IID DIID_IXMLElementCollection = {0x65725580,0x9B5D,0x11D0,{0x9B,0xFE,0x00,0xC0,0x4F,0xC9,0x9C,0x8E}};
const IID DIID_IXMLDocument = {0xF52E2B61,0x18A1,0x11D1,{0xB1,0x05,0x00,0x80,0x5F,0x49,0x91,0x6B}};
const IID DIID_IXMLElement = {0x3F7F31AC,0xE15F,0x11D0,{0x9C,0x25,0x00,0xC0,0x4F,0xC9,0x9C,0x8E}};
const IID IID_IXMLDocument2 = {0x2B8DE2FE,0x8D2D,0x11D1,{0xB2,0xFC,0x00,0xC0,0x4F,0xD9,0x15,0xA9}};
const IID DIID_IXMLElement2 = {0x2B8DE2FF,0x8D2D,0x11D1,{0xB2,0xFC,0x00,0xC0,0x4F,0xD9,0x15,0xA9}};
const IID DIID_IXMLAttribute = {0xD4D4A0FC,0x3B73,0x11D1,{0xB2,0xB4,0x00,0xC0,0x4F,0xB9,0x25,0x96}};
const IID IID_IXMLError = {0x948C5AD3,0xC58D,0x11D0,{0x9C,0x0B,0x00,0xC0,0x4F,0xC9,0x9C,0x8E}};
const CLSID CLASS_XMLDocument = {0xCFC399AF,0xD876,0x11D0,{0x9C,0x10,0x00,0xC0,0x4F,0xC9,0x9C,0x8E}};


// Constants for enum tagDOMNodeType
#ifndef _tagDOMNodeType_
#define _tagDOMNodeType_
typedef enum _tagDOMNodeType {
  NODE_INVALID = 0x00000000,
  NODE_ELEMENT = 0x00000001,
  NODE_ATTRIBUTE = 0x00000002,
  NODE_TEXT = 0x00000003,
  NODE_CDATA_SECTION = 0x00000004,
  NODE_ENTITY_REFERENCE = 0x00000005,
  NODE_ENTITY = 0x00000006,
  NODE_PROCESSING_INSTRUCTION = 0x00000007,
  NODE_COMMENT = 0x00000008,
  NODE_DOCUMENT = 0x00000009,
  NODE_DOCUMENT_TYPE = 0x0000000A,
  NODE_DOCUMENT_FRAGMENT = 0x0000000B,
  NODE_NOTATION = 0x0000000C
} tagDOMNodeType;
#endif

typedef struct __xml_error {
    unsigned int _nLine;
    BSTR _pchBuf;
    unsigned int _cchBuf;
    unsigned int _ich;
    BSTR _pszFound;
    BSTR _pszExpected;
    unsigned long _reserved1;
    unsigned long _reserved2;
} _xml_error, *P_xml_error;

// Constants for enum tagXMLEMEM_TYPE
#ifndef _tagXMLEMEM_TYPE_
#define _tagXMLEMEM_TYPE_
typedef enum _tagXMLEMEM_TYPE {
  XMLELEMTYPE_ELEMENT = 0x00000000,
  XMLELEMTYPE_TEXT = 0x00000001,
  XMLELEMTYPE_COMMENT = 0x00000002,
  XMLELEMTYPE_DOCUMENT = 0x00000003,
  XMLELEMTYPE_DTD = 0x00000004,
  XMLELEMTYPE_PI = 0x00000005,
  XMLELEMTYPE_OTHER = 0x00000006
} tagXMLEMEM_TYPE;
#endif
// *********************************************************************//
// Forward declaration of types defined in TypeLibrary
// *********************************************************************//

struct IXMLDOMImplementation;
struct IXMLDOMNode;
struct IXMLDOMNodeList;
struct IXMLDOMNamedNodeMap;
struct IXMLDOMDocument;
struct IXMLDOMDocumentType;
struct IXMLDOMElement;
struct IXMLDOMAttribute;
struct IXMLDOMDocumentFragment;
struct IXMLDOMCharacterData;
struct IXMLDOMText;
struct IXMLDOMComment;
struct IXMLDOMCDATASection;
struct IXMLDOMProcessingInstruction;
struct IXMLDOMEntityReference;
struct IXMLDOMParseError;
struct IXMLDOMNotation;
struct IXMLDOMEntity;
struct IXTLRuntime;
struct XMLDOMDocumentEvents;
struct IXMLHttpRequest;
struct IXMLDSOControl;
struct IXMLElementCollection;
struct IXMLDocument;
struct IXMLElement;
struct IXMLDocument2;
struct IXMLElement2;
struct IXMLAttribute;
struct IXMLError;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library
// *********************************************************************//

class DOMDocument;
class DOMFreeThreadedDocument;
class XMLHTTPRequest;
class XMLDSOControl;
class XMLDocument;

typedef long DOMNodeType;
typedef long XMLELEM_TYPE;

struct IXMLDOMImplementation : IDispatch
{
    virtual HRESULT __stdcall hasFeature(BSTR feature, BSTR version, VARIANT_BOOL *hasFeature) = 0;
};

struct IXMLDOMImplementationDisp : IDispatch
{
    BOOL hasFeature(LPTSTR feature, LPTSTR version);
};

struct IXMLDOMNode : IDispatch
{
    virtual HRESULT __stdcall get_nodeName(BSTR *name) = 0;
    virtual HRESULT __stdcall get_nodeValue(VARIANT *value) = 0;
    virtual HRESULT __stdcall set_nodeValue(VARIANT value) = 0;
    virtual HRESULT __stdcall get_nodeType(long *type) = 0;
    virtual HRESULT __stdcall get_parentNode(IXMLDOMNode* *parent) = 0;
    virtual HRESULT __stdcall get_childNodes(IXMLDOMNodeList* *childList) = 0;
    virtual HRESULT __stdcall get_firstChild(IXMLDOMNode* *firstChild) = 0;
    virtual HRESULT __stdcall get_lastChild(IXMLDOMNode* *lastChild) = 0;
    virtual HRESULT __stdcall get_previousSibling(IXMLDOMNode* *previousSibling) = 0;
    virtual HRESULT __stdcall get_nextSibling(IXMLDOMNode* *nextSibling) = 0;
    virtual HRESULT __stdcall get_attributes(IXMLDOMNamedNodeMap* *attributeMap) = 0;
    virtual HRESULT __stdcall insertBefore(IXMLDOMNode* newChild, VARIANT refChild,
      IXMLDOMNode* *outNewChild) = 0;
    virtual HRESULT __stdcall replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild,
      IXMLDOMNode* *outOldChild) = 0;
    virtual HRESULT __stdcall removeChild(IXMLDOMNode* childNode, IXMLDOMNode* *oldChild) = 0;
    virtual HRESULT __stdcall appendChild(IXMLDOMNode* newChild, IXMLDOMNode* *outNewChild) = 0;
    virtual HRESULT __stdcall hasChildNodes(VARIANT_BOOL *hasChild) = 0;
    virtual HRESULT __stdcall get_ownerDocument(IXMLDOMDocument* *DOMDocument) = 0;
    virtual HRESULT __stdcall cloneNode(VARIANT_BOOL deep, IXMLDOMNode* *cloneRoot) = 0;
    virtual HRESULT __stdcall get_nodeTypeString(BSTR *nodeType) = 0;
    virtual HRESULT __stdcall get_text(BSTR *text) = 0;
    virtual HRESULT __stdcall set_text(BSTR text) = 0;
    virtual HRESULT __stdcall get_specified(VARIANT_BOOL *isSpecified) = 0;
    virtual HRESULT __stdcall get_definition(IXMLDOMNode* *definitionNode) = 0;
    virtual HRESULT __stdcall get_nodeTypedValue(VARIANT *typedValue) = 0;
    virtual HRESULT __stdcall set_nodeTypedValue(VARIANT typedValue) = 0;
    virtual HRESULT __stdcall get_dataType(VARIANT *dataTypeName) = 0;
    virtual HRESULT __stdcall set_dataType(VARIANT dataTypeName) = 0;
    virtual HRESULT __stdcall get_xml(BSTR *xmlString) = 0;
    virtual HRESULT __stdcall transformNode(IXMLDOMNode* stylesheet, BSTR *xmlString) = 0;
    virtual HRESULT __stdcall selectNodes(BSTR queryString, IXMLDOMNodeList* *resultList) = 0;
    virtual HRESULT __stdcall selectSingleNode(BSTR queryString, IXMLDOMNode* *resultNode) = 0;
    virtual HRESULT __stdcall get_parsed(VARIANT_BOOL *isParsed) = 0;
    virtual HRESULT __stdcall get_namespaceURI(BSTR *namespaceURI) = 0;
    virtual HRESULT __stdcall get_prefix(BSTR *prefixString) = 0;
    virtual HRESULT __stdcall get_baseName(BSTR *nameString) = 0;
    virtual HRESULT __stdcall transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject) = 0;
};

struct IXMLDOMNodeDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
};

struct IXMLDOMNodeList : IDispatch
{
    virtual HRESULT __stdcall get_item(long index, IXMLDOMNode* *listItem) = 0;
    virtual HRESULT __stdcall get_length(long *listLength) = 0;
    virtual HRESULT __stdcall nextNode(IXMLDOMNode* *nextItem) = 0;
    virtual HRESULT __stdcall reset() = 0;
    virtual HRESULT __stdcall get__newEnum(IUnknown * *ppUnk) = 0;
};

struct IXMLDOMNodeListDisp : IDispatch
{
    IXMLDOMNode* get_item(long index);
    long get_length();
    IXMLDOMNode* nextNode();
    void reset();
    IUnknown * get__newEnum();
};

struct IXMLDOMNamedNodeMap : IDispatch
{
    virtual HRESULT __stdcall getNamedItem(BSTR name, IXMLDOMNode* *namedItem) = 0;
    virtual HRESULT __stdcall setNamedItem(IXMLDOMNode* newItem, IXMLDOMNode* *nameItem) = 0;
    virtual HRESULT __stdcall removeNamedItem(BSTR name, IXMLDOMNode* *namedItem) = 0;
    virtual HRESULT __stdcall get_item(long index, IXMLDOMNode* *listItem) = 0;
    virtual HRESULT __stdcall get_length(long *listLength) = 0;
    virtual HRESULT __stdcall getQualifiedItem(BSTR baseName, BSTR namespaceURI,
      IXMLDOMNode* *qualifiedItem) = 0;
    virtual HRESULT __stdcall removeQualifiedItem(BSTR baseName, BSTR namespaceURI,
      IXMLDOMNode* *qualifiedItem) = 0;
    virtual HRESULT __stdcall nextNode(IXMLDOMNode* *nextItem) = 0;
    virtual HRESULT __stdcall reset() = 0;
    virtual HRESULT __stdcall get__newEnum(IUnknown * *ppUnk) = 0;
};

struct IXMLDOMNamedNodeMapDisp : IDispatch
{
    IXMLDOMNode* getNamedItem(LPTSTR name);
    IXMLDOMNode* setNamedItem(IXMLDOMNode* newItem);
    IXMLDOMNode* removeNamedItem(LPTSTR name);
    IXMLDOMNode* get_item(long index);
    long get_length();
    IXMLDOMNode* getQualifiedItem(LPTSTR baseName, LPTSTR namespaceURI);
    IXMLDOMNode* removeQualifiedItem(LPTSTR baseName, LPTSTR namespaceURI);
    IXMLDOMNode* nextNode();
    void reset();
    IUnknown * get__newEnum();
};

struct IXMLDOMDocument : IXMLDOMNode
{
    virtual HRESULT __stdcall get_doctype(IXMLDOMDocumentType* *documentType) = 0;
    virtual HRESULT __stdcall get_implementation(IXMLDOMImplementation* *impl) = 0;
    virtual HRESULT __stdcall get_documentElement(IXMLDOMElement* *DOMElement) = 0;
    virtual HRESULT __stdcall set_documentElement(IXMLDOMElement* DOMElement) = 0;
    virtual HRESULT __stdcall createElement(BSTR tagName, IXMLDOMElement* *element) = 0;
    virtual HRESULT __stdcall createDocumentFragment(IXMLDOMDocumentFragment* *docFrag) = 0;
    virtual HRESULT __stdcall createTextNode(BSTR data, IXMLDOMText* *text) = 0;
    virtual HRESULT __stdcall createComment(BSTR data, IXMLDOMComment* *comment) = 0;
    virtual HRESULT __stdcall createCDATASection(BSTR data, IXMLDOMCDATASection* *cdata) = 0;
    virtual HRESULT __stdcall createProcessingInstruction(BSTR target, BSTR data,
      IXMLDOMProcessingInstruction* *pi) = 0;
    virtual HRESULT __stdcall createAttribute(BSTR name, IXMLDOMAttribute* *attribute) = 0;
    virtual HRESULT __stdcall createEntityReference(BSTR name, IXMLDOMEntityReference* *entityRef) = 0;
    virtual HRESULT __stdcall getElementsByTagName(BSTR tagName, IXMLDOMNodeList* *resultList) = 0;
    virtual HRESULT __stdcall createNode(VARIANT type, BSTR name, BSTR namespaceURI,
      IXMLDOMNode* *node) = 0;
    virtual HRESULT __stdcall nodeFromID(BSTR idString, IXMLDOMNode* *node) = 0;
    virtual HRESULT __stdcall load(VARIANT xmlSource, VARIANT_BOOL *isSuccessful) = 0;
    virtual HRESULT __stdcall get_readyState(long *value) = 0;
    virtual HRESULT __stdcall get_parseError(IXMLDOMParseError* *errorObj) = 0;
    virtual HRESULT __stdcall get_url(BSTR *urlString) = 0;
    virtual HRESULT __stdcall get_async(VARIANT_BOOL *isAsync) = 0;
    virtual HRESULT __stdcall set_async(VARIANT_BOOL isAsync) = 0;
    virtual HRESULT __stdcall abort() = 0;
    virtual HRESULT __stdcall loadXML(BSTR bstrXML, VARIANT_BOOL *isSuccessful) = 0;
    virtual HRESULT __stdcall save(VARIANT destination) = 0;
    virtual HRESULT __stdcall get_validateOnParse(VARIANT_BOOL *isValidating) = 0;
    virtual HRESULT __stdcall set_validateOnParse(VARIANT_BOOL isValidating) = 0;
    virtual HRESULT __stdcall get_resolveExternals(VARIANT_BOOL *isResolving) = 0;
    virtual HRESULT __stdcall set_resolveExternals(VARIANT_BOOL isResolving) = 0;
    virtual HRESULT __stdcall get_preserveWhiteSpace(VARIANT_BOOL *isPreserving) = 0;
    virtual HRESULT __stdcall set_preserveWhiteSpace(VARIANT_BOOL isPreserving) = 0;
    virtual HRESULT __stdcall set_onreadystatechange(VARIANT Value) = 0;
    virtual HRESULT __stdcall set_ondataavailable(VARIANT Value) = 0;
    virtual HRESULT __stdcall set_ontransformnode(VARIANT Value) = 0;
};

struct IXMLDOMDocumentDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
    IXMLDOMDocumentType* get_doctype();
    IXMLDOMImplementation* get_implementation();
    IXMLDOMElement* get_documentElement();
    void set_documentElement(IXMLDOMElement* Value);
    IXMLDOMElement* createElement(LPTSTR tagName);
    IXMLDOMDocumentFragment* createDocumentFragment();
    IXMLDOMText* createTextNode(LPTSTR data);
    IXMLDOMComment* createComment(LPTSTR data);
    IXMLDOMCDATASection* createCDATASection(LPTSTR data);
    IXMLDOMProcessingInstruction* createProcessingInstruction(LPTSTR target, LPTSTR data);
    IXMLDOMAttribute* createAttribute(LPTSTR name);
    IXMLDOMEntityReference* createEntityReference(LPTSTR name);
    IXMLDOMNodeList* getElementsByTagName(LPTSTR tagName);
    IXMLDOMNode* createNode(VARIANT type, LPTSTR name, LPTSTR namespaceURI);
    IXMLDOMNode* nodeFromID(LPTSTR idString);
    BOOL load(VARIANT xmlSource);
    long get_readyState();
    IXMLDOMParseError* get_parseError();
    LPTSTR get_url();
    BOOL get_async();
    void set_async(BOOL Value);
    void abort();
    BOOL loadXML(LPTSTR bstrXML);
    void save(VARIANT destination);
    BOOL get_validateOnParse();
    void set_validateOnParse(BOOL Value);
    BOOL get_resolveExternals();
    void set_resolveExternals(BOOL Value);
    BOOL get_preserveWhiteSpace();
    void set_preserveWhiteSpace(BOOL Value);
    void set_onreadystatechange(VARIANT Value);
    void set_ondataavailable(VARIANT Value);
    void set_ontransformnode(VARIANT Value);
};

struct IXMLDOMDocumentType : IXMLDOMNode
{
    virtual HRESULT __stdcall get_name(BSTR *rootName) = 0;
    virtual HRESULT __stdcall get_entities(IXMLDOMNamedNodeMap* *entityMap) = 0;
    virtual HRESULT __stdcall get_notations(IXMLDOMNamedNodeMap* *notationMap) = 0;
};

struct IXMLDOMDocumentTypeDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
    LPTSTR get_name();
    IXMLDOMNamedNodeMap* get_entities();
    IXMLDOMNamedNodeMap* get_notations();
};

struct IXMLDOMElement : IXMLDOMNode
{
    virtual HRESULT __stdcall get_tagName(BSTR *tagName) = 0;
    virtual HRESULT __stdcall getAttribute(BSTR name, VARIANT *value) = 0;
    virtual HRESULT __stdcall setAttribute(BSTR name, VARIANT value) = 0;
    virtual HRESULT __stdcall removeAttribute(BSTR name) = 0;
    virtual HRESULT __stdcall getAttributeNode(BSTR name, IXMLDOMAttribute* *attributeNode) = 0;
    virtual HRESULT __stdcall setAttributeNode(IXMLDOMAttribute* DOMAttribute,
      IXMLDOMAttribute* *attributeNode) = 0;
    virtual HRESULT __stdcall removeAttributeNode(IXMLDOMAttribute* DOMAttribute,
      IXMLDOMAttribute* *attributeNode) = 0;
    virtual HRESULT __stdcall getElementsByTagName(BSTR tagName, IXMLDOMNodeList* *resultList) = 0;
    virtual HRESULT __stdcall normalize() = 0;
};

struct IXMLDOMElementDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
    LPTSTR get_tagName();
    VARIANT getAttribute(LPTSTR name);
    void setAttribute(LPTSTR name, VARIANT value);
    void removeAttribute(LPTSTR name);
    IXMLDOMAttribute* getAttributeNode(LPTSTR name);
    IXMLDOMAttribute* setAttributeNode(IXMLDOMAttribute* DOMAttribute);
    IXMLDOMAttribute* removeAttributeNode(IXMLDOMAttribute* DOMAttribute);
    IXMLDOMNodeList* getElementsByTagName(LPTSTR tagName);
    void normalize();
};

struct IXMLDOMAttribute : IXMLDOMNode
{
    virtual HRESULT __stdcall get_name(BSTR *attributeName) = 0;
    virtual HRESULT __stdcall get_value(VARIANT *attributeValue) = 0;
    virtual HRESULT __stdcall set_value(VARIANT attributeValue) = 0;
};

struct IXMLDOMAttributeDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
    LPTSTR get_name();
    VARIANT get_value();
    void set_value(VARIANT Value);
};

struct IXMLDOMDocumentFragment : IXMLDOMNode
{
};

struct IXMLDOMDocumentFragmentDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
};

struct IXMLDOMCharacterData : IXMLDOMNode
{
    virtual HRESULT __stdcall get_data(BSTR *data) = 0;
    virtual HRESULT __stdcall set_data(BSTR data) = 0;
    virtual HRESULT __stdcall get_length(long *dataLength) = 0;
    virtual HRESULT __stdcall substringData(long offset, long count, BSTR *data) = 0;
    virtual HRESULT __stdcall appendData(BSTR data) = 0;
    virtual HRESULT __stdcall insertData(long offset, BSTR data) = 0;
    virtual HRESULT __stdcall deleteData(long offset, long count) = 0;
    virtual HRESULT __stdcall replaceData(long offset, long count, BSTR data) = 0;
};

struct IXMLDOMCharacterDataDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
    LPTSTR get_data();
    void set_data(LPTSTR Value);
    long get_length();
    LPTSTR substringData(long offset, long count);
    void appendData(LPTSTR data);
    void insertData(long offset, LPTSTR data);
    void deleteData(long offset, long count);
    void replaceData(long offset, long count, LPTSTR data);
};

struct IXMLDOMText : IXMLDOMCharacterData
{
    virtual HRESULT __stdcall splitText(long offset, IXMLDOMText* *rightHandTextNode) = 0;
};

struct IXMLDOMTextDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
    LPTSTR get_data();
    void set_data(LPTSTR Value);
    long get_length();
    LPTSTR substringData(long offset, long count);
    void appendData(LPTSTR data);
    void insertData(long offset, LPTSTR data);
    void deleteData(long offset, long count);
    void replaceData(long offset, long count, LPTSTR data);
    IXMLDOMText* splitText(long offset);
};

struct IXMLDOMComment : IXMLDOMCharacterData
{
};

struct IXMLDOMCommentDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
    LPTSTR get_data();
    void set_data(LPTSTR Value);
    long get_length();
    LPTSTR substringData(long offset, long count);
    void appendData(LPTSTR data);
    void insertData(long offset, LPTSTR data);
    void deleteData(long offset, long count);
    void replaceData(long offset, long count, LPTSTR data);
};

struct IXMLDOMCDATASection : IXMLDOMText
{
};

struct IXMLDOMCDATASectionDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
    LPTSTR get_data();
    void set_data(LPTSTR Value);
    long get_length();
    LPTSTR substringData(long offset, long count);
    void appendData(LPTSTR data);
    void insertData(long offset, LPTSTR data);
    void deleteData(long offset, long count);
    void replaceData(long offset, long count, LPTSTR data);
    IXMLDOMText* splitText(long offset);
};

struct IXMLDOMProcessingInstruction : IXMLDOMNode
{
    virtual HRESULT __stdcall get_target(BSTR *name) = 0;
    virtual HRESULT __stdcall get_data(BSTR *value) = 0;
    virtual HRESULT __stdcall set_data(BSTR value) = 0;
};

struct IXMLDOMProcessingInstructionDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
    LPTSTR get_target();
    LPTSTR get_data();
    void set_data(LPTSTR Value);
};

struct IXMLDOMEntityReference : IXMLDOMNode
{
};

struct IXMLDOMEntityReferenceDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
};

struct IXMLDOMParseError : IDispatch
{
    virtual HRESULT __stdcall get_errorCode(long *errorCode) = 0;
    virtual HRESULT __stdcall get_url(BSTR *urlString) = 0;
    virtual HRESULT __stdcall get_reason(BSTR *reasonString) = 0;
    virtual HRESULT __stdcall get_srcText(BSTR *sourceString) = 0;
    virtual HRESULT __stdcall get_line(long *lineNumber) = 0;
    virtual HRESULT __stdcall get_linepos(long *linePosition) = 0;
    virtual HRESULT __stdcall get_filepos(long *filePosition) = 0;
};

struct IXMLDOMParseErrorDisp : IDispatch
{
    long get_errorCode();
    LPTSTR get_url();
    LPTSTR get_reason();
    LPTSTR get_srcText();
    long get_line();
    long get_linepos();
    long get_filepos();
};

struct IXMLDOMNotation : IXMLDOMNode
{
    virtual HRESULT __stdcall get_publicId(VARIANT *publicId) = 0;
    virtual HRESULT __stdcall get_systemId(VARIANT *systemId) = 0;
};

struct IXMLDOMNotationDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
    VARIANT get_publicId();
    VARIANT get_systemId();
};

struct IXMLDOMEntity : IXMLDOMNode
{
    virtual HRESULT __stdcall get_publicId(VARIANT *publicId) = 0;
    virtual HRESULT __stdcall get_systemId(VARIANT *systemId) = 0;
    virtual HRESULT __stdcall get_notationName(BSTR *name) = 0;
};

struct IXMLDOMEntityDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
    VARIANT get_publicId();
    VARIANT get_systemId();
    LPTSTR get_notationName();
};

struct IXTLRuntime : IXMLDOMNode
{
    virtual HRESULT __stdcall uniqueID(IXMLDOMNode* pNode, long *pID) = 0;
    virtual HRESULT __stdcall depth(IXMLDOMNode* pNode, long *pDepth) = 0;
    virtual HRESULT __stdcall childNumber(IXMLDOMNode* pNode, long *pNumber) = 0;
    virtual HRESULT __stdcall ancestorChildNumber(BSTR bstrNodeName, IXMLDOMNode* pNode,
      long *pNumber) = 0;
    virtual HRESULT __stdcall absoluteChildNumber(IXMLDOMNode* pNode, long *pNumber) = 0;
    virtual HRESULT __stdcall formatIndex(long lIndex, BSTR bstrFormat, BSTR *pbstrFormattedString) = 0;
    virtual HRESULT __stdcall formatNumber(double dblNumber, BSTR bstrFormat, BSTR *pbstrFormattedString) = 0;
    virtual HRESULT __stdcall formatDate(VARIANT varDate, BSTR bstrFormat, VARIANT varDestLocale,
      BSTR *pbstrFormattedString) = 0;
    virtual HRESULT __stdcall formatTime(VARIANT varTime, BSTR bstrFormat, VARIANT varDestLocale,
      BSTR *pbstrFormattedString) = 0;
};

struct IXTLRuntimeDisp : IDispatch
{
    LPTSTR get_nodeName();
    VARIANT get_nodeValue();
    void set_nodeValue(VARIANT Value);
    long get_nodeType();
    IXMLDOMNode* get_parentNode();
    IXMLDOMNodeList* get_childNodes();
    IXMLDOMNode* get_firstChild();
    IXMLDOMNode* get_lastChild();
    IXMLDOMNode* get_previousSibling();
    IXMLDOMNode* get_nextSibling();
    IXMLDOMNamedNodeMap* get_attributes();
    IXMLDOMNode* insertBefore(IXMLDOMNode* newChild, VARIANT refChild);
    IXMLDOMNode* replaceChild(IXMLDOMNode* newChild, IXMLDOMNode* oldChild);
    IXMLDOMNode* removeChild(IXMLDOMNode* childNode);
    IXMLDOMNode* appendChild(IXMLDOMNode* newChild);
    BOOL hasChildNodes();
    IXMLDOMDocument* get_ownerDocument();
    IXMLDOMNode* cloneNode(BOOL deep);
    LPTSTR get_nodeTypeString();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    BOOL get_specified();
    IXMLDOMNode* get_definition();
    VARIANT get_nodeTypedValue();
    void set_nodeTypedValue(VARIANT Value);
    VARIANT get_dataType();
    void set_dataType(VARIANT Value);
    LPTSTR get_xml();
    LPTSTR transformNode(IXMLDOMNode* stylesheet);
    IXMLDOMNodeList* selectNodes(LPTSTR queryString);
    IXMLDOMNode* selectSingleNode(LPTSTR queryString);
    BOOL get_parsed();
    LPTSTR get_namespaceURI();
    LPTSTR get_prefix();
    LPTSTR get_baseName();
    void transformNodeToObject(IXMLDOMNode* stylesheet, VARIANT outputObject);
    long uniqueID(IXMLDOMNode* pNode);
    long depth(IXMLDOMNode* pNode);
    long childNumber(IXMLDOMNode* pNode);
    long ancestorChildNumber(LPTSTR bstrNodeName, IXMLDOMNode* pNode);
    long absoluteChildNumber(IXMLDOMNode* pNode);
    LPTSTR formatIndex(long lIndex, LPTSTR bstrFormat);
    LPTSTR formatNumber(double dblNumber, LPTSTR bstrFormat);
    LPTSTR formatDate(VARIANT varDate, LPTSTR bstrFormat, VARIANT varDestLocale);
    LPTSTR formatTime(VARIANT varTime, LPTSTR bstrFormat, VARIANT varDestLocale);
};

struct XMLDOMDocumentEvents : IDispatch
{
    void ondataavailable();
    void onreadystatechange();
};

IXMLDOMDocument* CreateDOMDocument();

IXMLDOMDocument* CreateDOMFreeThreadedDocument();

struct IXMLHttpRequest : IDispatch
{
    virtual HRESULT __stdcall open(BSTR bstrMethod, BSTR bstrUrl, VARIANT varAsync,
      VARIANT bstrUser, VARIANT bstrPassword) = 0;
    virtual HRESULT __stdcall setRequestHeader(BSTR bstrHeader, BSTR bstrValue) = 0;
    virtual HRESULT __stdcall getResponseHeader(BSTR bstrHeader, BSTR *pbstrValue) = 0;
    virtual HRESULT __stdcall getAllResponseHeaders(BSTR *pbstrHeaders) = 0;
    virtual HRESULT __stdcall send(VARIANT varBody) = 0;
    virtual HRESULT __stdcall abort() = 0;
    virtual HRESULT __stdcall get_status(long *plStatus) = 0;
    virtual HRESULT __stdcall get_statusText(BSTR *pbstrStatus) = 0;
    virtual HRESULT __stdcall get_responseXML(IDispatch * *ppBody) = 0;
    virtual HRESULT __stdcall get_responseText(BSTR *pbstrBody) = 0;
    virtual HRESULT __stdcall get_responseBody(VARIANT *pvarBody) = 0;
    virtual HRESULT __stdcall get_responseStream(VARIANT *pvarBody) = 0;
    virtual HRESULT __stdcall get_readyState(long *plState) = 0;
    virtual HRESULT __stdcall set_onreadystatechange(IDispatch * Value) = 0;
};

struct IXMLHttpRequestDisp : IDispatch
{
    void open(LPTSTR bstrMethod, LPTSTR bstrUrl, VARIANT varAsync, VARIANT bstrUser,
      VARIANT bstrPassword);
    void setRequestHeader(LPTSTR bstrHeader, LPTSTR bstrValue);
    LPTSTR getResponseHeader(LPTSTR bstrHeader);
    LPTSTR getAllResponseHeaders();
    void send(VARIANT varBody);
    void abort();
    long get_status();
    LPTSTR get_statusText();
    IDispatch * get_responseXML();
    LPTSTR get_responseText();
    VARIANT get_responseBody();
    VARIANT get_responseStream();
    long get_readyState();
    void set_onreadystatechange(IDispatch * Value);
};

IXMLHttpRequest* CreateXMLHTTPRequest();

struct IXMLDSOControl : IDispatch
{
    virtual HRESULT __stdcall get_XMLDocument(IXMLDOMDocument* *ppDoc) = 0;
    virtual HRESULT __stdcall set_XMLDocument(IXMLDOMDocument* ppDoc) = 0;
    virtual HRESULT __stdcall get_JavaDSOCompatible(long *fJavaDSOCompatible) = 0;
    virtual HRESULT __stdcall set_JavaDSOCompatible(long fJavaDSOCompatible) = 0;
    virtual HRESULT __stdcall get_readyState(long *state) = 0;
};

struct IXMLDSOControlDisp : IDispatch
{
    IXMLDOMDocument* get_XMLDocument();
    void set_XMLDocument(IXMLDOMDocument* Value);
    long get_JavaDSOCompatible();
    void set_JavaDSOCompatible(long Value);
    long get_readyState();
};

IXMLDSOControl* CreateXMLDSOControl();

struct IXMLElementCollection : IDispatch
{
    virtual HRESULT __stdcall set_length(long p) = 0;
    virtual HRESULT __stdcall get_length(long *p) = 0;
    virtual HRESULT __stdcall get__newEnum(IUnknown * *ppUnk) = 0;
    virtual HRESULT __stdcall item(VARIANT var1, VARIANT var2, IDispatch * *ppDisp) = 0;
};

struct IXMLElementCollectionDisp : IDispatch
{
    void set_length(long Value);
    long get_length();
    IUnknown * get__newEnum();
    IDispatch * item(VARIANT var1, VARIANT var2);
};

struct IXMLDocument : IDispatch
{
    virtual HRESULT __stdcall get_root(IXMLElement* *p) = 0;
    virtual HRESULT __stdcall get_fileSize(BSTR *p) = 0;
    virtual HRESULT __stdcall get_fileModifiedDate(BSTR *p) = 0;
    virtual HRESULT __stdcall get_fileUpdatedDate(BSTR *p) = 0;
    virtual HRESULT __stdcall get_url(BSTR *p) = 0;
    virtual HRESULT __stdcall set_url(BSTR p) = 0;
    virtual HRESULT __stdcall get_mimeType(BSTR *p) = 0;
    virtual HRESULT __stdcall get_readyState(long *pl) = 0;
    virtual HRESULT __stdcall get_charset(BSTR *p) = 0;
    virtual HRESULT __stdcall set_charset(BSTR p) = 0;
    virtual HRESULT __stdcall get_version(BSTR *p) = 0;
    virtual HRESULT __stdcall get_doctype(BSTR *p) = 0;
    virtual HRESULT __stdcall get_dtdURL(BSTR *p) = 0;
    virtual HRESULT __stdcall createElement(VARIANT vType, VARIANT var1, IXMLElement* *ppElem) = 0;
};

struct IXMLDocumentDisp : IDispatch
{
    IXMLElement* get_root();
    LPTSTR get_fileSize();
    LPTSTR get_fileModifiedDate();
    LPTSTR get_fileUpdatedDate();
    LPTSTR get_url();
    void set_url(LPTSTR Value);
    LPTSTR get_mimeType();
    long get_readyState();
    LPTSTR get_charset();
    void set_charset(LPTSTR Value);
    LPTSTR get_version();
    LPTSTR get_doctype();
    LPTSTR get_dtdURL();
    IXMLElement* createElement(VARIANT vType, VARIANT var1);
};

struct IXMLElement : IDispatch
{
    virtual HRESULT __stdcall get_tagName(BSTR *p) = 0;
    virtual HRESULT __stdcall set_tagName(BSTR p) = 0;
    virtual HRESULT __stdcall get_parent(IXMLElement* *ppParent) = 0;
    virtual HRESULT __stdcall setAttribute(BSTR strPropertyName, VARIANT PropertyValue) = 0;
    virtual HRESULT __stdcall getAttribute(BSTR strPropertyName, VARIANT *PropertyValue) = 0;
    virtual HRESULT __stdcall removeAttribute(BSTR strPropertyName) = 0;
    virtual HRESULT __stdcall get_children(IXMLElementCollection* *pp) = 0;
    virtual HRESULT __stdcall get_type(long *plType) = 0;
    virtual HRESULT __stdcall get_text(BSTR *p) = 0;
    virtual HRESULT __stdcall set_text(BSTR p) = 0;
    virtual HRESULT __stdcall addChild(IXMLElement* pChildElem, long lIndex, long lReserved) = 0;
    virtual HRESULT __stdcall removeChild(IXMLElement* pChildElem) = 0;
};

struct IXMLElementDisp : IDispatch
{
    LPTSTR get_tagName();
    void set_tagName(LPTSTR Value);
    IXMLElement* get_parent();
    void setAttribute(LPTSTR strPropertyName, VARIANT PropertyValue);
    VARIANT getAttribute(LPTSTR strPropertyName);
    void removeAttribute(LPTSTR strPropertyName);
    IXMLElementCollection* get_children();
    long get_type();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    void addChild(IXMLElement* pChildElem, long lIndex, long lReserved);
    void removeChild(IXMLElement* pChildElem);
};

struct IXMLDocument2 : IDispatch
{
    virtual HRESULT __stdcall get_root(IXMLElement2* *p) = 0;
    virtual HRESULT __stdcall get_fileSize(BSTR *p) = 0;
    virtual HRESULT __stdcall get_fileModifiedDate(BSTR *p) = 0;
    virtual HRESULT __stdcall get_fileUpdatedDate(BSTR *p) = 0;
    virtual HRESULT __stdcall get_url(BSTR *p) = 0;
    virtual HRESULT __stdcall set_url(BSTR p) = 0;
    virtual HRESULT __stdcall get_mimeType(BSTR *p) = 0;
    virtual HRESULT __stdcall get_readyState(long *pl) = 0;
    virtual HRESULT __stdcall get_charset(BSTR *p) = 0;
    virtual HRESULT __stdcall set_charset(BSTR p) = 0;
    virtual HRESULT __stdcall get_version(BSTR *p) = 0;
    virtual HRESULT __stdcall get_doctype(BSTR *p) = 0;
    virtual HRESULT __stdcall get_dtdURL(BSTR *p) = 0;
    virtual HRESULT __stdcall createElement(VARIANT vType, VARIANT var1, IXMLElement2* *ppElem) = 0;
    virtual HRESULT __stdcall get_async(VARIANT_BOOL *pf) = 0;
    virtual HRESULT __stdcall set_async(VARIANT_BOOL pf) = 0;
};

struct IXMLElement2 : IDispatch
{
    virtual HRESULT __stdcall get_tagName(BSTR *p) = 0;
    virtual HRESULT __stdcall set_tagName(BSTR p) = 0;
    virtual HRESULT __stdcall get_parent(IXMLElement2* *ppParent) = 0;
    virtual HRESULT __stdcall setAttribute(BSTR strPropertyName, VARIANT PropertyValue) = 0;
    virtual HRESULT __stdcall getAttribute(BSTR strPropertyName, VARIANT *PropertyValue) = 0;
    virtual HRESULT __stdcall removeAttribute(BSTR strPropertyName) = 0;
    virtual HRESULT __stdcall get_children(IXMLElementCollection* *pp) = 0;
    virtual HRESULT __stdcall get_type(long *plType) = 0;
    virtual HRESULT __stdcall get_text(BSTR *p) = 0;
    virtual HRESULT __stdcall set_text(BSTR p) = 0;
    virtual HRESULT __stdcall addChild(IXMLElement2* pChildElem, long lIndex, long lReserved) = 0;
    virtual HRESULT __stdcall removeChild(IXMLElement2* pChildElem) = 0;
    virtual HRESULT __stdcall get_attributes(IXMLElementCollection* *pp) = 0;
};

struct IXMLElement2Disp : IDispatch
{
    LPTSTR get_tagName();
    void set_tagName(LPTSTR Value);
    IXMLElement2* get_parent();
    void setAttribute(LPTSTR strPropertyName, VARIANT PropertyValue);
    VARIANT getAttribute(LPTSTR strPropertyName);
    void removeAttribute(LPTSTR strPropertyName);
    IXMLElementCollection* get_children();
    long get_type();
    LPTSTR get_text();
    void set_text(LPTSTR Value);
    void addChild(IXMLElement2* pChildElem, long lIndex, long lReserved);
    void removeChild(IXMLElement2* pChildElem);
    IXMLElementCollection* get_attributes();
};

struct IXMLAttribute : IDispatch
{
    virtual HRESULT __stdcall get_name(BSTR *n) = 0;
    virtual HRESULT __stdcall get_value(BSTR *v) = 0;
};

struct IXMLAttributeDisp : IDispatch
{
    LPTSTR get_name();
    LPTSTR get_value();
};

struct IXMLError : IUnknown
{
    virtual HRESULT __stdcall GetErrorInfo(_xml_error *pErrorReturn) = 0;
};

IXMLDocument2* CreateXMLDocument();

#ifdef __cplusplus
}
#endif
#endif
