/*
 * RichEdit - ITextDocument implementation
 *
 * Copyright 2013 by Caibin Chen
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA 02110-1301, USA
 */
#define COBJMACROS              /* For COM objects interfaces definition */

#include "txtdoc.h"

#include "editor.h"

WINE_DEFAULT_DEBUG_CHANNEL(richedit);

/* Struct definitions */

struct tagReTxtDoc {
  IUnknown *outerObj;
  ME_TextEditor *editor;
  ITextDocument iTextDocumentIface;
};

typedef struct {
  ITextRange iTextRangeIface;
  LONG ref;
  long first;
  long lim;
} ReTxtRange;

/* Forward definitions */
static ReTxtRange *ReTxtRange_create(long first, long lim);

/* ITextDocument implementation */

static inline ReTxtDoc *impl_from_ITextDocument(ITextDocument *iface)
{
  return CONTAINING_RECORD(iface, ReTxtDoc, iTextDocumentIface);
}

/* ITextDocument - IUnknown */
static ULONG WINAPI ITextDocument_fnAddRef(ITextDocument *iface)
{
  ReTxtDoc *This = impl_from_ITextDocument(iface);
  TRACE("(%p)\n", iface);
  /* COM aggregation - delegate to outer IUnknown object. */
  return IUnknown_AddRef(This->outerObj);
}

static ULONG WINAPI ITextDocument_fnRelease(ITextDocument *iface)
{
  ReTxtDoc *This = impl_from_ITextDocument(iface);
  TRACE("(%p)\n", iface);
  /* COM aggregation - delegate to outer IUnknown object. */
  return IUnknown_Release(This->outerObj);
}

static HRESULT WINAPI ITextDocument_fnQueryInterface(ITextDocument *iface, REFIID riid, void **ppv)
{
  ReTxtDoc *This = impl_from_ITextDocument(iface);
  TRACE("(%p)->(%s, %p)\n", iface, debugstr_guid(riid), ppv);

  /* COM aggregation - delegate to outer IUnknown object. */
  return IUnknown_QueryInterface(This->outerObj, riid, ppv);
}

/* ITextDocument - IDispatch */

static HRESULT WINAPI ITextDocument_fnGetTypeInfoCount(ITextDocument *iface, UINT *pctinfo)
{
  TRACE("(%p)->(%p)\n", iface, pctinfo);
  /* Ruturn 0 since type info is not implemented */
  *pctinfo = 0;

  return S_OK;
}

static HRESULT WINAPI ITextDocument_fnGetTypeInfo(ITextDocument *iface, UINT iTInfo, LCID lcid,
    ITypeInfo **ppTInfo)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnGetIDsOfNames(ITextDocument *iface, REFIID riid,
    LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnInvoke(
    ITextDocument *me,
    DISPID dispIdMember,
    REFIID riid,
    LCID lcid,
    WORD wFlags,
    DISPPARAMS *pDispParams,
    VARIANT *pVarResult,
    EXCEPINFO *pExcepInfo,
    UINT *puArgErr)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

/* ITextDocument itself */

static HRESULT WINAPI ITextDocument_fnGetName(ITextDocument *iface, BSTR *pName)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnGetSelection(ITextDocument *me, ITextSelection **ppSel)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnGetStoryCount(ITextDocument *iface, LONG *pCount)
{
  TRACE("(%p)->(%p)\n", iface, pCount);
  *pCount = 1; /* GetStory is not implemented. */
  return S_OK;
}

static HRESULT WINAPI ITextDocument_fnGetStoryRanges(ITextDocument *iface,
    ITextStoryRanges **ppStories)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnGetSaved(ITextDocument *iface, LONG *pValue)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnSetSaved(ITextDocument *iface, LONG Value)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnGetDefaultTabStop(ITextDocument *iface, float *pValue)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnSetDefaultTabStop(ITextDocument *iface, float Value)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnNew(ITextDocument *iface)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnOpen(ITextDocument *iface, VARIANT *pVar, LONG Flags,
    LONG CodePage)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnSave(ITextDocument *iface, VARIANT *pVar, LONG Flags,
    LONG CodePage)
{
  FIXME("not implement\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnFreeze(ITextDocument *iface, LONG *pCount)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnUnfreeze(ITextDocument *iface, LONG *pCount)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnBeginEditCollection(ITextDocument *iface)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnEndEditCollection(ITextDocument *iface)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnUndo(ITextDocument *iface, LONG Count, LONG *prop)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnRedo(ITextDocument *iface, LONG Count, LONG *prop)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextDocument_fnRange(ITextDocument *iface, LONG cp1, LONG cp2,
    ITextRange **ppRange)
{
  ReTxtDoc *This = impl_from_ITextDocument(iface);
  ReTxtRange *range;

  TRACE("(%p)->(%d,%d)\n", This, cp1, cp2);

  if (ppRange == NULL) {
    return E_INVALIDARG;
  }

  range = ReTxtRange_create(cp1, cp2);
  if (range == NULL) {
    return E_OUTOFMEMORY;
  }

  *ppRange = &range->iTextRangeIface;

  return S_OK;
}

static HRESULT WINAPI ITextDocument_fnRangeFromPoint(ITextDocument *me, LONG x, LONG y,
    ITextRange **ppRange)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static const ITextDocumentVtbl vtbl = {
  /* IUnknown */
  ITextDocument_fnQueryInterface,
  ITextDocument_fnAddRef,
  ITextDocument_fnRelease,
  /* IDispatch */
  ITextDocument_fnGetTypeInfoCount,
  ITextDocument_fnGetTypeInfo,
  ITextDocument_fnGetIDsOfNames,
  ITextDocument_fnInvoke,
  /* ITextDocument */
  ITextDocument_fnGetName,
  ITextDocument_fnGetSelection,
  ITextDocument_fnGetStoryCount,
  ITextDocument_fnGetStoryRanges,
  ITextDocument_fnGetSaved,
  ITextDocument_fnSetSaved,
  ITextDocument_fnGetDefaultTabStop,
  ITextDocument_fnSetDefaultTabStop,
  ITextDocument_fnNew,
  ITextDocument_fnOpen,
  ITextDocument_fnSave,
  ITextDocument_fnFreeze,
  ITextDocument_fnUnfreeze,
  ITextDocument_fnBeginEditCollection,
  ITextDocument_fnEndEditCollection,
  ITextDocument_fnUndo,
  ITextDocument_fnRedo,
  ITextDocument_fnRange,
  ITextDocument_fnRangeFromPoint
};

/* ITextRange implementation */

static inline ReTxtRange *impl_from_ITextRange(ITextRange *iface)
{
  return CONTAINING_RECORD(iface, ReTxtRange, iTextRangeIface);
}

static HRESULT WINAPI ITextRange_fnQueryInterface(ITextRange* iface, REFIID riid, void **ppvObject)
{
  ReTxtRange *This = impl_from_ITextRange(iface);
  TRACE("(%p)->(%s,%p)\n", This, debugstr_guid(riid), ppvObject);
  if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch)
      || IsEqualIID(riid, &IID_ITextRange)) {
    *ppvObject = iface;
    ITextRange_AddRef(iface);
    return S_OK;
  } else {
    return E_NOINTERFACE;
  }
}

ULONG WINAPI ITextRange_fnAddRef(ITextRange* iface)
{
  ReTxtRange *This = impl_from_ITextRange(iface);
  ULONG ref = InterlockedIncrement(&This->ref);

  TRACE("(%p) ref=%d\n", This, ref);

  return ref;
}

ULONG WINAPI ITextRange_fnRelease(ITextRange* iface)
{
  ReTxtRange *This = impl_from_ITextRange(iface);
  ULONG ref = InterlockedDecrement(&This->ref);

  TRACE("(%p) ref=%d\n", This, ref);

  if (!ref)
  {
    heap_free(This);
  }
  return ref;
}

/* ITextRange - IDispatch methods */
static HRESULT WINAPI ITextRange_fnGetTypeInfoCount(ITextRange* iface, UINT *pctinfo)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetTypeInfo(ITextRange* iface, UINT iTInfo, LCID lcid,
    ITypeInfo **ppTInfo)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetIDsOfNames(ITextRange* iface, REFIID riid, LPOLESTR *rgszNames,
    UINT cNames, LCID lcid, DISPID *rgDispId)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnInvoke(ITextRange* iface, DISPID dispIdMember, REFIID riid, LCID lcid,
    WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo,
    UINT *puArgErr)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

/* ITextRange methods */
static HRESULT WINAPI ITextRange_fnGetText(ITextRange* iface, BSTR *pbstr)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetText(ITextRange* iface, BSTR bstr)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetChar(ITextRange* iface, LONG *pch)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetChar(ITextRange* iface, LONG ch)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetDuplicate(ITextRange* iface, ITextRange **ppRange)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetFormattedText(ITextRange* iface, ITextRange **ppRange)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetFormattedText(ITextRange* iface, ITextRange *pRange)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetStart(ITextRange* iface, LONG *pcpFirst)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetStart(ITextRange* iface, LONG cpFirst)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetEnd(ITextRange* iface, LONG *pcpLim)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetEnd(ITextRange* iface, LONG cpLim)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetFont(ITextRange* iface, ITextFont **pFont)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetFont(ITextRange* iface, ITextFont *pFont)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetPara(ITextRange* iface, ITextPara **ppPara)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetPara(ITextRange* iface, ITextPara *pPara)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetStoryLength(ITextRange* iface, LONG *pcch)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetStoryType(ITextRange* iface, LONG *pValue)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnCollapse(ITextRange* iface, LONG bStart)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnExpand(ITextRange* iface, LONG Unit, LONG *pDelta)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetIndex(ITextRange* iface, LONG Unit, LONG *pIndex)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetIndex(ITextRange* iface, LONG Unit, LONG Index, LONG Extend)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetRange(ITextRange* iface, LONG cpActive, LONG cpOther)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnInRange(ITextRange* iface, ITextRange *pRange, LONG *pb)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnInStory(ITextRange* iface, ITextRange *pRange, LONG *pb)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnIsEqual(ITextRange* iface, ITextRange *pRange, LONG *pb)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSelect(ITextRange* iface)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnStartOf(ITextRange* iface, LONG Unit, LONG Extend, LONG *pDelta)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnEndOf(ITextRange* iface, LONG Unit, LONG Extend, LONG *pDelta)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMove(ITextRange* iface, LONG Unit, LONG Count, LONG *pDelta)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveStart(ITextRange* iface, LONG Unit, LONG Count, LONG *pDelta)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveEnd(ITextRange* iface, LONG Unit, LONG Count, LONG *pDelta)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveWhile(ITextRange* iface, VARIANT *Cset, LONG Count, LONG *pDelta) {
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveStartWhile(ITextRange* iface, VARIANT *Cset, LONG Count, LONG *pDelta)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveEndWhile(ITextRange* iface, VARIANT *Cset, LONG Count, LONG *pDelta)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveUntil(ITextRange* iface, VARIANT *Cset, LONG Count, LONG *pDelta)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveStartUntil(ITextRange* iface, VARIANT *Cset, LONG Count, LONG *pDelta)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnMoveEndUntil(ITextRange* iface, VARIANT *Cset, LONG Count, LONG *pDelta)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnFindText(ITextRange* iface, BSTR bstr, LONG cch, LONG Flags, LONG *pLength)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnFindTextStart(ITextRange* iface, BSTR bstr, LONG cch, LONG Flags,
    LONG *pLength)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnFindTextEnd(ITextRange* iface, BSTR bstr, LONG cch, LONG Flags,
    LONG *pLength)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnDelete(ITextRange* iface, LONG Unit, LONG Count, LONG *pDelta)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnCut(ITextRange* iface, VARIANT *pVar)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnCopy(ITextRange* iface, VARIANT *pVar)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnPaste(ITextRange* iface, VARIANT *pVar, LONG Format)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnCanPaste(ITextRange* iface, VARIANT *pVar, LONG Format, LONG *pb)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnCanEdit(ITextRange* iface, LONG *pb)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnChangeCase(ITextRange* iface, LONG Type)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetPoint(ITextRange* iface, LONG Type, LONG *cx, LONG *cy)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnSetPoint(ITextRange* iface, LONG x, LONG y, LONG Type, LONG Extend)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnScrollIntoView(ITextRange* iface, LONG Value)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextRange_fnGetEmbeddedObject(ITextRange* iface, IUnknown **ppv)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

ITextRangeVtbl rangeVtbl = {
  ITextRange_fnQueryInterface,
  ITextRange_fnAddRef,
  ITextRange_fnRelease,
  ITextRange_fnGetTypeInfoCount,
  ITextRange_fnGetTypeInfo,
  ITextRange_fnGetIDsOfNames,
  ITextRange_fnInvoke,
  ITextRange_fnGetText,
  ITextRange_fnSetText,
  ITextRange_fnGetChar,
  ITextRange_fnSetChar,
  ITextRange_fnGetDuplicate,
  ITextRange_fnGetFormattedText,
  ITextRange_fnSetFormattedText,
  ITextRange_fnGetStart,
  ITextRange_fnSetStart,
  ITextRange_fnGetEnd,
  ITextRange_fnSetEnd,
  ITextRange_fnGetFont,
  ITextRange_fnSetFont,
  ITextRange_fnGetPara,
  ITextRange_fnSetPara,
  ITextRange_fnGetStoryLength,
  ITextRange_fnGetStoryType,
  ITextRange_fnCollapse,
  ITextRange_fnExpand,
  ITextRange_fnGetIndex,
  ITextRange_fnSetIndex,
  ITextRange_fnSetRange,
  ITextRange_fnInRange,
  ITextRange_fnInStory,
  ITextRange_fnIsEqual,
  ITextRange_fnSelect,
  ITextRange_fnStartOf,
  ITextRange_fnEndOf,
  ITextRange_fnMove,
  ITextRange_fnMoveStart,
  ITextRange_fnMoveEnd,
  ITextRange_fnMoveWhile,
  ITextRange_fnMoveStartWhile,
  ITextRange_fnMoveEndWhile,
  ITextRange_fnMoveUntil,
  ITextRange_fnMoveStartUntil,
  ITextRange_fnMoveEndUntil,
  ITextRange_fnFindText,
  ITextRange_fnFindTextStart,
  ITextRange_fnFindTextEnd,
  ITextRange_fnDelete,
  ITextRange_fnCut,
  ITextRange_fnCopy,
  ITextRange_fnPaste,
  ITextRange_fnCanPaste,
  ITextRange_fnCanEdit,
  ITextRange_fnChangeCase,
  ITextRange_fnGetPoint,
  ITextRange_fnSetPoint,
  ITextRange_fnScrollIntoView,
  ITextRange_fnGetEmbeddedObject,
};

static ReTxtRange *ReTxtRange_create(long first, long lim)
{
  ReTxtRange *This = heap_alloc(sizeof *This);
  This->iTextRangeIface.lpVtbl = &rangeVtbl;
  This->ref = 1;
  This->first = first;
  This->lim = lim;
  return This;
}

ReTxtDoc *ReTxtDoc_create(IUnknown *outerObj, ME_TextEditor *editor)
{
  ReTxtDoc *This = heap_alloc(sizeof *This);
  TRACE("(%p,%p)\n", outerObj, editor);
  if (!This)
    return NULL;

  This->outerObj = outerObj;
  This->editor = editor;
  This->iTextDocumentIface.lpVtbl = &vtbl;
  return This;
}

ITextDocument *ReTxtDoc_get_ITextDocument(ReTxtDoc *document)
{
  TRACE("(%p)\n", document);
  return &document->iTextDocumentIface;
}

void ReTxtDoc_destroy(ReTxtDoc *document)
{
  TRACE("(%p)\n", document);
  document->outerObj = NULL;
  document->editor = NULL;
  heap_free(document);
}
