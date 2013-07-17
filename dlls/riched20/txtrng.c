/*
 * RichEdit - ITextRange implementation
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

#include "txtrng.h"
#include "editor.h"

WINE_DEFAULT_DEBUG_CHANNEL(richedit);

struct tagReTxtRng {
  ITextRange iTextRangeIface;
  LONG ref;
  ME_Cursor *first;
  ME_Cursor *lim;
  ME_TextEditor *editor;
};

static void ReTxtRng_destroy(ReTxtRng *txtRng);

static inline ReTxtRng *impl_from_ITextRange(ITextRange *iface)
{
  return CONTAINING_RECORD(iface, ReTxtRng, iTextRangeIface);
}

static HRESULT WINAPI ITextRange_fnQueryInterface(ITextRange* iface, REFIID riid, void **ppvObject)
{
  ReTxtRng *This = impl_from_ITextRange(iface);
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
  ReTxtRng *This = impl_from_ITextRange(iface);
  ULONG ref = InterlockedIncrement(&This->ref);

  TRACE("(%p) ref=%d\n", This, ref);

  return ref;
}

ULONG WINAPI ITextRange_fnRelease(ITextRange* iface)
{
  ReTxtRng *This = impl_from_ITextRange(iface);
  ULONG ref = InterlockedDecrement(&This->ref);

  TRACE("(%p) ref=%d\n", This, ref);

  if (!ref)
  {
    ReTxtRng_destroy(This);
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

ReTxtRng *ReTxtRng_create(ME_TextEditor *editor, long first, long lim)
{
  ReTxtRng *This;

  TRACE("editor: %p, first: %ld, lim: %ld\n", editor, first, lim);
  if (editor == NULL)
    ERR("editor cannot be NULL\n");
  if (first < 0 || first > lim)
    ERR("invalid first or lim: %ld - %ld\n", first, lim);

  This = heap_alloc(sizeof *This);
  This->iTextRangeIface.lpVtbl = &rangeVtbl;
  This->ref = 1;
  This->editor = editor;
  This->first = heap_alloc(sizeof(*This->first));
  This->lim = heap_alloc(sizeof(*This->lim));
  return This;
}

ITextRange *ReTxtRng_get_ITextRange(ReTxtRng *txtRng)
{
  return &txtRng->iTextRangeIface;
}

static void ReTxtRng_destroy(ReTxtRng *txtRng)
{
  heap_free(txtRng->first);
  heap_free(txtRng->lim);
  heap_free(txtRng);
}
