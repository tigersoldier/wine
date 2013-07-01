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

struct tagReTxtDoc {
  IUnknown *outerObj;
  ME_TextEditor *editor;
  ITextDocument iTextDocumentIface;
};

static inline ReTxtDoc *impl_from_interface(ITextDocument *iface)
{
  return CONTAINING_RECORD(iface, ReTxtDoc, iTextDocumentIface);
}

/* Implementation of IUnknown */
static ULONG WINAPI ITextDocument_fnAddRef(ITextDocument *iface)
{
  ReTxtDoc *This = impl_from_interface(iface);
  TRACE("(%p)\n", iface);
  /* COM aggregation - delegate to outer IUnknown object. */
  return IUnknown_AddRef(This->outerObj);
}

static ULONG WINAPI ITextDocument_fnRelease(ITextDocument *iface)
{
  ReTxtDoc *This = impl_from_interface(iface);;
  TRACE("(%p)\n", iface);
  /* COM aggregation - delegate to outer IUnknown object. */
  return IUnknown_Release(This->outerObj);
}

static HRESULT WINAPI ITextDocument_fnQueryInterface(ITextDocument *iface, REFIID riid, void **ppv)
{
  ReTxtDoc *This = impl_from_interface(iface);;
  TRACE("(%p)->(%s, %p)\n", iface, debugstr_guid(riid), ppv);

  /* COM aggregation - delegate to outer IUnknown object. */
  return IUnknown_QueryInterface(This->outerObj, riid, ppv);
}

/* Implementation of IDispatch */

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
  FIXME("not implemented\n");
  return E_NOTIMPL;
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
