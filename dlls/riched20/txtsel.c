/*
 * RichEdit - ITextSelection implementation
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

#include "txtsel.h"

#include "editor.h"
#include "txtrng.h"

WINE_DEFAULT_DEBUG_CHANNEL(richedit);

struct tagReTxtSel {
  ITextSelection iTextSelectionIface;
  ReTxtRng *txtRng;
  ME_TextEditor *editor;
};

static void ReTxtSel_destroy(ReTxtSel *txtSel);

/*** IUnknown methods ***/
static inline ReTxtSel *impl_from_ITextSelection(ITextSelection *iface)
{
  return CONTAINING_RECORD(iface, ReTxtSel, iTextSelectionIface);
}

static HRESULT WINAPI ITextSelection_fnQueryInterface(
                                                      ITextSelection *iface,
                                                      REFIID riid,
                                                      void **ppvObj)
{
  *ppvObj = NULL;
  if (IsEqualGUID(riid, &IID_IUnknown)
      || IsEqualGUID(riid, &IID_IDispatch)
      || IsEqualGUID(riid, &IID_ITextRange)
      || IsEqualGUID(riid, &IID_ITextSelection))
  {
    *ppvObj = iface;
    ITextSelection_AddRef(iface);
    return S_OK;
  }

  return E_NOINTERFACE;
}

static ULONG WINAPI ITextSelection_fnAddRef(ITextSelection *iface)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  /* Delegate to ITextRange */
  ITextRange *rangeIface = ReTxtRng_get_ITextRange(This->txtRng);
  return ITextRange_AddRef(rangeIface);
}

static ULONG WINAPI ITextSelection_fnRelease(ITextSelection *iface)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);
  /* Delegate to ITextRange */
  ITextRange *rangeIface = ReTxtRng_get_ITextRange(This->txtRng);
  ULONG ref = ITextRange_Release(rangeIface);
  if (ref == 0)
  {
    This->txtRng = NULL;
    ReTxtSel_destroy(This);
  }
  return ref;
}

/*** IDispatch methods ***/
static HRESULT WINAPI ITextSelection_fnGetTypeInfoCount(ITextSelection *iface,
                                                        UINT *pctinfo)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetTypeInfo(ITextSelection *iface,
                                                   UINT iTInfo,
                                                   LCID lcid,
                                                   ITypeInfo **ppTInfo)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetIDsOfNames(ITextSelection *iface,
                                                     REFIID riid,
                                                     LPOLESTR *rgszNames,
                                                     UINT cNames,
                                                     LCID lcid,
                                                     DISPID *rgDispId)
{
  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnInvoke(ITextSelection *iface,
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

/*** ITextRange methods ***/
static HRESULT WINAPI ITextSelection_fnGetText(ITextSelection *iface, BSTR *pbstr)
{
  ReTxtSel *This;
  ITextRange *rangeIface;

  TRACE("(%p)\n", pbstr);
  This = impl_from_ITextSelection(iface);
  rangeIface = ReTxtRng_get_ITextRange(This->txtRng);
  return ITextRange_GetText(rangeIface, pbstr);
}

static HRESULT WINAPI ITextSelection_fnSetText(ITextSelection *iface, BSTR bstr)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetChar(ITextSelection *iface, LONG *pch)
{
  ReTxtSel *This;
  ITextRange *rangeIface;

  TRACE("(%p)\n", pch);
  This = impl_from_ITextSelection(iface);
  rangeIface = ReTxtRng_get_ITextRange(This->txtRng);
  return ITextRange_GetChar(rangeIface, pch);
}

static HRESULT WINAPI ITextSelection_fnSetChar(ITextSelection *iface, LONG ch)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetDuplicate(ITextSelection *iface,
                                                    ITextRange **ppRange)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetFormattedText(ITextSelection *iface,
                                                        ITextRange **ppRange)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSetFormattedText(ITextSelection *iface, ITextRange *pRange)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetStart(ITextSelection *iface, LONG *pcpFirst)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  ITextRange *rangeIface = ReTxtRng_get_ITextRange(This->txtRng);
  return ITextRange_GetStart(rangeIface, pcpFirst);
}

static HRESULT WINAPI ITextSelection_fnSetStart(ITextSelection *iface, LONG cpFirst)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetEnd(ITextSelection *iface, LONG *pcpLim)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  ITextRange *rangeIface = ReTxtRng_get_ITextRange(This->txtRng);
  return ITextRange_GetEnd(rangeIface, pcpLim);
}

static HRESULT WINAPI ITextSelection_fnSetEnd(ITextSelection *iface, LONG cpLim)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetFont(ITextSelection *iface, ITextFont **pFont)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  ITextRange *rangeIface = ReTxtRng_get_ITextRange(This->txtRng);
  return ITextRange_GetFont(rangeIface, pFont);
}

static HRESULT WINAPI ITextSelection_fnSetFont(ITextSelection *iface, ITextFont *pFont)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetPara(ITextSelection *iface, ITextPara **ppPara)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);
  ITextRange *rangeIface = ReTxtRng_get_ITextRange(This->txtRng);

  return ITextRange_GetPara(rangeIface, ppPara);
}

static HRESULT WINAPI ITextSelection_fnSetPara(ITextSelection *iface, ITextPara *pPara)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetStoryLength(ITextSelection *iface, LONG *pcch)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);
  ITextRange *rangeIface = ReTxtRng_get_ITextRange(This->txtRng);

  return ITextRange_GetStoryLength(rangeIface, pcch);
}

static HRESULT WINAPI ITextSelection_fnGetStoryType(ITextSelection *iface, LONG *pValue)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);
  ITextRange *rangeIface = ReTxtRng_get_ITextRange(This->txtRng);

  return ITextRange_GetStoryType(rangeIface, pValue);
}

static HRESULT WINAPI ITextSelection_fnCollapse(ITextSelection *iface, LONG bStart)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnExpand(ITextSelection *iface, LONG Unit, LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetIndex(ITextSelection *iface, LONG Unit, LONG *pIndex)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  ITextRange *rangeIface = ReTxtRng_get_ITextRange(This->txtRng);

  return ITextRange_GetIndex(rangeIface, Unit, pIndex);
}

static HRESULT WINAPI ITextSelection_fnSetIndex(ITextSelection *iface,
                                                LONG Unit,
                                                LONG Index,
                                                LONG Extend)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSetRange(ITextSelection *iface,
                                                LONG cpActive,
                                                LONG cpOther)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnInRange(ITextSelection *iface,
                                               ITextRange *pRange,
                                               LONG *pb)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnInStory(ITextSelection *iface,
                                               ITextRange *pRange,
                                               LONG *pb)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnIsEqual(ITextSelection *iface,
                                               ITextRange *pRange,
                                               LONG *pb)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSelect(ITextSelection *iface)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnStartOf(ITextSelection *iface,
                                               LONG Unit,
                                               LONG Extend,
                                               LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnEndOf(ITextSelection *iface,
                                             LONG Unit,
                                             LONG Extend,
                                             LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMove(ITextSelection *iface,
                                            LONG Unit,
                                            LONG Count,
                                            LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveStart(ITextSelection *iface,
                                                 LONG Unit,
                                                 LONG Count,
                                                 LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveEnd(ITextSelection *iface,
                                               LONG Unit,
                                               LONG Count,
                                               LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveWhile(ITextSelection *iface,
                                                 VARIANT *Cset,
                                                 LONG Count,
                                                 LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveStartWhile(ITextSelection *iface,
                                                      VARIANT *Cset,
                                                      LONG Count,
                                                      LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveEndWhile(ITextSelection *iface,
                                                    VARIANT *Cset,
                                                    LONG Count,
                                                    LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveUntil(ITextSelection *iface,
                                                 VARIANT *Cset,
                                                 LONG Count,
                                                 LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveStartUntil(ITextSelection *iface,
                                                      VARIANT *Cset,
                                                      LONG Count,
                                                      LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveEndUntil(ITextSelection *iface,
                                                    VARIANT *Cset,
                                                    LONG Count,
                                                    LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnFindText(ITextSelection *iface,
                                                BSTR bstr,
                                                LONG cch,
                                                LONG Flags,
                                                LONG *pLength)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnFindTextStart(ITextSelection *iface,
                                                     BSTR bstr,
                                                     LONG cch,
                                                     LONG Flags,
                                                     LONG *pLength)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnFindTextEnd(ITextSelection *iface,
                                                   BSTR bstr,
                                                   LONG cch,
                                                   LONG Flags,
                                                   LONG *pLength)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnDelete(ITextSelection *iface,
                                              LONG Unit,
                                              LONG Count,
                                              LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnCut(ITextSelection *iface, VARIANT *pVar)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnCopy(ITextSelection *iface, VARIANT *pVar)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnPaste(ITextSelection *iface, VARIANT *pVar, LONG Format)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnCanPaste(ITextSelection *iface,
                                                VARIANT *pVar,
                                                LONG Format,
                                                LONG *pb)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnCanEdit(ITextSelection *iface, LONG *pb)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnChangeCase(ITextSelection *iface, LONG Type)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetPoint(ITextSelection *iface,
                                                LONG Type,
                                                LONG *cx,
                                                LONG *cy)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSetPoint(ITextSelection *iface,
                                                LONG x,
                                                LONG y,
                                                LONG Type,
                                                LONG Extend)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnScrollIntoView(ITextSelection *iface, LONG Value)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetEmbeddedObject(ITextSelection *iface,
                                                         IUnknown **ppv)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

/*** ITextSelection methods ***/
static HRESULT WINAPI ITextSelection_fnGetFlags(ITextSelection *iface, LONG *pFlags)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnSetFlags(ITextSelection *iface, LONG Flags)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnGetType(ITextSelection *iface, LONG *pType)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveLeft(ITextSelection *iface,
                                                LONG Unit,
                                                LONG Count,
                                                LONG Extend,
                                                LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveRight(ITextSelection *iface,
                                                 LONG Unit,
                                                 LONG Count,
                                                 LONG Extend,
                                                 LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveUp(ITextSelection *iface,
                                              LONG Unit,
                                              LONG Count,
                                              LONG Extend,
                                              LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnMoveDown(ITextSelection *iface,
                                                LONG Unit,
                                                LONG Count,
                                                LONG Extend,
                                                LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnHomeKey(ITextSelection *iface,
                                               LONG Unit,
                                               LONG Extend,
                                               LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnEndKey(ITextSelection *iface,
                                              LONG Unit,
                                              LONG Extend,
                                              LONG *pDelta)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static HRESULT WINAPI ITextSelection_fnTypeText(ITextSelection *iface, BSTR bstr)
{
  ReTxtSel *This = impl_from_ITextSelection(iface);

  if (This->editor == NULL)
    return CO_E_RELEASED;

  FIXME("not implemented\n");
  return E_NOTIMPL;
}

static const ITextSelectionVtbl selectionVtbl = {
  ITextSelection_fnQueryInterface,
  ITextSelection_fnAddRef,
  ITextSelection_fnRelease,
  ITextSelection_fnGetTypeInfoCount,
  ITextSelection_fnGetTypeInfo,
  ITextSelection_fnGetIDsOfNames,
  ITextSelection_fnInvoke,
  ITextSelection_fnGetText,
  ITextSelection_fnSetText,
  ITextSelection_fnGetChar,
  ITextSelection_fnSetChar,
  ITextSelection_fnGetDuplicate,
  ITextSelection_fnGetFormattedText,
  ITextSelection_fnSetFormattedText,
  ITextSelection_fnGetStart,
  ITextSelection_fnSetStart,
  ITextSelection_fnGetEnd,
  ITextSelection_fnSetEnd,
  ITextSelection_fnGetFont,
  ITextSelection_fnSetFont,
  ITextSelection_fnGetPara,
  ITextSelection_fnSetPara,
  ITextSelection_fnGetStoryLength,
  ITextSelection_fnGetStoryType,
  ITextSelection_fnCollapse,
  ITextSelection_fnExpand,
  ITextSelection_fnGetIndex,
  ITextSelection_fnSetIndex,
  ITextSelection_fnSetRange,
  ITextSelection_fnInRange,
  ITextSelection_fnInStory,
  ITextSelection_fnIsEqual,
  ITextSelection_fnSelect,
  ITextSelection_fnStartOf,
  ITextSelection_fnEndOf,
  ITextSelection_fnMove,
  ITextSelection_fnMoveStart,
  ITextSelection_fnMoveEnd,
  ITextSelection_fnMoveWhile,
  ITextSelection_fnMoveStartWhile,
  ITextSelection_fnMoveEndWhile,
  ITextSelection_fnMoveUntil,
  ITextSelection_fnMoveStartUntil,
  ITextSelection_fnMoveEndUntil,
  ITextSelection_fnFindText,
  ITextSelection_fnFindTextStart,
  ITextSelection_fnFindTextEnd,
  ITextSelection_fnDelete,
  ITextSelection_fnCut,
  ITextSelection_fnCopy,
  ITextSelection_fnPaste,
  ITextSelection_fnCanPaste,
  ITextSelection_fnCanEdit,
  ITextSelection_fnChangeCase,
  ITextSelection_fnGetPoint,
  ITextSelection_fnSetPoint,
  ITextSelection_fnScrollIntoView,
  ITextSelection_fnGetEmbeddedObject,
  ITextSelection_fnGetFlags,
  ITextSelection_fnSetFlags,
  ITextSelection_fnGetType,
  ITextSelection_fnMoveLeft,
  ITextSelection_fnMoveRight,
  ITextSelection_fnMoveUp,
  ITextSelection_fnMoveDown,
  ITextSelection_fnHomeKey,
  ITextSelection_fnEndKey,
  ITextSelection_fnTypeText
};

ReTxtSel *ReTxtSel_create(ME_TextEditor *editor)
{
  ReTxtSel *txtSel;
  ME_Cursor *startCursor;
  ME_Cursor *endCursor;

  txtSel = heap_alloc(sizeof *txtSel);
  txtSel->editor = editor;
  ME_GetSelection(editor, &startCursor, &endCursor);
  txtSel->txtRng = ReTxtRng_createWithCursors(editor, startCursor, endCursor);
  txtSel->iTextSelectionIface.lpVtbl = &selectionVtbl;
  return txtSel;
}

ITextSelection *ReTxtSel_get_ITextSelection(ReTxtSel *txtSel)
{
  return &txtSel->iTextSelectionIface;
}

static void ReTxtSel_destroy(ReTxtSel *txtSel)
{
  /* txtSel->txtRng should be self-destroyed in ITextSelection_fnRelease(). */
  heap_free(txtSel);
}

void ReTxtSel_releaseEditor(ReTxtSel *txtSel)
{
  if (txtSel == NULL)
  {
    ERR("txtSel == NULL\n");
    return;
  }
  txtSel->editor = NULL;
  ReTxtRng_releaseEditor(txtSel->txtRng);
}
