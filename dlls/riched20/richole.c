/*
 * RichEdit GUIDs and OLE interface
 *
 * Copyright 2004 by Krzysztof Foltman
 * Copyright 2004 Aric Stewart
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

#include <stdarg.h>

#define NONAMELESSUNION
#define NONAMELESSSTRUCT
#define COBJMACROS

#include "windef.h"
#include "winbase.h"
#include "wingdi.h"
#include "winuser.h"
#include "ole2.h"
#include "richole.h"
#include "editor.h"
#include "tom.h"
#include "txtdoc.h"
#include "wine/debug.h"


WINE_DEFAULT_DEBUG_CHANNEL(richedit);

/* there is no way to be consistent across different sets of headers - mingw, Wine, Win32 SDK*/

#include "initguid.h"
DEFINE_GUID(IID_ITextServices, 0x8d33f740, 0xcf58, 0x11ce, 0xa8, 0x9d, 0x00, 0xaa, 0x00, 0x6c, 0xad, 0xc5);
DEFINE_GUID(IID_ITextHost, 0x13e670f4,0x1a5a,0x11cf,0xab,0xeb,0x00,0xaa,0x00,0xb6,0x5e,0xa1);
DEFINE_GUID(IID_ITextHost2, 0x13e670f5,0x1a5a,0x11cf,0xab,0xeb,0x00,0xaa,0x00,0xb6,0x5e,0xa1);
DEFINE_GUID(IID_ITextDocument, 0x8cc497c0, 0xa1df, 0x11ce, 0x80, 0x98, 0x00, 0xaa, 0x00, 0x47, 0xbe, 0x5d);
DEFINE_GUID(IID_ITextRange, 0x8cc497c2, 0xa1df, 0x11ce, 0x80, 0x98, 0x00, 0xaa, 0x00, 0x47, 0xbe, 0x5d);
DEFINE_GUID(IID_ITextSelection, 0x8cc497c1, 0xa1df, 0x11ce, 0x80, 0x98, 0x00, 0xaa, 0x00, 0x47, 0xbe, 0x5d);

typedef struct IOleClientSiteImpl IOleClientSiteImpl;

typedef struct IRichEditOleImpl {
    IRichEditOle IRichEditOle_iface;
    ReTxtDoc *txtDoc;
    LONG ref;

    ME_TextEditor *editor;
    IOleClientSiteImpl *clientSite;
} IRichEditOleImpl;

struct IOleClientSiteImpl {
    IOleClientSite IOleClientSite_iface;
    LONG ref;

    IRichEditOleImpl *reOle;
};

static inline IRichEditOleImpl *impl_from_IRichEditOle(IRichEditOle *iface)
{
    return CONTAINING_RECORD(iface, IRichEditOleImpl, IRichEditOle_iface);
}

static HRESULT WINAPI
IRichEditOle_fnQueryInterface(IRichEditOle *me, REFIID riid, LPVOID *ppvObj)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);

    TRACE("%p %s\n", This, debugstr_guid(riid) );

    *ppvObj = NULL;
    if (IsEqualGUID(riid, &IID_IUnknown) ||
        IsEqualGUID(riid, &IID_IRichEditOle))
        *ppvObj = &This->IRichEditOle_iface;
    else if (IsEqualGUID(riid, &IID_ITextDocument))
        *ppvObj = ReTxtDoc_get_ITextDocument(This->txtDoc);
    if (*ppvObj)
    {
        IRichEditOle_AddRef(me);
        return S_OK;
    }
    FIXME("%p: unhandled interface %s\n", This, debugstr_guid(riid) );

    return E_NOINTERFACE;
}

static ULONG WINAPI
IRichEditOle_fnAddRef(IRichEditOle *me)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    ULONG ref = InterlockedIncrement( &This->ref );

    TRACE("%p ref = %u\n", This, ref);

    return ref;
}

static ULONG WINAPI
IRichEditOle_fnRelease(IRichEditOle *me)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    ULONG ref = InterlockedDecrement(&This->ref);

    TRACE ("%p ref=%u\n", This, ref);

    if (!ref)
    {
        TRACE ("Destroying %p\n", This);
        IOleClientSite_Release(&This->clientSite->IOleClientSite_iface);
        ReTxtDoc_destroy(This->txtDoc);
        heap_free(This);
    }
    return ref;
}

static HRESULT WINAPI
IRichEditOle_fnActivateAs(IRichEditOle *me, REFCLSID rclsid, REFCLSID rclsidAs)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnContextSensitiveHelp(IRichEditOle *me, BOOL fEnterMode)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnConvertObject(IRichEditOle *me, LONG iob,
               REFCLSID rclsidNew, LPCSTR lpstrUserTypeNew)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static inline IOleClientSiteImpl *impl_from_IOleClientSite(IOleClientSite *iface)
{
    return CONTAINING_RECORD(iface, IOleClientSiteImpl, IOleClientSite_iface);
}

static HRESULT WINAPI
IOleClientSite_fnQueryInterface(IOleClientSite *me, REFIID riid, LPVOID *ppvObj)
{
    TRACE("%p %s\n", me, debugstr_guid(riid) );

    *ppvObj = NULL;
    if (IsEqualGUID(riid, &IID_IUnknown) ||
        IsEqualGUID(riid, &IID_IOleClientSite))
        *ppvObj = me;
    if (*ppvObj)
    {
        IOleClientSite_AddRef(me);
        return S_OK;
    }
    FIXME("%p: unhandled interface %s\n", me, debugstr_guid(riid) );

    return E_NOINTERFACE;
}

static ULONG WINAPI IOleClientSite_fnAddRef(IOleClientSite *iface)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    return InterlockedIncrement(&This->ref);
}

static ULONG WINAPI IOleClientSite_fnRelease(IOleClientSite *iface)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    ULONG ref = InterlockedDecrement(&This->ref);
    if (ref == 0)
        heap_free(This);
    return ref;
}

static HRESULT WINAPI IOleClientSite_fnSaveObject(IOleClientSite *iface)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("stub %p\n", iface);
    return E_NOTIMPL;
}


static HRESULT WINAPI IOleClientSite_fnGetMoniker(IOleClientSite *iface, DWORD dwAssign,
        DWORD dwWhichMoniker, IMoniker **ppmk)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("stub %p\n", iface);
    return E_NOTIMPL;
}

static HRESULT WINAPI IOleClientSite_fnGetContainer(IOleClientSite *iface,
        IOleContainer **ppContainer)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("stub %p\n", iface);
    return E_NOTIMPL;
}

static HRESULT WINAPI IOleClientSite_fnShowObject(IOleClientSite *iface)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("stub %p\n", iface);
    return E_NOTIMPL;
}

static HRESULT WINAPI IOleClientSite_fnOnShowWindow(IOleClientSite *iface, BOOL fShow)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("stub %p\n", iface);
    return E_NOTIMPL;
}

static HRESULT WINAPI IOleClientSite_fnRequestNewObjectLayout(IOleClientSite *iface)
{
    IOleClientSiteImpl *This = impl_from_IOleClientSite(iface);
    if (!This->reOle)
        return CO_E_RELEASED;

    FIXME("stub %p\n", iface);
    return E_NOTIMPL;
}

static const IOleClientSiteVtbl ocst = {
    IOleClientSite_fnQueryInterface,
    IOleClientSite_fnAddRef,
    IOleClientSite_fnRelease,
    IOleClientSite_fnSaveObject,
    IOleClientSite_fnGetMoniker,
    IOleClientSite_fnGetContainer,
    IOleClientSite_fnShowObject,
    IOleClientSite_fnOnShowWindow,
    IOleClientSite_fnRequestNewObjectLayout
};

static IOleClientSiteImpl *
CreateOleClientSite(IRichEditOleImpl *reOle)
{
    IOleClientSiteImpl *clientSite = heap_alloc(sizeof *clientSite);
    if (!clientSite)
        return NULL;

    clientSite->IOleClientSite_iface.lpVtbl = &ocst;
    clientSite->ref = 1;
    clientSite->reOle = reOle;
    return clientSite;
}

static HRESULT WINAPI
IRichEditOle_fnGetClientSite(IRichEditOle *me,
               LPOLECLIENTSITE *lplpolesite)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);

    TRACE("%p,%p\n",This, lplpolesite);

    if(!lplpolesite)
        return E_INVALIDARG;
    *lplpolesite = &This->clientSite->IOleClientSite_iface;
    IOleClientSite_AddRef(*lplpolesite);
    return S_OK;
}

static HRESULT WINAPI
IRichEditOle_fnGetClipboardData(IRichEditOle *me, CHARRANGE *lpchrg,
               DWORD reco, LPDATAOBJECT *lplpdataobj)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    ME_Cursor start;
    int nChars;

    TRACE("(%p,%p,%d)\n",This, lpchrg, reco);
    if(!lplpdataobj)
        return E_INVALIDARG;
    if(!lpchrg) {
        int nFrom, nTo, nStartCur = ME_GetSelectionOfs(This->editor, &nFrom, &nTo);
        start = This->editor->pCursors[nStartCur];
        nChars = nTo - nFrom;
    } else {
        ME_CursorFromCharOfs(This->editor, lpchrg->cpMin, &start);
        nChars = lpchrg->cpMax - lpchrg->cpMin;
    }
    return ME_GetDataObject(This->editor, &start, nChars, lplpdataobj);
}

static LONG WINAPI IRichEditOle_fnGetLinkCount(IRichEditOle *me)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnGetObject(IRichEditOle *me, LONG iob,
               REOBJECT *lpreobject, DWORD dwFlags)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static LONG WINAPI
IRichEditOle_fnGetObjectCount(IRichEditOle *me)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return 0;
}

static HRESULT WINAPI
IRichEditOle_fnHandsOffStorage(IRichEditOle *me, LONG iob)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnImportDataObject(IRichEditOle *me, LPDATAOBJECT lpdataobj,
               CLIPFORMAT cf, HGLOBAL hMetaPict)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnInPlaceDeactivate(IRichEditOle *me)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnInsertObject(IRichEditOle *me, REOBJECT *reo)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    TRACE("(%p,%p)\n", This, reo);

    if (reo->cbStruct < sizeof(*reo)) return STG_E_INVALIDPARAMETER;

    ME_InsertOLEFromCursor(This->editor, reo, 0);
    ME_CommitUndo(This->editor);
    ME_UpdateRepaint(This->editor, FALSE);
    return S_OK;
}

static HRESULT WINAPI IRichEditOle_fnSaveCompleted(IRichEditOle *me, LONG iob,
               LPSTORAGE lpstg)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnSetDvaspect(IRichEditOle *me, LONG iob, DWORD dvaspect)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static HRESULT WINAPI IRichEditOle_fnSetHostNames(IRichEditOle *me,
               LPCSTR lpstrContainerApp, LPCSTR lpstrContainerObj)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p %s %s\n",This, lpstrContainerApp, lpstrContainerObj);
    return E_NOTIMPL;
}

static HRESULT WINAPI
IRichEditOle_fnSetLinkAvailable(IRichEditOle *me, LONG iob, BOOL fAvailable)
{
    IRichEditOleImpl *This = impl_from_IRichEditOle(me);
    FIXME("stub %p\n",This);
    return E_NOTIMPL;
}

static const IRichEditOleVtbl revt = {
    IRichEditOle_fnQueryInterface,
    IRichEditOle_fnAddRef,
    IRichEditOle_fnRelease,
    IRichEditOle_fnGetClientSite,
    IRichEditOle_fnGetObjectCount,
    IRichEditOle_fnGetLinkCount,
    IRichEditOle_fnGetObject,
    IRichEditOle_fnInsertObject,
    IRichEditOle_fnConvertObject,
    IRichEditOle_fnActivateAs,
    IRichEditOle_fnSetHostNames,
    IRichEditOle_fnSetLinkAvailable,
    IRichEditOle_fnSetDvaspect,
    IRichEditOle_fnHandsOffStorage,
    IRichEditOle_fnSaveCompleted,
    IRichEditOle_fnInPlaceDeactivate,
    IRichEditOle_fnContextSensitiveHelp,
    IRichEditOle_fnGetClipboardData,
    IRichEditOle_fnImportDataObject
};

LRESULT CreateIRichEditOle(ME_TextEditor *editor, LPVOID *ppObj)
{
    IRichEditOleImpl *reo;

    reo = heap_alloc(sizeof(IRichEditOleImpl));
    if (!reo)
        return 0;

    reo->IRichEditOle_iface.lpVtbl = &revt;
    reo->txtDoc = ReTxtDoc_create((IUnknown*) &reo->IRichEditOle_iface, editor);
    if (!reo->txtDoc)
    {
        heap_free(reo);
        return 0;
    }
    reo->ref = 1;
    reo->editor = editor;
    reo->clientSite = CreateOleClientSite(reo);
    if (!reo->clientSite)
    {
        ITextDocument_Release(ReTxtDoc_get_ITextDocument(reo->txtDoc));
        heap_free(reo);
        return 0;
    }
    TRACE("Created %p\n",reo);
    *ppObj = reo;

    return 1;
}

static void convert_sizel(const ME_Context *c, const SIZEL* szl, SIZE* sz)
{
  /* sizel is in .01 millimeters, sz in pixels */
  sz->cx = MulDiv(szl->cx, c->dpi.cx, 2540);
  sz->cy = MulDiv(szl->cy, c->dpi.cy, 2540);
}

/******************************************************************************
 * ME_GetOLEObjectSize
 *
 * Sets run extent for OLE objects.
 */
void ME_GetOLEObjectSize(const ME_Context *c, ME_Run *run, SIZE *pSize)
{
  IDataObject*  ido;
  FORMATETC     fmt;
  STGMEDIUM     stgm;
  DIBSECTION    dibsect;
  ENHMETAHEADER emh;

  assert(run->nFlags & MERF_GRAPHICS);
  assert(run->ole_obj);

  if (run->ole_obj->sizel.cx != 0 || run->ole_obj->sizel.cy != 0)
  {
    convert_sizel(c, &run->ole_obj->sizel, pSize);
    if (c->editor->nZoomNumerator != 0)
    {
      pSize->cx = MulDiv(pSize->cx, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
      pSize->cy = MulDiv(pSize->cy, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
    }
    return;
  }

  IOleObject_QueryInterface(run->ole_obj->poleobj, &IID_IDataObject, (void**)&ido);
  fmt.cfFormat = CF_BITMAP;
  fmt.ptd = NULL;
  fmt.dwAspect = DVASPECT_CONTENT;
  fmt.lindex = -1;
  fmt.tymed = TYMED_GDI;
  if (IDataObject_GetData(ido, &fmt, &stgm) != S_OK)
  {
    fmt.cfFormat = CF_ENHMETAFILE;
    fmt.tymed = TYMED_ENHMF;
    if (IDataObject_GetData(ido, &fmt, &stgm) != S_OK)
    {
      FIXME("unsupported format\n");
      pSize->cx = pSize->cy = 0;
      IDataObject_Release(ido);
      return;
    }
  }

  switch (stgm.tymed)
  {
  case TYMED_GDI:
    GetObjectW(stgm.u.hBitmap, sizeof(dibsect), &dibsect);
    pSize->cx = dibsect.dsBm.bmWidth;
    pSize->cy = dibsect.dsBm.bmHeight;
    if (!stgm.pUnkForRelease) DeleteObject(stgm.u.hBitmap);
    break;
  case TYMED_ENHMF:
    GetEnhMetaFileHeader(stgm.u.hEnhMetaFile, sizeof(emh), &emh);
    pSize->cx = emh.rclBounds.right - emh.rclBounds.left;
    pSize->cy = emh.rclBounds.bottom - emh.rclBounds.top;
    if (!stgm.pUnkForRelease) DeleteEnhMetaFile(stgm.u.hEnhMetaFile);
    break;
  default:
    FIXME("Unsupported tymed %d\n", stgm.tymed);
    break;
  }
  IDataObject_Release(ido);
  if (c->editor->nZoomNumerator != 0)
  {
    pSize->cx = MulDiv(pSize->cx, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
    pSize->cy = MulDiv(pSize->cy, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
  }
}

void ME_DrawOLE(ME_Context *c, int x, int y, ME_Run *run,
                ME_Paragraph *para, BOOL selected)
{
  IDataObject*  ido;
  FORMATETC     fmt;
  STGMEDIUM     stgm;
  DIBSECTION    dibsect;
  ENHMETAHEADER emh;
  HDC           hMemDC;
  SIZE          sz;
  BOOL          has_size;

  assert(run->nFlags & MERF_GRAPHICS);
  assert(run->ole_obj);
  if (IOleObject_QueryInterface(run->ole_obj->poleobj, &IID_IDataObject, (void**)&ido) != S_OK)
  {
    FIXME("Couldn't get interface\n");
    return;
  }
  has_size = run->ole_obj->sizel.cx != 0 || run->ole_obj->sizel.cy != 0;
  fmt.cfFormat = CF_BITMAP;
  fmt.ptd = NULL;
  fmt.dwAspect = DVASPECT_CONTENT;
  fmt.lindex = -1;
  fmt.tymed = TYMED_GDI;
  if (IDataObject_GetData(ido, &fmt, &stgm) != S_OK)
  {
    fmt.cfFormat = CF_ENHMETAFILE;
    fmt.tymed = TYMED_ENHMF;
    if (IDataObject_GetData(ido, &fmt, &stgm) != S_OK)
    {
      FIXME("Couldn't get storage medium\n");
      IDataObject_Release(ido);
      return;
    }
  }
  switch (stgm.tymed)
  {
  case TYMED_GDI:
    GetObjectW(stgm.u.hBitmap, sizeof(dibsect), &dibsect);
    hMemDC = CreateCompatibleDC(c->hDC);
    SelectObject(hMemDC, stgm.u.hBitmap);
    if (has_size)
    {
      convert_sizel(c, &run->ole_obj->sizel, &sz);
    } else {
      sz.cx = MulDiv(dibsect.dsBm.bmWidth, c->dpi.cx, 96);
      sz.cy = MulDiv(dibsect.dsBm.bmHeight, c->dpi.cy, 96);
    }
    if (c->editor->nZoomNumerator != 0)
    {
      sz.cx = MulDiv(sz.cx, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
      sz.cy = MulDiv(sz.cy, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
    }
    if (sz.cx == dibsect.dsBm.bmWidth && sz.cy == dibsect.dsBm.bmHeight)
    {
      BitBlt(c->hDC, x, y - sz.cy,
             dibsect.dsBm.bmWidth, dibsect.dsBm.bmHeight,
             hMemDC, 0, 0, SRCCOPY);
    } else {
      StretchBlt(c->hDC, x, y - sz.cy, sz.cx, sz.cy,
                 hMemDC, 0, 0, dibsect.dsBm.bmWidth,
                 dibsect.dsBm.bmHeight, SRCCOPY);
    }
    DeleteDC(hMemDC);
    if (!stgm.pUnkForRelease) DeleteObject(stgm.u.hBitmap);
    break;
  case TYMED_ENHMF:
    GetEnhMetaFileHeader(stgm.u.hEnhMetaFile, sizeof(emh), &emh);
    if (has_size)
    {
      convert_sizel(c, &run->ole_obj->sizel, &sz);
    } else {
      sz.cy = MulDiv(emh.rclBounds.bottom - emh.rclBounds.top, c->dpi.cx, 96);
      sz.cx = MulDiv(emh.rclBounds.right - emh.rclBounds.left, c->dpi.cy, 96);
    }
    if (c->editor->nZoomNumerator != 0)
    {
      sz.cx = MulDiv(sz.cx, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
      sz.cy = MulDiv(sz.cy, c->editor->nZoomNumerator, c->editor->nZoomDenominator);
    }

    {
      RECT    rc;

      rc.left = x;
      rc.top = y - sz.cy;
      rc.right = x + sz.cx;
      rc.bottom = y;
      PlayEnhMetaFile(c->hDC, stgm.u.hEnhMetaFile, &rc);
    }
    if (!stgm.pUnkForRelease) DeleteEnhMetaFile(stgm.u.hEnhMetaFile);
    break;
  default:
    FIXME("Unsupported tymed %d\n", stgm.tymed);
    selected = FALSE;
    break;
  }
  if (selected && !c->editor->bHideSelection)
    PatBlt(c->hDC, x, y - sz.cy, sz.cx, sz.cy, DSTINVERT);
  IDataObject_Release(ido);
}

void ME_DeleteReObject(REOBJECT* reo)
{
    if (reo->poleobj)   IOleObject_Release(reo->poleobj);
    if (reo->pstg)      IStorage_Release(reo->pstg);
    if (reo->polesite)  IOleClientSite_Release(reo->polesite);
    FREE_OBJ(reo);
}

void ME_CopyReObject(REOBJECT* dst, const REOBJECT* src)
{
    *dst = *src;

    if (dst->poleobj)   IOleObject_AddRef(dst->poleobj);
    if (dst->pstg)      IStorage_AddRef(dst->pstg);
    if (dst->polesite)  IOleClientSite_AddRef(dst->polesite);
}
