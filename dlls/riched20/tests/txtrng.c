/*
 * Tests for ITextRange.
 *
 * Copyright 2013 Caibin Chen
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

#define COBJMACROS

#include <stdarg.h>

#include <windef.h>
#include <winbase.h>
#include <wingdi.h>
#include <ole2.h>
#include <richedit.h>
#include <richole.h>
#include <tom.h>
#include <wine/test.h>

static const size_t MAX_LEN = 256;

static HWND new_window(LPCTSTR lpClassName, DWORD dwStyle, HWND parent);
static HWND new_richedit(HWND parent);
static BOOL clear_text(ITextDocument *txtDoc);
static int insert_text_at_start(ITextDocument *txtDoc, const char *text);
static BOOL set_text(ITextDocument *txtDoc, const char *text);
static ITextRange* get_range(ITextDocument *txtDoc, long cpFirst, long cpLim);
static int set_range_text(ITextRange *txtRng, const char *text);

static HMODULE hmoduleRichEdit;

static HWND new_window(LPCTSTR lpClassName, DWORD dwStyle, HWND parent)
{
  HWND hwnd
    = CreateWindow(lpClassName, NULL,
                   dwStyle | WS_POPUP | WS_HSCROLL | WS_VSCROLL | WS_VISIBLE,
                   0, 0, 200, 60, parent, NULL, hmoduleRichEdit, NULL);
  ok(hwnd != NULL, "class: %s, error: %d\n", lpClassName, (int) GetLastError());
  return hwnd;
}

static HWND new_richedit(HWND parent)
{
  return new_window(RICHEDIT_CLASS, ES_MULTILINE, parent);
}

static BOOL clear_text(ITextDocument *txtDoc)
{
  HRESULT hres;
  LONG storyLength;
  ITextRange *txtRng;
  BOOL ret = TRUE;

  txtRng = get_range(txtDoc, 0, 0);
  if (txtRng == NULL)
  {
    ok(0, "Get range 0, 0 failed\n");
    return FALSE;
  }

  hres = ITextRange_GetStoryLength(txtRng, &storyLength);
  if (hres != S_OK)
  {
    ok(0, "ITextRange_GetStoryLength failed\n");
    ret = FALSE;
  }

  if (storyLength == 1) {
    /* The only character is the last CR, which can not be deleted. */
    ITextRange_Release(txtRng);
    return ret;
  }

  hres = ITextRange_SetEnd(txtRng, storyLength);
  if (hres != S_OK)
  {
    ok(0, "ITextRange_SetEnd to %d failed\n", storyLength);
    ret = FALSE;
  }

  hres = ITextRange_Delete(txtRng, tomCharacter, 0, NULL);
  if (hres != S_OK)
  {
    ok(0, "ITextRange_Delete failed\n");
    ret = FALSE;
  }

  ITextRange_Release(txtRng);
  return ret;
}

/**
 * @return Length of the wide string inserted. 0 on failure.
 */
static int insert_text_at_start(ITextDocument *txtDoc, const char *text)
{
  ITextRange *txtRng = NULL;
  int len;

  txtRng = get_range(txtDoc, 0, 0);
  if (!txtRng) {
    skip("Cannot insert text at start: get range failed\n");
    return 0;
  }

  len = set_range_text(txtRng, text);
  ok(len > 0, "Cannot insert text at start: set range text failed\n");

  if (txtRng) ITextRange_Release(txtRng);
  return len;
}

/**
 * @return Length of the wide string set to txtDoc. 0 on failure.
 */
static int set_text(ITextDocument *txtDoc, const char *text)
{
  int len;
  if (!clear_text(txtDoc))
  {
    skip("Cannot set text, clear_text failed\n");
    return 0;
  }

  len = insert_text_at_start(txtDoc, text);
  if (len == 0)
  {
    skip("Cannot set text, insert_text_at_start failed\n");
    return 0;
  }

  return len;
}

static ITextRange* get_range(ITextDocument *txtDoc, long cpFirst, long cpLim)
{
  ITextRange *txtRng = NULL;
  HRESULT hres;
  hres = ITextDocument_Range(txtDoc, cpFirst, cpLim, &txtRng);
  if (hres == S_OK)
  {
    return txtRng;
  }
  else
  {
    return NULL;
  }
}

/**
 * @return Length of the wide string set to the range. 0 on failure.
 */
static int set_range_text(ITextRange *txtRng, const char *text)
{
  WCHAR wstr[MAX_LEN];
  int len;
  HRESULT hres;
  BSTR bstr;

  len = MultiByteToWideChar(CP_UTF8,
                            /* dwFlags */ 0,
                            text,
                            /*cbMultiByte */ -1,
                            wstr,
                            MAX_LEN);
  /* MultiByteToWideChar return value counts NUL charactor in the length */
  len--;

  bstr = SysAllocString(wstr);
  ok(bstr != NULL, "Bstr allocation failed\n");
  hres = ITextRange_SetText(txtRng, bstr);
  SysFreeString(bstr);

  if (hres != S_OK)
  {
    skip("ITextRange_SetText \"%s\" failed\n", text);
    return 0;
  }
  return len;
}

static void test_SetText(ITextDocument *txtDoc)
{
  const char *text = "Text to set";
  BSTR retText = NULL;
  int len;
  ITextRange *txtRng = NULL;
  LONG retCpLim;

  ok(clear_text(txtDoc), "clear_text failed\n");

  txtRng = get_range(txtDoc, 0, 0);
  if (txtRng == NULL)
  {
    ok(0, "get_range failed\n");
    return;
  }

  len = set_range_text(txtRng, text);
  ok(len == strlen(text),
     "Set text length mismatch, expected %d, got %d\n",
     (int) strlen(text), len);

  ITextRange_GetText(txtRng, &retText);
  ok(len == SysStringLen(retText),
     "Length mismatch, expected: %d, got %d\n", len, SysStringLen(retText));
  SysFreeString(retText);

  ITextRange_GetEnd(txtRng, &retCpLim);
  ok(len == retCpLim,
     "Range should end at the end of the text, expected %d, got %d\n", len, retCpLim);

  if (txtRng) ITextRange_Release(txtRng);
  if (retText) SysFreeString(retText);
}

static BOOL internal_test_range_start_end(ITextDocument *txtDoc,
                                          LONG cpStart, LONG cpLimit,
                                          LONG expectedStart, LONG expectedEnd)
{
  LONG retCpStart;
  LONG retCpLimit;
  ITextRange *txtRng;
  HRESULT hres;
  BOOL ret = TRUE;

  txtRng = get_range(txtDoc, cpStart, cpLimit);
  if (txtRng == NULL)
  {
    ok(0, "get_range failed\n");
    return FALSE;
  }

  hres = ITextRange_GetStart(txtRng, &retCpStart);
  if (hres != S_OK)
  {
    ok(0, "ITextRange_GetStart failed\n");
    ret = FALSE;
  }
  if (retCpStart != expectedStart)
  {
    ok(0, "Start point mismatch, expected %d, got %d\n",
       expectedStart, retCpStart);
    ret = FALSE;
  }

  hres = ITextRange_GetEnd(txtRng, &retCpLimit);
  if (hres != S_OK)
  {
    ok(0, "ITextRange_GetEnd failed\n");
    ret = FALSE;
  }
  if (retCpLimit != expectedEnd)
  {
    ok(0, "End point mismatch, expected %d, got %d\n",
       expectedEnd, retCpLimit);
    ret = FALSE;
  }

  ITextRange_Release(txtRng);
  return ret;
}

static void test_range_start_end(ITextDocument *txtDoc)
{
  const char *text = "Text to set";
  int len;

  len = set_text(txtDoc, text);
  len = set_text(txtDoc, text);
  ok(len == strlen(text), "Length mismatch, expected %d, got %d\n",
     (int) strlen(text), len);

  /* The last CR is can be in a range */
  ok(internal_test_range_start_end(txtDoc, 0, len + 1, 0, len + 1), "rang failed\n");
  /* A range can include any text range inside the story */
  ok(internal_test_range_start_end(txtDoc,
                                   len / 3, len * 2 / 3,
                                   len / 3, len * 2 / 3),
     "range failed\n");
  /* Range cannot beyond the story. */
  ok(internal_test_range_start_end(txtDoc,
                                   -1, len * 10,
                                   0, len + 1),
     "range failed\n");
}

START_TEST(txtrng)
{
  IRichEditOle *reOle = NULL;
  ITextDocument *txtDoc = NULL;
  HRESULT hres;
  LRESULT res;
  HWND w;

  /* Must explicitly LoadLibrary(). The test has no references to functions in
   * RICHED20.DLL, so the linker doesn't actually link to it. */
  hmoduleRichEdit = LoadLibrary("RICHED20.DLL");
  ok(hmoduleRichEdit != NULL, "error: %d\n", (int) GetLastError());

  w = new_richedit(NULL);
  if (!w) {
    skip("Couldn't create window\n");
    return;
  }

  res = SendMessage(w, EM_GETOLEINTERFACE, 0, (LPARAM) &reOle);
  ok(res, "SendMessage\n");
  ok(reOle != NULL, "EM_GETOLEINTERFACE\n");

  hres = IRichEditOle_QueryInterface(reOle, &IID_ITextDocument,
                                     (void **) &txtDoc);
  ok(hres == S_OK, "IRichEditOle_QueryInterface\n");
  ok(txtDoc != NULL, "IRichEditOle_QueryInterface\n");

  test_SetText(txtDoc);
  test_range_start_end(txtDoc);

  ITextDocument_Release(txtDoc);
  IRichEditOle_Release(reOle);
  DestroyWindow(w);
}
