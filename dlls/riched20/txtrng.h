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

#ifndef __TXTRNG_H
#define __TXTRNG_H

#include <tom.h>

#include "editstr.h"

typedef struct tagReTxtRng ReTxtRng;

ReTxtRng *ReTxtRng_create(ME_TextEditor *editor,
                          long first,
                          long lim);
ReTxtRng *ReTxtRng_createWithCursors(ME_TextEditor *editor,
                                     ME_Cursor *start,
                                     ME_Cursor *end);
ITextRange *ReTxtRng_get_ITextRange(ReTxtRng *txtRng);
/**
 * Mark the attached editor as released so that all calls related to
 * the editor will return CO_E_RELEASED.
 */
void ReTxtRng_releaseEditor(ReTxtRng *txtRng);

#endif /* __TXTRNG_H */
