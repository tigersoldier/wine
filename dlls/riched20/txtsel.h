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

#ifndef __TXTSEL_H
#define __TXTSEL_H

#include <tom.h>

#include "editstr.h"

typedef struct tagReTxtSel ReTxtSel;

ReTxtSel *ReTxtSel_create(ME_TextEditor *editor);
ITextSelection *ReTxtSel_get_ITextSelection(ReTxtSel *txtSel);
/**
 * Mark the attached editor as released so that all calls related to
 * the editor will return CO_E_RELEASED.
 */
void ReTxtSel_releaseEditor(ReTxtSel *txtSel);

#endif /* __TXTSEL_H */
