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

#ifndef __TXTDOC_H
#define __TXTDOC_H

#include <tom.h>

#include "editstr.h"

typedef struct tagReTxtDoc ReTxtDoc;

/**
 * Create an ReTxtDoc object, which acts as an inner object of COM aggression.
 *
 * The ReTxtDoc object will delegate all of its IUnknown calls to the specified
 * {@code outerObj}.
 *
 * @param outerObj The outer objects that creates and delegates ITextDocument
 *        method calls to the created ReTxtDoc.
 * @param editor The editor implementation.
 *
 * @return An ReTxtDoc object that should be destroyed with {@code ReTxtDoc_destroy}
 */
ReTxtDoc *ReTxtDoc_create(IUnknown *outerObj, ME_TextEditor *editor);
ITextDocument *ReTxtDoc_get_ITextDocument(ReTxtDoc *document);
void ReTxtDoc_destroy(ReTxtDoc *document);

#endif /* __TXTDOC_H */
