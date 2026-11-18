/*
 * Styled MS Office Viewer, standalone spreadsheet renderer.
 * Copyright (C) 2026 Zoroaster1x
 *
 * This program is free software: you can redistribute it and/or modify it under
 * the terms of the GNU General Public License as published by the Free Software
 * Foundation, either version 3 of the License, or (at your option) any later
 * version.
 *
 * This program is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
 * details.
 */

// The entry point for hosts that are not Obsidian. It exposes the same parser
// and grid renderer the plugin uses, with no Obsidian API in the graph. The
// source modules are CommonJS, so they arrive as default imports and are
// re-exported under the names a host expects.

import readModule from "../spreadsheet/read.js";
import renderModule from "../spreadsheet/render.js";

export const readWorkbook = readModule.readWorkbook;
export const createRenderer = renderModule.createRenderer;
