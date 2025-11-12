//@target illustrator

(function () {
    'use strict';

    // === EDIT ME: Scripts to display =====================================
    // Provide friendly names, paths, and colors (RGB 0-1). Paths can be:
    //  - Relative (e.g. 'MiniScripts/true_darts.jsx') resolved next to this panel
    //  - Absolute (e.g. 'D:/Scripts/foo.jsx')
    var SCRIPTS = [
        { name: "True Dart", path: "../MiniScripts/true_darts.jsx", color: [0.62, 0.35, 0.71], text: [1, 1, 1] },
        { name: "Equal Paths", path: "../MiniScripts/make_paths_equal.jsx", color: [0.32, 0.51, 0.78], text: [1, 1, 1] },
        { name: "Add Midpoint", path: "../MiniScripts/add_midpoint.jsx", color: [0.85, 0.33, 0.19], text: [1, 1, 1] },
        { name: "Create Dart", path: "../MiniScripts/create_dart.jsx", color: [0.20, 0.6, 0.45], text: [1, 1, 1] },
        { name: "Draw & Label", path: "../MiniScripts/draw_and_label.jsx", color: [0.50, 0.35, 0.85], text: [1, 1, 1] },
        { name: "Mark", path: "../MiniScripts/mark_along_path.jsx", color: [0.95, 0.55, 0.15], text: [0.1, 0.1, 0.1] }
    ];
    // =====================================================================

    // === OPTIONAL LAYOUT TWEAKS ==========================================
    // Edit these numbers if you want larger (or smaller) buttons.
    var BOX_SIZE = {
        minColumnWidth: 75,    // Minimum width of each button in pixels
        maxColumnWidth: 140,   // Maximum width allowed when the panel is wide
        rowHeight: 25          // Button height in pixels
    };
    // =====================================================================

    // === Constants & shared state ========================================
    var PANEL_FILE = File($.fileName);
    var PANEL_DIR = PANEL_FILE.parent;
    var GRID_MARGIN = 14;
    var GRID_SPACING = 8;
    var TEXT_PADDING_H = 14;
    var ELLIPSIS = '\u2026';
    var APP_MAJOR = parseInt(app.version, 10) || 0;
    var ILLUSTRATOR_TARGET = 'illustrator-' + APP_MAJOR;

    if (typeof BridgeTalk === 'undefined') {
        alert('BridgeTalk is not available. Enable BridgeTalk in Illustrator and try again.');
        return;
    }

    var state = {
        entries: [],
        buttons: [],
        rowHeight: clamp(BOX_SIZE.rowHeight || 50, 20, 120),
        scrollOffset: 0,
        columnWidth: clamp(BOX_SIZE.minColumnWidth || 150, 70, 400),
        minColumnWidth: clamp(BOX_SIZE.minColumnWidth || 150, 70, 400),
        maxColumnWidth: clamp(BOX_SIZE.maxColumnWidth || BOX_SIZE.minColumnWidth || 150, 90, 520),
        contentHeight: 0,
        viewHeight: 0,
        maxScroll: 0,
        suppressScrollbarEvent: false,
        window: null,
        viewport: null,
        scrollbar: null,
        font: null,
        smallFont: null
    };

    state.font = getFont('Segoe UI', 'bold', 11);
    state.smallFont = getFont('Segoe UI', 'regular', 10);

    main();

    // === Entry point =====================================================
    function main() {
        try {
            hydrateEntries();
            buildUI();
            buildGrid();
            if (state.window && state.window.layout) {
                state.window.layout.layout(true);
                state.window.layout.resize();
            }
            layoutButtons();
            state.window.center();
            state.window.show();
        } catch (err) {
            alert('My Script Panel failed to load:\n' + err);
        }
    }

    // === Data setup ======================================================
    function hydrateEntries() {
        state.entries = [];
        for (var i = 0; i < SCRIPTS.length; i++) {
            var entry = SCRIPTS[i];
            var file = resolveScriptPath(entry.path || '');
            var normalizedColor = normalizeColor(entry.color, [0.25, 0.25, 0.25]);
            var normalizedText = normalizeColor(entry.text, [1, 1, 1]);
            state.entries.push({
                name: entry.name || file.name || 'Untitled Script',
                path: entry.path || '',
                color: normalizedColor,
                text: normalizedText,
                file: file,
                exists: fileExists(file),
                absolutePath: file.fsName || entry.path || '',
                isBinary: /\.jsxbin$/i.test(file.name || '')
            });
        }
    }

    // === UI construction =================================================
    function buildUI() {
        var win = new Window('palette', 'My Script Panel', undefined, { resizeable: true });
        state.window = win;
        win.orientation = 'column';
        win.alignChildren = ['fill', 'fill'];
        win.spacing = 10;
        win.margins = [12, 12, 12, 12];
        win.minimumSize = [320, 320];

        var scrollArea = win.add('group');
        scrollArea.orientation = 'row';
        scrollArea.alignChildren = ['fill', 'fill'];
        scrollArea.alignment = ['fill', 'fill'];
        scrollArea.spacing = 6;

        var viewport = scrollArea.add('panel', undefined, '');
        viewport.alignment = ['fill', 'fill'];
        viewport.minimumSize.height = 320;
        viewport.margins = [0, 0, 0, 0];
        viewport.spacing = 0;
        viewport.text = '';
        viewport.orientation = 'stack';
        state.viewport = viewport;

        try {
            viewport.addEventListener('mousewheel', handleMouseWheel);
        } catch (wheelErr) {
            viewport.onMouseWheel = handleMouseWheel;
        }

        var scrollbar = scrollArea.add('scrollbar', undefined, 0, 0, 100);
        scrollbar.alignment = ['right', 'fill'];
        scrollbar.preferredSize.width = 12;
        scrollbar.stepdelta = state.rowHeight;
        scrollbar.jumpdelta = state.rowHeight * 2;
        scrollbar.minvalue = 0;
        scrollbar.maxvalue = 0;
        scrollbar.enabled = false;
        scrollbar.onChanging = function () {
            if (state.suppressScrollbarEvent) {
                return;
            }
            setScrollOffset(this.value);
        };
        scrollbar.onChange = scrollbar.onChanging;
        state.scrollbar = scrollbar;

        win.onResizing = win.onResize = function () {
            if (win.layout) {
                win.layout.layout(true);
                win.layout.resize();
            }
            layoutButtons();
        };
    }

    // === Grid construction ===============================================
    function buildGrid() {
        clearGrid();
        for (var i = 0; i < state.entries.length; i++) {
            createButton(state.entries[i]);
        }
        if (state.entries.length === 0) {
            var emptyMsg = state.viewport.add('statictext', undefined, 'Add entries to SCRIPTS to begin.');
            emptyMsg.graphics.font = state.smallFont;
        }
    }

    function clearGrid() {
        if (!state.viewport) {
            return;
        }
        while (state.viewport.children.length) {
            state.viewport.remove(state.viewport.children[0]);
        }
        state.buttons = [];
    }

    function createButton(entry) {
        var btn = state.viewport.add('button', undefined, entry.name);
        btn.margins = 0;
        btn.bounds = [0, 0, state.columnWidth, state.rowHeight];
        btn.minimumSize = [state.minColumnWidth, state.rowHeight];
        btn.maximumSize = [state.maxColumnWidth, state.rowHeight];
        btn.preferredSize = [state.columnWidth, state.rowHeight];
        btn.alignment = ['left', 'top'];
        btn._entry = entry;
        btn._displayText = entry.name;
        btn.enabled = entry.exists;
        if (btn.graphics) {
            btn.graphics.font = state.font;
        }

        if (entry.exists) {
            btn.onClick = function () {
                runEntry(this._entry);
            };
        } else {
            btn.onClick = function () {
                alert('File not found:\n' + (this._entry ? this._entry.path : 'Unknown'));
            };
        }

        var buttonRecord = {
            control: btn,
            entry: entry,
            setBounds: function (x, y, width) {
                var textWidth = Math.max(20, width - (TEXT_PADDING_H * 2));
                var graphics = this.control.graphics;
                var truncated = truncateLabel(entry.name, textWidth, graphics, state.font);
                this.control._displayText = truncated;
                this.control.text = truncated;
                this.control.bounds = [x, y, x + width, y + state.rowHeight];
                this.control.preferredSize = [width, state.rowHeight];
            }
        };

        state.buttons.push(buttonRecord);
    }

    // === Layout & scrolling ==============================================
    function layoutButtons() {
        if (!state.viewport) {
            return;
        }
        var size = getInnerSize(state.viewport);
        var innerWidth = Math.max(2 * state.minColumnWidth + GRID_SPACING, size.width);
        var availableWidth = innerWidth - (GRID_MARGIN * 2);
        var unclamped = Math.floor((availableWidth - GRID_SPACING) / 2);
        var columnWidth = clamp(unclamped, state.minColumnWidth, state.maxColumnWidth);
        var contentWidth = (columnWidth * 2) + GRID_SPACING;
        var horizontalStart = GRID_MARGIN + Math.max(0, Math.floor((availableWidth - contentWidth) / 2));

        var totalRows = Math.ceil(state.buttons.length / 2);
        var rowsHeight = totalRows > 0 ? (totalRows * state.rowHeight) + ((totalRows - 1) * GRID_SPACING) : 0;
        var contentHeight = rowsHeight + (GRID_MARGIN * 2);

        state.columnWidth = columnWidth;
        state.contentHeight = contentHeight;
        state.viewHeight = size.height;
        state.maxScroll = Math.max(0, contentHeight - size.height);
        if (state.scrollOffset > state.maxScroll) {
            state.scrollOffset = state.maxScroll;
        }

        var currentScroll = state.scrollOffset;
        for (var i = 0; i < state.buttons.length; i++) {
            var button = state.buttons[i];
            var columnIndex = i % 2;
            var rowIndex = Math.floor(i / 2);
            var x = horizontalStart + (columnIndex * (columnWidth + GRID_SPACING));
            var y = GRID_MARGIN + (rowIndex * (state.rowHeight + GRID_SPACING)) - currentScroll;
            if (typeof button.setBounds === 'function') {
                button.setBounds(x, y, columnWidth);
            } else if (button.control) {
                button.control.bounds = [x, y, x + columnWidth, y + state.rowHeight];
            }
        }

        updateScrollbarControl();
        if (state.viewport.graphics && typeof state.viewport.graphics.invalidate === 'function') {
            state.viewport.graphics.invalidate();
        } else if (state.viewport && typeof state.viewport.notify === 'function') {
            try {
                state.viewport.notify('onDraw');
            } catch (viewportErr) {
                // Ignore when host lacks notify support.
            }
        }
    }

    function setScrollOffset(value) {
        var clamped = clamp(value, 0, state.maxScroll);
        if (clamped === state.scrollOffset) {
            return;
        }
        state.scrollOffset = clamped;
        layoutButtons();
    }

    function updateScrollbarControl() {
        if (!state.scrollbar) {
            return;
        }
        state.suppressScrollbarEvent = true;
        state.scrollbar.minvalue = 0;
        state.scrollbar.maxvalue = Math.max(0, state.maxScroll);
        state.scrollbar.enabled = state.maxScroll > 0;
        state.scrollbar.value = state.scrollOffset;
        state.suppressScrollbarEvent = false;
    }

    function handleMouseWheel(event) {
        var delta = 0;
        if (event.detail) {
            delta = event.detail;
        } else if (event.wheelDelta) {
            delta = -event.wheelDelta;
        } else if (event.delta) {
            delta = event.delta;
        }
        if (!delta) {
            return;
        }
        var direction = delta > 0 ? 1 : -1;
        setScrollOffset(state.scrollOffset + direction * state.rowHeight);
        if (event.preventDefault) {
            event.preventDefault();
        }
    }

    // === Script execution ================================================
    function runEntry(entry) {
        if (!entry || !entry.file || !entry.file.exists) {
            alert('File not found:\n' + (entry ? entry.path : 'Unknown'));
            return;
        }
        try {
            var body = buildBridgeBody(entry);
            var bt = new BridgeTalk();
            bt.target = ILLUSTRATOR_TARGET;
            bt.body = body;
            bt.onError = function (msg) {
                var details = (msg && msg.body) ? msg.body : 'Unknown BridgeTalk error.';
                alert('BridgeTalk error while running "' + entry.name + '":\n' + details);
            };
            if (!bt.send()) {
                throw new Error('BridgeTalk refused the message.');
            }
        } catch (err) {
            alert('Could not run "' + entry.name + '":\n' + err);
        }
    }

    function buildBridgeBody(entry) {
        var file = entry.file;
        var contents = readFileContents(file, entry.isBinary);
        if (entry.isBinary) {
            return contents;
        }
        var encoded = encodeURIComponent(contents);
        var safe = encoded.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
        var parts = [];
        parts.push('(function(){');
        parts.push('    var __jsx = decodeURIComponent("' + safe + '");');
        parts.push('    eval(__jsx);');
        parts.push('})();');
        return parts.join('\n');
    }

    function readFileContents(file, binary) {
        if (!file.exists) {
            throw new Error('File not found: ' + file.fsName);
        }
        var mode = binary ? 'rb' : 'r';
        try {
            file.encoding = binary ? 'BINARY' : 'UTF8';
            if (!file.open(mode)) {
                throw new Error('Unable to open file.');
            }
            var data = file.read();
            file.close();
            return data;
        } catch (err) {
            try {
                file.close();
            } catch (closeErr) {
                // Ignore cleanup errors.
            }
            throw err;
        }
    }

    // === Helpers =========================================================
    function fileExists(file) {
        try {
            return file && file.exists;
        } catch (err) {
            return false;
        }
    }

    function normalizeColor(value, fallback) {
        if (!value || value.length < 3) {
            return fallback.slice(0);
        }
        return [
            clamp(Number(value[0]) || 0, 0, 1),
            clamp(Number(value[1]) || 0, 0, 1),
            clamp(Number(value[2]) || 0, 0, 1)
        ];
    }

    function clamp(num, min, max) {
        if (num < min) {
            return min;
        }
        if (num > max) {
            return max;
        }
        return num;
    }

    function truncateLabel(label, maxWidth, graphics, font) {
        var ellipsis = ELLIPSIS;
        if (!graphics || !graphics.measureString) {
            var approxChars = Math.max(3, Math.floor(maxWidth / 7));
            if (label.length <= approxChars) {
                return label;
            }
            return label.substring(0, approxChars - 1) + ellipsis;
        }
        graphics.font = font;
        if (graphics.measureString(label)[0] <= maxWidth) {
            return label;
        }
        for (var len = label.length - 1; len > 0; len--) {
            var test = label.substring(0, len) + ellipsis;
            if (graphics.measureString(test)[0] <= maxWidth) {
                return test;
            }
        }
        return ellipsis;
    }

    function resolveScriptPath(spec) {
        if (!spec) {
            return new File('');
        }
        if (spec instanceof File) {
            return spec;
        }
        var cleaned = String(spec).replace(/\\/g, '/');
        if (cleaned.indexOf('~/') === 0) {
            cleaned = Folder('~').fsName.replace(/\\/g, '/') + cleaned.substring(1);
        }
        if (isAbsolutePath(cleaned)) {
            return new File(cleaned);
        }
        var basePath = PANEL_DIR ? PANEL_DIR.fsName.replace(/\\/g, '/') : '';
        if (basePath && basePath.charAt(basePath.length - 1) !== '/') {
            basePath += '/';
        }
        return new File(basePath + cleaned);
    }

    function isAbsolutePath(path) {
        return /^(?:[a-zA-Z]:\/|\/)/.test(path);
    }

    function getInnerSize(control) {
        var bounds = control && control.innerBounds ? control.innerBounds : control.bounds;
        return {
            width: bounds[2] - bounds[0],
            height: bounds[3] - bounds[1]
        };
    }

    function getFont(name, style, size) {
        try {
            return ScriptUI.newFont(name, style, size);
        } catch (err) {
            try {
                return ScriptUI.newFont('Arial', style, size);
            } catch (err2) {
                return ScriptUI.newFont('dialog', 'regular', size);
            }
        }
    }
})();
