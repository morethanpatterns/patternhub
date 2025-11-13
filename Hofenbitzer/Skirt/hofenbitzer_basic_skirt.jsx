/*
 Basic Skirt by Guido Hofenbitzer
 ------------------------------------------------
 - Center front/back, waist/hip/hem guides
 - ScriptUI measurement panel (cm) with default size 38 inputs
 - Artboard resizes with 10 cm margin around content
*/

(function () {
    var SHOW_MEASUREMENT_DIALOG = true;

    var params = {
        HiC: 97,
        WaC: 72,
        HiD: 21,
        MoL: 50,
        HipEase: 3,
        WaistEase: 2,
        FrontDartLength: 10,
        BackDartLength1: 14.5,
        BackDartLength2: 13,
        HipProfile: "Normal",
        WaistShaping: 1,
        SideDart: null,
        FrontDart: null,
        BackDart1: null,
        BackDart2: null,
        HiDEase: 0,
        MoLEase: 0,
        SideDartOverride: false,
        FrontDartOverride: false,
        BackDart1Override: false,
        BackDart2Override: false
    };

    var measurementResults = {};
    var easeResults = {};
    var finalResults = {};
    var measurementLabels = {};
    var measurementOrder = [];
    var measurementPalette = null;
    var shouldShowMeasurementPalette = false;

    var DART_KEYS = ['SideDart','FrontDart','BackDart1','BackDart2'];

    var CM_TO_PT = 28.346456692913385;
    function cm(v) { return v * CM_TO_PT; }

    function formatNumber(val) {
        var num = (typeof val === 'number' && isFinite(val)) ? val : 0;
        var rounded = Math.round(num * 100) / 100;
        var str = rounded.toFixed(2);
        str = str.replace(/\.00$/, "");
        str = str.replace(/(\.\d[1-9])0$/, "$1");
        return str;
    }
    function formatCm(val) { return formatNumber(val); }
    function formatEase(val) {
        var prefix = val > 0 ? "+" : (val < 0 ? "" : "");
        return prefix + formatNumber(val);
    }

    function formatReferenceValue(value) {
        if (value === null || value === undefined) return "-";
        if (typeof value === 'number') {
            if (!isFinite(value)) return "-";
            return formatNumber(value);
        }
        var parsed = Number(value);
        if (!isNaN(parsed)) return formatNumber(parsed);
        return String(value);
    }

    var STROKE_PT = 1;
    var DASH_PT = [25, 12];
    var LABEL_OFFSET_PT = cm(0.5);
    var HIP_CURVE_HANDLE_PT = cm(10.5);
    var MARGIN_CM = 10;
    var LABEL_FONT_SIZE_PT = 12;
    var MARKER_RADIUS_CM = 0.25;
    var NUMBER_FONT_SIZE_PT = 10;
    var MIN_SECOND_BACK_DART_CM = 0.05;

    function rgb(r,g,b){ var c=new RGBColor(); c.red=r; c.green=g; c.blue=b; return c; }
    var COL_BLACK = rgb(0, 0, 0);
    var COL_WHITE = rgb(255, 255, 255);

    if (SHOW_MEASUREMENT_DIALOG) {
        var dialogResult = showMeasurementDialog(params);
        if (!dialogResult) return;
        params = dialogResult;
    }
    params.HipProfile = params.HipProfile || "Normal";
    params.WaistShaping = (params.HipProfile === "Curvy") ? 1.5 : 1;

    var derived = computeDerived(params);

    var frameWidthPt = cm(derived.HiW);
    var frameHeightPt = cm(params.MoL);
    var contentWidthPt = frameWidthPt;
    var contentHeightPt = frameHeightPt;
    var MARGIN_PT = cm(MARGIN_CM);
    var widthPt = contentWidthPt + MARGIN_PT * 2;
    var heightPt = contentHeightPt + MARGIN_PT * 2;

    var doc;
    if (app.documents.length === 0 || (app.activeDocument && app.activeDocument.pageItems.length > 0)) {
        doc = app.documents.add(DocumentColorSpace.RGB, widthPt, heightPt);
    } else {
        doc = app.activeDocument;
    }
    try {
        doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect = [0, heightPt, widthPt, 0];
    } catch (eRect) {}

    function unlockLayerHierarchy(layer) {
        if (!layer) return;
        try { layer.locked = false; } catch (eLock) {}
        try { layer.visible = true; } catch (eVis) {}
        for (var i = 0; i < layer.layers.length; i++) {
            unlockLayerHierarchy(layer.layers[i]);
        }
        var items = layer.pageItems;
        for (var j = 0; j < items.length; j++) {
            try { items[j].locked = false; } catch (eItemLock) {}
            try { items[j].hidden = false; } catch (eItemHide) {}
        }
    }

    function clearDocumentLayers() {
        for (var li = doc.layers.length - 1; li >= 0; li--) {
            unlockLayerHierarchy(doc.layers[li]);
            doc.layers[li].remove();
        }
    }

    function resetLayer(name, layerColor) {
        try {
            var existing = doc.layers.getByName(name);
            unlockLayerHierarchy(existing);
            existing.remove();
        } catch (eExisting) {}
        var layer = doc.layers.add();
        layer.name = name;
        layer.visible = true;
        layer.locked = false;
        if (layerColor) layer.layerColor = layerColor;
        return layer;
    }

    function purgeExtraLayers(allowedNames) {
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var layer = doc.layers[i];
            if (!allowedNames[layer.name]) {
                unlockLayerHierarchy(layer);
                try { layer.remove(); } catch (eRemove) {}
            }
        }
    }

    function ensureLayerWritable(layer) {
        if (!layer) return;
        var current = layer;
        while (current) {
            try { current.locked = false; } catch (eLayerLock) {}
            try { current.visible = true; } catch (eLayerVis) {}
            if (current.parent && current.parent.typename === 'Layer') {
                current = current.parent;
            } else {
                current = null;
            }
        }
        try { doc.activeLayer = layer; } catch (eActive) {}
    }

    function setupLayers() {
        var stack = {};
        stack.basicFrame = resetLayer("Basic Frame", COL_BLACK);
        stack.labelsParent = resetLayer("Labels, Markers & Numbers", COL_BLACK);
        stack.frontConstruction = resetLayer("Front Construction", COL_BLACK);
        stack.backConstruction = resetLayer("Back Construction", COL_BLACK);
        stack.dartsShaping = resetLayer("Darts & Shaping", COL_BLACK);

        stack.labels = stack.labelsParent.layers.add();
        stack.labels.name = "Labels";
        stack.markers = stack.labelsParent.layers.add();
        stack.markers.name = "Markers";
        stack.numbers = stack.labelsParent.layers.add();
        stack.numbers.name = "Numbers";

        stack.dartsLayer = stack.dartsShaping.layers.add();
        stack.dartsLayer.name = "Darts";
        stack.shapingLayer = stack.dartsShaping.layers.add();
        stack.shapingLayer.name = "Shaping";
        return stack;
    }

    function buildPaletteStaticText(group, label, width, justification) {
        var st = group.add('statictext', undefined, label);
        if (width !== undefined) {
            st.preferredSize.width = width;
            st.minimumSize.width = width;
        }
        if (justification) {
            try { st.justify = justification; } catch (eJust) {}
        }
        return st;
    }

    function collectMeasurementRows() {
        var rows = [];
        for (var i = 0; i < measurementOrder.length; i++) {
            var id = measurementOrder[i];
            rows.push({
                id: measurementLabels[id] || id,
                meas: formatReferenceValue(measurementResults[id]),
                ease: formatReferenceValue(easeResults[id]),
                finalValue: formatReferenceValue(finalResults[id])
            });
        }
        return rows;
    }

    function resolveSummaryValue(def, dataSnapshot, derivedSnapshot) {
        if (!def || !def.key) return null;
        var val = null;
        switch (def.type) {
            case 'derived':
            case 'construction':
                val = derivedSnapshot[def.key];
                break;
            case 'readonly':
            case 'derivedInput':
            case 'input':
            case 'dropdown':
            case 'static':
            default:
                val = dataSnapshot[def.key];
                break;
        }
        if (typeof val === 'number') {
            return isFinite(val) ? val : null;
        }
        var parsed = parseFloat(val);
        return isNaN(parsed) ? null : parsed;
    }

    function assignMeasurementSummaryRows(rowDefs, dataSnapshot) {
        measurementResults = {};
        easeResults = {};
        finalResults = {};
        measurementLabels = {};
        measurementOrder = [];
        var derivedSnapshot = computeDerived(dataSnapshot);
        var idUsage = {};

        for (var i = 0; i < rowDefs.length; i++) {
            var rowDef = rowDefs[i];
            if (!rowDef || !rowDef.main) continue;
            var entryLabel = rowDef.main.label || ('Row ' + (rowDef.number || (i + 1)));
            var baseId = rowDef.main.key || (rowDef.construction && rowDef.construction.key) || entryLabel.replace(/\s+/g, '') || ('Row' + (i + 1));
            var suffix = idUsage[baseId] || 0;
            idUsage[baseId] = suffix + 1;
            var entryId = (suffix === 0) ? baseId : (baseId + '_' + suffix);
            var measureVal = resolveSummaryValue(rowDef.main, dataSnapshot, derivedSnapshot);
            var easeVal = resolveSummaryValue(rowDef.ease, dataSnapshot, derivedSnapshot);
            var finalVal = resolveSummaryValue(rowDef.construction, dataSnapshot, derivedSnapshot);
            if (measureVal === null && easeVal === null && finalVal === null) continue;
            measurementOrder.push(entryId);
            measurementLabels[entryId] = entryLabel;
            measurementResults[entryId] = measureVal;
            easeResults[entryId] = easeVal;
            finalResults[entryId] = finalVal;
        }
    }

    function closeMeasurementPalette() {
        try {
            if (measurementPalette && typeof measurementPalette.close === 'function') {
                measurementPalette.close();
            }
        } catch (eCloseLocal) {}
        try {
            var globalPalette = $.global.hofenbitzerSkirtMeasurementPalette;
            if (globalPalette && globalPalette.window && typeof globalPalette.window.close === 'function') {
                globalPalette.window.close();
            }
            delete $.global.hofenbitzerSkirtMeasurementPalette;
        } catch (eCloseGlobal) {}
        measurementPalette = null;
    }

    function showMeasurementPaletteWindow(profileName) {
        closeMeasurementPalette();
        var palette = new Window('palette', 'Measurement Summary');
        palette.orientation = 'column';
        palette.alignChildren = ['fill', 'top'];
        palette.spacing = 8;
        palette.margins = 12;
        palette.preferredSize.width = 420;

        var metaGroup = palette.add('group');
        metaGroup.alignment = 'fill';
        metaGroup.add('statictext', undefined, 'Hip profile: ' + (profileName || '-'));

        var headerGroup = palette.add('group');
        headerGroup.alignment = 'fill';
        headerGroup.spacing = 6;
        var colWidths = [180, 80, 80, 80];
        buildPaletteStaticText(headerGroup, 'Measurement', colWidths[0], 'left');
        buildPaletteStaticText(headerGroup, 'Measure', colWidths[1], 'right');
        buildPaletteStaticText(headerGroup, 'Ease', colWidths[2], 'right');
        buildPaletteStaticText(headerGroup, 'Final', colWidths[3], 'right');

        var rows = collectMeasurementRows();
        if (!rows.length) {
            var emptyRow = palette.add('group');
            emptyRow.alignment = 'fill';
            emptyRow.add('statictext', undefined, 'No measurements to display.');
        } else {
            for (var i = 0; i < rows.length; i++) {
                var row = rows[i];
                var rowGroup = palette.add('group');
                rowGroup.alignment = 'fill';
                rowGroup.spacing = 6;
                buildPaletteStaticText(rowGroup, row.id, colWidths[0], 'left');
                buildPaletteStaticText(rowGroup, row.meas, colWidths[1], 'right');
                buildPaletteStaticText(rowGroup, row.ease, colWidths[2], 'right');
                buildPaletteStaticText(rowGroup, row.finalValue, colWidths[3], 'right');
            }
        }

        palette.onClose = function () {
            measurementPalette = null;
            try { delete $.global.hofenbitzerSkirtMeasurementPalette; } catch (eDel) {}
        };

        measurementPalette = palette;
        $.global.hofenbitzerSkirtMeasurementPalette = { window: palette };
        try { palette.show(); } catch (eShow) {}
    }

    clearDocumentLayers();
    var layers = setupLayers();
    purgeExtraLayers({
        "Basic Frame": true,
        "Labels, Markers & Numbers": true,
        "Front Construction": true,
        "Back Construction": true,
        "Darts & Shaping": true
    });

    function getStrokeColorForLayer(layer) {
        return COL_BLACK;
    }

    function drawLine(layer, x1, y1, x2, y2, strokeColor, dashArray) {
        if (!layer) return null;
        ensureLayerWritable(layer);
        var path = layer.pathItems.add();
        path.stroked = true;
        path.strokeWidth = STROKE_PT;
        path.strokeColor = strokeColor || getStrokeColorForLayer(layer);
        path.strokeDashes = (dashArray && dashArray.length) ? dashArray : [];
        path.strokeCap = StrokeCap.BUTTENDCAP;
        path.strokeJoin = StrokeJoin.MITERENDJOIN;
        path.filled = false;
        path.closed = false;
        path.setEntirePath([[x1, y1], [x2, y2]]);
        return path;
    }

    function centerTextFrame(tf, anchor) {
        if (!tf || !anchor) return;
        try {
            var bounds = tf.visibleBounds; // [x1, y1, x2, y2]
            var currentCenterX = (bounds[0] + bounds[2]) / 2;
            var currentCenterY = (bounds[1] + bounds[3]) / 2;
            tf.translate(anchor[0] - currentCenterX, anchor[1] - currentCenterY);
        } catch (eBounds) {
            try {
                tf.left = anchor[0] - tf.width / 2;
                tf.top = anchor[1] + tf.height / 2;
            } catch (eFallback) {}
        }
    }

    function addParallelLineLabel(text, pA, pB, options) {
        if (!text || !pA || !pB) return null;
        options = options || {};
        var targetLayer = options.layer || layers.labels;
        if (!targetLayer || !targetLayer.textFrames) return null;
        ensureLayerWritable(targetLayer);

        var dx = pB[0] - pA[0];
        var dy = pB[1] - pA[1];
        var length = Math.sqrt(dx * dx + dy * dy);
        if (length === 0) length = 1;

        var side = options.side || 1;
        var offset = options.offsetPt || 0;
        var nx = -dy / length;
        var ny = dx / length;
        var midX = (pA[0] + pB[0]) / 2;
        var midY = (pA[1] + pB[1]) / 2;
        var position = [midX + nx * offset * side, midY + ny * offset * side];

        var tf = targetLayer.textFrames.add();
        tf.contents = text;
        try { tf.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (eJust) {}
        try { tf.textRange.characterAttributes.size = LABEL_FONT_SIZE_PT; } catch (eSize) {}
        try { tf.textRange.characterAttributes.fillColor = COL_BLACK; } catch (eFill) {}
        try { tf.textRange.characterAttributes.textFont = app.textFonts.getByName('ArialMT'); } catch (eFont) {}

        centerTextFrame(tf, position);

        var angleDeg = Math.atan2(dy, dx) * 180 / Math.PI;
        if (angleDeg < -90 || angleDeg > 90) {
            angleDeg += 180;
        }
        try { tf.rotate(angleDeg); } catch (eRotate) {}
        centerTextFrame(tf, position);

        return tf;
    }

    function drawCurveWithHandle(layer, startPt, endPt, labelText, options) {
        if (!layer || !startPt || !endPt) return null;
        options = options || {};
        ensureLayerWritable(layer);
        var path = layer.pathItems.add();
        path.stroked = true;
        path.strokeWidth = STROKE_PT;
        path.strokeColor = options.strokeColor || getStrokeColorForLayer(layer);
        path.strokeDashes = (options.dash && options.dash.length) ? options.dash : [];
        path.filled = false;
        path.closed = false;
        path.setEntirePath([startPt, endPt]);

        if (path.pathPoints.length === 2) {
            var startPoint = path.pathPoints[0];
            startPoint.pointType = PointType.SMOOTH;
            startPoint.leftDirection = startPt;
            startPoint.rightDirection = startPt;

            var endPoint = path.pathPoints[1];
            endPoint.pointType = PointType.SMOOTH;
            var handle = options.handlePoint || endPt;
            endPoint.leftDirection = handle;
            endPoint.rightDirection = endPt;
        }

        if (labelText) {
            addParallelLineLabel(labelText, startPt, endPt, {
                offsetPt: LABEL_OFFSET_PT,
                layer: layers.labels,
                side: options.side || 1
            });
        }

        return path;
    }

    function placeMarker(x, y, number) {
        if (!layers.markers || !layers.numbers) return null;
        ensureLayerWritable(layers.markers);
        var radius = cm(MARKER_RADIUS_CM);
        var circle = layers.markers.pathItems.ellipse(y + radius, x - radius, radius * 2, radius * 2);
        circle.stroked = false;
        circle.filled = true;
        circle.fillColor = COL_BLACK;

        if (typeof number === 'number') {
            ensureLayerWritable(layers.numbers);
            var tf = layers.numbers.textFrames.add();
            tf.contents = String(number);
            try { tf.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (eJust) {}
            try { tf.textRange.characterAttributes.size = NUMBER_FONT_SIZE_PT; } catch (eSize) {}
            try { tf.textRange.characterAttributes.fillColor = COL_WHITE; } catch (eFill) {}
            try { tf.textRange.characterAttributes.textFont = app.textFonts.getByName('Arial-BoldMT'); } catch (eFont) {}
            centerTextFrame(tf, [x, y]);
        }
    }

    // ---------------------------------
    // Drafting references (markers + frame)
    // ---------------------------------
    var activeArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
    var artRect = activeArtboard.artboardRect; // [left, top, right, bottom]
    var artLeft = artRect[0];
    var artTop = artRect[1];

    function ptOffset(dx, dy) {
        return [artLeft + MARGIN_PT + dx, artTop - MARGIN_PT - dy];
    }

    function lerpPoint(a, b, t) {
        if (!a || !b) return [0, 0];
        if (t === undefined) t = 0;
        return [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t];
    }

    function horizontalIntersection(pA, pB, targetY) {
        if (!pA || !pB) return [0, targetY];
        var dy = pB[1] - pA[1];
        if (Math.abs(dy) < 0.0001) return [pA[0], targetY];
        var t = (targetY - pA[1]) / dy;
        return [pA[0] + (pB[0] - pA[0]) * t, targetY];
    }

    var L_mol = cm(params.MoL);
    var L_hiW = cm(derived.HiW);
    var L_hiD = cm(params.HiD);
    var L_halfHiW = L_hiW / 2;

    var P1 = ptOffset(0, 0);
    var P2 = ptOffset(0, L_mol);
    var P4 = ptOffset(L_hiW, 0);
    var P5 = ptOffset(L_hiW, L_mol);
    var P3 = ptOffset(0, L_hiD);
    var P6 = ptOffset(L_hiW, L_hiD);
    var P7 = ptOffset(L_halfHiW, 0);
    var P8 = ptOffset(L_halfHiW, L_mol);
    var P9 = ptOffset(L_halfHiW, L_hiD);

    var waistShapingPt = cm(determineWaistShaping(params.HipProfile));
    var P10 = [P7[0], P7[1] + waistShapingPt];

    var halfSideDartPt = cm(Math.max(0, (derived.SideDart || 0) / 2));
    var P11 = [P10[0] - halfSideDartPt, P10[1]];
    var P12 = [P10[0] + halfSideDartPt, P10[1]];

    var waistLineY = P1[1];
    var frontHipIntersection = horizontalIntersection([P11[0], P11[1]], [P9[0], P9[1]], waistLineY);
    var frontDartAnchorOffset = cm((params.WaC || 0) / 10);
    var P13 = [frontHipIntersection[0] - frontDartAnchorOffset, waistLineY];
    var P13dashHalf = cm(2.5);
    var P13dashOffset = cm((params.HipProfile === 'Curvy') ? 0.7 : 0.5);
    var P13TopY = P13[1] + P13dashOffset;
    var P13dashLeft = [P13[0] - P13dashHalf, P13TopY];
    var P13dashRight = [P13[0] + P13dashHalf, P13TopY];
    var frontDartLengthPt = cm(Math.max(10, params.FrontDartLength || 10));
    var P13Base = [P13[0], P13[1] - frontDartLengthPt];
    var halfFrontDartPt = cm(Math.max(0, (derived.FrontDart || 0) / 2));
    var P13TopLeft = [P13[0] - halfFrontDartPt, P13TopY];
    var P13TopRight = [P13[0] + halfFrontDartPt, P13TopY];

    var waistY = P1[1];
    var cbWaistPoint = [P4[0], waistY];
    var backHipWaistPoint = horizontalIntersection([P12[0], P12[1]], [P9[0], P9[1]], waistY);
    var rawBackDart1 = Math.max(0, derived.BackDart1 || 0);
    var rawBackDart2 = Math.max(0, derived.BackDart2 || 0);
    var hasSecondBackDart = rawBackDart2 > MIN_SECOND_BACK_DART_CM;
    var P14 = lerpPoint(cbWaistPoint, backHipWaistPoint, hasSecondBackDart ? (1 / 3) : 0.5);
    var backDartLength1Pt = cm(Math.max(0, params.BackDartLength1 || 0));
    var P14Base = [P14[0], P14[1] - backDartLength1Pt];
    var halfBackDart1Pt = cm(rawBackDart1 / 2);
    var P14Left = [P14[0] - halfBackDart1Pt, P14[1]];
    var P14Right = [P14[0] + halfBackDart1Pt, P14[1]];
    var P14UpperLeft = P14Left;
    var P14UpperRight = P14Right;
    var singleBackDashLeft = null;
    var singleBackDashRight = null;
    if (!hasSecondBackDart) {
        var singleDashHalf = cm(2.5);
        var singleDashOffset = cm((params.HipProfile === 'Curvy') ? 0.5 : 0.3);
        var singleGuideY = P14[1] + singleDashOffset;
        singleBackDashLeft = [P14[0] - singleDashHalf, singleGuideY];
        singleBackDashRight = [P14[0] + singleDashHalf, singleGuideY];
        P14UpperLeft = [P14[0] - halfBackDart1Pt, singleGuideY];
        P14UpperRight = [P14[0] + halfBackDart1Pt, singleGuideY];
    }
    var halfBackDart2Pt = cm(rawBackDart2 / 2);
    var backDartLength2Pt = cm(Math.max(0, params.BackDartLength2 || 0));
    var P15 = null;
    var P15Base = null;
    var P15Left = null;
    var P15Right = null;
    var P15dashLeft = null;
    var P15dashRight = null;
    var P15TopOffset = null;
    var P15TopY = null;
    if (hasSecondBackDart) {
        P15 = lerpPoint(P14Left, backHipWaistPoint, 0.5);
        P15Base = [P15[0], P15[1] - backDartLength2Pt];
        var secondDashHalf = cm(2.5);
        P15TopOffset = cm((params.HipProfile === 'Curvy') ? 0.5 : 0.3);
        P15TopY = P15[1] + P15TopOffset;
        P15dashLeft = [P15[0] - secondDashHalf, P15TopY];
        P15dashRight = [P15[0] + secondDashHalf, P15TopY];
        P15Left = [P15[0] - halfBackDart2Pt, P15TopY];
        P15Right = [P15[0] + halfBackDart2Pt, P15TopY];
    }

    // CF line (1 to 2)
    var line12 = drawLine(layers.basicFrame, P1[0], P1[1], P2[0], P2[1], null, []);
    try { line12.name = 'Centre Front Line'; } catch (eCFName) {}
    addParallelLineLabel('Centre Front (CF)', P1, P2, { offsetPt: LABEL_OFFSET_PT, layer: layers.labels, side: 1 });
    // Top horizontal (1 to 4)
    var line14 = drawLine(layers.basicFrame, P1[0], P1[1], P4[0], P4[1], null, []);
    try { line14.name = 'Waist Line'; } catch (eWaistName) {}
    addParallelLineLabel('Waist Line', P1, P4, { offsetPt: LABEL_OFFSET_PT, layer: layers.labels, side: 1 });
    // Right vertical (4 to 5)
    var line45 = drawLine(layers.basicFrame, P4[0], P4[1], P5[0], P5[1], null, []);
    try { line45.name = 'Centre Back Line'; } catch (eCBName) {}
    addParallelLineLabel('Centre Back (CB)', P4, P5, { offsetPt: LABEL_OFFSET_PT, layer: layers.labels, side: -1 });
    // Bottom horizontal (5 to 2)
    var line25 = drawLine(layers.basicFrame, P5[0], P5[1], P2[0], P2[1], null, []);
    try { line25.name = 'Hem Line'; } catch (eHemName) {}
    addParallelLineLabel('Hem Line', P5, P2, { offsetPt: LABEL_OFFSET_PT, layer: layers.labels, side: -1 });
    // Mid vertical (7 to 8)
    var line78 = drawLine(layers.basicFrame, P7[0], P7[1], P8[0], P8[1], null, []);
    try { line78.name = 'Side Line'; } catch (eMidName) {}
    addParallelLineLabel('Side Line', P7, P8, { offsetPt: LABEL_OFFSET_PT, layer: layers.labels, side: 1 });

    // Hip depth guide (3 horizontally across to meet 4-5 line) dashed
    var line36 = drawLine(layers.basicFrame, P3[0], P3[1], P6[0], P6[1], null, DASH_PT);
    try { line36.name = 'Hip Line'; } catch (eHipName) {}
    addParallelLineLabel('Hip Line', P3, P6, { offsetPt: LABEL_OFFSET_PT, layer: layers.labels, side: 1 });

    // Extended top from 7 upwards to point 10
    var line710 = drawLine(layers.basicFrame, P7[0], P7[1], P10[0], P10[1], null, []);
    try { line710.name = 'Waist Shaping Guide'; } catch (eShapeName) {}

    // Side dart halves from 10
    if (halfSideDartPt > 0) {
        var line10Left = drawLine(layers.dartsLayer, P10[0], P10[1], P11[0], P11[1], null, []);
        try { line10Left.name = 'Side Dart Left'; } catch (eSDL) {}
        var line10Right = drawLine(layers.dartsLayer, P10[0], P10[1], P12[0], P12[1], null, []);
        try { line10Right.name = 'Side Dart Right'; } catch (eSDR) {}
        var hipHandlePoint = [P7[0], P9[1] + HIP_CURVE_HANDLE_PT];
        var frontHipCurve = drawCurveWithHandle(layers.shapingLayer, [P11[0], P11[1]], [P9[0], P9[1]], 'Front Hip Curve', { side: -1, handlePoint: hipHandlePoint });
        try { frontHipCurve.name = 'Front Hip Curve'; } catch (eFrontName) {}
        var backHipCurve = drawCurveWithHandle(layers.shapingLayer, [P12[0], P12[1]], [P9[0], P9[1]], 'Back Hip Curve', { side: 1, handlePoint: hipHandlePoint });
        try { backHipCurve.name = 'Back Hip Curve'; } catch (eBackName) {}
    }

    // Waist shaping dash (12 cm centered at P10)
    var P10Left = [P10[0] - cm(6), P10[1]];
    var P10Right = [P10[0] + cm(6), P10[1]];
    var waistGuide = drawLine(layers.basicFrame, P10Left[0], P10Left[1], P10Right[0], P10Right[1], null, DASH_PT);
    try { waistGuide.name = 'Upper Waist Shaping Guide'; } catch (eGuideName) {}
    var frontGuide = drawLine(layers.dartsLayer, P13dashLeft[0], P13dashLeft[1], P13dashRight[0], P13dashRight[1], null, DASH_PT);
    try { frontGuide.name = 'Front Dart Guide'; } catch (eFrontGuide) {}
    var line13Down = drawLine(layers.dartsLayer, P13[0], P13[1], P13Base[0], P13Base[1], null, DASH_PT);
    try { line13Down.name = 'Front Dart Centre'; } catch (eDartCenter) {}
    if (halfFrontDartPt > 0) {
        var line13Left = drawLine(layers.dartsLayer, P13TopLeft[0], P13TopLeft[1], P13Base[0], P13Base[1], null, []);
        try { line13Left.name = 'Front Dart Left'; } catch (eDartLeft) {}
        var line13Right = drawLine(layers.dartsLayer, P13TopRight[0], P13TopRight[1], P13Base[0], P13Base[1], null, []);
        try { line13Right.name = 'Front Dart Right'; } catch (eDartRight) {}
        addParallelLineLabel('Front Dart', P13TopLeft, P13TopRight, { offsetPt: LABEL_OFFSET_PT, layer: layers.labels, side: 1 });
    }

    if (singleBackDashLeft && singleBackDashRight) {
        var backGuide = drawLine(layers.dartsLayer, singleBackDashLeft[0], singleBackDashLeft[1], singleBackDashRight[0], singleBackDashRight[1], null, DASH_PT);
        try { backGuide.name = 'Back Dart Guide'; } catch (eBackGuide) {}
    }
    if (backDartLength1Pt > 0) {
        var line14Down = drawLine(layers.dartsLayer, P14[0], P14[1], P14Base[0], P14Base[1], null, DASH_PT);
        try { line14Down.name = '1st Back Dart Centre'; } catch (eBackCenter) {}
    }
    if (halfBackDart1Pt > 0) {
        var line14Left = drawLine(layers.dartsLayer, P14UpperLeft[0], P14UpperLeft[1], P14Base[0], P14Base[1], null, []);
        try { line14Left.name = 'First Back Dart Left'; } catch (eBackLeft) {}
        var line14Right = drawLine(layers.dartsLayer, P14UpperRight[0], P14UpperRight[1], P14Base[0], P14Base[1], null, []);
        try { line14Right.name = 'First Back Dart Right'; } catch (eBackRight) {}
        addParallelLineLabel('1st Back Dart', P14UpperLeft, P14UpperRight, { offsetPt: LABEL_OFFSET_PT, layer: layers.labels, side: 1 });
    }

    if (hasSecondBackDart && P15dashLeft && P15dashRight) {
        drawLine(layers.dartsLayer, P15dashLeft[0], P15dashLeft[1], P15dashRight[0], P15dashRight[1], null, DASH_PT);
    }
    if (hasSecondBackDart && backDartLength2Pt > 0 && P15 && P15Base) {
        var line15Down = drawLine(layers.dartsLayer, P15[0], P15[1], P15Base[0], P15Base[1], null, DASH_PT);
        try { line15Down.name = '2nd Back Dart Centre'; } catch (eBackCenter2) {}
    }
    if (hasSecondBackDart && halfBackDart2Pt > 0 && P15Left && P15Right && P15Base) {
        var line15Left = drawLine(layers.dartsLayer, P15Left[0], P15Left[1], P15Base[0], P15Base[1], null, []);
        try { line15Left.name = 'Second Back Dart Left'; } catch (eBack2Left) {}
        var line15Right = drawLine(layers.dartsLayer, P15Right[0], P15Right[1], P15Base[0], P15Base[1], null, []);
        try { line15Right.name = 'Second Back Dart Right'; } catch (eBack2Right) {}
        addParallelLineLabel('2nd Back Dart', P15Left, P15Right, { offsetPt: LABEL_OFFSET_PT, layer: layers.labels, side: 1 });
    }

    // Markers
    placeMarker(P1[0], P1[1], 1);
    placeMarker(P2[0], P2[1], 2);
    placeMarker(P3[0], P3[1], 3);
    placeMarker(P4[0], P4[1], 4);
    placeMarker(P5[0], P5[1], 5);
    placeMarker(P6[0], P6[1], 6);
    placeMarker(P7[0], P7[1], 7);
    placeMarker(P8[0], P8[1], 8);
    placeMarker(P9[0], P9[1], 9);
    placeMarker(P10[0], P10[1], 10);
    if (halfSideDartPt > 0) {
        placeMarker(P11[0], P11[1], 11);
        placeMarker(P12[0], P12[1], 12);
    }
    placeMarker(P13[0], P13[1], 13);
    placeMarker(P14[0], P14[1], 14);
    if (hasSecondBackDart && P15) {
        placeMarker(P15[0], P15[1], 15);
    }

    if (shouldShowMeasurementPalette) {
        showMeasurementPaletteWindow(params.HipProfile || 'Normal');
    } else {
        closeMeasurementPalette();
    }

    function determineWaistShaping(profile) {
        return (profile === "Curvy") ? 1.5 : 1;
    }

    function determineSideDartBase(profile, WaDif) {
        if (profile === "Curvy") return Math.max(0, WaDif / 2 + 1);
        if (profile === "Flat") return Math.max(0, WaDif / 2 - 1);
        return Math.max(0, WaDif / 2);
    }

    function computeDerived(p) {
        var result = {};
        var hipWidth = (Number(p.HiC) + Number(p.HipEase || 0)) / 2;
        var waistWidth = (Number(p.WaC) + Number(p.WaistEase || 0)) / 2;
        if (!isFinite(hipWidth)) hipWidth = 0;
        if (!isFinite(waistWidth)) waistWidth = 0;
        var waistDiffTarget = Math.max(0, hipWidth - waistWidth);

        result.HiW = hipWidth;
        result.WaW = waistWidth;

        var sideBase = determineSideDartBase(p.HipProfile || "Normal", waistDiffTarget);
        var manualSide = Number(p.SideDart);
        var sideVal = (p.SideDartOverride && isFinite(manualSide)) ? manualSide : sideBase;
        result.SideDart = sideVal;

        var frontAuto = Math.min(waistDiffTarget * 0.2, 2.5);
        var manualFront = Number(p.FrontDart);
        var frontVal = (p.FrontDartOverride && isFinite(manualFront)) ? manualFront : frontAuto;
        result.FrontDart = frontVal;

        var back1Auto = Math.min(waistDiffTarget * 0.3, 4.5);
        var manualBack1 = Number(p.BackDart1);
        var back1Val = (p.BackDart1Override && isFinite(manualBack1)) ? manualBack1 : back1Auto;
        result.BackDart1 = back1Val;

        var remainder = waistDiffTarget - sideVal - frontVal - back1Val;
        var back2Auto = Math.max(0, remainder);
        var manualBack2 = Number(p.BackDart2);
        var back2Val = (p.BackDart2Override && isFinite(manualBack2)) ? manualBack2 : back2Auto;
        result.BackDart2 = back2Val;

        var actualWaistDiff = Math.max(0, sideVal + frontVal + back1Val + back2Val);
        result.DartSum = actualWaistDiff;

        var manualSum = 0;
        for (var md = 0; md < DART_KEYS.length; md++) {
            var k = DART_KEYS[md];
            if (p[k + 'Override']) {
                manualSum += Math.max(0, Number(p[k]) || 0);
            }
        }

        result.HiDCons = (Number(p.HiD) || 0) + (Number(p.HiDEase) || 0);
        result.MoLCons = (Number(p.MoL) || 0) + (Number(p.MoLEase) || 0);
        result.WaistDiffTarget = waistDiffTarget;
        result.ManualDartSum = manualSum;
        result.WaDif = waistDiffTarget - manualSum;
        return result;
    }

    function showMeasurementDialog(initial) {
        var data = {
            HiC: initial.HiC,
            WaC: initial.WaC,
            HiD: initial.HiD,
            MoL: initial.MoL,
            SideDart: initial.SideDart,
            FrontDart: initial.FrontDart,
            BackDart1: initial.BackDart1,
            BackDart2: initial.BackDart2,
            HiDEase: initial.HiDEase || 0,
            MoLEase: initial.MoLEase || 0,
            HipEase: initial.HipEase,
            WaistEase: initial.WaistEase,
            FrontDartLength: initial.FrontDartLength,
            BackDartLength1: initial.BackDartLength1,
            BackDartLength2: initial.BackDartLength2,
        HipProfile: initial.HipProfile || 'Normal',
        SideDartOverride: initial.SideDartOverride || false,
        FrontDartOverride: initial.FrontDartOverride || false,
        BackDart1Override: initial.BackDart1Override || false,
        BackDart2Override: initial.BackDart2Override || false
        };
        data.WaistShaping = determineWaistShaping(data.HipProfile);

        var dlg = new Window('dialog', 'Measurement Panel in cm, default size is 38');
        dlg.orientation = 'column';
        dlg.alignChildren = 'fill';
        dlg.spacing = 14;

        var DIALOG_COL_WIDTHS = [260, 170, 230];
        var ROW_GAP = 8;

        var panel = dlg.add('panel', undefined, 'Inputs (cm)');
        panel.orientation = 'column';
        panel.alignChildren = 'fill';
        panel.margins = 16;
        panel.spacing = 10;

        var header = panel.add('group');
        header.orientation = 'row';
        header.alignChildren = ['left', 'center'];
        header.spacing = 16;
        var headerLabels = ['#', 'Main measurement', 'Ease', 'Construction measurement'];
        for (var h = 0; h < headerLabels.length; h++) {
            var hdr = header.add('statictext', undefined, headerLabels[h]);
            if (h === 0) {
                hdr.minimumSize.width = 30;
            } else {
                hdr.minimumSize.width = DIALOG_COL_WIDTHS[h - 1];
            }
        }

        var rowsContainer = panel.add('group');
        rowsContainer.orientation = 'column';
        rowsContainer.alignChildren = 'fill';
        rowsContainer.spacing = ROW_GAP;

        var inputs = {};
        var derivedFields = {};
        var derivedInputFields = {};
        var lastCommittedValues = {};
        var manualOverrides = {};
        var overrideKeys = DART_KEYS.slice();
        var dartsAutoDisabled = false;
        for (var mk = 0; mk < overrideKeys.length; mk++) {
            var k = overrideKeys[mk];
            if (data[k + 'Override']) {
                manualOverrides[k] = true;
                dartsAutoDisabled = true;
            }
        }

        function activateManualDartMode(activeKey) {
            if (dartsAutoDisabled) return;
            dartsAutoDisabled = true;
            for (var i = 0; i < overrideKeys.length; i++) {
                var key = overrideKeys[i];
                manualOverrides[key] = true;
                data[key + 'Override'] = true;
                var currentVal = parseFloat(inputs[key] ? inputs[key].text : data[key]);
                if (!isNaN(currentVal)) {
                    data[key] = currentVal;
                    lastCommittedValues[key] = currentVal;
                } else {
                    data[key] = 0;
                    lastCommittedValues[key] = 0;
                    if (inputs[key]) inputs[key].text = '';
                }
            }
            updateDerived();
        }

        function allDartInputsEmpty() {
            for (var i = 0; i < overrideKeys.length; i++) {
                var key = overrideKeys[i];
                var field = inputs[key];
                if (!field) return false;
                var txt = (field.text || '').replace(/\s+/g, '');
                if (txt.length > 0) return false;
            }
            return true;
        }

        function restoreAutomaticDartMode() {
            dartsAutoDisabled = false;
            for (var i = 0; i < overrideKeys.length; i++) {
                var key = overrideKeys[i];
                manualOverrides[key] = false;
                data[key + 'Override'] = false;
                delete data[key];
                lastCommittedValues[key] = 0;
            }
            updateDerived();
        }

        function maybeRestoreAutoAfterClear() {
            if (!dartsAutoDisabled) return;
            if (allDartInputsEmpty()) {
                restoreAutomaticDartMode();
            }
        }

        function styleReadOnlyBox(box) {
            if (!box) return;
            box.enabled = false;
            try {
                var brush = box.graphics.newBrush(box.graphics.BrushType.SOLID_COLOR, [0.92, 0.92, 0.92, 1]);
                box.graphics.backgroundColor = brush;
            } catch (eBrush) {}
        }

        function createEditBox(parent, value, readOnly) {
            var init = (value === null || value === undefined) ? '' : String(value);
            var edit = parent.add('edittext', undefined, init);
            edit.characters = 8;
            if (readOnly) styleReadOnlyBox(edit);
            return edit;
        }

        function createDropdownControl(parent, def) {
            var dropdown = parent.add('dropdownlist', undefined, def.options || []);
            dropdown.selection = dropdown.items[0] || null;
            if (data[def.key]) {
                for (var di = 0; di < dropdown.items.length; di++) {
                    if (dropdown.items[di].text === data[def.key]) {
                        dropdown.selection = dropdown.items[di];
                        break;
                    }
                }
            }
            dropdown.onChange = function () {
                if (dropdown.selection) data[def.key] = dropdown.selection.text;
                data.WaistShaping = determineWaistShaping(data.HipProfile);
                if (inputs.WaistShaping) inputs.WaistShaping.text = formatNumber(data.WaistShaping);
                updateDerived();
            };
            return dropdown;
        }

        var dialogRows = [
            { number: 1, main: { type: 'input', key: 'HiC', label: 'HiC' }, ease: { type: 'input', key: 'HipEase', label: 'Ease' }, construction: { type: 'construction', key: 'HiW', label: 'HiW' } },
            { number: 2, main: { type: 'input', key: 'WaC', label: 'WaC' }, ease: { type: 'input', key: 'WaistEase', label: 'Ease' }, construction: { type: 'construction', key: 'WaW', label: 'WaW' } },
            { number: 3, main: { type: 'input', key: 'HiD', label: 'HiD' }, ease: null, construction: null },
            { number: 4, main: { type: 'input', key: 'MoL', label: 'MoL' }, ease: null, construction: null },
            { number: 5, main: { type: 'derived', key: 'WaDif', label: 'WaDif' }, ease: null, construction: null },
            { number: 6, main: { type: 'derivedInput', key: 'SideDart', label: 'Side dart' }, ease: null, construction: null },
            { number: 7, main: { type: 'derivedInput', key: 'FrontDart', label: 'Front dart' }, ease: null, construction: null },
            { number: 8, main: { type: 'derivedInput', key: 'BackDart1', label: '1st back dart' }, ease: null, construction: null },
            { number: 9, main: { type: 'derivedInput', key: 'BackDart2', label: '2nd back dart' }, ease: null, construction: null },
            { number: 10, main: { type: 'input', key: 'FrontDartLength', label: 'Front dart length (8-10 cm)' }, ease: null, construction: null },
            { number: 11, main: { type: 'input', key: 'BackDartLength1', label: '1st back dart length (13-16 cm)' }, ease: null, construction: null },
            { number: 12, main: { type: 'input', key: 'BackDartLength2', label: '2nd back dart length (12-14 cm)' }, ease: null, construction: null },
            { number: 13, main: { type: 'readonly', key: 'WaistShaping', label: 'Waist shaping (1-1.5 cm)' }, ease: null, construction: null },
            { number: 14, main: { type: 'dropdown', key: 'HipProfile', label: 'Hip profile', options: ['Flat', 'Normal', 'Curvy'] }, ease: null, construction: null }
        ];

        for (var r = 0; r < dialogRows.length; r++) {
            var rowDef = dialogRows[r];
            var rowGroup = rowsContainer.add('group');
            rowGroup.orientation = 'row';
            rowGroup.alignChildren = ['left', 'center'];
            rowGroup.spacing = 16;
            rowGroup.margins = [0, 4, 0, 4];

            var numberCell = rowGroup.add('group');
            numberCell.minimumSize.width = 30;
            var numText = '';
            if (rowDef.number !== null && rowDef.number !== undefined && rowDef.number !== '') {
                numText = rowDef.number + '.';
            }
            numberCell.add('statictext', undefined, numText);

            createCell(rowDef.main, rowGroup, DIALOG_COL_WIDTHS[0]);
            createCell(rowDef.ease, rowGroup, DIALOG_COL_WIDTHS[1]);
            createCell(rowDef.construction, rowGroup, DIALOG_COL_WIDTHS[2]);
        }

        var summaryGroup = dlg.add('group');
        summaryGroup.alignment = 'left';
        summaryGroup.alignChildren = ['left', 'center'];
        summaryGroup.margins = [0, 6, 0, 0];
        var summaryCheckbox = summaryGroup.add('checkbox', undefined, 'Show measurement summary');
        summaryCheckbox.value = shouldShowMeasurementPalette ? true : false;

        var buttonGroup = dlg.add('group');
        buttonGroup.alignment = 'right';
        buttonGroup.spacing = 8;
        var okBtn = buttonGroup.add('button', undefined, 'OK', { name: 'ok' });
        var cancelBtn = buttonGroup.add('button', undefined, 'Cancel', { name: 'cancel' });

        cancelBtn.onClick = function () { dlg.close(0); };
        okBtn.onClick = function () {
            for (var key in inputs) {
                if (inputs.hasOwnProperty(key)) {
                    var val = parseFloat(inputs[key].text);
                    if (!isNaN(val)) data[key] = val;
                }
            }
            dlg.close(1);
        };

        function createCell(def, parent, width) {
            if (!def) return;
            var cell = parent.add('group');
            cell.orientation = 'row';
            cell.alignChildren = ['left', 'center'];
            cell.spacing = 4;
            if (width) cell.minimumSize.width = width;

            if (def.type !== 'label') {
                cell.add('statictext', undefined, def.label + ':');
            }

            switch (def.type) {
                case 'input':
                    var initial = (data[def.key] !== null && data[def.key] !== undefined) ? formatNumber(data[def.key]) : '';
                    var edit = createEditBox(cell, initial, false);
                    inputs[def.key] = edit;
                    edit.onChanging = function () {
                        var val = parseFloat(edit.text);
                        if (!isNaN(val)) data[def.key] = val;
                        updateDerived();
                    };
                    break;
                case 'dropdown':
                    createDropdownControl(cell, def);
                    break;
                case 'derived':
                    registerDerivedField(def.key, createEditBox(cell, '', true));
                    break;
                case 'construction':
                    registerDerivedField(def.key, createEditBox(cell, '', true));
                    break;
                case 'static':
                    createEditBox(cell, def.value, true);
                    break;
                case 'derivedInput':
                    var derivedSnapshot = computeDerived(data)[def.key] || 0;
                    var initialValue = data[def.key + 'Override'] ? data[def.key] : derivedSnapshot;
                    var derivedEdit = createEditBox(cell, formatNumber(initialValue), false);
                    inputs[def.key] = derivedEdit;
                    derivedInputFields[def.key] = derivedEdit;
                    (function (key, editControl) {
                        lastCommittedValues[key] = isFinite(data[key]) ? data[key] : derivedSnapshot;
                        function handleValue(isFinal) {
                            if (!dartsAutoDisabled) activateManualDartMode(key);
                            var val = parseFloat(editControl.text);
                            var trimmed = (editControl.text || '').replace(/\s+/g, '');
                            if (!isNaN(val)) {
                                data[key] = val;
                                manualOverrides[key] = true;
                                data[key + 'Override'] = true;
                                lastCommittedValues[key] = val;
                                updateDerived();
                            } else if (trimmed.length === 0) {
                                if (dartsAutoDisabled) {
                                    data[key] = 0;
                                    manualOverrides[key] = true;
                                    data[key + 'Override'] = true;
                                    lastCommittedValues[key] = 0;
                                    updateDerived();
                                    if (isFinal) maybeRestoreAutoAfterClear();
                                } else if (isFinal) {
                                    editControl.text = '';
                                    delete data[key];
                                    manualOverrides[key] = false;
                                    data[key + 'Override'] = false;
                                    lastCommittedValues[key] = 0;
                                    updateDerived();
                                }
                            } else if (isFinal) {
                                editControl.text = formatNumber(lastCommittedValues[key] || 0);
                            }
                        }
                        editControl.onChanging = function () { handleValue(false); };
                        editControl.onChange = function () { handleValue(true); };
                    })(def.key, derivedEdit);
                    break;
                case 'readonly':
                    var readonlyBox = createEditBox(cell, formatNumber(data[def.key]), true);
                    inputs[def.key] = readonlyBox;
                    break;
                case 'label':
                    var labelText = cell.add('statictext', undefined, def.label);
                    try { labelText.graphics.font = ScriptUI.newFont('dialog', 'bold', 11); } catch (eFontBold) {}
                    break;
            }
        }

        function registerDerivedField(key, element) {
            if (!derivedFields[key]) derivedFields[key] = [];
            derivedFields[key].push(element);
        }

        function updateDerived() {
            var d = computeDerived(data);
            for (var key in derivedFields) {
                if (derivedFields.hasOwnProperty(key)) {
                    var val = d[key];
                    var formatted = formatCm(val || 0);
                    var targets = derivedFields[key];
                    for (var i = 0; i < targets.length; i++) {
                        targets[i].text = formatted;
                    }
                }
            }
            for (var key in derivedInputFields) {
                if (derivedInputFields.hasOwnProperty(key)) {
                    if (manualOverrides[key]) continue;
                    var derivedVal = d[key];
                    var formattedInput = formatCm(derivedVal || 0);
                    derivedInputFields[key].text = formattedInput;
                    data[key] = derivedVal;
                    data[key + 'Override'] = false;
                }
            }
        }

        if (inputs.WaistShaping) inputs.WaistShaping.text = formatNumber(data.WaistShaping);
        updateDerived();
        var result = dlg.show();
        if (result === 1) {
            shouldShowMeasurementPalette = summaryCheckbox.value ? true : false;
            assignMeasurementSummaryRows(dialogRows, data);
            return data;
        }
        return null;
    }

})();
