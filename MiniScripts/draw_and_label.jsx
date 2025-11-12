(function () {
    if (app.documents.length === 0) {
        alert('Open a document before running this script.');
        return;
    }

    var DEFAULT_LABEL_PREFIX = 'Waist';
    var DEFAULT_LENGTH = '8';
    var DEFAULT_DIRECTION_INDEX = 3;
    var DEFAULT_UNIT_INDEX = 0;
    var DEFAULT_AUTO_LINE_LENGTH_PT = 144;
    var LABEL_OFFSET_PT = 13;
    var DEFAULT_FONT_SIZE = 12;

    function trim(str) {
        return str.replace(/^\s+|\s+$/g, '');
    }

    try {
        var doc = app.activeDocument;

        var options = showOptionsDialog();
        if (!options) {
            return;
        }

        var canGroupUndo = typeof app.beginUndoGroup === 'function' && typeof app.endUndoGroup === 'function';
        if (canGroupUndo) {
            app.beginUndoGroup('Draw and Label a Line');
        }

        var guidesLayer = ensureLayer(doc, options.labelPrefix || 'Guides');
        var basePath = resolveBasePath(doc, guidesLayer, DEFAULT_AUTO_LINE_LENGTH_PT);

        var baseStroke = extractStroke(basePath, doc.defaultStrokeColor);
        var strokeColor = baseStroke.color;
        var strokeWidth = baseStroke.width;

        var ptScale = options.unit === 'cm' ? 28.3464567 : 72;
        var lengthPts = options.length * ptScale;
        var directionVector = getDirectionVector(options.direction, lengthPts);
        if (!directionVector) {
            throw new Error('Unsupported direction: ' + options.direction);
        }

        var unitSuffix = options.unit === 'cm' ? 'cm' : 'in';
        var lengthDisplay = options.lengthText && options.lengthText.length ? options.lengthText : String(options.length);
        var measurementText = lengthDisplay + unitSuffix;
        var labelText = measurementText;
        if (options.labelPrefix && options.labelPrefix.length) {
            labelText = options.labelPrefix + ' (' + measurementText + ')';
        }

        var points = getPointsToProcess(basePath);
        var createdPaths = [];
        var shouldNameLines = !(options.joinLines && basePath);

        for (var i = 0; i < points.length; i++) {
            var start = points[i];
            var end = [
                start[0] + directionVector.dx,
                start[1] + directionVector.dy
            ];

            var newPath = guidesLayer.pathItems.add();
            newPath.stroked = true;
            newPath.filled = false;
            newPath.strokeWidth = strokeWidth;
            try {
                newPath.strokeColor = strokeColor;
            } catch (eColor) {
                newPath.strokeColor = makeBlack();
            }
            newPath.closed = false;
            newPath.setEntirePath([start, end]);
            if (shouldNameLines) {
                if (options.labelPrefix && options.labelPrefix.length) {
                    newPath.name = options.labelPrefix + ' Line';
                } else {
                    newPath.name = 'Endpoint Line (' + options.directionLabel + ')';
                }
            }

            placeLabel(guidesLayer, start, end, options.direction, labelText, options.labelPrefix, options.fontSize);
            createdPaths.push(newPath);
        }

        if (options.joinLines) {
            basePath = joinCreatedPaths(doc, basePath, createdPaths);
        }

        if (canGroupUndo) {
            app.endUndoGroup();
        }
    } catch (err) {
        try {
            if (canGroupUndo) {
                app.endUndoGroup();
            }
        } catch (e) {}
        alert('Sorry, something went wrong: ' + err.message);
    }

    function showOptionsDialog() {
        var DIRECTION_OPTIONS = [
            { label: 'Bottom', key: 'top' },
            { label: 'Top', key: 'bottom' },
            { label: 'Left', key: 'left' },
            { label: 'Right', key: 'right' }
        ];

        var UNIT_OPTIONS = [
            { label: 'Inches', key: 'in' },
            { label: 'Centimeters', key: 'cm' }
        ];

        var dialog = new Window('dialog', 'Draw and Label a Line');
        dialog.orientation = 'column';
        dialog.alignChildren = ['fill', 'top'];
        dialog.margins = 16;
        dialog.spacing = 12;

        var directionGroup = dialog.add('group');
        directionGroup.add('statictext', undefined, 'Direction');
        var directionDropdown = directionGroup.add('dropdownlist', undefined, []);
        for (var d = 0; d < DIRECTION_OPTIONS.length; d++) {
            var item = directionDropdown.add('item', DIRECTION_OPTIONS[d].label);
            item.data = DIRECTION_OPTIONS[d];
        }
        directionDropdown.selection = directionDropdown.items[DEFAULT_DIRECTION_INDEX];

        var unitGroup = dialog.add('group');
        unitGroup.add('statictext', undefined, 'Units');
        var unitDropdown = unitGroup.add('dropdownlist', undefined, []);
        for (var u = 0; u < UNIT_OPTIONS.length; u++) {
            var unitItem = unitDropdown.add('item', UNIT_OPTIONS[u].label);
            unitItem.data = UNIT_OPTIONS[u];
        }
        unitDropdown.selection = unitDropdown.items[DEFAULT_UNIT_INDEX];

        var lengthGroup = dialog.add('group');
        lengthGroup.add('statictext', undefined, 'Length');
        var lengthInput = lengthGroup.add('edittext', undefined, DEFAULT_LENGTH);
        lengthInput.characters = 6;
        lengthInput.active = true;

        var labelGroup = dialog.add('group');
        labelGroup.add('statictext', undefined, 'Label text');
        var labelInput = labelGroup.add('edittext', undefined, DEFAULT_LABEL_PREFIX);
        labelInput.characters = 12;

        var fontSizeGroup = dialog.add('group');
        fontSizeGroup.add('statictext', undefined, 'Font size');
        var fontSizeInput = fontSizeGroup.add('edittext', undefined, String(DEFAULT_FONT_SIZE));
        fontSizeInput.characters = 6;

        var joinGroup = dialog.add('group');
        var joinCheckbox = joinGroup.add('checkbox', undefined, 'Join guides to base line');
        joinCheckbox.value = false;

        var buttonGroup = dialog.add('group');
        buttonGroup.alignment = ['right', 'bottom'];
        var okButton = buttonGroup.add('button', undefined, 'Draw', { name: 'ok' });
        buttonGroup.add('button', undefined, 'Cancel', { name: 'cancel' });

        var result = null;

        okButton.onClick = function () {
            var lengthText = trim(lengthInput.text || '');
            var value = parseFloat(lengthText);
            if (isNaN(value) || value <= 0) {
                alert('Length must be a positive number.');
                lengthInput.active = true;
                return;
            }

            var fontSizeText = trim(fontSizeInput.text || '');
            var fontSizeValue = parseFloat(fontSizeText);
            if (isNaN(fontSizeValue) || fontSizeValue <= 0) {
                alert('Font size must be a positive number.');
                fontSizeInput.active = true;
                return;
            }

            var directionData = directionDropdown.selection ? directionDropdown.selection.data : DIRECTION_OPTIONS[DEFAULT_DIRECTION_INDEX];
            var unitData = unitDropdown.selection ? unitDropdown.selection.data : UNIT_OPTIONS[DEFAULT_UNIT_INDEX];

            result = {
                direction: directionData.key,
                directionLabel: directionData.label,
                unit: unitData.key,
                length: value,
                lengthText: lengthText,
                labelPrefix: trim(labelInput.text || ''),
                joinLines: joinCheckbox.value === true,
                fontSize: fontSizeValue
            };
            dialog.close(1);
        };

        var status = dialog.show();
        if (status !== 1) {
            return null;
        }
        return result;
    }

    function resolveBasePath(doc, guidesLayer, defaultLengthPts) {
        var selection = doc.selection;
        if (!selection || selection.length === 0) {
            return createDefaultBaseLine(doc, guidesLayer, defaultLengthPts);
        }
        if (selection.length !== 1) {
            throw new Error('Select a single straight line (two anchor points) before running this script.');
        }
        var item = selection[0];
        if (!(item instanceof PathItem)) {
            throw new Error('Selected item must be a PathItem.');
        }
        if (item.pathPoints.length !== 2) {
            throw new Error('Selected path must have exactly two anchor points.');
        }
        return item;
    }

    function createDefaultBaseLine(doc, guidesLayer, lengthPts) {
        var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var rect = artboard.artboardRect;
        var centerX = (rect[0] + rect[2]) / 2;
        var centerY = (rect[1] + rect[3]) / 2;
        var half = lengthPts / 2;

        var path = guidesLayer.pathItems.add();
        path.stroked = true;
        path.strokeWidth = 1;
        var stroke = cloneStrokeColor(doc.defaultStrokeColor);
        try {
            path.strokeColor = stroke;
        } catch (eColor) {
            path.strokeColor = makeBlack();
        }
        path.filled = false;
        path.closed = false;
        path.setEntirePath([
            [centerX - half, centerY],
            [centerX + half, centerY]
        ]);
        path.name = 'Auto Base Line';
        return path;
    }

    function getDirectionVector(direction, lengthPts) {
        switch (direction) {
            case 'top':
                return { dx: 0, dy: -lengthPts };
            case 'bottom':
                return { dx: 0, dy: lengthPts };
            case 'left':
                return { dx: -lengthPts, dy: 0 };
            case 'right':
                return { dx: lengthPts, dy: 0 };
            default:
                return null;
        }
    }

    function placeLabel(layer, start, end, direction, labelText, baseName, fontSize) {
        if (!labelText || !labelText.length) return;
        var isVertical = direction === 'top' || direction === 'bottom';
        var midX = (start[0] + end[0]) / 2;
        var midY = (start[1] + end[1]) / 2;

        var anchor;
        if (isVertical) {
            anchor = [midX + LABEL_OFFSET_PT, midY];
        } else {
            anchor = [midX, midY + LABEL_OFFSET_PT];
        }

        var tf = layer.textFrames.add();
        tf.contents = labelText;
        try {
            tf.textRange.paragraphAttributes.justification = Justification.CENTER;
        } catch (eJust) {}
        var sizeToUse = (typeof fontSize === 'number' && !isNaN(fontSize) && fontSize > 0) ? fontSize : DEFAULT_FONT_SIZE;
        try {
            tf.textRange.characterAttributes.size = sizeToUse;
        } catch (eSize) {}
        try {
            tf.textRange.characterAttributes.fillColor = makeBlack();
        } catch (eFill) {}

        centerTextFrame(tf, anchor);

        if (isVertical) {
            try {
                tf.rotate(-90);
            } catch (eRot) {}
            centerTextFrame(tf, anchor);
        }

        try {
            if (baseName && baseName.length) {
                tf.name = baseName + ' Label';
            } else {
                var labelDirection = direction.charAt(0).toUpperCase() + direction.slice(1);
                tf.name = 'Endpoint Label (' + labelDirection + ')';
            }
        } catch (eName) {}
    }

    function getPointsToProcess(path) {
        var pts = path.pathPoints;
        var selected = [];
        for (var i = 0; i < pts.length; i++) {
            if (isAnchorSelected(pts[i])) {
                selected.push(toArray(pts[i].anchor));
            }
        }
        if (selected.length > 0) {
            return selected;
        }
        var all = [];
        for (var j = 0; j < pts.length; j++) {
            all.push(toArray(pts[j].anchor));
        }
        return all;
    }

    function isAnchorSelected(point) {
        try {
            var state = point.selected;
            return state === PathPointSelection.ANCHORPOINT || state === PathPointSelection.PATHPOINT;
        } catch (e) {
            return false;
        }
    }

    function centerTextFrame(tf, anchor) {
        if (!tf || !anchor) return;
        try {
            var bounds = tf.visibleBounds; // [left, top, right, bottom]
            var width = bounds[2] - bounds[0];
            var height = bounds[1] - bounds[3];
            var left = anchor[0] - width / 2;
            var top = anchor[1] + height / 2;
            tf.position = [left, top];
        } catch (eBounds) {
            try {
                tf.position = [anchor[0], anchor[1]];
            } catch (ePos) {}
        }
    }

    function ensureLayer(doc, name) {
        var target = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                target = doc.layers[i];
                break;
            }
        }
        if (!target) {
            target = doc.layers.add();
            target.name = name;
        }
        try {
            target.locked = false;
        } catch (eLock) {}
        try {
            target.visible = true;
        } catch (eVis) {}
        return target;
    }

    function extractStroke(path, fallbackColor) {
        var width = 1;
        var color = makeBlack();

        if (path && path.stroked) {
            width = path.strokeWidth || 1;
            color = cloneStrokeColor(path.strokeColor);
        } else if (fallbackColor) {
            color = cloneStrokeColor(fallbackColor);
        }

        return {
            color: color,
            width: width
        };
    }

    function makeBlack() {
        var color = new RGBColor();
        color.red = 0;
        color.green = 0;
        color.blue = 0;
        return color;
    }

    function cloneStrokeColor(color) {
        if (!color || color.typename === 'NoColor') {
            return makeBlack();
        }
        try {
            switch (color.typename) {
                case 'RGBColor':
                    var rgb = new RGBColor();
                    rgb.red = color.red;
                    rgb.green = color.green;
                    rgb.blue = color.blue;
                    return rgb;
                case 'CMYKColor':
                    var cmyk = new CMYKColor();
                    cmyk.cyan = color.cyan;
                    cmyk.magenta = color.magenta;
                    cmyk.yellow = color.yellow;
                    cmyk.black = color.black;
                    return cmyk;
                case 'GrayColor':
                    var gray = new GrayColor();
                    gray.gray = color.gray;
                    return gray;
                default:
                    return color;
            }
        } catch (eClone) {
            return makeBlack();
        }
    }

    function toArray(anchor) {
        return [anchor[0], anchor[1]];
    }

    function joinCreatedPaths(doc, basePath, newPaths) {
        if (!newPaths || !newPaths.length) return basePath;
        var current = basePath;
        for (var i = 0; i < newPaths.length; i++) {
            var path = newPaths[i];
            if (!path) continue;
            try {
                doc.selection = null;
                current.selected = true;
                path.selected = true;
                app.executeMenuCommand('join');
                if (doc.selection && doc.selection.length === 1 && doc.selection[0] instanceof PathItem) {
                    current = doc.selection[0];
                }
            } catch (eJoin) {}
        }
        try {
            doc.selection = null;
        } catch (eSel) {}
        return current;
    }
})();

