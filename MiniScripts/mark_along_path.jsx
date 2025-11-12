(function () {
    // Usage: select one open path to place marks along it, then run and follow prompts for distance/options.
    if (app.documents.length === 0) {
        alert('Open a document with a single path selected before running this script.');
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;
    if (!sel || sel.length === 0) {
        alert('Select a single open path before running this script.');
        return;
    }

    var targetPath = null;
    for (var i = 0; i < sel.length; i++) {
        if (sel[i].typename === 'PathItem') {
            if (targetPath) {
                alert('Please select only one path.');
                return;
            }
            targetPath = sel[i];
        }
    }

    if (!targetPath) {
        alert('Selection must contain a PathItem.');
        return;
    }

    if (targetPath.closed) {
        alert('This script works with open paths. Please select an open path.');
        return;
    }

    var MARKER_TYPES = {
        LINE: 'line',
        DOT: 'dot'
    };
    var DEFAULT_DOT_DIAMETER_IN = 0.05;

    function trim(str) {
        return str.replace(/^\s+|\s+$/g, '');
    }

    function cm(val) {
        return val * 28.3464566929;
    }

    function inches(val) {
        return val * 72;
    }

    function makeBlack() {
        var color = new RGBColor();
        color.red = 0;
        color.green = 0;
        color.blue = 0;
        return color;
    }

    function cloneStrokeColor(source) {
        try {
            var color = source.strokeColor;
            if (!color || color.typename === 'NoColor') {
                return makeBlack();
            }
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
        } catch (err) {
            return makeBlack();
        }
    }

    function showOptionsDialog() {
        var dialog = new Window('dialog', 'Mark Along Path');
        dialog.orientation = 'column';
        dialog.alignChildren = ['fill', 'top'];
        dialog.margins = 16;
        dialog.spacing = 12;

        var sideGroup = dialog.add('group');
        sideGroup.add('statictext', undefined, 'Start from');
        var sideDropdown = sideGroup.add('dropdownlist', undefined, []);
        var leftItem = sideDropdown.add('item', 'Left/Top');
        leftItem.side = 'left';
        leftItem.data = { key: 'left', label: 'Left/Top' };
        var rightItem = sideDropdown.add('item', 'Right/Bottom');
        rightItem.side = 'right';
        rightItem.data = { key: 'right', label: 'Right/Bottom' };
        sideDropdown.selection = 0;

        var unitGroup = dialog.add('group');
        unitGroup.add('statictext', undefined, 'Units');
        var unitDropdown = unitGroup.add('dropdownlist', undefined, []);
        unitDropdown.add('item', 'Inches').unit = 'in';
        unitDropdown.add('item', 'Centimeters').unit = 'cm';
        unitDropdown.selection = 0;

        var distanceGroup = dialog.add('group');
        distanceGroup.add('statictext', undefined, 'Distance');
        var distanceInput = distanceGroup.add('edittext', undefined, '1');
        distanceInput.characters = 6;
        distanceInput.active = true;

        var markerGroup = dialog.add('group');
        markerGroup.add('statictext', undefined, 'Marker');
        var markerDropdown = markerGroup.add('dropdownlist', undefined, []);
        markerDropdown.add('item', 'Black Round (filled)').markerType = MARKER_TYPES.DOT;
        markerDropdown.add('item', 'Line marker').markerType = MARKER_TYPES.LINE;
        markerDropdown.selection = 0;

        var sizeGroup = dialog.add('group');
        sizeGroup.visible = false;
        sizeGroup.add('statictext', undefined, 'Dot size (in)');
        var sizeInput = sizeGroup.add('edittext', undefined, DEFAULT_DOT_DIAMETER_IN.toString());
        sizeInput.characters = 6;
        sizeInput.enabled = false;

        var buttonGroup = dialog.add('group');
        buttonGroup.alignment = ['right', 'bottom'];
        var okButton = buttonGroup.add('button', undefined, 'Place Mark', {
            name: 'ok'
        });
        var cancelButton = buttonGroup.add('button', undefined, 'Cancel', {
            name: 'cancel'
        });

        var result = null;

        okButton.onClick = function () {
            var distanceText = trim(distanceInput.text || '');
            if (!distanceText.length) {
                alert('Enter a distance to mark along the path.');
                distanceInput.active = true;
                return;
            }

            var value = parseFloat(distanceText);
            if (isNaN(value) || value < 0) {
                alert('Distance must be a non-negative number.');
                distanceInput.active = true;
                return;
            }

            var sizeText = trim(sizeInput.text || '');
            var sizeValue = null;
            var isDot = markerDropdown.selection ? markerDropdown.selection.markerType === MARKER_TYPES.DOT : true;
            if (isDot) {
                if (!sizeText.length) {
                    alert('Dot size must be provided for round markers.');
                    sizeInput.active = true;
                    return;
                }
                sizeValue = parseFloat(sizeText);
                if (isNaN(sizeValue) || sizeValue <= 0) {
                    alert('Dot size must be a positive number.');
                    sizeInput.active = true;
                    return;
                }
            }

            var selectedSide = sideDropdown.selection ? sideDropdown.selection.data : null;
            var sideKey = selectedSide && selectedSide.key ? selectedSide.key : 'left';
            var sideLabel = selectedSide && selectedSide.label ? selectedSide.label : (sideKey === 'left' ? 'Left/Top' : 'Right/Bottom');

            result = {
                side: sideKey,
                sideLabel: sideLabel,
                unit: unitDropdown.selection ? unitDropdown.selection.unit : 'in',
                distance: value,
                distanceLabel: distanceText,
                markerType: markerDropdown.selection ? markerDropdown.selection.markerType : MARKER_TYPES.DOT,
                dotDiameterIn: isDot ? sizeValue : null
            };
            dialog.close(1);
        };

        cancelButton.onClick = function () {
            dialog.close(0);
        };

        markerDropdown.onChange = function () {
            var isDot = markerDropdown.selection ? markerDropdown.selection.markerType === MARKER_TYPES.DOT : true;
            sizeGroup.visible = isDot;
            sizeInput.enabled = isDot;
        };

        markerDropdown.onChange();

        var status = dialog.show();
        if (status !== 1) return null;
        return result;
    }

    var options = showOptionsDialog();
    if (!options) return;

    var sideInput = options.side;
    var sideLabel = options.sideLabel || (sideInput === 'left' ? 'Left/Top' : 'Right/Bottom');
    var unitInput = options.unit;
    var distanceVal = options.distance;
    var distanceLabel = options.distanceLabel;
    var markerType = options.markerType;
    var dotDiameterIn = options.dotDiameterIn != null ? options.dotDiameterIn : DEFAULT_DOT_DIAMETER_IN;

    var pointsPerUnit = unitInput === 'cm' ? 28.3464566929 : 72;
    var distancePts = distanceVal * pointsPerUnit;

    var pathPoints = targetPath.pathPoints;
    if (pathPoints.length < 2) {
        alert('The path must have at least two points.');
        return;
    }

    function pointFromArray(arr) {
        return {
            x: arr[0],
            y: arr[1]
        };
    }

    function vectorsEqual(a, b) {
        return Math.abs(a[0] - b[0]) < 0.01 && Math.abs(a[1] - b[1]) < 0.01;
    }

    function dist(a, b) {
        var dx = a.x - b.x;
        var dy = a.y - b.y;
        return Math.sqrt(dx * dx + dy * dy);
    }

    function lerp(a, b, t) {
        return {
            x: a.x + (b.x - a.x) * t,
            y: a.y + (b.y - a.y) * t
        };
    }

    function bezierSplit(p0, p1, p2, p3, t) {
        var p01 = lerp(p0, p1, t);
        var p12 = lerp(p1, p2, t);
        var p23 = lerp(p2, p3, t);
        var p012 = lerp(p01, p12, t);
        var p123 = lerp(p12, p23, t);
        var p0123 = lerp(p012, p123, t);
        return {
            left: [p0, p01, p012, p0123],
            right: [p0123, p123, p23, p3]
        };
    }

    function bezierLength(p0, p1, p2, p3, tol, depth) {
        if (tol === undefined) tol = 0.5;
        if (depth === undefined) depth = 0;
        var chord = dist(p0, p3);
        var cont = dist(p0, p1) + dist(p1, p2) + dist(p2, p3);
        if (depth > 10 || Math.abs(cont - chord) <= tol) {
            return (chord + cont) * 0.5;
        }
        var split = bezierSplit(p0, p1, p2, p3, 0.5);
        return bezierLength(split.left[0], split.left[1], split.left[2], split.left[3], tol / 2, depth + 1) +
            bezierLength(split.right[0], split.right[1], split.right[2], split.right[3], tol / 2, depth + 1);
    }

    function bezierPoint(p0, p1, p2, p3, t) {
        var mt = 1 - t;
        var mt2 = mt * mt;
        var t2 = t * t;
        return {
            x: mt2 * mt * p0.x + 3 * mt2 * t * p1.x + 3 * mt * t2 * p2.x + t2 * t * p3.x,
            y: mt2 * mt * p0.y + 3 * mt2 * t * p1.y + 3 * mt * t2 * p2.y + t2 * t * p3.y
        };
    }

    function bezierDerivative(p0, p1, p2, p3, t) {
        var mt = 1 - t;
        return {
            x: 3 * mt * mt * (p1.x - p0.x) + 6 * mt * t * (p2.x - p1.x) + 3 * t * t * (p3.x - p2.x),
            y: 3 * mt * mt * (p1.y - p0.y) + 6 * mt * t * (p2.y - p1.y) + 3 * t * t * (p3.y - p2.y)
        };
    }

    function normalize(vec) {
        var len = Math.sqrt(vec.x * vec.x + vec.y * vec.y);
        if (len === 0) {
            return {
                x: 1,
                y: 0
            };
        }
        return {
            x: vec.x / len,
            y: vec.y / len
        };
    }

    function makeSegment(startPoint, endPoint, outboundHandle, inboundHandle) {
        var startAnchor = pointFromArray(startPoint.anchor);
        var endAnchor = pointFromArray(endPoint.anchor);
        var startControl = outboundHandle ? pointFromArray(outboundHandle) : startAnchor;
        var endControl = inboundHandle ? pointFromArray(inboundHandle) : endAnchor;
        var isLine = vectorsEqual(startPoint.anchor, (outboundHandle || startPoint.anchor)) &&
            vectorsEqual(endPoint.anchor, (inboundHandle || endPoint.anchor));
        var seg = {
            p0: startAnchor,
            p1: startControl,
            p2: endControl,
            p3: endAnchor,
            isLine: isLine
        };
        seg.length = isLine ? dist(seg.p0, seg.p3) : bezierLength(seg.p0, seg.p1, seg.p2, seg.p3);
        return seg;
    }

    function computeSegments(pathPoints, fromStart) {
        var segments = [];
        if (fromStart) {
            for (var s = 0; s < pathPoints.length - 1; s++) {
                segments.push(makeSegment(
                    pathPoints[s],
                    pathPoints[s + 1],
                    pathPoints[s].rightDirection,
                    pathPoints[s + 1].leftDirection
                ));
            }
        } else {
            for (var r = pathPoints.length - 1; r > 0; r--) {
                segments.push(makeSegment(
                    pathPoints[r],
                    pathPoints[r - 1],
                    pathPoints[r].leftDirection,
                    pathPoints[r - 1].rightDirection
                ));
            }
        }
        return segments;
    }

    var firstPoint = pathPoints[0].anchor;
    var lastPoint = pathPoints[pathPoints.length - 1].anchor;
    var traverseForward;

    var deltaX = Math.abs(firstPoint[0] - lastPoint[0]);
    var deltaY = Math.abs(firstPoint[1] - lastPoint[1]);
    var isVertical = deltaY > deltaX;

    if (isVertical) {
        if (sideInput === 'left') { // Left/Top option maps to Top for vertical paths
            traverseForward = firstPoint[1] >= lastPoint[1];
        } else { // Right/Bottom option maps to Bottom for vertical paths
            traverseForward = firstPoint[1] <= lastPoint[1];
        }
    } else {
        if (sideInput === 'left') {
            traverseForward = firstPoint[0] <= lastPoint[0];
        } else {
            traverseForward = firstPoint[0] >= lastPoint[0];
        }
    }

    var segments = traverseForward ? computeSegments(pathPoints, true) : computeSegments(pathPoints, false);

    if (segments.length === 0) {
        alert('Unable to determine path segments.');
        return;
    }

    var totalLength = 0;
    for (var sl = 0; sl < segments.length; sl++) totalLength += segments[sl].length;

    if (distancePts > totalLength) {
        alert('Distance exceeds path length. The mark will be placed at the end instead.');
        distancePts = totalLength;
    }

    var remaining = distancePts;
    var segment = null;
    for (var segIndex = 0; segIndex < segments.length; segIndex++) {
        var segLen = segments[segIndex].length;
        if (remaining <= segLen) {
            segment = segments[segIndex];
            break;
        }
        remaining -= segLen;
    }

    if (!segment) {
        segment = segments[segments.length - 1];
        remaining = segment.length;
    }

    var t;
    if (segment.isLine) {
        t = segment.length === 0 ? 0 : remaining / segment.length;
    } else {
        var low = 0;
        var high = 1;
        for (var it = 0; it < 25; it++) {
            var mid = (low + high) / 2;
            var split = bezierSplit(segment.p0, segment.p1, segment.p2, segment.p3, mid);
            var len = bezierLength(split.left[0], split.left[1], split.left[2], split.left[3]);
            if (len < remaining) {
                low = mid;
            } else {
                high = mid;
            }
        }
        t = (low + high) / 2;
    }

    var markPoint = segment.isLine ? lerp(segment.p0, segment.p3, t) : bezierPoint(segment.p0, segment.p1, segment.p2, segment.p3, t);
    var tangent;
    if (segment.isLine) {
        tangent = {
            x: segment.p3.x - segment.p0.x,
            y: segment.p3.y - segment.p0.y
        };
    } else {
        tangent = bezierDerivative(segment.p0, segment.p1, segment.p2, segment.p3, t);
    }
    tangent = normalize(tangent);
    var normal = normalize({
        x: -tangent.y,
        y: tangent.x
    });

    function createLineMark(container, centerPoint, normalVec, strokeColor, strokeWidth) {
        var tickLength = 12;
        var half = tickLength / 2;
        var tickStart = {
            x: centerPoint.x - normalVec.x * half,
            y: centerPoint.y - normalVec.y * half
        };
        var tickEnd = {
            x: centerPoint.x + normalVec.x * half,
            y: centerPoint.y + normalVec.y * half
        };

        var item = container.pathItems.add();
        item.stroked = true;
        item.filled = false;
        item.strokeWidth = strokeWidth > 0 ? strokeWidth : 1;
        item.strokeColor = strokeColor;
        item.closed = false;
        item.setEntirePath([
            [tickStart.x, tickStart.y],
            [tickEnd.x, tickEnd.y]
        ]);
        return item;
    }

    function createDotMark(container, centerPoint, diameterIn) {
        var diameter = inches(diameterIn);
        var radius = diameter / 2;
        var top = centerPoint.y + radius;
        var left = centerPoint.x - radius;
        var item = container.pathItems.ellipse(top, left, diameter, diameter);
        item.stroked = true;
        item.strokeWidth = 1;
        item.strokeColor = makeBlack();
        item.filled = true;
        item.fillColor = makeBlack();
        item.closed = true;
        return item;
    }

    var parentContainer = targetPath.parent;
    var strokeColor = cloneStrokeColor(targetPath);
    var strokeWidth = targetPath.strokeWidth || 1;

    var markItem;
    if (markerType === MARKER_TYPES.LINE) {
        markItem = createLineMark(parentContainer, markPoint, normal, strokeColor, strokeWidth);
    } else {
        markItem = createDotMark(parentContainer, markPoint, dotDiameterIn);
    }

    var unitLabel = unitInput === 'cm' ? 'cm' : 'in';
    markItem.name = 'Mark ' + distanceLabel + ' ' + unitLabel + ' from ' + sideLabel;
    markItem.selected = true;

})();
