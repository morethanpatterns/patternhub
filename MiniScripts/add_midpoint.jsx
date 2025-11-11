#target illustrator

var BEZIER_TOLERANCE = 0.25;
var COLOR_CLONERS = {
    'RGBColor': function (color) {
        var rgb = new RGBColor();
        rgb.red = color.red;
        rgb.green = color.green;
        rgb.blue = color.blue;
        return rgb;
    },
    'CMYKColor': function (color) {
        var cmyk = new CMYKColor();
        cmyk.cyan = color.cyan;
        cmyk.magenta = color.magenta;
        cmyk.yellow = color.yellow;
        cmyk.black = color.black;
        return cmyk;
    },
    'GrayColor': function (color) {
        var gray = new GrayColor();
        gray.gray = color.gray;
        return gray;
    }
};

(function () {
    // Usage: select a single path (open or closed) to insert a midpoint at the centre.
    var doc = ensureDocument();
    if (!doc) return;

    var targetPath = findTargetPath(doc);
    if (!targetPath) return;

    var pointCount = targetPath.pathPoints.length;
    if (pointCount < 2) {
        fail('The path must contain at least two points.');
        return;
    }

    var points = extractPoints(targetPath);
    var isClosed = targetPath.closed;
    var segments = buildSegments(points, isClosed);
    if (segments.length === 0) {
        fail('Unable to evaluate the selected path.');
        return;
    }

    var totalLength = 0;
    for (var i = 0; i < segments.length; i++) totalLength += segments[i].length;
    if (totalLength < 1e-6) {
        fail('The path is too short to insert a midpoint.');
        return;
    }

    var targetLength = totalLength / 2;
    var segmentInfo = locateSegment(segments, targetLength);
    if (!segmentInfo) {
        fail('Failed to locate the midpoint on this path.');
        return;
    }

    var insertion = computeInsertion(points, segmentInfo);
    if (!insertion) {
        fail('Unable to insert a midpoint on this path.');
        return;
    }

    var rebuilt = rebuildPath(targetPath, insertion.points, isClosed);
    if (!rebuilt) {
        fail('Failed to rebuild the path after inserting the midpoint.');
        return;
    }

    rebuilt.selected = true;
    report('Midpoint inserted successfully.');
})();

function ensureDocument() {
    if (app.documents.length === 0) {
        fail('Open a document before running this script.');
        return null;
    }
    try {
        return app.activeDocument || null;
    } catch (e) {
        fail('Unable to access the active document.');
        return null;
    }
}

function findTargetPath(doc) {
    var selection = doc.selection;
    var paths = [];
    if (selection && selection.length > 0) {
        collectPaths(selection, paths);
    }
    if (paths.length === 1) return paths[0];
    if (paths.length === 0) {
        fail('Select exactly one editable path with at least two points.');
    } else {
        fail('Select only one path before running this script.');
    }
    return null;
}

function collectPaths(items, output) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (!item) continue;
        switch (item.typename) {
            case 'PathItem':
                if (isEligiblePath(item)) output.push(item);
                break;
            case 'GroupItem':
                collectPaths(item.pageItems, output);
                break;
            case 'CompoundPathItem':
                collectPaths(item.pathItems, output);
                break;
            default:
                break;
        }
    }
}

function isEligiblePath(path) {
    return !path.guides && !path.locked && path.editable && path.pathPoints.length >= 2;
}

function extractPoints(path) {
    var pts = [];
    var ppoints = path.pathPoints;
    for (var i = 0; i < ppoints.length; i++) {
        var pt = ppoints[i];
        pts.push({
            anchor: [pt.anchor[0], pt.anchor[1]],
            left: [pt.leftDirection[0], pt.leftDirection[1]],
            right: [pt.rightDirection[0], pt.rightDirection[1]],
            type: pt.pointType
        });
    }
    return pts;
}

function clonePoints(points) {
    var clones = [];
    for (var i = 0; i < points.length; i++) {
        clones.push({
            anchor: points[i].anchor.slice(),
            left: points[i].left.slice(),
            right: points[i].right.slice(),
            type: points[i].type
        });
    }
    return clones;
}

function buildSegments(points, isClosed) {
    var segments = [];
    var count = points.length;
    for (var i = 0; i < count; i++) {
        var nextIndex = (i + 1);
        if (nextIndex >= count) {
            if (!isClosed) break;
            nextIndex = 0;
        }
        var start = points[i];
        var end = points[nextIndex];
        var isCurve = !pointsEqual(start.anchor, start.right) || !pointsEqual(end.anchor, end.left);
        var length = isCurve ? bezierLength(start.anchor, start.right, end.left, end.anchor) : distance(start.anchor, end.anchor);
        segments.push({
            index: i,
            nextIndex: nextIndex,
            length: length,
            isCurve: isCurve,
            start: {
                anchor: start.anchor.slice(),
                right: start.right.slice()
            },
            end: {
                anchor: end.anchor.slice(),
                left: end.left.slice()
            }
        });
    }
    return segments;
}

function locateSegment(segments, targetLength) {
    var accumulated = 0;
    for (var i = 0; i < segments.length; i++) {
        var seg = segments[i];
        if (accumulated + seg.length >= targetLength) {
            return {
                segment: seg,
                distanceIntoSegment: targetLength - accumulated
            };
        }
        accumulated += seg.length;
    }
    return segments.length ? { segment: segments[segments.length - 1], distanceIntoSegment: segments[segments.length - 1].length } : null;
}

function computeInsertion(points, info) {
    var seg = info.segment;
    var remaining = info.distanceIntoSegment;
    if (seg.length <= 0) return null;
    var updatedPoints = clonePoints(points);
    var t;
    if (seg.isCurve) {
        t = solveTForLength(seg.start.anchor, seg.start.right, seg.end.left, seg.end.anchor, remaining / seg.length);
        if (t === null) t = 0.5;
        var split = bezierSplitFull(seg.start.anchor, seg.start.right, seg.end.left, seg.end.anchor, t);
        var insertIndex = seg.index + 1;
        updatedPoints[seg.index].right = split.left[1].slice();
        updatedPoints[seg.index].type = PointType.SMOOTH;
        updatedPoints[seg.nextIndex].left = split.right[2].slice();
        updatedPoints[seg.nextIndex].type = PointType.SMOOTH;
        var newPoint = {
            anchor: split.left[3].slice(),
            left: split.left[2].slice(),
            right: split.right[1].slice(),
            type: PointType.SMOOTH
        };
        updatedPoints.splice(insertIndex, 0, newPoint);
        return { index: insertIndex, points: updatedPoints };
    } else {
        t = remaining / seg.length;
        var startAnchor = seg.start.anchor;
        var endAnchor = seg.end.anchor;
        var anchor = [
            startAnchor[0] + (endAnchor[0] - startAnchor[0]) * t,
            startAnchor[1] + (endAnchor[1] - startAnchor[1]) * t
        ];
        var newPointLine = {
            anchor: anchor,
            left: anchor.slice(),
            right: anchor.slice(),
            type: PointType.CORNER
        };
        var insertIndexLine = seg.index + 1;
        updatedPoints.splice(insertIndexLine, 0, newPointLine);
        return { index: insertIndexLine, points: updatedPoints };
    }
}

function rebuildPath(original, pointData, isClosed) {
    var parent = original.parent;
    var attributes = captureAttributes(original);
    var anchors = [];
    for (var i = 0; i < pointData.length; i++) {
        anchors.push(pointData[i].anchor);
    }

    var newPath = parent.pathItems.add();
    newPath.stroked = attributes.stroked;
    newPath.strokeWidth = attributes.strokeWidth;
    if (attributes.strokeColor) newPath.strokeColor = attributes.strokeColor;
    newPath.filled = attributes.filled;
    if (attributes.fillColor) newPath.fillColor = attributes.fillColor;
    newPath.closed = isClosed;
    newPath.name = attributes.name;

    newPath.setEntirePath(anchors);

    for (var p = 0; p < pointData.length; p++) {
        var pt = newPath.pathPoints[p];
        pt.leftDirection = pointData[p].left.slice();
        pt.rightDirection = pointData[p].right.slice();
        pt.pointType = pointData[p].type;
    }

    try { original.remove(); } catch (eRemove) {}
    return newPath;
}

function captureAttributes(path) {
    return {
        stroked: path.stroked,
        strokeWidth: path.strokeWidth,
        filled: path.filled,
        strokeColor: path.stroked ? cloneColor(path.strokeColor) : null,
        fillColor: path.filled ? cloneColor(path.fillColor) : null,
        name: path.name || ''
    };
}

function distance(a, b) {
    var dx = a[0] - b[0];
    var dy = a[1] - b[1];
    return Math.sqrt(dx * dx + dy * dy);
}

function pointsEqual(a, b) {
    return Math.abs(a[0] - b[0]) < 0.01 && Math.abs(a[1] - b[1]) < 0.01;
}

function lerp(a, b, t) {
    return [
        a[0] + (b[0] - a[0]) * t,
        a[1] + (b[1] - a[1]) * t
    ];
}

function bezierSplitFull(p0, p1, p2, p3, t) {
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

function bezierLength(p0, p1, p2, p3) {
    return adaptiveBezierLength(p0, p1, p2, p3, BEZIER_TOLERANCE);
}

function adaptiveBezierLength(p0, p1, p2, p3, tolerance) {
    var chord = distance(p0, p3);
    var contNet = distance(p0, p1) + distance(p1, p2) + distance(p2, p3);
    if (Math.abs(contNet - chord) <= tolerance) {
        return (contNet + chord) * 0.5;
    }
    var split = bezierSplitFull(p0, p1, p2, p3, 0.5);
    return (
        adaptiveBezierLength(split.left[0], split.left[1], split.left[2], split.left[3], tolerance) +
        adaptiveBezierLength(split.right[0], split.right[1], split.right[2], split.right[3], tolerance)
    );
}

function solveTForLength(p0, p1, p2, p3, fraction) {
    fraction = Math.max(0, Math.min(1, fraction));
    var target = bezierLength(p0, p1, p2, p3) * fraction;
    var low = 0;
    var high = 1;
    for (var i = 0; i < 25; i++) {
        var mid = (low + high) / 2;
        var len = bezierPartialLength(p0, p1, p2, p3, mid);
        if (Math.abs(len - target) < 0.5) return mid;
        if (len < target) {
            low = mid;
        } else {
            high = mid;
        }
    }
    return (low + high) / 2;
}

function bezierPartialLength(p0, p1, p2, p3, t) {
    var split = bezierSplitFull(p0, p1, p2, p3, t);
    return adaptiveBezierLength(split.left[0], split.left[1], split.left[2], split.left[3], BEZIER_TOLERANCE);
}

function report(message) {
    $.writeln('[AddMidpoint] ' + message);
}

function fail(message) {
    report(message);
    alert(message);
}

function cloneColor(color) {
    if (!color || color.typename === 'NoColor') return null;
    var cloner = COLOR_CLONERS[color.typename];
    if (!cloner) return color;
    try {
        return cloner(color);
    } catch (e) {
        return null;
    }
}
