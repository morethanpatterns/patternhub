var POINTS_PER_INCH = 72;
var CM_PER_INCH = 2.54;
var POINTS_PER_CM = POINTS_PER_INCH / CM_PER_INCH;

function main() {
    if (app.documents.length === 0) {
        alert("Open a document and select a straight path before running the script.");
        return;
    }

    var doc = app.activeDocument;
    if (doc.selection.length !== 1 || !(doc.selection[0] instanceof PathItem)) {
        alert("Select exactly one path.");
        return;
    }

    var basePath = doc.selection[0];
    if (basePath.pathPoints.length !== 2) {
        alert("The path must have exactly two anchor points.");
        return;
    }

    var rawStart = toPoint(basePath.pathPoints[0].anchor);
    var rawEnd = toPoint(basePath.pathPoints[1].anchor);

    var horizontal = Math.abs(rawEnd.x - rawStart.x) >= Math.abs(rawEnd.y - rawStart.y);
    var start, end;
    if (horizontal) {
        if (rawStart.x <= rawEnd.x) {
            start = rawStart;
            end = rawEnd;
        } else {
            start = rawEnd;
            end = rawStart;
        }
    } else {
        if (rawStart.y <= rawEnd.y) {
            start = rawStart;
            end = rawEnd;
        } else {
            start = rawEnd;
            end = rawStart;
        }
    }

    var lineVec = subtract(end, start);
    var lineLength = magnitude(lineVec);

    if (lineLength < 1) {
        alert("The selected path is too short to place a dart.");
        return;
    }

    var defaults = {
        widthInches: 1,
        lengthInches: 3,
        widthCm: 2,
        lengthCm: 10,
        offsetInches: 1,
        offsetCm: 2,
        joinLegs: false,
        addCenterLine: false,
        unit: "in"
    };

    var inputs = showDartDialog(defaults);
    if (!inputs) {
        return;
    }

    if (inputs.widthPts > lineLength) {
        alert("Dart width must be smaller than the selected path length.");
        return;
    }

    var halfWidth = inputs.widthPts / 2;
    var centerDistance;
    if (inputs.positionMode === "center") {
        centerDistance = lineLength / 2;
    } else {
        var offsetPts = inputs.offsetPts;
        var offsetMagnitude = Math.abs(offsetPts);
        if (offsetMagnitude > lineLength) {
            alert("Offset exceeds the selected path length.");
            return;
        }
        if (horizontal) {
            centerDistance = offsetPts >= 0
                ? offsetMagnitude
                : lineLength - offsetMagnitude;
        } else {
            centerDistance = offsetPts >= 0
                ? lineLength - offsetMagnitude
                : offsetMagnitude;
        }
    }

    if (centerDistance < halfWidth || centerDistance > lineLength - halfWidth) {
        alert("Dart width and offset exceed the usable length of the path.");
        return;
    }

    var unitDir = normalize(lineVec);
    var centerPoint = add(start, scale(unitDir, centerDistance));
    var leftPoint = add(centerPoint, scale(unitDir, -halfWidth));
    var rightPoint = add(centerPoint, scale(unitDir, halfWidth));

    var normal = perpendicular(unitDir);
    var oppositeNormal = scale(normal, -1);
    var rightVector;
    var leftVector;

    if (Math.abs(normal.x) > 0.0001) {
        rightVector = normal.x > 0 ? normal : oppositeNormal;
    } else if (Math.abs(normal.y) > 0.0001) {
        rightVector = normal.y > 0 ? normal : oppositeNormal;
    } else {
        rightVector = normal;
    }

    leftVector = rightVector === normal ? oppositeNormal : normal;

    var upVector = leftVector;
    if (rightVector.y < upVector.y) {
        upVector = rightVector;
    }

    var apexVector;
    switch (inputs.direction) {
        case "Up":
            // Illustrator increases Y downward, so invert to make "Up" visually upward.
            apexVector = scale(upVector, -1);
            break;
        case "Down":
            apexVector = upVector;
            break;
        case "Left":
            apexVector = leftVector;
            break;
        case "Right":
            apexVector = rightVector;
            break;
        default:
            apexVector = scale(upVector, -1);
            break;
    }
    var apexPoint = add(centerPoint, scale(apexVector, inputs.lengthPts));

    var dartGroup = drawDart(
        basePath.layer,
        leftPoint,
        rightPoint,
        apexPoint,
        inputs.joinLegs,
        inputs.addCenterLine,
        centerPoint
    );

    doc.selection = [dartGroup];
    app.redraw();
}

function showDartDialog(defaults) {
    var dlg = new Window("dialog", "Create Dart");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];

    var unitGroup = dlg.add("group");
    unitGroup.add("statictext", undefined, "Units");
    var unitDropdown = unitGroup.add("dropdownlist", undefined, ["Inches", "Centimeters"]);
    unitDropdown.selection = defaults.unit === "cm" ? 1 : 0;
    var currentUnitKey = unitDropdown.selection.index === 1 ? "cm" : "in";

    var widthGroup = dlg.add("group");
    var widthLabel = widthGroup.add("statictext", undefined, "");
    widthLabel.minimumSize.width = 140;
    var widthInput = widthGroup.add("edittext", undefined, "");
    widthInput.characters = 6;

    var lengthGroup = dlg.add("group");
    var lengthLabel = lengthGroup.add("statictext", undefined, "");
    lengthLabel.minimumSize.width = 140;
    var lengthInput = lengthGroup.add("edittext", undefined, "");
    lengthInput.characters = 6;

    var positionPanel = dlg.add("panel", undefined, "Position from start point");
    positionPanel.orientation = "column";
    positionPanel.alignChildren = ["left", "top"];

    var radioCenter = positionPanel.add("radiobutton", undefined, "Use center of selected line");
    var radioOffset = positionPanel.add("radiobutton", undefined, "Offset distance");
    radioOffset.preferredSize.width = 200;
    radioCenter.value = true;

    var offsetGroup = positionPanel.add("group");
    var offsetLabel = offsetGroup.add("statictext", undefined, "");
    offsetLabel.minimumSize.width = 100;
    var offsetInput = offsetGroup.add("edittext", undefined, "");
    offsetInput.characters = 6;
    offsetInput.enabled = false;

    var offsetHint = positionPanel.add("statictext", undefined, "Positive offsets measure from left/top; negatives measure from right/bottom.");
    offsetHint.alignment = "left";

    var measurementDefaults = {
        "in": {
            width: defaults.widthInches,
            length: defaults.lengthInches
        },
        "cm": {
            width: defaults.widthCm != null ? defaults.widthCm : pointsToUnit(unitToPoints(defaults.widthInches, "in"), "cm"),
            length: defaults.lengthCm != null ? defaults.lengthCm : pointsToUnit(unitToPoints(defaults.lengthInches, "in"), "cm")
        }
    };

    var offsetPtsDefault = unitToPoints(defaults.offsetInches, "in");
    var offsetDefaults = {
        "in": pointsToUnit(offsetPtsDefault, "in"),
        "cm": defaults.offsetCm != null ? defaults.offsetCm : pointsToUnit(offsetPtsDefault, "cm")
    };

    widthInput.text = formatNumber(measurementDefaults[currentUnitKey].width);
    lengthInput.text = formatNumber(measurementDefaults[currentUnitKey].length);
    offsetInput.text = formatNumber(offsetDefaults[currentUnitKey]);

    unitDropdown.onChange = function () {
        if (!unitDropdown.selection) {
            return;
        }
        setUnit(unitDropdown.selection.index === 1 ? "cm" : "in");
    };

    radioCenter.onClick = toggleOffset;
    radioOffset.onClick = toggleOffset;

    function toggleOffset() {
        offsetInput.enabled = radioOffset.value;
    }

    toggleOffset();

    function setUnit(newUnitKey) {
        if (newUnitKey === currentUnitKey) {
            updateUnitLabels();
            return;
        }

        var previousUnit = currentUnitKey;
        var widthVal = parseFloat(widthInput.text);
        var lengthVal = parseFloat(lengthInput.text);
        var offsetVal = parseFloat(offsetInput.text);

        var widthPts = isNaN(widthVal) ? unitToPoints(measurementDefaults[previousUnit].width, previousUnit) : unitToPoints(widthVal, previousUnit);
        var lengthPts = isNaN(lengthVal) ? unitToPoints(measurementDefaults[previousUnit].length, previousUnit) : unitToPoints(lengthVal, previousUnit);
        var offsetPts = isNaN(offsetVal) ? offsetPtsDefault : unitToPoints(offsetVal, previousUnit);

        currentUnitKey = newUnitKey;

        var defaultWidth = measurementDefaults[currentUnitKey].width;
        var defaultLength = measurementDefaults[currentUnitKey].length;

        var shouldUseDefaults =
            (previousUnit === "in" && newUnitKey === "cm" && !isNaN(widthVal) && !isNaN(lengthVal) &&
                widthVal === measurementDefaults["in"].width && lengthVal === measurementDefaults["in"].length) ||
            (previousUnit === "cm" && newUnitKey === "in" && !isNaN(widthVal) && !isNaN(lengthVal) &&
                widthVal === measurementDefaults["cm"].width && lengthVal === measurementDefaults["cm"].length);

        if (shouldUseDefaults) {
            widthInput.text = formatNumber(defaultWidth);
            lengthInput.text = formatNumber(defaultLength);
        } else {
            widthInput.text = formatNumber(pointsToUnit(widthPts, currentUnitKey));
            lengthInput.text = formatNumber(pointsToUnit(lengthPts, currentUnitKey));
        }

        var defaultOffset = offsetDefaults[currentUnitKey];
        var shouldUseOffsetDefault =
            (previousUnit === "in" && newUnitKey === "cm" && !isNaN(offsetVal) && offsetVal === offsetDefaults["in"]) ||
            (previousUnit === "cm" && newUnitKey === "in" && !isNaN(offsetVal) && offsetVal === offsetDefaults["cm"]);

        if (shouldUseOffsetDefault) {
            offsetInput.text = formatNumber(defaultOffset);
        } else {
            offsetInput.text = formatNumber(pointsToUnit(offsetPts, currentUnitKey));
        }

        updateUnitLabels();
    }

    function updateUnitLabels() {
        var unitLabel = currentUnitKey === "cm" ? "cm" : "in";
        widthLabel.text = "Dart width (" + unitLabel + ")";
        lengthLabel.text = "Dart length (" + unitLabel + ")";
        radioOffset.text = "Offset distance (" + unitLabel + ")";
        offsetLabel.text = "Distance (" + unitLabel + ")";
    }

    updateUnitLabels();

    var directionGroup = dlg.add("group");
    directionGroup.add("statictext", undefined, "Direction");
    var directionDropdown = directionGroup.add("dropdownlist", undefined, ["Up", "Down", "Left", "Right"]);
    directionDropdown.selection = 1;

    var joinCheckbox = dlg.add("checkbox", undefined, "Join dart legs into a single path");
    joinCheckbox.value = !!defaults.joinLegs;

    var centerCheckbox = dlg.add("checkbox", undefined, "Add center line");
    centerCheckbox.value = !!defaults.addCenterLine;

    var buttons = dlg.add("group");
    buttons.alignment = "right";
    var okBtn = buttons.add("button", undefined, "OK");
    var cancelBtn = buttons.add("button", undefined, "Cancel");

    var result = null;

    okBtn.onClick = function () {
        var widthVal = parseFloat(widthInput.text);
        var lengthVal = parseFloat(lengthInput.text);
        if (!isPositive(widthVal)) {
            alert("Enter a positive dart width.");
            return;
        }
        if (!isPositive(lengthVal)) {
            alert("Enter a positive dart length.");
            return;
        }

        var widthPts = unitToPoints(widthVal, currentUnitKey);
        var lengthPts = unitToPoints(lengthVal, currentUnitKey);

        var mode = radioCenter.value ? "center" : "offset";
        var offsetPts = 0;
        if (mode === "offset") {
            var offsetVal = parseFloat(offsetInput.text);
            if (isNaN(offsetVal)) {
                alert("Enter a numeric offset distance.");
                return;
            }
            if (offsetVal === 0) {
                alert("Offset distance cannot be zero.");
                return;
            }
            offsetPts = unitToPoints(offsetVal, currentUnitKey);
        }

        var selectedDirection = directionDropdown.selection ? directionDropdown.selection.text : "Up";

        result = {
            widthPts: widthPts,
            lengthPts: lengthPts,
            positionMode: mode,
            offsetPts: offsetPts,
            direction: selectedDirection,
            joinLegs: joinCheckbox.value,
            addCenterLine: centerCheckbox.value,
            unit: currentUnitKey
        };

        dlg.close(1);
    };

    cancelBtn.onClick = function () {
        dlg.close(0);
    };

    var confirmed = dlg.show();
    if (confirmed !== 1) {
        return null;
    }

    return result;
}

function drawDart(layer, leftPoint, rightPoint, apexPoint, joinLegs, addCenterLine, centerPoint) {
    var newGroup = layer.groupItems.add();
    newGroup.name = joinLegs ? "Dart (joined)" : "Dart";

    var col = black();

    if (joinLegs) {
        createPath(newGroup, [leftPoint, apexPoint, rightPoint], col, false);
    } else {
        createPath(newGroup, [leftPoint, apexPoint], col, false);
        createPath(newGroup, [rightPoint, apexPoint], col, false);
    }

    if (addCenterLine) {
        createPath(newGroup, [centerPoint, apexPoint], col, false, [16, 8]);
    }

    return newGroup;
}

function createPath(parent, points, color, closed, dashArray) {
    var path = parent.pathItems.add();
    var coords = [];
    for (var i = 0; i < points.length; i++) {
        coords.push([points[i].x, points[i].y]);
    }
    path.setEntirePath(coords);
    path.stroked = true;
    path.strokeWidth = 1;
    path.strokeColor = color;
    path.filled = false;
    path.closed = !!closed;
    if (dashArray && dashArray.length) {
        path.strokeDashes = dashArray;
        path.strokeDashOffset = 0;
    } else {
        path.strokeDashes = [];
    }
}

function black() {
    var c = new RGBColor();
    c.red = 0;
    c.green = 0;
    c.blue = 0;
    return c;
}

function toPoint(arr) {
    return {
        x: arr[0],
        y: arr[1]
    };
}

function add(p1, p2) {
    return {
        x: p1.x + p2.x,
        y: p1.y + p2.y
    };
}

function subtract(p1, p2) {
    return {
        x: p1.x - p2.x,
        y: p1.y - p2.y
    };
}

function scale(p, value) {
    return {
        x: p.x * value,
        y: p.y * value
    };
}

function magnitude(p) {
    return Math.sqrt(p.x * p.x + p.y * p.y);
}

function normalize(p) {
    var len = magnitude(p);
    if (len === 0) {
        return {
            x: 0,
            y: 0
        };
    }
    return {
        x: p.x / len,
        y: p.y / len
    };
}

function perpendicular(p) {
    return {
        x: -p.y,
        y: p.x
    };
}

function isPositive(value) {
    return !isNaN(value) && value > 0;
}

function formatNumber(value) {
    return value.toFixed(2);
}

function pointsToInches(value) {
    return value / POINTS_PER_INCH;
}

function unitToPoints(value, unitKey) {
    if (unitKey === "cm") {
        return value * POINTS_PER_CM;
    }
    return value * POINTS_PER_INCH;
}

function pointsToUnit(value, unitKey) {
    if (unitKey === "cm") {
        return value / POINTS_PER_CM;
    }
    return value / POINTS_PER_INCH;
}

main();
