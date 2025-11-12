var CENTER_LINE_DASH = [25, 12];

function main() {
  if (!app.documents.length) {
    alert("Open a document and select the dart legs.");
    return;
  }

  var doc = app.activeDocument;
  var sel = doc.selection;
  if (!sel || sel.length !== 2) {
    alert("Select exactly two open paths (fixed leg first, rotating leg second).");
    return;
  }

  var legA = sel[0];
  var legB = sel[1];
  if (!isTwoPointOpenPath(legA) || !isTwoPointOpenPath(legB)) {
    alert("Both selections must be open paths with exactly two anchor points.");
    return;
  }

  var options = promptForDartOptions();
  if (!options) return;

  var tolerance = 0.01;
  var apexInfo = findCommonApex(legA, legB, tolerance);
  if (!apexInfo) {
    alert("Selected paths do not share a common apex.");
    return;
  }

  var roleInfo = assignDartLegRoles(
    doc,
    legA,
    legB,
    apexInfo,
    preferenceFromFoldDirection(apexInfo, options.direction)
  );
  if (!roleInfo) {
    alert("Unable to determine which dart leg should stay fixed. Check that one leg is closer to an existing seam.");
    return;
  }

  var fixedLeg = roleInfo.fixedLeg;
  var rotatingLeg = roleInfo.rotatingLeg;
  var apex = roleInfo.apex.slice();
  var fixedOpen = roleInfo.fixedOpen.slice();
  var rotatingOpen = roleInfo.rotatingOpen.slice();

  var seamInfo = roleInfo.seamInfo || findNearestSeam(doc, fixedLeg, rotatingLeg, fixedOpen);
  var seamPoint = seamInfo ? seamInfo.point.slice() : fixedOpen.slice();
  var seamSegmentStart = seamInfo ? seamInfo.segmentStart.slice() : fixedOpen.slice();
  var seamSegmentEnd = seamInfo ? seamInfo.segmentEnd.slice() : apex.slice();

  var targetLayer = fixedLeg.layer;
  if (targetLayer.locked) {
    alert("Unlock the layer containing the dart before running the script.");
    return;
  }

  var strokeColor = fixedLeg.stroked ? fixedLeg.strokeColor : doc.defaultStrokeColor;
  var strokeWidth = fixedLeg.strokeWidth || 1;

  var baseline = drawLine(targetLayer, seamPoint, rotatingOpen, {
    strokeColor: strokeColor,
    strokeWidth: strokeWidth,
  });
  var baselineMid = midpoint(seamPoint, rotatingOpen);
  var triangle = drawTriangle(targetLayer, apex, baselineMid, rotatingOpen, strokeColor, strokeWidth);

  var centerLine = null;
  if (options.includeCenterLine) {
    centerLine = drawLine(targetLayer, apex, baselineMid, {
      strokeColor: strokeColor,
      strokeWidth: strokeWidth,
      strokeDashes: CENTER_LINE_DASH.slice(0),
    });
  }
  var foldNote = annotateFoldDirection(triangle, centerLine, options.direction);

  var rotation = angleBetween(apex, rotatingOpen, fixedOpen);
  rotatePathAround(rotatingLeg, apex, rotation);
  rotatePathAround(triangle, apex, rotation);
  if (centerLine) rotatePathAround(centerLine, apex, rotation);

  snapRotatedLegToSeam(rotatingLeg, apex, seamSegmentStart, seamSegmentEnd);

  var centerTarget = adjustTriangleCenterEdge(triangle, apex, seamSegmentStart, seamSegmentEnd);
  if (centerLine) updateCenterLineEndpoint(centerLine, apex, centerTarget);

  rotatePathAround(rotatingLeg, apex, -rotation);
  rotatePathAround(triangle, apex, -rotation);
  if (centerLine) rotatePathAround(centerLine, apex, -rotation);

  addHelperAnchorToTriangle(triangle, fixedOpen);

  if (!options.makeSinglePath) {
    convertTriangleToLines(triangle, targetLayer, strokeColor, strokeWidth, foldNote);
  }

  safeRemovePath(fixedLeg);
  safeRemovePath(rotatingLeg);
  safeRemovePath(baseline);
}

function promptForDartOptions() {
  var defaults = promptForDartOptions._last || {
    direction: "left",
    includeCenterLine: true,
    makeSinglePath: false,
  };

  var dialogResult = showDartOptionsDialog(defaults);
  var result = dialogResult === undefined ? promptForDartOptionsFallback(defaults) : dialogResult;
  if (!result) return null;
  promptForDartOptions._last = result;
  return result;
}

function showDartOptionsDialog(defaults) {
  if (
    typeof Window === "undefined" ||
    (typeof Window !== "function" && typeof Window !== "object")
  ) {
    return undefined;
  }

  var dlg = new Window("dialog", "True Dart");
  dlg.orientation = "column";
  dlg.alignChildren = "fill";
  dlg.margins = 18;

  var dirPanel = dlg.add("panel", undefined, "Fold direction");
  dirPanel.orientation = "column";
  dirPanel.alignChildren = "left";
  dirPanel.margins = 12;

  var lastDir = (defaults.direction || "left").toLowerCase();
  var optUp = dirPanel.add("radiobutton", undefined, "Up");
  var optDown = dirPanel.add("radiobutton", undefined, "Down");
  var optLeft = dirPanel.add("radiobutton", undefined, "Left");
  var optRight = dirPanel.add("radiobutton", undefined, "Right");

  switch (lastDir) {
    case "down":
      optDown.value = true;
      break;
    case "left":
      optLeft.value = true;
      break;
    case "right":
      optRight.value = true;
      break;
    default:
      optUp.value = true;
      break;
  }

  var optionsPanel = dlg.add("group");
  optionsPanel.orientation = "column";
  optionsPanel.alignChildren = "left";
  optionsPanel.margins = 0;
  optionsPanel.spacing = 6;

  var centerCheck = optionsPanel.add("checkbox", undefined, "Add middle (dashed) line");
  centerCheck.value = defaults.includeCenterLine !== false;

  var singlePathCheck = optionsPanel.add("checkbox", undefined, "Make dart a single path");
  singlePathCheck.value = !!defaults.makeSinglePath;

  var buttons = dlg.add("group");
  buttons.alignment = "right";
  buttons.add("button", undefined, "OK", { name: "ok" });
  buttons.add("button", undefined, "Cancel", { name: "cancel" });

  var result = dlg.show();
  if (result !== 1) return null;

  var direction = "up";
  if (optDown.value) direction = "down";
  else if (optLeft.value) direction = "left";
  else if (optRight.value) direction = "right";

  return {
    direction: direction,
    includeCenterLine: !!centerCheck.value,
    makeSinglePath: !!singlePathCheck.value,
  };
}

function promptForDartOptionsFallback(defaults) {
  var directionInput = textPrompt(
    "Dart fold direction? (Up/Down/Left/Right — enter U/D/L/R)",
    (defaults.direction || "L").charAt(0).toUpperCase()
  );
  if (directionInput === null) return null;
  var direction = normalizeFoldDirection(directionInput, defaults.direction || "left");

  var centerInput = textPrompt(
    "Add middle (dashed) line? (Y/N)",
    defaults.includeCenterLine !== false ? "Y" : "N"
  );
  if (centerInput === null) return null;
  var includeCenterLine = parseYesNo(centerInput, defaults.includeCenterLine !== false);

  var singleInput = textPrompt(
    "Make dart a single path? (Y/N)",
    defaults.makeSinglePath ? "Y" : "N"
  );
  if (singleInput === null) return null;
  var makeSinglePath = parseYesNo(singleInput, !!defaults.makeSinglePath);

  return {
    direction: direction,
    includeCenterLine: includeCenterLine,
    makeSinglePath: makeSinglePath,
  };
}

function textPrompt(message, defaultValue) {
  if (typeof prompt === "function") {
    return prompt(message, defaultValue);
  }
  if (typeof $.global !== "undefined" && typeof $.global.prompt === "function") {
    return $.global.prompt(message, defaultValue);
  }
  return null;
}

function parseYesNo(value, defaultValue) {
  var key = stringifyValue(value);
  if (!key) return defaultValue ? true : false;
  if (key === "y" || key === "yes") return true;
  if (key === "n" || key === "no") return false;
  return defaultValue ? true : false;
}

function preferenceFromFoldDirection(apexInfo, foldDirection) {
  if (!apexInfo) return "auto";
  var fold = normalizeFoldDirection(foldDirection, "up");
  var openA = apexInfo.fixedOpen;
  var openB = apexInfo.rotatingOpen;
  if (!openA || !openB) return "auto";
  var tolerance = 0.5;

  if (fold === "up" || fold === "down") {
    var aY = openA[1];
    var bY = openB[1];
    if (Math.abs(aY - bY) <= tolerance) return "auto";
    return fold === "up" ? (aY <= bY ? "first" : "second") : aY >= bY ? "first" : "second";
  }

  if (fold === "left" || fold === "right") {
    var aX = openA[0];
    var bX = openB[0];
    if (Math.abs(aX - bX) <= tolerance) return "auto";
    return fold === "left" ? (aX <= bX ? "first" : "second") : aX >= bX ? "first" : "second";
  }

  return "auto";
}

function assignDartLegRoles(doc, legA, legB, apexInfo, preference) {
  if (!apexInfo) return null;

  var candidateA = {
    leg: legA,
    open: apexInfo.fixedOpen.slice(),
  };
  var candidateB = {
    leg: legB,
    open: apexInfo.rotatingOpen.slice(),
  };

  var seamA = findNearestSeam(doc, candidateA.leg, candidateB.leg, candidateA.open);
  var seamB = findNearestSeam(doc, candidateB.leg, candidateA.leg, candidateB.open);

  var fixedCandidate = candidateA;
  var rotatingCandidate = candidateB;
  var seamInfo = seamA;

  if (preference === "first") {
    // keep defaults
  } else if (preference === "second") {
    fixedCandidate = candidateB;
    rotatingCandidate = candidateA;
    seamInfo = seamB;
  } else {
    if (shouldSwapFixedCandidate(seamA, seamB)) {
      fixedCandidate = candidateB;
      rotatingCandidate = candidateA;
      seamInfo = seamB;
    }
  }

  return {
    apex: apexInfo.apex.slice(),
    fixedLeg: fixedCandidate.leg,
    rotatingLeg: rotatingCandidate.leg,
    fixedOpen: fixedCandidate.open.slice(),
    rotatingOpen: rotatingCandidate.open.slice(),
    seamInfo: seamInfo,
  };
}

function shouldSwapFixedCandidate(seamA, seamB) {
  if (seamA && seamB) {
    var diff = seamA.distance - seamB.distance;
    if (Math.abs(diff) <= 0.01) return false;
    return seamB.distance < seamA.distance;
  }
  if (!seamA && seamB) return true;
  return false;
}

function drawLine(layer, start, end, options) {
  var line = layer.pathItems.add();
  line.setEntirePath([start, end]);
  line.stroked = true;
  line.filled = false;
  line.strokeWidth = (options && options.strokeWidth) || 1;
  if (options && options.strokeColor) {
    line.strokeColor = options.strokeColor;
  }
  if (options && options.strokeDashes) {
    line.strokeDashes = options.strokeDashes;
    line.strokeDashOffset = 0;
  }
  return line;
}

function drawTriangle(layer, apex, mid, rotatingEnd, strokeColor, strokeWidth) {
  var tri = layer.pathItems.add();
  tri.setEntirePath([apex, mid, rotatingEnd, apex]);
  tri.closed = true;
  tri.filled = false;
  tri.stroked = true;
  tri.strokeWidth = strokeWidth;
  if (strokeColor) tri.strokeColor = strokeColor;
  return tri;
}

function angleBetween(origin, fromPoint, toPoint) {
  var v1 = [fromPoint[0] - origin[0], fromPoint[1] - origin[1]];
  var v2 = [toPoint[0] - origin[0], toPoint[1] - origin[1]];
  var a1 = Math.atan2(v1[1], v1[0]);
  var a2 = Math.atan2(v2[1], v2[0]);
  var angle = a2 - a1;
  while (angle > Math.PI) angle -= 2 * Math.PI;
  while (angle < -Math.PI) angle += 2 * Math.PI;
  return angle;
}

function rotatePathAround(pathItem, origin, radians) {
  var cosA = Math.cos(radians);
  var sinA = Math.sin(radians);
  var pts = pathItem.pathPoints;
  for (var i = 0; i < pts.length; i++) {
    rotatePointInPlace(pts[i], origin, cosA, sinA);
  }
}

function rotatePointInPlace(point, origin, cosA, sinA) {
  var anchor = rotatePoint(point.anchor, origin, cosA, sinA);
  point.anchor = anchor;
  point.leftDirection = rotatePoint(point.leftDirection, origin, cosA, sinA);
  point.rightDirection = rotatePoint(point.rightDirection, origin, cosA, sinA);
}

function rotatePoint(point, origin, cosA, sinA) {
  var x = point[0] - origin[0];
  var y = point[1] - origin[1];
  var rx = x * cosA - y * sinA;
  var ry = x * sinA + y * cosA;
  return [origin[0] + rx, origin[1] + ry];
}

function snapRotatedLegToSeam(rotatedLeg, apex, seamStart, seamEnd) {
  if (!rotatedLeg || rotatedLeg.pathPoints.length < 2) return;
  var p0 = rotatedLeg.pathPoints[0];
  var p1 = rotatedLeg.pathPoints[1];
  var movingPoint = pointsClose(p0.anchor, apex, 0.01) ? p1 : p0;
  var direction = [
    movingPoint.anchor[0] - apex[0],
    movingPoint.anchor[1] - apex[1],
  ];
  var seamDir = [seamEnd[0] - seamStart[0], seamEnd[1] - seamStart[1]];
  var intersection = intersectLines(apex, direction, seamStart, seamDir);
  if (!intersection) {
    intersection = projectPointToSegment(movingPoint.anchor, seamStart, seamEnd);
  }
  movingPoint.anchor = intersection.slice();
  movingPoint.leftDirection = intersection.slice();
  movingPoint.rightDirection = intersection.slice();
}

function adjustTriangleCenterEdge(triangle, apex, seamStart, seamEnd) {
  if (!triangle || triangle.pathPoints.length < 3) return null;
  var centerPoint = triangle.pathPoints[1];
  var direction = [
    centerPoint.anchor[0] - apex[0],
    centerPoint.anchor[1] - apex[1],
  ];
  var seamDir = [seamEnd[0] - seamStart[0], seamEnd[1] - seamStart[1]];
  var intersection = intersectLines(apex, direction, seamStart, seamDir);
  if (!intersection) {
    intersection = projectPointToSegment(centerPoint.anchor, seamStart, seamEnd);
  }
  centerPoint.anchor = intersection.slice();
  centerPoint.leftDirection = intersection.slice();
  centerPoint.rightDirection = intersection.slice();
  return intersection.slice();
}

function updateCenterLineEndpoint(centerLine, apex, targetPoint) {
  if (!centerLine || !targetPoint) return;
  var pts = centerLine.pathPoints;
  if (!pts || pts.length < 2) return;
  var p0 = pts[0];
  var p1 = pts[1];
  var tol = 0.01;
  var apexIsP0 = pointsClose(p0.anchor, apex, tol);
  var apexIsP1 = pointsClose(p1.anchor, apex, tol);
  var movingPoint;
  if (apexIsP0 && !apexIsP1) {
    movingPoint = p1;
  } else if (apexIsP1 && !apexIsP0) {
    movingPoint = p0;
  } else {
    var dist0 = distanceBetween(p0.anchor, apex);
    var dist1 = distanceBetween(p1.anchor, apex);
    movingPoint = dist0 > dist1 ? p0 : p1;
  }
  movingPoint.anchor = targetPoint.slice();
  movingPoint.leftDirection = targetPoint.slice();
  movingPoint.rightDirection = targetPoint.slice();
}

function addHelperAnchorToTriangle(triangle, targetPoint) {
  if (!triangle || triangle.pathPoints.length < 3) return;
  var apexAnchor = triangle.pathPoints[0].anchor.slice();
  var midAnchor = triangle.pathPoints[1].anchor.slice();
  var rotatingAnchor = triangle.pathPoints[2].anchor.slice();

  var helperAnchor = [
    apexAnchor[0] + (midAnchor[0] - apexAnchor[0]) * 0.5,
    apexAnchor[1] + (midAnchor[1] - apexAnchor[1]) * 0.5,
  ];

  triangle.setEntirePath([apexAnchor, helperAnchor, midAnchor, rotatingAnchor, apexAnchor]);
  triangle.closed = true;

  var helperPoint = triangle.pathPoints[1];
  helperPoint.anchor = targetPoint.slice();
  helperPoint.leftDirection = targetPoint.slice();
  helperPoint.rightDirection = targetPoint.slice();
}

function annotateFoldDirection(triangle, centerLine, foldDirection) {
  var clean = (foldDirection || "up").toString().toLowerCase();
  var readable = foldDirectionLabel(clean);
  var note = "Dart fold direction: " + readable;
  try {
    if (triangle) triangle.note = note;
  } catch (err) {}
  try {
    if (centerLine) centerLine.note = note;
  } catch (err) {}
  return note;
}

function foldDirectionLabel(value) {
  switch (value) {
    case "down":
      return "Down";
    case "left":
      return "Left";
    case "right":
      return "Right";
    case "up":
    default:
      return "Up";
  }
}

function convertTriangleToLines(triangle, layer, strokeColor, strokeWidth, note) {
  if (!triangle || !triangle.pathPoints || triangle.pathPoints.length < 3) return;
  var pts = triangle.pathPoints;
  var anchors = [];
  for (var i = 0; i < pts.length; i++) {
    var anchor = pts[i].anchor.slice();
    if (i === 0 || !pointsClose(anchor, anchors[anchors.length - 1], 0.01)) {
      anchors.push(anchor);
    }
  }
  if (anchors.length >= 2 && pointsClose(anchors[0], anchors[anchors.length - 1], 0.01)) {
    anchors.pop();
  }
  safeRemovePath(triangle);

  for (var j = 0; j < anchors.length; j++) {
    var start = anchors[j];
    var end = anchors[(j + 1) % anchors.length];
    var line = drawLine(layer, start, end, {
      strokeColor: strokeColor,
      strokeWidth: strokeWidth,
    });
    if (note) {
      try {
        line.note = note;
      } catch (err) {}
    }
  }
}

function safeRemovePath(item) {
  if (!item || !item.remove) return;
  try {
    item.remove();
  } catch (err) {}
}

function findNearestSeam(doc, fixedLeg, rotatingLeg, referencePoint) {
  var best = null;
  var bestScore = Number.MAX_VALUE;
  var items = doc.pathItems;

  for (var i = 0; i < items.length; i++) {
    var path = items[i];
    if (
      path === fixedLeg ||
      path === rotatingLeg ||
      path.guides ||
      path.clipping ||
      path.hidden ||
      path.locked ||
      !path.pathPoints ||
      path.pathPoints.length < 2
    ) {
      continue;
    }

    var projection = projectPointOntoPath(referencePoint, path);
    if (!projection) continue;

    var score = projection.distance;
    if (path.layer !== fixedLeg.layer) score *= 1.1;

    if (score < bestScore) {
      bestScore = score;
      best = {
        path: path,
        point: projection.point.slice(),
        segmentStart: projection.segmentStart.slice(),
        segmentEnd: projection.segmentEnd.slice(),
        distance: projection.distance,
        score: score,
      };
    }
  }

  return best;
}

function intersectLines(p, r, q, s) {
  var rxs = r[0] * s[1] - r[1] * s[0];
  if (Math.abs(rxs) < 1e-6) return null;
  var qp = [q[0] - p[0], q[1] - p[1]];
  var t = (qp[0] * s[1] - qp[1] * s[0]) / rxs;
  return [p[0] + t * r[0], p[1] + t * r[1]];
}

function projectPointOntoPath(point, path) {
  var pts = path.pathPoints;
  if (!pts || pts.length < 2) return null;
  var limit = path.closed ? pts.length : pts.length - 1;
  var best = null;
  var bestDistance = Number.MAX_VALUE;

  for (var i = 0; i < limit; i++) {
    var start = pts[i].anchor.slice();
    var end = pts[(i + 1) % pts.length].anchor.slice();
    var projected = projectPointToSegment(point, start, end);
    var distance = distanceBetween(point, projected);
    if (distance < bestDistance) {
      bestDistance = distance;
      best = {
        point: projected,
        distance: distance,
        segmentStart: start,
        segmentEnd: end,
      };
    }
  }

  return best;
}

function projectPointToSegment(point, a, b) {
  var abx = b[0] - a[0];
  var aby = b[1] - a[1];
  var lengthSq = abx * abx + aby * aby;
  if (lengthSq === 0) return [a[0], a[1]];
  var t = ((point[0] - a[0]) * abx + (point[1] - a[1]) * aby) / lengthSq;
  if (t < 0) t = 0;
  if (t > 1) t = 1;
  return [a[0] + abx * t, a[1] + aby * t];
}

function midpoint(a, b) {
  return [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2];
}

function distanceBetween(a, b) {
  var dx = a[0] - b[0];
  var dy = a[1] - b[1];
  return Math.sqrt(dx * dx + dy * dy);
}

function pointsClose(a, b, tolerance) {
  return distanceBetween(a, b) <= tolerance;
}

function isTwoPointOpenPath(item) {
  return item && item.typename === "PathItem" && !item.closed && item.pathPoints.length === 2;
}

function findCommonApex(fixedLeg, rotatingLeg, tolerance) {
  var fixedPts = [
    fixedLeg.pathPoints[0].anchor.slice(),
    fixedLeg.pathPoints[1].anchor.slice(),
  ];
  var rotatingPts = [
    rotatingLeg.pathPoints[0].anchor.slice(),
    rotatingLeg.pathPoints[1].anchor.slice(),
  ];

  for (var i = 0; i < fixedPts.length; i++) {
    for (var j = 0; j < rotatingPts.length; j++) {
      if (pointsClose(fixedPts[i], rotatingPts[j], tolerance)) {
        return {
          apex: fixedPts[i],
          fixedOpen: fixedPts[1 - i],
          rotatingOpen: rotatingPts[1 - j],
        };
      }
    }
  }
  return null;
}

function normalizeFoldDirection(value, defaultValue) {
  if (!value) return defaultValue || "up";
  var key = stringifyValue(value);
  if (!key) return defaultValue || "up";
  if (key === "up" || key === "u") return "up";
  if (key === "down" || key === "d") return "down";
  if (key === "left" || key === "l") return "left";
  if (key === "right" || key === "r") return "right";
  return defaultValue || "up";
}

function stringifyValue(value) {
  if (value === undefined || value === null) return "";
  if (typeof value === "string") return trimLower(value);
  try {
    return trimLower(value.toString());
  } catch (err) {
    try {
      return trimLower(String(value));
    } catch (err2) {
      return "";
    }
  }
}

function trimLower(str) {
  if (str === undefined || str === null) return "";
  var asString = "" + str;
  var trimmed = asString.replace(/^[\s\u00a0]+|[\s\u00a0]+$/g, "");
  return trimmed.toLowerCase();
}

main();
