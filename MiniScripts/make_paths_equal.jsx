/**
 * Make Paths Equal
 *
 * Duplicates the length of the first selected open line onto the second item
 * without changing direction. You can work in two ways:
 *   • Select two two-point open paths (pick the reference first, target second).
 *   • Select two individual endpoints (anchors) belonging to such paths.
 * In endpoint mode, the target anchor you select is the one that moves; the
 * opposite anchor on that path stays fixed while the selected point slides
 * along the existing direction to match the reference length.
 *
 * Usage:
 *   1. Select two open paths (first = reference, second = target).
 *   2. Run this script. The second path will adopt the first path's length.
 */

var MATCH_TOLERANCE = 0.01;
var LENGTH_EPSILON = 1e-6;

(function () {
  var doc = ensureDocument();
  if (!doc) return;

  var parsed = parseSelection(doc.selection);
  if (!parsed) return;

  var reference = parsed.reference;
  var target = parsed.target;
  var refData = getLineData(reference);
  var targetData = getLineData(target);

  if (refData.length <= LENGTH_EPSILON || targetData.length <= LENGTH_EPSILON) {
    fail("Paths must have a non-zero length.");
    return;
  }

  var lockedIndex = parsed.lockedIndex;
  if (lockedIndex === null) {
    lockedIndex = findLockedAnchor(refData, targetData, MATCH_TOLERANCE);
    if (lockedIndex === 2) {
      report("Both endpoints already match the reference length; no change made.");
      return;
    }
    if (lockedIndex < 0) lockedIndex = 0;
  }

  adjustTargetLength(target, targetData, refData.length, lockedIndex);
  report("Target path updated to match the reference length.");
})();

function isLinearPath(item) {
  return (
    item.typename === "PathItem" &&
    !item.guides &&
    !item.clipping &&
    !item.locked &&
    item.editable &&
    !item.closed &&
    item.pathPoints.length === 2
  );
}

function isPathPoint(item) {
  return item && item.typename === "PathPoint";
}

function getPathPointIndex(path, point) {
  for (var i = 0; i < path.pathPoints.length; i++) {
    if (path.pathPoints[i] === point) {
      return i;
    }
  }
  return -1;
}

function parseSelection(sel) {
  if (!sel || sel.length !== 2) {
    fail("Select exactly two items. Pick the reference first, target second.");
    return null;
  }

  // Respect the user's pick order.
  var referenceInput = sel[0];
  var targetInput = sel[1];

  if (isPathPoint(referenceInput) && isPathPoint(targetInput)) {
    var reference = referenceInput.parent;
    var target = targetInput.parent;

    if (!isLinearPath(reference) || !isLinearPath(target)) {
      fail("Both endpoints must come from simple open paths with two points.");
      return null;
    }

    var movingIndex = getPathPointIndex(target, targetInput);
    if (movingIndex < 0) {
      fail("Unable to resolve the selected target anchor.");
      return null;
    }

    var lockedIndex = movingIndex === 0 ? 1 : 0;
    return {
      reference: reference,
      target: target,
      lockedIndex: lockedIndex,
    };
  }

  if (isLinearPath(referenceInput) && isLinearPath(targetInput)) {
    return {
      reference: referenceInput,
      target: targetInput,
      lockedIndex: null,
    };
  }

  fail(
    "Unsupported selection. Select two eligible paths or two endpoints from those paths."
  );
  return null;
}

function findLockedAnchor(refData, targetData, tolerance) {
  var startMatches =
    pointsClose(targetData.start, refData.start, tolerance) ||
    pointsClose(targetData.start, refData.end, tolerance);
  var endMatches =
    pointsClose(targetData.end, refData.start, tolerance) ||
    pointsClose(targetData.end, refData.end, tolerance);

  if (startMatches && endMatches) return 2;
  if (startMatches) return 0;
  if (endMatches) return 1;
  return -1;
}

function pointsClose(a, b, tolerance) {
  var dx = a[0] - b[0];
  var dy = a[1] - b[1];
  return Math.sqrt(dx * dx + dy * dy) <= tolerance;
}

function getLineData(path) {
  var p0 = path.pathPoints[0].anchor;
  var p1 = path.pathPoints[1].anchor;
  var dx = p1[0] - p0[0];
  var dy = p1[1] - p0[1];
  var len = Math.sqrt(dx * dx + dy * dy);
  var dir = len === 0 ? [0, 0] : [dx / len, dy / len];
  return {
    start: [p0[0], p0[1]],
    end: [p1[0], p1[1]],
    direction: dir,
    length: len,
  };
}

function ensureDocument() {
  if (!app.documents.length) {
    fail("Open a document and select eligible paths first.");
    return null;
  }
  try {
    return app.activeDocument || null;
  } catch (e) {
    fail("Unable to access the active document.");
    return null;
  }
}

function adjustTargetLength(target, targetData, referenceLength, lockedIndex) {
  var movingIndex = lockedIndex === 0 ? 1 : 0;
  var fixedAnchor = lockedIndex === 0 ? targetData.start : targetData.end;
  var direction =
    lockedIndex === 0
      ? targetData.direction
      : [-targetData.direction[0], -targetData.direction[1]];

  var newAnchor = [
    fixedAnchor[0] + direction[0] * referenceLength,
    fixedAnchor[1] + direction[1] * referenceLength,
  ];

  var point = target.pathPoints[movingIndex];
  point.anchor = newAnchor;
  point.leftDirection = newAnchor.slice();
  point.rightDirection = newAnchor.slice();
}

function report(message) {
  $.writeln("[MakePathsEqual] " + message);
}

function fail(message) {
  report(message);
  alert(message);
}
