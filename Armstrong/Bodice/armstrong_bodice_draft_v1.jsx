/*
 Helen Joseph-Armstrong's Front Bodice Draft
 ------------------------------------------------
 - Measurement dialog (inches) with default size 12 inputs
 - Fraction-to-decimal helper note for sixteenths
*/

(function () {
    var SHOW_MEASUREMENT_DIALOG = true;

    var params = {
        size: 12,
        fullLength: 18,
        acrossShoulder: 7.9375,
        centreFrontLength: 14.875,
        bustArc: 10.375,
        shoulderSlope: 18.125,
        bustDepth: 9.6875,
        shoulderLength: 5.375,
        bustSpan: 4.0625,
        acrossChest: 6.9375,
        dartPlacement: 3.4375,
        newStrap: 18.1875,
        sideLength: 8.5,
        waistArc: 7.375,
        fullLengthBack: 17.875,
        acrossShoulderBack: 8.1875,
        centreFrontLengthBack: 17,
        bustArcBack: 9,
        shoulderSlopeBack: 17.375,
        shoulderLengthBack: 5.375,
        bustSpanBack: 4.0625,
        acrossChestBack: 7.1875,
        dartPlacementBack: 3.4375,
        sideLengthBack: 8.5,
        waistArcBack: 7,
        backNeck: 3.125,
        bustCup: "B Cup",
        showMeasurementPalette: false
    };

    var IN_TO_PT = 72;
    function inches(val) { return val * IN_TO_PT; }

    var CM_TO_PT = 28.346456692913385;
    function cm(val) { return val * CM_TO_PT; }

    var ARTBOARD_MARGIN_CM = 10;
    var STROKE_PT = 1;
    var DASH_PATTERN_PT = [25, 12];
    var MARKER_RADIUS_CM = 0.45;
    var LABEL_FONT_SIZE_PT = 10;
    var POINT_LABEL_OFFSET_IN = 0.15;
    var DART_PLACEMENT_DROP_IN = 0.1875;
    var BACK_GAP_CM = 10;
    var BACK_OFFSET_IN = BACK_GAP_CM / 2.54;

    var BUST_CUP_OFFSETS = {
        "A": 0.875,
        "B": 1.25,
        "C": 1.5,
        "D": 1.75
    };

    function resolveBustCupOffset(cup) {
        var fallback = BUST_CUP_OFFSETS.B;
        if (!cup) return fallback;
        var normalized = String(cup).toUpperCase();
        if (normalized.indexOf("CUP") !== -1) {
            normalized = normalized.replace("CUP", "");
        }
        normalized = trimWhitespace(normalized.replace(/[^A-Z]/g, ""));
        if (!normalized) return fallback;
        var key = normalized.charAt(0);
        if (BUST_CUP_OFFSETS.hasOwnProperty(key)) {
            return BUST_CUP_OFFSETS[key];
        }
        return fallback;
    }

    if (SHOW_MEASUREMENT_DIALOG) {
        var dialogResult = showMeasurementDialog(params);
        if (!dialogResult) return;
        params = dialogResult;
    }

    if (params.showMeasurementPalette === false) {
        closeMeasurementPalette();
    } else {
        showMeasurementPalette(params);
    }

    // Drafting instructions will append to the block below as steps are defined.

    var doc = ensureDocument();
    if (!doc) return;

    clearDocumentLayers(doc);
    var layerStack = setupLayers(doc);
    var artboardIndex = doc ? doc.artboards.getActiveArtboardIndex() : -1;
    var draftItems = [];
    function trackDraftItem(item) {
        if (!item) return;
        draftItems.push(item);
    }
    // Pattern coordinate system: X grows to the left of A, Y grows downward from A.
    var origin = determineOrigin(doc, artboardIndex);

    var points = {};
    var currentLineColor = "black";
    var measurementPalette = null;

    function closeMeasurementPalette() {
        try {
            if (measurementPalette && typeof measurementPalette.close === "function") {
                measurementPalette.close();
            }
        } catch (errCloseLocal) {}
        try {
            var globalPalette = $.global.armstrongMeasurementPalette;
            if (globalPalette && globalPalette.window && typeof globalPalette.window.close === "function") {
                globalPalette.window.close();
            }
            delete $.global.armstrongMeasurementPalette;
        } catch (errRemoveGlobal) {}
        measurementPalette = null;
    }

    // Drafting instructions
    var pointAResult = markPoint("A", { x: 0, y: 0 }, "A");
    var pointA = pointAResult.point;
    var fullLengthPlusEighth = params.fullLength + 0.125;
    var pointBResult = markPoint("B", { x: 0, y: fullLengthPlusEighth }, "B");
    var pointB = pointBResult.point;
    var abLine = drawStraightLine(layerStack.front, pointA, pointB, "solid", { name: "Full Length" });
    trackDraftItem(abLine);

    // Square from A to the left: Across Shoulder minus 1/8 to establish point C.
    var acrossShoulderMinusEighth = params.acrossShoulder - 0.125;
    var pointCResult = markPoint("C", {
        x: pointA.x + acrossShoulderMinusEighth,
        y: pointA.y
    }, "C");
    var pointC = pointCResult.point;
    var acLine = drawStraightLine(layerStack.front, pointA, pointC, "solid", { name: "Front Across Shoulder" });
    trackDraftItem(acLine);

    // From B mark upwards Centre Front Length to establish point D.
    var pointDResult = markPoint("D", {
        x: pointB.x,
        y: pointB.y - params.centreFrontLength
    }, "D");
    var pointD = pointDResult.point;
    var bdLine = drawStraightLine(layerStack.front, pointB, pointD, "solid", { name: "B-D", color: "blue" });
    trackDraftItem(bdLine);

    // Square dashed left of D by 4 inches.
    var dLeftPoint = {
        x: pointD.x + 4,
        y: pointD.y
    };
    var dDashLine = drawStraightLine(layerStack.foundationFront, pointD, dLeftPoint, "dashed", { name: "Front Neck Guide 1" });
    trackDraftItem(dDashLine);

    // Square left of B by Bust Arc plus 1/4 to establish point E.
    var bustArcPlusQuarter = params.bustArc + 0.25;
    var pointEResult = markPoint("E", {
        x: pointB.x + bustArcPlusQuarter,
        y: pointB.y
    }, "E");
    var pointE = pointEResult.point;
    var beLine = drawStraightLine(layerStack.front, pointB, pointE, "solid", { name: "Bust Arc" });
    trackDraftItem(beLine);

    // Square up from E by 11 inches.
    var eUpPoint = {
        x: pointE.x,
        y: pointE.y - 11
    };
    var eVerticalLine = drawStraightLine(layerStack.front, pointE, eUpPoint, "solid", { name: "Side Guide Line" });
    trackDraftItem(eVerticalLine);

    // Square down 4 inches from C (guide).
    var cGuidePoint = {
        x: pointC.x,
        y: pointC.y + 4
    };
    currentLineColor = "black";
    var cGuideLine = drawStraightLine(layerStack.front, pointC, cGuidePoint, "dashed", { color: "black", name: "Shoulder Slope Guide" });
    trackDraftItem(cGuideLine);

    // Draw the shoulder slope from B to the 4-inch guide and mark G.
    currentLineColor = "blue";
    var shoulderSlopeLength = params.shoulderSlope;
    var horizontalSeparation = Math.abs(pointC.x - pointB.x);
    var verticalSpanSquared = (shoulderSlopeLength * shoulderSlopeLength) - (horizontalSeparation * horizontalSeparation);
    var verticalSpan = (verticalSpanSquared > 0) ? Math.sqrt(verticalSpanSquared) : 0;
    var gY = pointB.y - verticalSpan;
    if (gY < pointC.y) gY = pointC.y;
    if (gY > cGuidePoint.y) gY = cGuidePoint.y;

    var pointGResult = markPoint("G", {
        x: pointC.x,
        y: gY
    }, "G", "Shoulder Slope");
    var pointG = pointGResult.point;
    var shoulderSlopeLine = drawStraightLine(layerStack.foundationFront, pointB, pointG, "solid", { name: "Shoulder Slope", color: "black" });
    trackDraftItem(shoulderSlopeLine);

    // Mark bust depth along the shoulder slope from G toward B (point H).
    var gbVectorX = pointB.x - pointG.x;
    var gbVectorY = pointB.y - pointG.y;
    var gbLength = Math.sqrt(gbVectorX * gbVectorX + gbVectorY * gbVectorY);
    var gbUnitX = gbLength ? gbVectorX / gbLength : 0;
    var gbUnitY = gbLength ? gbVectorY / gbLength : 0;
    var pointHPosition = {
        x: pointG.x + gbUnitX * params.bustDepth,
        y: pointG.y + gbUnitY * params.bustDepth
    };
    var pointHResult = markPoint("H", pointHPosition, "H");
    var pointH = pointHResult.point;

    // Mark I on the across shoulder line at shoulder length from G.
    var acVectorX = pointC.x - pointA.x;
    var acVectorY = pointC.y - pointA.y;
    var wX = pointA.x - pointG.x;
    var wY = pointA.y - pointG.y;
    var aQuad = acVectorX * acVectorX + acVectorY * acVectorY;
    var bQuad = 2 * (wX * acVectorX + wY * acVectorY);
    var cQuad = wX * wX + wY * wY - params.shoulderLength * params.shoulderLength;
    var discriminant = bQuad * bQuad - 4 * aQuad * cQuad;
    var t = 0;
    if (discriminant < 0) {
        discriminant = 0;
    }
    if (aQuad !== 0) {
        var sqrtDisc = Math.sqrt(discriminant);
        var t1 = (-bQuad + sqrtDisc) / (2 * aQuad);
        var t2 = (-bQuad - sqrtDisc) / (2 * aQuad);
        function clamp01(val) { return val >= 0 && val <= 1; }
        if (clamp01(t1) && clamp01(t2)) {
            t = (Math.abs(t1 - 1) < Math.abs(t2 - 1)) ? t1 : t2;
        } else if (clamp01(t1)) {
            t = t1;
        } else if (clamp01(t2)) {
            t = t2;
        } else {
            t = Math.max(0, Math.min(1, t1));
        }
    }
    var pointIPosition = {
        x: pointA.x + acVectorX * t,
        y: pointA.y + acVectorY * t
    };
    var pointIResult = markPoint("I", pointIPosition, "I");
    var pointI = pointIResult.point;
    var giLine = drawStraightLine(layerStack.front, pointG, pointI, "solid", { name: "Shoulder Length", color: "blue" });
    trackDraftItem(giLine);

    var diStartControl = movePoint(pointD, 1.8, 0);
    var diEndControl = movePoint(pointI, 0, 1.8);
    var diArc = drawInwardArc(layerStack.front, pointD, pointI, 0, {
        name: "D-I Arc",
        color: "blue",
        controlStartAbsolute: diStartControl,
        controlEndAbsolute: diEndControl
    });
    trackDraftItem(diArc);

    // Drop a right-angled line from I to the D guide (perpendicular to GI).
    var giVecX = pointI.x - pointG.x;
    var giVecY = pointI.y - pointG.y;
    var giLength = Math.sqrt(giVecX * giVecX + giVecY * giVecY);
    var perpX = giLength ? (-giVecY / giLength) : 0;
    var perpY = giLength ? (giVecX / giLength) : 1;
    var tDrop = (perpY !== 0) ? (pointD.y - pointI.y) / perpY : 0;
    var iDropPoint = {
        x: pointI.x + perpX * tDrop,
        y: pointI.y + perpY * tDrop
    };
    var iDropLine = drawStraightLine(layerStack.foundationFront, pointI, iDropPoint, "dashed", { name: "Front Neck Guide 2", color: "black" });
    trackDraftItem(iDropLine);

    // Mark J from H horizontally to meet the full length line at H's vertical level.
    var pointJPosition = {
        x: pointB.x,
        y: pointH.y
    };
    var pointJResult = markPoint("J", pointJPosition, "J");
    var pointJ = pointJResult.point;

    // Mark L from D downward along the full length line by half of D to J.
    var dToJLength = pointJ.y - pointD.y;
    var pointLPosition = {
        x: pointB.x,
        y: pointD.y + dToJLength * 0.5
    };
    var pointLResult = markPoint("L", pointLPosition, "L");
    var pointL = pointLResult.point;

    // Square out left from J by Bust Span plus 1/4 to mark K.
    var bustSpanPlusQuarter = params.bustSpan + 0.25;
    var pointKPosition = {
        x: pointJ.x + bustSpanPlusQuarter,
        y: pointJ.y
    };
    var pointKResult = markPoint("K", pointKPosition, "K");
    var pointK = pointKResult.point;
    var kDropPoint = {
        x: pointK.x,
        y: pointK.y + 0.625
    };
    var kDropLine = drawStraightLine(layerStack.foundationFront, pointK, kDropPoint, "dashed", { name: "Dart Reduction", color: "black" });
    trackDraftItem(kDropLine);
    var jkLine = drawStraightLine(layerStack.foundationFront, pointJ, pointK, "solid", { name: "Bust Span", color: "black" });
    trackDraftItem(jkLine);

    // Square out left from L by Across Chest plus 1/4 to mark M.
    var acrossChestPlusQuarter = params.acrossChest + 0.25;
    var pointMPosition = {
        x: pointL.x + acrossChestPlusQuarter,
        y: pointL.y
    };
    var pointMResult = markPoint("M", pointMPosition, "M");
    var pointM = pointMResult.point;
    var lmLine = drawStraightLine(layerStack.foundationFront, pointL, pointM, "solid", { name: "Across Chest", color: "black" });
    trackDraftItem(lmLine);

    // Vertical guideline at M: 2 inches up, 1 inch down.
    var mGuideTop = {
        x: pointM.x,
        y: pointM.y - 2
    };
    var mGuideBottom = {
        x: pointM.x,
        y: pointM.y + 1
    };
    var mGuideLine = drawStraightLine(layerStack.foundationFront, mGuideTop, mGuideBottom, "dashed", { name: "Front Armhole Guide Line", color: "black" });
    trackDraftItem(mGuideLine);

    // Mark F to the left of B by Dart Placement and drop 3/16.
    var pointFTop = {
        x: pointB.x + params.dartPlacement,
        y: pointB.y
    };
    var pointFPosition = {
        x: pointFTop.x,
        y: pointFTop.y + DART_PLACEMENT_DROP_IN
    };
    var fDropLine = drawStraightLine(layerStack.foundationFront, pointFTop, pointFPosition, "solid", { name: "Dart Placement Guide", color: "black" });
    trackDraftItem(fDropLine);
    var pointFResult = markPoint("F", pointFPosition, "F");
    var pointF = pointFResult.point;
    var bfLine = drawStraightLine(layerStack.front, pointB, pointF, "solid", { name: "Front Waist Line 2", color: "blue" });
    trackDraftItem(bfLine);

    var kfStartPoint = kDropPoint;
    var kfLine = drawStraightLine(layerStack.front, kfStartPoint, pointF, "solid", { name: "K-F", color: "blue" });
    trackDraftItem(kfLine);

    // Mark N from I to the Side Guide Line using New Strap + 1/8.
    var newStrapPlusEighth = params.newStrap + 0.125;
    var dxSide = pointE.x - pointI.x;
    var dySquared = newStrapPlusEighth * newStrapPlusEighth - dxSide * dxSide;
    if (dySquared < 0) dySquared = 0;
    var dySide = Math.sqrt(dySquared);
    var candidateNY1 = pointI.y + dySide;
    var candidateNY2 = pointI.y - dySide;
    var sideUpperY = pointE.y;
    var sideLowerY = pointE.y - 11;
    function clampToSide(y) {
        return y >= Math.min(sideLowerY, sideUpperY) && y <= Math.max(sideLowerY, sideUpperY);
    }
    var selectedNY = candidateNY1;
    if (!clampToSide(selectedNY) && clampToSide(candidateNY2)) {
        selectedNY = candidateNY2;
    } else if (clampToSide(candidateNY1) && clampToSide(candidateNY2)) {
        // Choose the candidate nearer to the top of the guide (pointE.y)
        selectedNY = (Math.abs(candidateNY1 - sideUpperY) < Math.abs(candidateNY2 - sideUpperY)) ? candidateNY1 : candidateNY2;
    } else if (!clampToSide(selectedNY)) {
        selectedNY = Math.max(Math.min(selectedNY, sideUpperY), sideLowerY);
    }
    var pointNPosition = {
        x: pointE.x,
        y: selectedNY
    };
    var pointNResult = markPoint("N", pointNPosition, "N");
    var pointN = pointNResult.point;
    var inLine = drawStraightLine(layerStack.foundationFront, pointI, pointN, "solid", { name: "New Strap", color: "black" });
    trackDraftItem(inLine);

    // Mark O on the side guide line from N with length Side Length (upwards).
    var pointOPosition = {
        x: pointE.x,
        y: pointN.y - params.sideLength
    };
    var pointOResult = markPoint("O", pointOPosition, "O");
    var pointO = pointOResult.point;
    var noLine = drawStraightLine(layerStack.foundationFront, pointN, pointO, "solid", { name: "Provisional Side Length", color: "black" });
    trackDraftItem(noLine);

    // Draw inward arc from G to O.
    var goArc = drawInwardArc(layerStack.front, pointG, pointO, 0, {
        name: "G-O Arc",
        color: "blue",
        controlStartRatio: { tangent: 0.5295, normal: 0.494 },
        controlStartOffset: { x: -0.15, y: 0 },
        controlEndRatio: { tangent: -0.1495, normal: 0.3325 }
    });
    trackDraftItem(goArc);

    // Mark P base 1.25 inches to the left of N, then project to side length direction.
    var bustCupOffset = resolveBustCupOffset(params.bustCup);
    var pointPBase = {
        x: pointN.x + bustCupOffset,
        y: pointN.y
    };
    var pointPPosition = null;
    var sideLengthTarget = params.sideLength;
    var onVecX = pointN.x - pointO.x;
    var onVecY = pointN.y - pointO.y;
    var onDistance = Math.sqrt(onVecX * onVecX + onVecY * onVecY);

    if (sideLengthTarget > 0 && onDistance > 0) {
        var r0 = sideLengthTarget;
        var r1 = bustCupOffset;
        if (onDistance <= r0 + r1 && onDistance >= Math.abs(r0 - r1)) {
            var aInt = (r0 * r0 - r1 * r1 + onDistance * onDistance) / (2 * onDistance);
            var hSq = r0 * r0 - aInt * aInt;
            if (hSq < 0) hSq = 0;
            var hInt = Math.sqrt(hSq);
            var baseX = pointO.x + aInt * (pointN.x - pointO.x) / onDistance;
            var baseY = pointO.y + aInt * (pointN.y - pointO.y) / onDistance;
            var offsetX = -(pointN.y - pointO.y) * (hInt / onDistance);
            var offsetY = (pointN.x - pointO.x) * (hInt / onDistance);
            var candidate1 = { x: baseX + offsetX, y: baseY + offsetY };
            var candidate2 = { x: baseX - offsetX, y: baseY - offsetY };
            var dist1 = (candidate1.x - pointPBase.x) * (candidate1.x - pointPBase.x) + (candidate1.y - pointPBase.y) * (candidate1.y - pointPBase.y);
            var dist2 = (candidate2.x - pointPBase.x) * (candidate2.x - pointPBase.x) + (candidate2.y - pointPBase.y) * (candidate2.y - pointPBase.y);
            pointPPosition = dist1 <= dist2 ? candidate1 : candidate2;
        }
    }

    if (!pointPPosition) {
        var opVecX = pointPBase.x - pointO.x;
        var opVecY = pointPBase.y - pointO.y;
        var opLength = Math.sqrt(opVecX * opVecX + opVecY * opVecY);
        if (opLength > 0 && sideLengthTarget > 0) {
            var opScale = sideLengthTarget / opLength;
            pointPPosition = {
                x: pointO.x + opVecX * opScale,
                y: pointO.y + opVecY * opScale
            };
        } else {
            pointPPosition = {
                x: pointPBase.x,
                y: pointPBase.y
            };
        }
    }

    var opLine = drawStraightLine(layerStack.front, pointO, pointPPosition, "solid", { name: "Side Length OP", color: "blue" });
    trackDraftItem(opLine);
    var pointPResult = markPoint("P", pointPPosition, "P");
    var pointP = pointPResult.point;

    // Mark Q from P along PF by (Waist Arc + 1/4) - BF.
    var waistArcPlusQuarter = params.waistArc + 0.25;
    var bToFValue = Math.abs(pointFTop.x - pointB.x);
    var qDistance = waistArcPlusQuarter - bToFValue;
    var pfVecX = pointF.x - pointP.x;
    var pfVecY = pointF.y - pointP.y;
    var pfLength = Math.sqrt(pfVecX * pfVecX + pfVecY * pfVecY);
    var pointQPosition = {
        x: pointP.x,
        y: pointP.y
    };
    if (pfLength > 0) {
        var scaleToQ = qDistance / pfLength;
        pointQPosition = {
            x: pointP.x + pfVecX * scaleToQ,
            y: pointP.y + pfVecY * scaleToQ
        };
    }
    var pointQTarget = {
        x: pointQPosition.x,
        y: pointQPosition.y
    };

    // Adjust Q target so KQ follows KF length while staying on the same direction.
    var kfVecX = pointF.x - kfStartPoint.x;
    var kfVecY = pointF.y - kfStartPoint.y;
    var kfLength = Math.sqrt(kfVecX * kfVecX + kfVecY * kfVecY);
    var kqStartPoint = kDropPoint || pointK;
    var kqVecX = pointQTarget.x - kqStartPoint.x;
    var kqVecY = pointQTarget.y - kqStartPoint.y;
    var kqLength = Math.sqrt(kqVecX * kqVecX + kqVecY * kqVecY);
    if (kfLength > 0 && kqLength > 0) {
        var kqScale = kfLength / kqLength;
        pointQTarget = {
            x: kqStartPoint.x + kqVecX * kqScale,
            y: kqStartPoint.y + kqVecY * kqScale
        };
    }
    var kqGuideLine = drawStraightLine(layerStack.front, kqStartPoint, pointQTarget, "solid", { name: "KQ Equivalent", color: "blue" });
    trackDraftItem(kqGuideLine);

    var pointQResult = markPoint("Q", pointQTarget, "Q");
    var pointQ = pointQResult.point;

    var pqLine = drawStraightLine(layerStack.front, pointP, pointQ, "solid", { name: "Front Waist Line 1", color: "blue" });
    trackDraftItem(pqLine);

    // Back foundation baseline (offset gap to the right of front A) aligned so BB shares front B's Y.
    var backFullLength = params.fullLengthBack || params.fullLength;
    var backFullLengthPlusEighth = backFullLength + 0.125;
    var backVerticalOffset = pointB.y - backFullLengthPlusEighth;
    var backOrigin = {
        x: pointA.x - BACK_OFFSET_IN,
        y: pointA.y + backVerticalOffset
    };
    function backCoord(x, y) {
        return { x: backOrigin.x + x, y: backOrigin.y + y };
    }

    currentLineColor = "black";
    var backPointAResult = markPoint("BA", backCoord(0, 0), "A");
    var backPointA = backPointAResult.point;
    var backPointBResult = markPoint("BB", backCoord(0, backFullLengthPlusEighth), "B");
    var backPointB = backPointBResult.point;
    var backABLine = drawStraightLine(layerStack.foundationBack, backPointA, backPointB, "solid", {
        name: "Back Full Length",
        color: "black",
        layer: layerStack.foundationBack
    });
    trackDraftItem(backABLine);
    var backDartPlacement = params.dartPlacementBack || params.dartPlacement || 0;
    var backPointIResult = markPoint("BI", backCoord(-backDartPlacement, backFullLengthPlusEighth), "I");
    var backPointI = backPointIResult.point;

    var backWaistArcValue = params.waistArcBack || params.waistArc || 0;
    var backJOffset = backWaistArcValue + 1.5 + 0.25;
    var backPointJResult = markPoint("BJ", backCoord(-backJOffset, backFullLengthPlusEighth), "J");
    var backPointJ = backPointJResult.point;

    var dartIntake = 1.5;
    var backPointKResult = markPoint("BK", backCoord(-(backDartPlacement + dartIntake), backFullLengthPlusEighth), "K");
    var backPointK = backPointKResult.point;

    var backPointLPosition = backCoord(-(backDartPlacement + dartIntake / 2), backFullLengthPlusEighth);
    var backPointLResult = markPoint("BL", backPointLPosition, "L");
    var backPointL = backPointLResult.point;

    var backPointMResult = markPoint("BM", backCoord(-backJOffset, backFullLengthPlusEighth + 0.1875), "M");
    var backPointM = backPointMResult.point;

    var centreBackLength = params.centreFrontLengthBack || params.centreFrontLength;
    var backPointDResult = markPoint("BD", backCoord(0, backFullLengthPlusEighth - centreBackLength), "D");
    var backPointD = backPointDResult.point;
    var backDBLine = drawStraightLine(layerStack.back, backPointD, backPointB, "solid", {
        name: "Centre Back",
        color: "blue",
        layer: layerStack.back
    });
    trackDraftItem(backDBLine);

    var backAcrossShoulder = params.acrossShoulderBack || params.acrossShoulder;
    var backPointCResult = markPoint("BC", backCoord(-backAcrossShoulder, 0), "C");
    var backPointC = backPointCResult.point;
    var backACLine = drawStraightLine(layerStack.foundationBack, backPointA, backPointC, "solid", {
        name: "Back Across Shoulder",
        color: "black",
        layer: layerStack.foundationBack
    });
    trackDraftItem(backACLine);

    var backDRight = {
        x: backPointD.x - 4,
        y: backPointD.y
    };
    var backDGuide = drawStraightLine(layerStack.foundationBack, backPointD, backDRight, "dashed", {
        name: "Back Neck Guide",
        color: "black",
        layer: layerStack.foundationBack
    });
    trackDraftItem(backDGuide);

    var backCDown = {
        x: backPointC.x,
        y: backPointC.y + 6
    };
    var backCGuide = drawStraightLine(layerStack.foundationBack, backPointC, backCDown, "dashed", {
        name: "Back Shoulder Slope Guide",
        color: "black",
        layer: layerStack.foundationBack
    });
    trackDraftItem(backCGuide);

    var backBustArc = params.bustArcBack || params.bustArc;
    var backPointEResult = markPoint("BE", backCoord(-(backBustArc + 0.75), backFullLengthPlusEighth), "E");
    var backPointE = backPointEResult.point;
    var backBEline = drawStraightLine(layerStack.foundationBack, backPointB, backPointE, "solid", {
        name: "Back Arc",
        color: "black",
        layer: layerStack.foundationBack
    });
    trackDraftItem(backBEline);
    var backEUp = backCoord(-(backBustArc + 0.75), backFullLengthPlusEighth - 10);
    var backEGuide = drawStraightLine(layerStack.foundationBack, backPointE, backEUp, "solid", {
        name: "Back Side Guide Line",
        color: "black",
        layer: layerStack.foundationBack
    });
    trackDraftItem(backEGuide);

    var backNeckPlusEighth = (params.backNeck || 0) + 0.125;
    var backPointFResult = markPoint("BF", backCoord(-backNeckPlusEighth, 0), "F");
    var backPointF = backPointFResult.point;

    var backDFVecX = backPointF.x - backPointD.x;
    var backDFVecY = backPointF.y - backPointD.y;
    var backDFLength = Math.sqrt(backDFVecX * backDFVecX + backDFVecY * backDFVecY);
    if (backDFLength > 0) {
        var backDFArc = drawBezierArc(layerStack.back, backPointD, backPointD, backPointF, {
            name: "Back Neck Curve",
            color: "blue",
            controlStartRatio: { tangent: 0.5164137931, normal: -0.1390344828 },
            controlEndRatio: { tangent: -0.1807448276, normal: -0.1913379310 }
        });
        trackDraftItem(backDFArc);
    }

    var dxSlope = backPointC.x - backPointB.x;
    var baseDy = backPointC.y - backPointB.y;
    var shoulderSlopeBack = (params.shoulderSlopeBack || params.shoulderSlope) + 0.125;
    var aSlope = 1;
    var bSlope = 2 * baseDy;
    var cSlope = baseDy * baseDy + dxSlope * dxSlope - shoulderSlopeBack * shoulderSlopeBack;
    var discriminantSlope = bSlope * bSlope - 4 * aSlope * cSlope;
    var tSlope = 0;
    if (discriminantSlope >= 0) {
        var sqrtDiscSlope = Math.sqrt(discriminantSlope);
        var t1Slope = (-bSlope + sqrtDiscSlope) / (2 * aSlope);
        var t2Slope = (-bSlope - sqrtDiscSlope) / (2 * aSlope);
        function withinBackGuide(val) {
            return val >= 0 && val <= 6;
        }
        if (withinBackGuide(t1Slope) && withinBackGuide(t2Slope)) {
            tSlope = Math.max(t1Slope, t2Slope);
        } else if (withinBackGuide(t1Slope)) {
            tSlope = t1Slope;
        } else if (withinBackGuide(t2Slope)) {
            tSlope = t2Slope;
        } else {
            tSlope = Math.max(0, Math.min(6, t1Slope));
        }
    }
    var backPointGPosition = backCoord(-backAcrossShoulder, tSlope);
    var backPointGResult = markPoint("BG", backPointGPosition, "G");
    var backPointG = backPointGResult.point;
    var backShoulderSlopeLine = drawStraightLine(layerStack.foundationBack, backPointB, backPointG, "solid", {
        name: "Back Shoulder Slope",
        color: "black",
        layer: layerStack.foundationBack
    });
    trackDraftItem(backShoulderSlopeLine);

    var backFGVecX = backPointG.x - backPointF.x;
    var backFGVecY = backPointG.y - backPointF.y;
    var backFGLen = Math.sqrt(backFGVecX * backFGVecX + backFGVecY * backFGVecY);
    var shoulderLengthPlusHalf = (params.shoulderLengthBack || params.shoulderLength || 0) + 0.5;
    var backFHScale = backFGLen ? (shoulderLengthPlusHalf / backFGLen) : 0;
    var backPointHPosition = {
        x: backPointF.x + backFGVecX * backFHScale,
        y: backPointF.y + backFGVecY * backFHScale
    };
    var backPointHResult = markPoint("BH", backPointHPosition, "H");
    var backPointH = backPointHResult.point;
    var backFHLine = drawStraightLine(layerStack.foundationBack, backPointF, backPointH, "solid", {
        name: "Back Shoulder Length",
        color: "black",
        layer: layerStack.foundationBack
    });
    trackDraftItem(backFHLine);

    var backPointPMid = {
        x: (backPointF.x + backPointH.x) / 2,
        y: (backPointF.y + backPointH.y) / 2
    };
    var backPointPResult = markPoint("BP", backPointPMid, "P");
    var backPointP = backPointPResult.point;

    var fhVecX = backPointH.x - backPointF.x;
    var fhVecY = backPointH.y - backPointF.y;
    var fhLength = Math.sqrt(fhVecX * fhVecX + fhVecY * fhVecY) || 1;
    var fhUnitX = fhVecX / fhLength;
    var fhUnitY = fhVecY / fhLength;

    var backPointRPosition = {
        x: backPointP.x - fhUnitX * 0.25,
        y: backPointP.y - fhUnitY * 0.25
    };
    var backPointRResult = markPoint("BR", backPointRPosition, "R");
    var backPointR = backPointRResult.point;
    var backPointaPosition = {
        x: backPointP.x + fhUnitX * 0.25,
        y: backPointP.y + fhUnitY * 0.25
    };
    var backPointaResult = markPoint("Ba", backPointaPosition, "a");
    var backPointa = backPointaResult.point;

    var backSideLength = params.sideLengthBack || params.sideLength || 0;
    var dxMN = backPointE.x - backPointM.x;
    var dySquaredMN = backSideLength * backSideLength - dxMN * dxMN;
    if (dySquaredMN < 0) dySquaredMN = 0;
    var dyMN = Math.sqrt(dySquaredMN);
    var backPointNPosition = {
        x: backPointE.x,
        y: backPointM.y - dyMN
    };
    if (backPointNPosition.y < backEUp.y) {
        backPointNPosition.y = backEUp.y;
    }
    var backPointNResult = markPoint("BN", backPointNPosition, "N");
    var backPointN = backPointNResult.point;
    var backMNLine = drawStraightLine(layerStack.back, backPointM, backPointN, "solid", {
        name: "Back Side Length",
        color: "blue",
        layer: layerStack.back
    });
    trackDraftItem(backMNLine);

    var backHNVecX = backPointN.x - backPointH.x;
    var backHNVecY = backPointN.y - backPointH.y;
    var backHNLength = Math.sqrt(backHNVecX * backHNVecX + backHNVecY * backHNVecY);
    if (backHNLength > 0) {
        var backHNArc = drawBezierArc(layerStack.back, backPointH, backPointH, backPointN, {
            name: "Back Armhole Curve",
            color: "blue",
            controlStartRatio: { tangent: 0.5836719092, normal: -0.3423131657 },
            controlStartOffset: { x: 0.3, y: 0 },
            controlEndRatio: { tangent: -0.1240008345, normal: -0.3063856393 }
        });
        trackDraftItem(backHNArc);
    }

    var backMNLength = Math.sqrt((backPointN.x - backPointM.x) * (backPointN.x - backPointM.x) +
        (backPointN.y - backPointM.y) * (backPointN.y - backPointM.y));
    var backOOffset = Math.max(0, backMNLength - 1);
    var backPointOPosition = {
        x: backPointL.x,
        y: backPointL.y - backOOffset
    };
    var backPointOResult = markPoint("BO", backPointOPosition, "O");
    var backPointO = backPointOResult.point;
    var backLOLine = drawStraightLine(layerStack.foundationBack, backPointL, backPointO, "dashed", {
        name: "Back Dart Leg Guide",
        color: "black",
        layer: layerStack.foundationBack
    });
    trackDraftItem(backLOLine);

    var backDLength = Math.sqrt(Math.pow(backPointB.x - backPointD.x, 2) + Math.pow(backPointB.y - backPointD.y, 2));
    var backPointSPosition = backCoord(0, backFullLengthPlusEighth - centreBackLength + backDLength / 4);
    var backPointSResult = markPoint("BS", backPointSPosition, "S");
    var backPointS = backPointSResult.point;
    var backAcrossChestBack = params.acrossChestBack || params.acrossChest;
    var backPointTPosition = backCoord(-(backAcrossChestBack + 0.25), backPointS.y);
    var backPointTResult = markPoint("BT", backPointTPosition, "T");
    var backPointT = backPointTResult.point;
    var backST = drawStraightLine(layerStack.foundationBack, backPointS, backPointT, "solid", {
        name: "Across Back",
        color: "black",
        layer: layerStack.foundationBack
    });
    trackDraftItem(backST);

    var backTUp = backCoord(-(backAcrossChestBack + 0.25), backPointT.y - 1);
    var backTDown = backCoord(-(backAcrossChestBack + 0.25), backPointT.y + 4);
    var backUGuide = drawStraightLine(layerStack.foundationBack, backTUp, backTDown, "dashed", {
        name: "Back Armhole Guideline",
        color: "black",
        layer: layerStack.foundationBack
    });
    trackDraftItem(backUGuide);

    var extensionLength = 0.125;
    var oiVecX = backPointI.x - backPointO.x;
    var oiVecY = backPointI.y - backPointO.y;
    var oiLen = Math.sqrt(oiVecX * oiVecX + oiVecY * oiVecY);
    var oiScale = oiLen ? (oiLen + extensionLength) / oiLen : 0;
    var backPointINew = {
        x: backPointO.x + oiVecX * oiScale,
        y: backPointO.y + oiVecY * oiScale
    };
    backPointI = definePoint("BI", backPointINew);
    moveMarker(backPointIResult.marker, backPointINew);
    backPointIResult.point = backPointI;
    var backOILine = drawStraightLine(layerStack.back, backPointO, backPointI, "solid", {
        name: "Back Waist Dart Left Leg",
        color: "blue",
        layer: layerStack.back
    });
    trackDraftItem(backOILine);

    var okVecX = backPointK.x - backPointO.x;
    var okVecY = backPointK.y - backPointO.y;
    var okLen = Math.sqrt(okVecX * okVecX + okVecY * okVecY);
    var okScale = okLen ? (okLen + extensionLength) / okLen : 0;
    var backPointKNew = {
        x: backPointO.x + okVecX * okScale,
        y: backPointO.y + okVecY * okScale
    };
    backPointK = definePoint("BK", backPointKNew);
    moveMarker(backPointKResult.marker, backPointKNew);
    backPointKResult.point = backPointK;
    var backOKLine = drawStraightLine(layerStack.back, backPointO, backPointK, "solid", {
        name: "Back Waist Dart Right Leg",
        color: "blue",
        layer: layerStack.back
    });
    trackDraftItem(backOKLine);

    var backPOVecX = backPointO.x - backPointP.x;
    var backPOVecY = backPointO.y - backPointP.y;
    var backPOLen = Math.sqrt(backPOVecX * backPOVecX + backPOVecY * backPOVecY);
    var backQDistance = 3;
    var backQScale = backPOLen ? (backQDistance / backPOLen) : 0;
    var backPointQPosition = {
        x: backPointP.x + backPOVecX * backQScale,
        y: backPointP.y + backPOVecY * backQScale
    };
    var backPointQResult = markPoint("BQ", backPointQPosition, "Q");
    var backPointQ = backPointQResult.point;

    var backQRVecX = backPointR.x - backPointQ.x;
    var backQRVecY = backPointR.y - backPointQ.y;
    var backQRBaseLength = Math.sqrt(backQRVecX * backQRVecX + backQRVecY * backQRVecY) || 1;
    var backDartLegExtension = 0.125;
    var backQRTargetLength = backQRBaseLength + backDartLegExtension;
    var backPointRExt = {
        x: backPointQ.x + backQRVecX / backQRBaseLength * backQRTargetLength,
        y: backPointQ.y + backQRVecY / backQRBaseLength * backQRTargetLength
    };
    var backQRLine = drawStraightLine(layerStack.back, backPointQ, backPointRExt, "solid", {
        name: "Back Shoulder Dart Left Leg",
        color: "blue",
        layer: layerStack.back
    });
    trackDraftItem(backQRLine);
    backPointR = definePoint("BR", backPointRExt);
    backPointRResult.point = backPointR;
    moveMarker(backPointRResult.marker, backPointRExt);

    var backQAVecX = backPointa.x - backPointQ.x;
    var backQAVecY = backPointa.y - backPointQ.y;
    var backQALength = Math.sqrt(backQAVecX * backQAVecX + backQAVecY * backQAVecY) || 1;
    var backPointaExt = {
        x: backPointQ.x + backQAVecX / backQALength * backQRTargetLength,
        y: backPointQ.y + backQAVecY / backQALength * backQRTargetLength
    };
    var backQALine = drawStraightLine(layerStack.back, backPointQ, backPointaExt, "solid", {
        name: "Back Shoulder Dart Right Leg",
        color: "blue",
        layer: layerStack.back
    });
    trackDraftItem(backQALine);
    backPointa = definePoint("Ba", backPointaExt);
    backPointaResult.point = backPointa;
    moveMarker(backPointaResult.marker, backPointaExt);

    var backAHLine = drawStraightLine(layerStack.back, backPointa, backPointH, "solid", {
        name: "Back Shoulder Line 2",
        color: "blue",
        layer: layerStack.back
    });
    trackDraftItem(backAHLine);

    var backFRLine = drawStraightLine(layerStack.back, backPointF, backPointR, "solid", {
        name: "Back Shoulder Line 1",
        color: "blue",
        layer: layerStack.back
    });
    trackDraftItem(backFRLine);

    var backBILine = drawStraightLine(layerStack.back, backPointB, backPointI, "solid", {
        name: "Back Waist Line 1",
        color: "blue",
        layer: layerStack.back
    });
    trackDraftItem(backBILine);

    var backKMLine = drawStraightLine(layerStack.back, backPointK, backPointM, "solid", {
        name: "Back Waist Line 2",
        color: "blue",
        layer: layerStack.back
    });
    trackDraftItem(backKMLine);

    var backPOLine = drawStraightLine(layerStack.foundationBack, backPointP, backPointO, "dashed", {
        name: "Back Dart Center Guide",
        color: "black",
        layer: layerStack.foundationBack
    });
    trackDraftItem(backPOLine);

    function formatNumber(num) {
        var rounded = Math.round(num * 10000) / 10000;
        var str = rounded.toFixed(4);
        str = str.replace(/\.?0+$/, "");
        return str;
    }

    function greatestCommonDivisor(a, b) {
        a = Math.abs(a);
        b = Math.abs(b);
        while (b !== 0) {
            var temp = b;
            b = a % b;
            a = temp;
        }
        return a || 1;
    }

    function formatInches(num) {
        if (!isFiniteNumber(num)) return "-";
        var negative = num < 0;
        var precision = 32;
        var rounded = Math.round(Math.abs(num) * precision) / precision;
        var whole = Math.floor(rounded);
        var fractional = rounded - whole;
        var numerator = Math.round(fractional * precision);
        if (numerator === precision) {
            whole += 1;
            numerator = 0;
        }
        var fractionText = "";
        if (numerator > 0) {
            var divisor = greatestCommonDivisor(numerator, precision);
            numerator = numerator / divisor;
            var denominator = precision / divisor;
            fractionText = numerator + "/" + denominator;
        }
        var text = "";
        if (negative) {
            text += "-";
        }
        if (fractionText && whole > 0) {
            text += whole + " " + fractionText;
        } else if (fractionText) {
            text += fractionText;
        } else {
            text += whole;
        }
        return text + "\"";
    }

    function getMeasurementDefinitions(defaults) {
        return [
            { label: "1. Full Length", front: { key: "fullLength", value: defaults.fullLength }, back: { key: "fullLengthBack", value: defaults.fullLengthBack } },
            { label: "2. Across Shoulder", front: { key: "acrossShoulder", value: defaults.acrossShoulder }, back: { key: "acrossShoulderBack", value: defaults.acrossShoulderBack } },
            { label: "3. Centre Front/Back Length", front: { key: "centreFrontLength", value: defaults.centreFrontLength }, back: { key: "centreFrontLengthBack", value: defaults.centreFrontLengthBack } },
            { label: "4. Bust/Back Arc", front: { key: "bustArc", value: defaults.bustArc }, back: { key: "bustArcBack", value: defaults.bustArcBack } },
            { label: "5. Shoulder Slope", front: { key: "shoulderSlope", value: defaults.shoulderSlope }, back: { key: "shoulderSlopeBack", value: defaults.shoulderSlopeBack } },
            { label: "6. Bust Depth", front: { key: "bustDepth", value: defaults.bustDepth }, back: null },
            { label: "7. Shoulder Length", front: { key: "shoulderLength", value: defaults.shoulderLength }, back: { key: "shoulderLengthBack", value: defaults.shoulderLengthBack } },
            { label: "8. Bust Span", front: { key: "bustSpan", value: defaults.bustSpan }, back: { key: "bustSpanBack", value: defaults.bustSpanBack } },
            { label: "9. Across Chest/Back", front: { key: "acrossChest", value: defaults.acrossChest }, back: { key: "acrossChestBack", value: defaults.acrossChestBack } },
            { label: "10. Dart Placement", front: { key: "dartPlacement", value: defaults.dartPlacement }, back: { key: "dartPlacementBack", value: defaults.dartPlacementBack } },
            { label: "11. New Strap", front: { key: "newStrap", value: defaults.newStrap }, back: null },
            { label: "12. Side Length", front: { key: "sideLength", value: defaults.sideLength }, back: { key: "sideLengthBack", value: defaults.sideLengthBack } },
            { label: "13. Waist Arc", front: { key: "waistArc", value: defaults.waistArc }, back: { key: "waistArcBack", value: defaults.waistArcBack } },
            { label: "14. Back Neck", front: null, back: { key: "backNeck", value: defaults.backNeck } }
        ];
    }

    function parseMeasurementInput(input) {
        if (input === null || input === undefined) return NaN;
        var text = String(input);
        var sanitizedChars = [];
        for (var ci = 0; ci < text.length; ci++) {
            var ch = text.charAt(ci);
            if (ch >= "0" && ch <= "9") {
                sanitizedChars.push(ch);
                continue;
            }
            if (ch === "." || ch === "/" || ch === "-") {
                sanitizedChars.push(ch);
                continue;
            }
            if (isWhitespaceChar(ch)) {
                sanitizedChars.push(" ");
            }
        }
        var sanitized = trimWhitespace(sanitizedChars.join(""));
        if (sanitized === "") return NaN;

        // Support whole numbers with fractional parts like "7 3/8" or "7-3/8".
        var wholePart = 0;
        var fractionPart = sanitized;
        var separatorMatch = sanitized.match(/^(-?\d+)[ -]+(\d+\/\d+)$/);
        if (separatorMatch) {
            wholePart = parseInt(separatorMatch[1], 10);
            fractionPart = separatorMatch[2];
        } else {
            var wholeOnly = sanitized.match(/^(-?\d+)$/);
            if (wholeOnly) {
                return parseFloat(sanitized);
            }
        }

        var fractionOnly = fractionPart.match(/^(-?\d+)\/(\d+)$/);
        if (fractionOnly) {
            var numerator = parseFloat(fractionOnly[1]);
            var denominator = parseFloat(fractionOnly[2]);
            if (!denominator) return NaN;
            return wholePart + numerator / denominator;
        }

        return parseFloat(sanitized);
    }

    function showMeasurementDialog(defaults) {
        var dlg = new Window("dialog", "Armstrong Front Bodice Draft - Size " + defaults.size);
        dlg.orientation = "column";
        dlg.alignChildren = "fill";
        dlg.spacing = 12;
        dlg.margins = 16;
        dlg.preferredSize.width = 440;

        var introGroup = dlg.add("group");
        introGroup.orientation = "column";
        introGroup.alignChildren = "left";
        introGroup.spacing = 8;
        introGroup.add("statictext", undefined, "Measurements in inches. Default size " + defaults.size + ".");

        var fractionsGroup = introGroup.add("group");
        fractionsGroup.orientation = "column";
        fractionsGroup.alignChildren = ["fill", "top"];
        fractionsGroup.spacing = 4;

        var fractionPairs = [
            ["1/16 = 0.0625", "1/8 = 0.125", "3/16 = 0.1875"],
            ["1/4 = 0.25", "5/16 = 0.3125", "3/8 = 0.375"],
            ["1/2 = 0.5", "3/4 = 0.75", "7/8 = 0.875"]
        ];
        var columnWidth = 100;
        for (var fr = 0; fr < fractionPairs.length; fr++) {
            var fracRow = fractionsGroup.add("group");
            fracRow.orientation = "row";
            fracRow.alignChildren = ["left", "center"];
            fracRow.spacing = 8;
            for (var fc = 0; fc < fractionPairs[fr].length; fc++) {
                var fracLabel = fracRow.add("statictext", undefined, fractionPairs[fr][fc]);
                fracLabel.preferredSize = [columnWidth, 15];
                fracLabel.maximumSize = [columnWidth, 15];
            }
        }

        dlg.add("panel", undefined, "");

        var optionsGroup = dlg.add("group");
        optionsGroup.orientation = "row";
        optionsGroup.alignChildren = ["left", "center"];
        optionsGroup.spacing = 10;
        optionsGroup.margins = [0, 0, 0, 6];
        var paletteCheckbox = optionsGroup.add("checkbox", undefined, "Show measurement summary");
        paletteCheckbox.value = defaults.showMeasurementPalette === true;

        var measurementGroup = dlg.add("group");
        measurementGroup.orientation = "column";
        measurementGroup.alignChildren = ["left", "top"];
        measurementGroup.spacing = 10;
        measurementGroup.margins = [0, 6, 0, 6];
        measurementGroup.preferredSize.width = 360;

        var labelWidth = 170;
        var valueWidth = 80;

        var bustCupRow = measurementGroup.add("group");
        bustCupRow.orientation = "row";
        bustCupRow.alignChildren = ["left", "center"];
        bustCupRow.spacing = 10;
        bustCupRow.margins = [0, 2, 0, 6];

        var bustCupLabel = bustCupRow.add("statictext", undefined, "Bust Cup");
        bustCupLabel.preferredSize = [labelWidth, 20];

        var bustCupDropdown = bustCupRow.add("dropdownlist", undefined, ["A Cup", "B Cup", "C Cup", "D Cup"]);
        bustCupDropdown.preferredSize = [valueWidth * 1.5, 24];
        var defaultCupText = defaults.bustCup || "B Cup";
        var selectedIndex = 0;
        for (var bc = 0; bc < bustCupDropdown.items.length; bc++) {
            if (bustCupDropdown.items[bc].text === defaultCupText) {
                selectedIndex = bc;
                break;
            }
        }
        bustCupDropdown.selection = bustCupDropdown.items[selectedIndex];

        var headerRow = measurementGroup.add("group");
        headerRow.orientation = "row";
        headerRow.alignChildren = ["left", "center"];
        headerRow.spacing = 10;
        headerRow.margins = [0, 4, 0, 4];

        var headerLabel = headerRow.add("statictext", undefined, "Measurement");
        headerLabel.preferredSize = [labelWidth, 18];

        var headerFront = headerRow.add("statictext", undefined, "Front");
        headerFront.preferredSize = [valueWidth, 18];

        var headerBack = headerRow.add("statictext", undefined, "Back");
        headerBack.preferredSize = [valueWidth, 18];

        var measurementDefs = getMeasurementDefinitions(defaults);

        var editFieldMap = {};
        for (var i = 0; i < measurementDefs.length; i++) {
            var def = measurementDefs[i];
            var row = measurementGroup.add("group");
            row.orientation = "row";
            row.alignChildren = ["left", "center"];
            row.spacing = 10;
            row.margins = [0, 2, 0, 2];

            var label = row.add("statictext", undefined, def.label);
            label.preferredSize = [labelWidth, 18];

            if (def.front && def.front.key) {
                var frontField = row.add("edittext", undefined, isFiniteNumber(def.front.value) ? formatNumber(def.front.value) : "");
                frontField.characters = 7;
                frontField.preferredSize = [valueWidth, 24];
                frontField.alignment = ["left", "center"];
                editFieldMap[def.front.key] = { field: frontField, label: def.label + " (Front)" };
            } else {
                var frontPlaceholder = row.add("statictext", undefined, "—");
                frontPlaceholder.preferredSize = [valueWidth, 18];
            }

            if (def.back && def.back.key) {
                var backField = row.add("edittext", undefined, isFiniteNumber(def.back.value) ? formatNumber(def.back.value) : "");
                backField.characters = 7;
                backField.preferredSize = [valueWidth, 24];
                backField.alignment = ["left", "center"];
                editFieldMap[def.back.key] = { field: backField, label: def.label + " (Back)" };
            } else {
                var backPlaceholder = row.add("statictext", undefined, "—");
                backPlaceholder.preferredSize = [valueWidth, 18];
            }
        }

        dlg.add("panel", undefined, "");

        var buttons = dlg.add("group");
        buttons.alignment = "right";
        buttons.spacing = 12;
        buttons.add("button", undefined, "OK", { name: "ok" });
        buttons.add("button", undefined, "Cancel", { name: "cancel" });

        dlg.center();
        var result = dlg.show();
        if (result !== 1) return null;

        var collected = { size: defaults.size };
        for (var key in editFieldMap) {
            if (!editFieldMap.hasOwnProperty(key)) continue;
            var entry = editFieldMap[key];
            var raw = entry.field.text;
            var parsed = parseMeasurementInput(raw);
            if (isNaN(parsed)) {
                alert("Enter a numeric value for " + entry.label + ".");
                try { entry.field.active = true; } catch (errActivate) {}
                return showMeasurementDialog(defaults);
            }
        collected[key] = Math.round(parsed * 10000) / 10000;
    }

        collected.bustCup = (bustCupDropdown && bustCupDropdown.selection) ? bustCupDropdown.selection.text : (defaults.bustCup || "B Cup");
        collected.showMeasurementPalette = paletteCheckbox.value ? true : false;

        return collected;
    }

    function showMeasurementPalette(values) {
        closeMeasurementPalette();
        var globalScope = $.global;

        var palette = new Window("palette", "Measurement Summary");
        palette.orientation = "column";
        palette.alignChildren = ["fill", "top"];
        palette.spacing = 12;
        palette.margins = 16;
        palette.preferredSize.width = 360;

        var cupGroup = palette.add("group");
        cupGroup.orientation = "row";
        cupGroup.alignChildren = ["left", "center"];
        cupGroup.spacing = 6;
        var cupLabel = cupGroup.add("statictext", undefined, "Bust Cup:");
        cupLabel.characters = 12;
        var cupValue = cupGroup.add("statictext", undefined, values.bustCup || "B Cup");
        cupValue.characters = 10;

        palette.add("panel", undefined, "");

        var defs = getMeasurementDefinitions(values);

        var tableGroup = palette.add("group");
        tableGroup.orientation = "column";
        tableGroup.alignChildren = ["fill", "top"];
        tableGroup.spacing = 6;

        var headerGroup = tableGroup.add("group");
        headerGroup.orientation = "row";
        headerGroup.alignChildren = ["left", "center"];
        headerGroup.spacing = 8;
        var headerMeasurement = headerGroup.add("statictext", undefined, "Measurement");
        headerMeasurement.preferredSize = [180, 18];
        var headerFront = headerGroup.add("statictext", undefined, "Front");
        headerFront.justify = "right";
        headerFront.preferredSize = [80, 18];
        var headerBack = headerGroup.add("statictext", undefined, "Back");
        headerBack.justify = "right";
        headerBack.preferredSize = [80, 18];

        var divider = tableGroup.add("panel", undefined, "");
        divider.alignment = "fill";

        var listGroup = tableGroup.add("group");
        listGroup.orientation = "column";
        listGroup.alignChildren = ["fill", "top"];
        listGroup.spacing = 2;

        for (var i = 0; i < defs.length; i++) {
            var def = defs[i];
            var row = listGroup.add("group");
            row.orientation = "row";
            row.alignChildren = ["left", "center"];
            row.spacing = 8;
            row.margins = [0, 2, 0, 2];

            var label = row.add("statictext", undefined, def.label);
            label.preferredSize = [180, 16];

            var frontText = "-";
            if (def.front && def.front.key) {
                frontText = isFiniteNumber(values[def.front.key]) ? formatInches(values[def.front.key]) : "-";
            }
            var frontField = row.add("statictext", undefined, frontText);
            frontField.preferredSize = [80, 16];
            frontField.justify = "right";

            var backText = "-";
            if (def.back && def.back.key) {
                backText = isFiniteNumber(values[def.back.key]) ? formatInches(values[def.back.key]) : "-";
            }
            var backField = row.add("statictext", undefined, backText);
            backField.preferredSize = [80, 16];
            backField.justify = "right";

            var rowDivider = listGroup.add("panel", undefined, "");
            rowDivider.alignment = "fill";
        }
        if (listGroup.children.length > 0) {
            listGroup.remove(listGroup.children[listGroup.children.length - 1]);
        }

        palette.onClose = function () {
            measurementPalette = null;
            try { delete $.global.armstrongMeasurementPalette; } catch (errDeleteGlobal) {}
        };

        palette.show();
        measurementPalette = palette;
        globalScope.armstrongMeasurementPalette = { window: palette };
    }

    function normalizePointName(name) {
        return String(name || "").replace(/\s+/g, "").toUpperCase();
    }

    function ensurePointObject(value) {
        if (!value) return { x: 0, y: 0 };
        if (typeof value.x === "number" && typeof value.y === "number") {
            return { x: value.x, y: value.y };
        }
        if (value.length && value.length >= 2) {
            return { x: value[0], y: value[1] };
        }
        return { x: 0, y: 0 };
    }

    function isFiniteNumber(value) {
        return typeof value === "number" && isFinite(value);
    }

    function determineOrigin(doc, artboardIndex) {
        var marginIn = 6;
        if (!doc || doc.artboards.length === 0) {
            return { x: inches(marginIn), y: inches(marginIn) };
        }
        var index = (artboardIndex >= 0 && artboardIndex < doc.artboards.length) ? artboardIndex : 0;
        var artboard = doc.artboards[index];
        var rect = artboard.artboardRect; // [left, top, right, bottom]
        var artTop = rect[1];
        var artRight = rect[2];
        return {
            x: artRight - inches(marginIn),
            y: artTop - inches(marginIn)
        };
    }

    function definePoint(name, coords) {
        var key = normalizePointName(name);
        var pt = ensurePointObject(coords);
        points[key] = { x: pt.x, y: pt.y };
        return points[key];
    }

    function getPoint(name) {
        var key = normalizePointName(name);
        var pt = points[key];
        if (!pt) return null;
        return { x: pt.x, y: pt.y };
    }

    function movePoint(base, dx, dy) {
        var pt = ensurePointObject(base);
        var deltaX = typeof dx === "number" ? dx : 0;
        var deltaY = typeof dy === "number" ? dy : 0;
        return { x: pt.x + deltaX, y: pt.y + deltaY };
    }

    function toArt(point) {
        var pt = ensurePointObject(point);
        // Pattern X increases leftward, pattern Y increases downward.
        return [origin.x - inches(pt.x), origin.y - inches(pt.y)];
    }

    function rgb(r, g, b) {
        var color = new RGBColor();
        color.red = r;
        color.green = g;
        color.blue = b;
        return color;
    }

    function blackColor() {
        return rgb(0, 0, 0);
    }

    function whiteColor() {
        return rgb(255, 255, 255);
    }

    function blueColor() {
        return rgb(0, 102, 204);
    }

    function resolveStrokeColor(name) {
        var key = (name || "black").toString().toLowerCase();
        switch (key) {
            case "blue": return blueColor();
            case "black":
            default: return blackColor();
        }
    }

    function resolveLayerColor(r, g, b) {
        var color = new RGBColor();
        color.red = r;
        color.green = g;
        color.blue = b;
        return color;
    }

    function ensureLayerWritable(layer) {
        if (!layer) return;
        var current = layer;
        while (current) {
            try { current.locked = false; } catch (errLock) {}
            try { current.visible = true; } catch (errVis) {}
            if (current.parent && current.parent.typename === "Layer") {
                current = current.parent;
            } else {
                current = null;
            }
        }
    }

    function unlockLayerHierarchy(layer) {
        if (!layer) return;
        try { layer.locked = false; } catch (errLayerLock) {}
        try { layer.visible = true; } catch (errLayerVisible) {}
        var sublayers = layer.layers;
        if (sublayers && sublayers.length) {
            for (var i = 0; i < sublayers.length; i++) {
                unlockLayerHierarchy(sublayers[i]);
            }
        }
        var items = layer.pageItems;
        if (items && items.length) {
            for (var j = 0; j < items.length; j++) {
                try { items[j].locked = false; } catch (errItemLock) {}
                try { items[j].hidden = false; } catch (errItemHide) {}
            }
        }
    }

    function clearDocumentLayers(doc) {
        if (!doc) return;
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            try {
                unlockLayerHierarchy(doc.layers[i]);
                doc.layers[i].remove();
            } catch (errRemove) {}
        }
    }

    function createNamedLayer(parent, name) {
        if (!parent || !parent.layers) return null;
        var layer = null;
        try {
            layer = parent.layers.add();
        } catch (errAddSub) {}
        if (!layer) return null;
        layer.name = name;
        layer.locked = false;
        layer.visible = true;
        return layer;
    }

    function setupLayers(doc) {
        if (!doc) return {};

        var stack = {};
        var foundation;
        if (doc.layers.length > 0) {
            foundation = doc.layers[doc.layers.length - 1];
            unlockLayerHierarchy(foundation);
        } else {
            foundation = doc.layers.add();
        }
        foundation.name = "Foundation";
        foundation.locked = false;
        foundation.visible = true;

        stack.foundation = foundation;
        stack.foundationFront = createNamedLayer(foundation, "Front");
        stack.foundationBack = createNamedLayer(foundation, "Back");

        var frontLayer = doc.layers.add();
        frontLayer.name = "Front Bodice";
        frontLayer.locked = false;
        frontLayer.visible = true;
        try { frontLayer.color = resolveLayerColor(255, 235, 0); } catch (errFrontColor) {}

        var backLayer = doc.layers.add();
        backLayer.name = "Back Bodice";
        backLayer.locked = false;
        backLayer.visible = true;

        var labelsParent = doc.layers.add();
        labelsParent.name = "Labels, Numbers & Markers";
        labelsParent.locked = false;
        labelsParent.visible = true;

        stack.front = frontLayer;
        stack.back = backLayer;
        stack.labelsParent = labelsParent;

        try { foundation.zOrder(ZOrderMethod.SENDTOBACK); } catch (errFoundationOrder) {}
        try { stack.labelsParent.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (errLabelsOrder) {}

        stack.labels = createNamedLayer(stack.labelsParent, "Labels");
        stack.markers = createNamedLayer(stack.labelsParent, "Markers");
        stack.numbers = createNamedLayer(stack.labelsParent, "Numbers");

        stack.all = [
            stack.foundation,
            stack.foundationFront,
            stack.foundationBack,
            stack.front,
            stack.back,
            stack.labelsParent,
            stack.labels,
            stack.markers,
            stack.numbers
        ];

        for (var i = 0; i < stack.all.length; i++) {
            ensureLayerWritable(stack.all[i]);
        }

        try { doc.activeLayer = stack.front; } catch (errActiveLayer) {}

        return stack;
    }

    function drawStraightLine(layer, startPoint, endPoint, style, options) {
        if (!layer) return null;
        var opts = {};
        if (typeof options === "string") {
            opts.name = options;
        } else if (options && typeof options === "object") {
            opts = options;
        }
        var startObj = ensurePointObject(startPoint);
        var endObj = ensurePointObject(endPoint);
        var startArt = toArt(startObj);
        var endArt = toArt(endObj);
        var strokeNameRaw = opts.color || currentLineColor || "black";
        var strokeKey = (strokeNameRaw || "").toString().toLowerCase();
        var targetLayer = layer;
        if (strokeKey === "black" && layerStack) {
            if (targetLayer === layerStack.front && layerStack.foundationFront) {
                targetLayer = layerStack.foundationFront;
            } else if (targetLayer === layerStack.back && layerStack.foundationBack) {
                targetLayer = layerStack.foundationBack;
            }
        }
        ensureLayerWritable(targetLayer);
        var path = targetLayer.pathItems.add();
        path.stroked = true;
        path.strokeWidth = STROKE_PT;
        try {
            path.strokeColor = resolveStrokeColor(strokeNameRaw);
        } catch (errStroke) {}
        try { path.strokeCap = StrokeCap.BUTTENDCAP; } catch (errCap) {}
        try { path.strokeJoin = StrokeJoin.MITERENDJOIN; } catch (errJoin) {}
        if (style === "dashed") {
            path.strokeDashes = DASH_PATTERN_PT.slice(0);
        } else {
            path.strokeDashes = [];
        }
        path.filled = false;
        path.closed = false;
        try {
            path.setEntirePath([startArt, endArt]);
        } catch (errPath) {}
        if (opts.name) {
            try { path.name = opts.name; } catch (errName) {}
        }
        return path;
    }

    function drawBezierArc(layer, startPoint, controlPoint, endPoint, options) {
        if (!layer) return null;
        var opts = {};
        if (typeof options === "string") {
            opts.name = options;
        } else if (options && typeof options === "object") {
            opts = options;
        }
        var start = ensurePointObject(startPoint);
        var ctrl = ensurePointObject(controlPoint);
        var end = ensurePointObject(endPoint);

        var strokeNameRaw = opts.color || currentLineColor || "black";
        var strokeKey = (strokeNameRaw || "").toString().toLowerCase();
        var targetLayer = layer;
        if (strokeKey === "black" && layerStack) {
            if (targetLayer === layerStack.front && layerStack.foundationFront) {
                targetLayer = layerStack.foundationFront;
            } else if (targetLayer === layerStack.back && layerStack.foundationBack) {
                targetLayer = layerStack.foundationBack;
            }
        }
        ensureLayerWritable(targetLayer);
        var path = targetLayer.pathItems.add();
        path.stroked = true;
        path.strokeWidth = STROKE_PT;
        try {
            path.strokeColor = resolveStrokeColor(strokeNameRaw);
        } catch (errArcStroke) {}
        path.filled = false;
        if (opts.dashed) {
            path.strokeDashes = DASH_PATTERN_PT.slice(0);
        } else {
            path.strokeDashes = [];
        }
        path.closed = false;
        try { path.strokeCap = StrokeCap.BUTTENDCAP; } catch (errArcCap) {}
        try { path.strokeJoin = StrokeJoin.ROUNDENDJOIN; } catch (errArcJoin) {}

        var startArt = toArt(start);
        var endArt = toArt(end);
        try { path.setEntirePath([startArt, endArt]); } catch (errArcSetPath) {}

        var customStartAbsolute = opts.controlStartAbsolute ? ensurePointObject(opts.controlStartAbsolute) : null;
        var customEndAbsolute = opts.controlEndAbsolute ? ensurePointObject(opts.controlEndAbsolute) : null;
        var customStartOffset = opts.controlStartOffset;
        var customEndOffset = opts.controlEndOffset;
        var startRatio = opts.controlStartRatio;
        var endRatio = opts.controlEndRatio;
        var dxChord = end.x - start.x;
        var dyChord = end.y - start.y;
        var chordLength = Math.sqrt(dxChord * dxChord + dyChord * dyChord);
        var tangentUnit = chordLength ? { x: dxChord / chordLength, y: dyChord / chordLength } : { x: 1, y: 0 };
        var normalUnit = { x: -tangentUnit.y, y: tangentUnit.x };

        function applyRatio(base, ratio) {
            var tangentComp = (ratio && isFiniteNumber(ratio.tangent)) ? ratio.tangent : 0;
            var normalComp = (ratio && isFiniteNumber(ratio.normal)) ? ratio.normal : 0;
            return {
                x: base.x + (tangentUnit.x * tangentComp + normalUnit.x * normalComp) * chordLength,
                y: base.y + (tangentUnit.y * tangentComp + normalUnit.y * normalComp) * chordLength
            };
        }

        var ctrl1;
        var ctrl2;
        if (customStartAbsolute) {
            ctrl1 = customStartAbsolute;
        } else if (startRatio && chordLength) {
            ctrl1 = applyRatio(start, startRatio);
        } else if (customStartOffset) {
            ctrl1 = movePoint(start, customStartOffset.x || 0, customStartOffset.y || 0);
        } else {
            ctrl1 = {
                x: start.x + (ctrl.x - start.x) * (2 / 3),
                y: start.y + (ctrl.y - start.y) * (2 / 3)
            };
        }
        if (customStartOffset && ctrl1) {
            ctrl1 = movePoint(ctrl1, customStartOffset.x || 0, customStartOffset.y || 0);
        }

        if (customEndAbsolute) {
            ctrl2 = customEndAbsolute;
        } else if (endRatio && chordLength) {
            ctrl2 = applyRatio(end, endRatio);
        } else if (customEndOffset) {
            ctrl2 = movePoint(end, customEndOffset.x || 0, customEndOffset.y || 0);
        } else {
            ctrl2 = {
                x: end.x + (ctrl.x - end.x) * (2 / 3),
                y: end.y + (ctrl.y - end.y) * (2 / 3)
            };
        }
        if (customEndOffset && ctrl2) {
            ctrl2 = movePoint(ctrl2, customEndOffset.x || 0, customEndOffset.y || 0);
        }
        var ctrl1Art = toArt(ctrl1);
        var ctrl2Art = toArt(ctrl2);

        var pointsArc = path.pathPoints;
        if (pointsArc.length >= 2) {
            pointsArc[0].leftDirection = startArt.slice(0);
            pointsArc[0].rightDirection = ctrl1Art.slice(0);
            pointsArc[1].rightDirection = endArt.slice(0);
            pointsArc[1].leftDirection = ctrl2Art.slice(0);
            try {
                pointsArc[0].pointType = PointType.SMOOTH;
                pointsArc[1].pointType = PointType.SMOOTH;
            } catch (errPointType) {}
        }

        if (opts.name) {
            try { path.name = opts.name; } catch (errArcName) {}
        }
        return path;
    }

    function drawInwardArc(layer, startPoint, endPoint, depth, options) {
        var start = ensurePointObject(startPoint);
        var end = ensurePointObject(endPoint);
        var dx = end.x - start.x;
        var dy = end.y - start.y;
        var length = Math.sqrt(dx * dx + dy * dy);
        if (!length) return null;
        var mid = {
            x: (start.x + end.x) * 0.5,
            y: (start.y + end.y) * 0.5
        };
        var nx = -dy / length;
        var ny = dx / length;
        var offset = (isFiniteNumber(depth) && depth !== 0) ? depth : 0.5;
        var inward = {
            x: mid.x + nx * offset,
            y: mid.y + ny * offset
        };
        var outward = {
            x: mid.x - nx * offset,
            y: mid.y - ny * offset
        };
        var control = inward;
        if (options && options.forceOutward) {
            control = outward;
        }
        return drawBezierArc(layer, start, control, end, options);
    }

    function createMarker(markersLayer, numbersLayer, position, label) {
        if (!markersLayer || !numbersLayer) {
            return { circle: null, text: null };
        }

        var coords = ensurePointObject(position);
        if (!isFiniteNumber(coords.x) || !isFiniteNumber(coords.y)) {
            return { circle: null, text: null };
        }

        var anchorArt = toArt(coords);
        var radiusPt = cm(MARKER_RADIUS_CM);

        ensureLayerWritable(markersLayer);
        ensureLayerWritable(numbersLayer);

        var circle = null;
        try {
            circle = markersLayer.pathItems.ellipse(
                anchorArt[1] + radiusPt,
                anchorArt[0] - radiusPt,
                radiusPt * 2,
                radiusPt * 2
            );
            circle.stroked = true;
            circle.strokeWidth = STROKE_PT;
            circle.strokeDashes = [];
            try { circle.strokeColor = blackColor(); } catch (errCircleStroke) {}
            circle.filled = true;
            try { circle.fillColor = blackColor(); } catch (errCircleFill) {}
        } catch (errCircle) {
            circle = null;
        }

        var tf = null;
        if (label !== undefined && label !== null && label !== "") {
            try {
                tf = numbersLayer.textFrames.add();
            } catch (errAddNumbersLayer) {
                tf = null;
                if (doc && doc.textFrames) {
                    try {
                        tf = doc.textFrames.add();
                        if (tf && tf.move) {
                            try {
                                tf.move(numbersLayer, ElementPlacement.PLACEATEND);
                            } catch (errMoveTf) {}
                        }
                    } catch (errAddDocTf) {
                        tf = null;
                    }
                }
            }

            if (tf) {
                tf.contents = String(label);
                try {
                    tf.textRange.characterAttributes.size = LABEL_FONT_SIZE_PT;
                } catch (errSize) {}
                try {
                    tf.textRange.characterAttributes.fillColor = whiteColor();
                } catch (errTextColor) {}
                try {
                    tf.textRange.paragraphAttributes.justification = Justification.CENTER;
                } catch (errJust) {
                    try { tf.textRange.justification = Justification.CENTER; } catch (errJustAlt) {}
                }
                try {
                    tf.position = anchorArt;
                } catch (errPos) {
                    try {
                        tf.left = anchorArt[0];
                        tf.top = anchorArt[1];
                    } catch (errPosFallback) {}
                }
                centerTextFrame(tf, anchorArt);
                try {
                    tf.zOrder(ZOrderMethod.BRINGTOFRONT);
                } catch (errZ) {}
            }
        }

        if (circle) trackDraftItem(circle);
        if (tf) trackDraftItem(tf);

        return {
            circle: circle,
            text: tf,
            center: coords
        };
    }

    function createPointLabel(name, point, targetLayer) {
        var labelName = normalizePointName(name);
        var layer = targetLayer || layerStack.labels;
        if (!layer) return null;
        var pt = ensurePointObject(point);
        var labelAnchor = movePoint(pt, POINT_LABEL_OFFSET_IN, -POINT_LABEL_OFFSET_IN);
        var artPos = toArt(labelAnchor);
        ensureLayerWritable(layer);
        var tf = null;
        try {
            tf = layer.textFrames.add();
            tf.contents = labelName;
            tf.textRange.characterAttributes.size = LABEL_FONT_SIZE_PT;
            try { tf.textRange.characterAttributes.fillColor = blackColor(); } catch (errLabelColor) {}
            tf.textRange.justification = Justification.LEFT;
            tf.left = artPos[0];
            tf.top = artPos[1];
        } catch (errLabel) {
            tf = null;
        }
        return tf;
    }

    function markPoint(pointName, position, numberLabel) {
        var coords = definePoint(pointName, position);
        var marker = createMarker(layerStack.markers, layerStack.numbers, coords, numberLabel);
        return {
            point: coords,
            marker: marker
        };
    }

    function moveMarker(markerObj, newPosition) {
        if (!markerObj) return;
        var newCoords = ensurePointObject(newPosition);
        var oldCoords = markerObj.center || newCoords;
        var oldArt = toArt(oldCoords);
        var newArt = toArt(newCoords);
        var dx = newArt[0] - oldArt[0];
        var dy = newArt[1] - oldArt[1];
        try {
            if (markerObj.circle) markerObj.circle.translate(dx, dy);
        } catch (errMoveCircle) {}
        try {
            if (markerObj.text) markerObj.text.translate(dx, dy);
        } catch (errMoveText) {}
        markerObj.center = newCoords;
    }

    function centerTextFrame(textFrame, center) {
        if (!textFrame || !center) return;
        try {
            var bounds = textFrame.visibleBounds;
            var currentCenter = [
                (bounds[0] + bounds[2]) / 2,
                (bounds[1] + bounds[3]) / 2
            ];
            textFrame.translate(center[0] - currentCenter[0], center[1] - currentCenter[1]);
        } catch (errBounds) {}
    }

    function isWhitespaceChar(ch) {
        return ch === " " || ch === "\t" || ch === "\r" || ch === "\n" || ch === "\f";
    }

    function trimWhitespace(str) {
        if (!str) return "";
        var start = 0;
        var end = str.length;
        while (start < end && isWhitespaceChar(str.charAt(start))) start++;
        while (end > start && isWhitespaceChar(str.charAt(end - 1))) end--;
        return str.substring(start, end);
    }

    function ensureDocument() {
        var widthPt = inches(36);
        var heightPt = inches(48);

        function addDocument() {
            try {
                return app.documents.add(DocumentColorSpace.RGB, widthPt, heightPt);
            } catch (errAddDoc) {
                return app.documents.add();
            }
        }

        if (app.documents.length === 0) {
            return addDocument();
        }

        var active = app.activeDocument;
        if (!active) {
            return addDocument();
        }

        var hasItems = false;
        try {
            hasItems = active.pageItems && active.pageItems.length > 0;
        } catch (errPageItems) {
            hasItems = false;
        }

        if (hasItems) {
            return addDocument();
        }

        return active;
    }

    function finalizeArtboardAndView(doc, artboardIndex, items, marginCm) {
        if (!doc || artboardIndex < 0) return;
        var workingItems = (items && items.length) ? items : collectAllPageItems(doc);
        cropArtboardAroundItems(doc, artboardIndex, workingItems, marginCm);
        focusArtboard(doc, artboardIndex);
    }

    function collectAllPageItems(doc) {
        if (!doc || doc.pageItems.length === 0) return [];
        var collected = [];
        for (var i = 0; i < doc.pageItems.length; i++) {
            collected.push(doc.pageItems[i]);
        }
        return collected;
    }

    function getItemBounds(item) {
        if (!item) return null;
        var bounds = null;
        try {
            bounds = item.visibleBounds;
        } catch (errVisible) {}
        if (!bounds) {
            try {
                bounds = item.geometricBounds;
            } catch (errGeo) {}
        }
        return bounds;
    }

    function unionBounds(existing, candidate) {
        if (!candidate) return existing;
        if (!existing) return candidate.slice(0);
        return [
            Math.min(existing[0], candidate[0]),
            Math.max(existing[1], candidate[1]),
            Math.max(existing[2], candidate[2]),
            Math.min(existing[3], candidate[3])
        ];
    }

    function cropArtboardAroundItems(doc, artboardIndex, items, marginCm) {
        if (!doc || artboardIndex < 0 || artboardIndex >= doc.artboards.length) return;
        if (!items || !items.length) return;
        var bounds = null;
        for (var i = 0; i < items.length; i++) {
            var itemBounds = getItemBounds(items[i]);
            bounds = unionBounds(bounds, itemBounds);
        }
        if (!bounds) return;
        var marginPt = cm(marginCm);
        var rect = [
            bounds[0] - marginPt,
            bounds[1] + marginPt,
            bounds[2] + marginPt,
            bounds[3] - marginPt
        ];
        try {
            doc.artboards[artboardIndex].artboardRect = rect;
        } catch (errCrop) {}
    }

    function focusArtboard(doc, artboardIndex) {
        if (!doc || artboardIndex < 0) return;
        try {
            doc.artboards.setActiveArtboardIndex(artboardIndex);
        } catch (errActive) {}
        try {
            doc.fitArtboardToWindow(artboardIndex);
        } catch (errFitWindow) {}
        try {
            app.executeMenuCommand("fitartboardinwindow");
        } catch (errMenuFitArtboard) {}
        try {
            app.executeMenuCommand("fitall");
        } catch (errMenuFitAll) {}
    }

    finalizeArtboardAndView(doc, artboardIndex, draftItems, ARTBOARD_MARGIN_CM);
})();
