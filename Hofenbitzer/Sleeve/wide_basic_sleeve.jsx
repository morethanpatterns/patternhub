// Wide Basic Sleeve - default size 38
(function () {
    var SHOW_MEASUREMENT_DIALOG = true;

    var CM_TO_PT = 28.346456692913385;
    var DEFAULT_MARGIN_CM = 10;
    var MARKER_RADIUS_CM = 0.25;
    var DASH_PATTERN_PT = [25, 12];
    var LABEL_OFFSET_CM = 0.5;
    var LABEL_FONT_PT = 12;
    var ARC_TARGET_LENGTH_CM = 20;
    var ARC_MAX_LENGTH_CM = 22;
    var ARC_MIN_SWEEP_RAD = 0.01;
    var ARC_MIN_START_ANGLE = -Math.PI / 2;
    var ARC_END_ANGLE = Math.PI / 2;
    var ARC_MARK4_UP_RATIO = 3;
    var ARC_MARK4_DOWN_RATIO = 2;
    var HANDLE_START_RATIO = 0.22;
    var HANDLE_MID_RATIO = 0.28;
    var HANDLE_END_RATIO = 0.24;
    var HANDLE_PERP_RATIO = 0.05;
    var HANDLE_START_MIN_CM = 1.2;
    var HANDLE_START_MAX_CM = 5.5;
    var HANDLE_MID_MIN_CM = 1.2;
    var HANDLE_MID_MAX_CM = 6;
    var HANDLE_END_MIN_CM = 1.2;
    var HANDLE_END_MAX_CM = 5;
    var HANDLE_PERP_MIN_CM = 0.2;
    var HANDLE_PERP_MAX_CM = 0.2;
    var measurementResults = {};
    var easeResults = {};
    var finalResults = {};
    var measurementLabelMap = {};
    var measurementPalette = null;
    var measurementPaletteGlobalKey = "wideBasicSleeveMeasurementPalette";

    function rgb(r, g, b) {
        var color = new RGBColor();
        color.red = r;
        color.green = g;
        color.blue = b;
        return color;
    }
    var COL_BLACK = rgb(0, 0, 0);
    var COL_WHITE = rgb(255, 255, 255);

    var params = {
        AhH: null,
        AhHEase: 0,
        fAh: 19.6,
        bAh: 23.9,
        AhC: null,
        fAhEase: 0,
        bAhEase: 0,
        AhCEase: 0,
        AL: 60,
        ALEase: 0,
        upAC: 28,
        upACEase: 9,
        WrC: 16,
        WrCEase: 6,
        CapEasePct: 3,
        CapEasePctEase: 0,
        CapCEase: 0,
        CapLineEase: 0,
        fAP: 4.5,
        bAP: 8.6,
        showMeasurementSummary: false
    };

    if (SHOW_MEASUREMENT_DIALOG) {
        var dialogResult = showMeasurementDialog(params);
        if (!dialogResult) return;
        params = dialogResult;
    }

    var derived = computeDerived(params);
    updateMeasurementSummaryData(params, derived);
    if (params.showMeasurementSummary) {
        showMeasurementPaletteWindow("Wide Basic Sleeve");
    } else {
        closeMeasurementPalette();
    }
    var layout = buildLayout(derived);
    var draftingContext = prepDocument(layout);
    if (draftingContext) {
        draftWideBasicSleeve(draftingContext, derived);
    }
    return;

    /* ----------------------------
        Document preparation
    ---------------------------- */
    function prepDocument(layout) {
        if (!layout) return null;
        var doc = getOrCreateDocument(layout.widthPt, layout.heightPt);
        doc.activate();
        resizeActiveArtboard(doc, layout.widthPt, layout.heightPt);

        var layerHandles = ensureBaseLayers(doc);
        removeExtraneousLayers(doc, {
            "Basic Frame": true,
            "Cap Shaping": true,
            "Sleeve Block": true,
            "Numbers, Markers & Labels": true,
            "Markers": true,
            "Labels": true,
            "Numbers": true
        });
        emptyLayer(layerHandles.basicFrame);
        emptyLayer(layerHandles.capShaping);
        emptyLayer(layerHandles.sleeveBlock);
        emptyLayer(layerHandles.markers);
        emptyLayer(layerHandles.labels);
        emptyLayer(layerHandles.numbers);

        return {
            doc: doc,
            layers: layerHandles,
            layout: layout
        };
    }

    function buildLayout(derived) {
        var marginCm = DEFAULT_MARGIN_CM;
        var bounds = computeContentBounds(derived);
        var contentWidthCm = Math.max(bounds.widthCm, 1);
        var contentHeightCm = Math.max(bounds.heightCm, 1);
        var totalWidthCm = contentWidthCm + marginCm * 2;
        var totalHeightCm = contentHeightCm + marginCm * 2;

        return {
            widthPt: cm(totalWidthCm),
            heightPt: cm(totalHeightCm),
            marginPt: cm(marginCm),
            contentWidthPt: cm(contentWidthCm),
            contentHeightPt: cm(contentHeightCm),
            bounds: bounds
        };
    }

    function getOrCreateDocument(widthPt, heightPt) {
        var doc;
        if (app.documents.length === 0 || (app.activeDocument && app.activeDocument.pageItems.length > 0)) {
            doc = app.documents.add(DocumentColorSpace.RGB, widthPt, heightPt);
        } else {
            doc = app.activeDocument;
        }
        return doc;
    }

    function resizeActiveArtboard(doc, widthPt, heightPt) {
        try {
            doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect = [0, heightPt, widthPt, 0];
        } catch (e) {
            // Ignore sizing failures
        }
    }

    function ensureBaseLayers(doc) {
        var annotationLayer = ensureLayer(doc, "Numbers, Markers & Labels");
        return {
            basicFrame: ensureLayer(doc, "Basic Frame"),
            capShaping: ensureLayer(doc, "Cap Shaping"),
            sleeveBlock: ensureLayer(doc, "Sleeve Block"),
            markers: ensureLayer(doc, "Markers", annotationLayer),
            labels: ensureLayer(doc, "Labels", annotationLayer),
            numbers: ensureLayer(doc, "Numbers", annotationLayer),
            annotationsParent: annotationLayer
        };
    }

    function removeExtraneousLayers(doc, keepMap) {
        keepMap = keepMap || {};
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var layer = doc.layers[i];
            var name = layer.name || "";
            if (!keepMap[name]) {
                try {
                    unlockLayerHierarchy(layer);
                    layer.remove();
                } catch (e) {}
            }
        }
    }

    function ensureLayer(doc, name, parentLayer) {
        var container = parentLayer ? parentLayer.layers : doc.layers;
        var layer;
        try {
            layer = container.getByName(name);
        } catch (e) {
            layer = container.add();
            layer.name = name;
        }
        layer.visible = true;
        layer.locked = false;
        return layer;
    }

    function emptyLayer(layer) {
        if (!layer) return;
        unlockLayerHierarchy(layer);
        deleteLayerContents(layer);
    }

    function unlockLayerHierarchy(layer) {
        try { layer.locked = false; } catch (e) {}
        try { layer.visible = true; } catch (e) {}
        for (var i = 0; i < layer.layers.length; i++) {
            unlockLayerHierarchy(layer.layers[i]);
        }
    }

    function deleteLayerContents(layer) {
        if (!layer) return;
        var subLayers = layer.layers;
        for (var i = subLayers.length - 1; i >= 0; i--) {
            deleteLayerContents(subLayers[i]);
        }
        var items = layer.pageItems;
        for (var j = items.length - 1; j >= 0; j--) {
            try { items[j].remove(); } catch (e) {}
        }
    }

    function ensureLayerWritable(layer) {
        if (!layer) return;
        try { layer.locked = false; } catch (e) {}
        try { layer.visible = true; } catch (e) {}
    }

    function draftWideBasicSleeve(context, derived) {
        if (!context || !context.doc || !derived) return;
        var layers = context.layers || {};
        var marginPt = (context.layout && context.layout.marginPt) ? context.layout.marginPt : cm(DEFAULT_MARGIN_CM);

        var artboard = context.doc.artboards[context.doc.artboards.getActiveArtboardIndex()];
        var rect = artboard.artboardRect;
        var artLeft = rect[0];
        var artTop = rect[1];

        var bounds = (context.layout && context.layout.bounds) ? context.layout.bounds : computeContentBounds(derived);
        var leftExtentPt = cm(bounds.leftExtentCm || 0);
        var topExtentPt = cm(bounds.topExtentCm || 0);

        var baselineLengthCm = (derived.SlW || 0) + (derived.CapLineEase || 0);
        if (!isFinite(baselineLengthCm) || baselineLengthCm <= 0) baselineLengthCm = 1;
        var centerX = artLeft + marginPt + leftExtentPt;
        var centerY = artTop - marginPt - topExtentPt;

        var mark1 = [centerX, centerY];
        var mark2 = [mark1[0] + cm(baselineLengthCm), mark1[1]];

        derived.CapLineCm = baselineLengthCm;
        var capLineLabel = "Cap Line (" + formatNumber(baselineLengthCm) + " cm)";
        var capLineOptions = {
            name: "Cap Line",
            labelLayer: layers.labels,
            labelText: capLineLabel,
            labelOffsetCm: LABEL_OFFSET_CM
        };
        drawDashedLine(layers.basicFrame, mark1, mark2, capLineOptions);
        if (layers.sleeveBlock) {
            drawDashedLine(layers.sleeveBlock, mark1, mark2, { name: capLineOptions.name });
        }

        placeMarker(layers.markers, layers.numbers, mark1, 1);
        placeMarker(layers.markers, layers.numbers, mark2, 2);

        var upwardArc = buildMark3Arc(mark1, derived, bounds, ARC_TARGET_LENGTH_CM);
        var mark4 = upwardArc ? buildMark4Intersection(mark2, upwardArc, derived) : null;

        if ((!mark4 || !upwardArc) && ARC_MAX_LENGTH_CM > ARC_TARGET_LENGTH_CM) {
            var extendedArc = buildMark3Arc(mark1, derived, bounds, ARC_MAX_LENGTH_CM);
            var extendedMark4 = extendedArc ? buildMark4Intersection(mark2, extendedArc, derived) : null;
            if (extendedArc && extendedMark4) {
                upwardArc = extendedArc;
                mark4 = extendedMark4;
            }
        }

        var mark4Point = mark4 ? mark4.point : null;
        if (mark4 && upwardArc) {
            rebalanceArcAroundMark4(upwardArc, mark4.angle);
        }
        if (mark4Point) {
            var backLineCm = pointDistanceCm(mark2, mark4Point);
            var frontLineCm = pointDistanceCm(mark1, mark4Point);
            drawDashedLine(
                layers.basicFrame,
                mark2,
                mark4Point,
                { name: "Back Line" }
            );
            drawDashedLine(
                layers.basicFrame,
                mark1,
                mark4Point,
                { name: "Front Line" }
            );
            placeMarker(layers.markers, layers.numbers, mark4Point, 4);
            buildVerticalAndSquares(layers, mark1, mark2, mark4Point, derived);
            addCapMarkers(layers.capShaping, layers.markers, layers.numbers, mark1, mark2, mark4Point, derived);
            addPerpendicularCapPoints(layers.capShaping, layers.markers, layers.numbers, mark1, mark2, mark4Point, derived);
            addCapArcs(layers, mark1, mark2, mark4Point, derived);
        }

        if (upwardArc) {
            drawArcPath(layers.basicFrame, upwardArc);
            placeMarker(layers.markers, layers.numbers, upwardArc.endPoint, 3);
        }
    }

    function drawDashedLine(targetLayer, startPt, endPt, options) {
        if (!targetLayer || !startPt || !endPt) return null;
        ensureLayerWritable(targetLayer);
        var path = targetLayer.pathItems.add();
        path.stroked = true;
        path.strokeWidth = 1;
        path.strokeColor = COL_BLACK;
        path.strokeDashes = DASH_PATTERN_PT.slice(0);
        path.strokeCap = StrokeCap.BUTTENDCAP;
        path.strokeJoin = StrokeJoin.MITERENDJOIN;
        path.filled = false;
        path.closed = false;
        path.setEntirePath([startPt, endPt]);
        if (options && options.name) {
            try { path.name = options.name; } catch (eName) {}
        }
        if (options && options.labelLayer && options.labelText) {
            ensureLayerWritable(options.labelLayer);
            var tf = options.labelLayer.textFrames.add();
            tf.contents = options.labelText;
            try { tf.textRange.characterAttributes.size = LABEL_FONT_PT; } catch (eSize) {}
            try { tf.textRange.characterAttributes.fillColor = COL_BLACK; } catch (eFill) {}
            var dx = endPt[0] - startPt[0];
            var dy = endPt[1] - startPt[1];
            var len = Math.sqrt(dx * dx + dy * dy);
            var offsetPt = cm(options.labelOffsetCm || 0);
            var midPoint = [
                (startPt[0] + endPt[0]) / 2,
                (startPt[1] + endPt[1]) / 2
            ];
            if (len > 0 && offsetPt !== 0) {
                var nx = -dy / len;
                var ny = dx / len;
                var normalSign = 1;
                if (options && typeof options.labelNormalSign === "number") {
                    normalSign = options.labelNormalSign >= 0 ? 1 : -1;
                } else if (ny < 0) {
                    normalSign = -1;
                }
                nx *= normalSign;
                ny *= normalSign;
                midPoint[0] += nx * offsetPt;
                midPoint[1] += ny * offsetPt;
            } else {
                midPoint[1] += offsetPt;
            }
            centerTextFrame(tf, midPoint);
            var angleDeg = Math.atan2(dy, dx) * 180 / Math.PI;
            if (angleDeg > 90) angleDeg -= 180;
            if (angleDeg < -90) angleDeg += 180;
            try {
                tf.rotate(angleDeg, true, true, true, true, Transformation.CENTER);
            } catch (eRot) {
                try { tf.rotate(angleDeg); } catch (ignored) {}
            }
        }
        return path;v
    }

    function drawSolidLine(targetLayer, startPt, endPt, options) {
        if (!targetLayer || !startPt || !endPt) return null;
        ensureLayerWritable(targetLayer);
        var path = targetLayer.pathItems.add();
        path.stroked = true;
        path.strokeWidth = (options && options.strokeWidth) || 1;
        path.strokeColor = COL_BLACK;
        path.strokeDashes = [];
        path.strokeCap = StrokeCap.BUTTENDCAP;
        path.strokeJoin = StrokeJoin.MITERENDJOIN;
        path.filled = false;
        path.closed = false;
        path.setEntirePath([startPt, endPt]);
        if (options && options.name) {
            try { path.name = options.name; } catch (eName) {}
        }
        if (options && options.labelLayer && options.labelText) {
            ensureLayerWritable(options.labelLayer);
            var tf = options.labelLayer.textFrames.add();
            tf.contents = options.labelText;
            try { tf.textRange.characterAttributes.size = LABEL_FONT_PT; } catch (eSize) {}
            try { tf.textRange.characterAttributes.fillColor = COL_BLACK; } catch (eFill) {}
            var dx = endPt[0] - startPt[0];
            var dy = endPt[1] - startPt[1];
            var len = Math.sqrt(dx * dx + dy * dy);
            var offsetPt = cm(options.labelOffsetCm || 0);
            var midPoint = [
                (startPt[0] + endPt[0]) / 2,
                (startPt[1] + endPt[1]) / 2
            ];
            if (len > 0 && offsetPt !== 0) {
                var nx = -dy / len;
                var ny = dx / len;
                var normalSign = 1;
                if (options && typeof options.labelNormalSign === "number") {
                    normalSign = options.labelNormalSign >= 0 ? 1 : -1;
                } else if (ny < 0) {
                    normalSign = -1;
                }
                nx *= normalSign;
                ny *= normalSign;
                midPoint[0] += nx * offsetPt;
                midPoint[1] += ny * offsetPt;
            } else {
                midPoint[1] += offsetPt;
            }
            centerTextFrame(tf, midPoint);
            var angleDeg = Math.atan2(dy, dx) * 180 / Math.PI;
            if (angleDeg > 90) angleDeg -= 180;
            if (angleDeg < -90) angleDeg += 180;
            try {
                tf.rotate(angleDeg, true, true, true, true, Transformation.CENTER);
            } catch (eRot) {
                try { tf.rotate(angleDeg); } catch (ignored) {}
            }
        }
        return path;
    }

    function pointDistanceCm(ptA, ptB) {
        if (!ptA || !ptB) return null;
        var dx = ptA[0] - ptB[0];
        var dy = ptA[1] - ptB[1];
        var distPt = Math.sqrt(dx * dx + dy * dy);
        if (!isFinite(distPt) || distPt <= 0) return null;
        return distPt / CM_TO_PT;
    }

    function buildLineLabelOptions(name, lengthCm, labelsLayer) {
        if (!labelsLayer) return { name: name };
        var value = isFinite(lengthCm) && lengthCm > 0 ? formatNumber(lengthCm) : "-";
        return {
            name: name,
            labelLayer: labelsLayer,
            labelText: name + " (" + value + " cm)",
            labelOffsetCm: LABEL_OFFSET_CM
        };
    }


    function buildVerticalAndSquares(layers, mark1, mark2, mark4Point, derived) {
        if (!layers || !mark1 || !mark2 || !mark4Point || !derived) return;
        var sleeveBlockLayer = layers.sleeveBlock || null;
        var sleeveLengthCm = derived.SlL || 0;
        if (!isFinite(sleeveLengthCm) || sleeveLengthCm <= 0) return;
        var sleeveLengthPt = cm(sleeveLengthCm);
        var verticalEnd = [mark4Point[0], mark4Point[1] - sleeveLengthPt];
        var sleeveLengthLabel = { name: "Sleeve Length" };
        drawDashedLine(
            layers.basicFrame,
            mark4Point,
            verticalEnd,
            sleeveLengthLabel
        );
        drawDashedLine(
            layers.sleeveBlock,
            mark4Point,
            verticalEnd,
            sleeveLengthLabel
        );
        placeMarker(layers.markers, layers.numbers, verticalEnd, 5);

        var mark6 = [mark4Point[0], mark4Point[1] - sleeveLengthPt * 0.6];
        placeMarker(layers.markers, layers.numbers, mark6, 6);
        var mid64 = midpoint(mark6, mark4Point);
        var sleeveWidthY = mid64 ? mid64[1] : null;

        var mark5 = verticalEnd;
        var mark1Down5 = [mark1[0], mark5[1]];
        var mark2Down5 = [mark2[0], mark5[1]];
        var mark1Down6 = [mark1[0], mark6[1]];
        var mark2Down6 = [mark2[0], mark6[1]];
        var elbowLineY = mark6 ? mark6[1] : (mark1Down6 ? mark1Down6[1] : null);

        drawDashedLine(layers.basicFrame, mark1, mark1Down5, { name: "1-7" });
        drawDashedLine(layers.basicFrame, mark2, mark2Down5, { name: "2-8" });
        drawDashedLine(layers.basicFrame, mark1Down5, mark2Down5, { name: "7-8" });
        drawDashedLine(layers.basicFrame, mark1Down6, mark2Down6, { name: "9-10" });

        var hemOptions = { name: "Hem Line", labelLayer: layers.labels, labelText: "Hem Line", labelOffsetCm: LABEL_OFFSET_CM * 2 };
        var hemLineStart = mark1Down5;
        var hemLineEnd = mark2Down5;
        var hemLineStartSleeve = mark1Down5;
        var hemLineEndSleeve = mark2Down5;

        var elbowLabel = buildLineLabelOptions("Elbow Line", derived.SlW, layers.labels);
        var elbowLineStart = null;
        var elbowLineEnd = null;
        var sleeveWidthStart = null;
        var sleeveWidthEnd = null;
        var leftSleeveCurve = null;
        var rightSleeveCurve = null;

        var mark15 = midpoint(mark1Down5, mark2Down5);
        var leftElbowIntersection = null;
        var rightElbowIntersection = null;
        if (mark15) {
            placeMarker(layers.markers, layers.numbers, mark15, 15);
            var halfHeW = (derived.HeW || 0) / 2;
            if (isFinite(halfHeW) && halfHeW > 0) {
                var halfHeWPts = cm(halfHeW);
                var left16 = [mark15[0] - halfHeWPts, mark15[1]];
                var right17 = [mark15[0] + halfHeWPts, mark15[1]];
                hemLineStart = left16;
                hemLineEnd = right17;
                hemLineStartSleeve = left16;
                hemLineEndSleeve = right17;
                drawDashedLine(layers.basicFrame, left16, mark1, {
                    name: "Left Sleeve Length"
                });
                if (elbowLineY !== null) {
                    if (!leftSleeveCurve) {
                        var leftStartHandle = [left16[0], left16[1] + cm(5)];
                        var dir1To16 = normalizeVector([left16[0] - mark1[0], left16[1] - mark1[1]]);
                        var mark1Handle = [
                            mark1[0] + dir1To16[0] * cm(12.47),
                            mark1[1] + dir1To16[1] * cm(12.47)
                        ];
                        leftSleeveCurve = {
                            start: left16,
                            startHandle: leftStartHandle,
                            endHandle: mark1Handle,
                            end: mark1
                        };
                    }
                    var leftIntersection = findBezierHorizontalIntersection(
                        leftSleeveCurve.start,
                        leftSleeveCurve.startHandle,
                        leftSleeveCurve.endHandle,
                        leftSleeveCurve.end,
                        elbowLineY
                    );
                    if (!leftIntersection) {
                        leftIntersection = findHorizontalIntersection(leftSleeveCurve.start, leftSleeveCurve.end, elbowLineY);
                    }
                    if (!leftIntersection) {
                        leftIntersection = projectHorizontalOntoLine(leftSleeveCurve.start, leftSleeveCurve.end, elbowLineY);
                    }
                    leftElbowIntersection = leftIntersection;
                }
                drawDashedLine(layers.basicFrame, right17, mark2, {
                    name: "Right Sleeve Length"
                });
                if (elbowLineY !== null) {
                    if (!rightSleeveCurve) {
                        var rightStartHandle = [right17[0], right17[1] + cm(5)];
                        var dir2To17 = normalizeVector([right17[0] - mark2[0], right17[1] - mark2[1]]);
                        var mark2Handle = [
                            mark2[0] + dir2To17[0] * cm(12.47),
                            mark2[1] + dir2To17[1] * cm(12.47)
                        ];
                        rightSleeveCurve = {
                            start: right17,
                            startHandle: rightStartHandle,
                            endHandle: mark2Handle,
                            end: mark2
                        };
                    }
                    var rightIntersection = findBezierHorizontalIntersection(
                        rightSleeveCurve.start,
                        rightSleeveCurve.startHandle,
                        rightSleeveCurve.endHandle,
                        rightSleeveCurve.end,
                        elbowLineY
                    );
                    if (!rightIntersection) {
                        rightIntersection = findHorizontalIntersection(rightSleeveCurve.start, rightSleeveCurve.end, elbowLineY);
                    }
                    if (!rightIntersection) {
                        rightIntersection = projectHorizontalOntoLine(rightSleeveCurve.start, rightSleeveCurve.end, elbowLineY);
                    }
                    rightElbowIntersection = rightIntersection;
                }
                placeMarker(layers.markers, layers.numbers, left16, 16);
                placeMarker(layers.markers, layers.numbers, right17, 17);

            }
            if (typeof sleeveWidthY === "number" && left16 && right17) {
                if (!leftSleeveCurve) {
                    var leftStartHandleWidth = [left16[0], left16[1] + cm(5)];
                    var dirWidth1 = normalizeVector([left16[0] - mark1[0], left16[1] - mark1[1]]);
                    var mark1HandleWidth = [
                        mark1[0] + dirWidth1[0] * cm(12.47),
                        mark1[1] + dirWidth1[1] * cm(12.47)
                    ];
                    leftSleeveCurve = {
                        start: left16,
                        startHandle: leftStartHandleWidth,
                        endHandle: mark1HandleWidth,
                        end: mark1
                    };
                }
                if (!rightSleeveCurve) {
                    var rightStartHandleWidth = [right17[0], right17[1] + cm(5)];
                    var dirWidth2 = normalizeVector([right17[0] - mark2[0], right17[1] - mark2[1]]);
                    var mark2HandleWidth = [
                        mark2[0] + dirWidth2[0] * cm(12.47),
                        mark2[1] + dirWidth2[1] * cm(12.47)
                    ];
                    rightSleeveCurve = {
                        start: right17,
                        startHandle: rightStartHandleWidth,
                        endHandle: mark2HandleWidth,
                        end: mark2
                    };
                }
                var leftWidthIntersection = findBezierHorizontalIntersection(
                    leftSleeveCurve.start,
                    leftSleeveCurve.startHandle,
                    leftSleeveCurve.endHandle,
                    leftSleeveCurve.end,
                    sleeveWidthY
                );
                if (!leftWidthIntersection) {
                    leftWidthIntersection = findHorizontalIntersection(leftSleeveCurve.start, leftSleeveCurve.end, sleeveWidthY);
                }
                if (!leftWidthIntersection) {
                    leftWidthIntersection = projectHorizontalOntoLine(leftSleeveCurve.start, leftSleeveCurve.end, sleeveWidthY);
                }
                var rightWidthIntersection = findBezierHorizontalIntersection(
                    rightSleeveCurve.start,
                    rightSleeveCurve.startHandle,
                    rightSleeveCurve.endHandle,
                    rightSleeveCurve.end,
                    sleeveWidthY
                );
                if (!rightWidthIntersection) {
                    rightWidthIntersection = findHorizontalIntersection(rightSleeveCurve.start, rightSleeveCurve.end, sleeveWidthY);
                }
                if (!rightWidthIntersection) {
                    rightWidthIntersection = projectHorizontalOntoLine(rightSleeveCurve.start, rightSleeveCurve.end, sleeveWidthY);
                }
                if (leftWidthIntersection && rightWidthIntersection) {
                    sleeveWidthStart = leftWidthIntersection;
                    sleeveWidthEnd = rightWidthIntersection;
                }
            }
        }
        if (leftElbowIntersection) elbowLineStart = leftElbowIntersection;
        if (rightElbowIntersection) elbowLineEnd = rightElbowIntersection;

        placeMarker(layers.markers, layers.numbers, mark1Down5, 7);
        placeMarker(layers.markers, layers.numbers, mark2Down5, 8);
        placeMarker(layers.markers, layers.numbers, mark1Down6, 9);
        placeMarker(layers.markers, layers.numbers, mark2Down6, 10);

        drawSolidLine(sleeveBlockLayer || layers.basicFrame, hemLineStartSleeve, hemLineEndSleeve, hemOptions);

        if (sleeveBlockLayer && leftSleeveCurve) {
            var leftArc = drawBezierSegment(
                sleeveBlockLayer,
                leftSleeveCurve.start,
                leftSleeveCurve.end,
                leftSleeveCurve.startHandle,
                leftSleeveCurve.endHandle,
                "Left Sleeve Line"
            );
            if (leftArc) {
                try { leftArc.strokeDashes = []; } catch (eLeftArc) {}
            }
        }
        if (sleeveBlockLayer && rightSleeveCurve) {
            var rightArc = drawBezierSegment(
                sleeveBlockLayer,
                rightSleeveCurve.start,
                rightSleeveCurve.end,
                rightSleeveCurve.startHandle,
                rightSleeveCurve.endHandle,
                "Right Sleeve Line"
            );
            if (rightArc) {
                try { rightArc.strokeDashes = []; } catch (eRightArc) {}
            }
        }

        if (!elbowLineStart) elbowLineStart = mark1Down6;
        if (!elbowLineEnd) elbowLineEnd = mark2Down6;
        if (sleeveBlockLayer && elbowLineStart && elbowLineEnd) {
            drawDashedLine(sleeveBlockLayer, elbowLineStart, elbowLineEnd, elbowLabel);
        }

        if (sleeveBlockLayer && sleeveWidthStart && sleeveWidthEnd) {
            var sleeveWidthCm = pointDistanceCm(sleeveWidthStart, sleeveWidthEnd);
            var sleeveWidthLabel = buildLineLabelOptions("Sleeve Width", sleeveWidthCm, layers.labels);
            drawDashedLine(sleeveBlockLayer, sleeveWidthStart, sleeveWidthEnd, sleeveWidthLabel);
        }
    }


    function computeCapMarkerPoints(mark1, mark2, mark4Point, derived) {
        if (!mark1 || !mark2 || !mark4Point || !derived) return null;
        var capLenCm = derived.CapLineCm || 0;
        var capLenPt = cm(capLenCm);
        if (!isFinite(capLenPt) || capLenPt <= 0) {
            capLenPt = Math.abs((mark2[0] || 0) - (mark1[0] || 0));
        }
        if (!isFinite(capLenPt) || capLenPt <= 0) return null;
        return {
            mark11: [mark4Point[0] - capLenPt / 8, mark4Point[1]],
            mark12: [mark4Point[0] + capLenPt / 5, mark4Point[1]],
            mark13: [mark2[0] - capLenPt / 14, mark2[1]],
            mark14: [mark1[0] + capLenPt / 12, mark1[1]]
        };
    }

    function addCapMarkers(capLayer, markerLayer, numberLayer, mark1, mark2, mark4Point, derived) {
        if (!capLayer || !markerLayer || !numberLayer) return;
        var points = computeCapMarkerPoints(mark1, mark2, mark4Point, derived);
        if (!points) return;
        var segments = [
            { start: mark4Point, end: points.mark11, label: 11, drawLine: true, name: "4-11" },
            { start: mark4Point, end: points.mark12, label: 12, drawLine: true, name: "4-12" },
            { start: mark2, end: points.mark13, label: 13, drawLine: false },
            { start: mark1, end: points.mark14, label: 14, drawLine: false }
        ];

        for (var i = 0; i < segments.length; i++) {
            var seg = segments[i];
            if (!seg.start || !seg.end) continue;
            if (seg.drawLine !== false) {
                drawDashedLine(capLayer, seg.start, seg.end, { name: seg.name || "" });
            }
            placeMarker(markerLayer, numberLayer, seg.end, seg.label);
        }

        if (points.mark11 && points.mark14) {
            drawDashedLine(capLayer, points.mark11, points.mark14, { name: "11-14" });
        }
        if (points.mark12 && points.mark13) {
            drawDashedLine(capLayer, points.mark12, points.mark13, { name: "12-13" });
        }
    }

    function addPerpendicularCapPoints(capLayer, markerLayer, numberLayer, mark1, mark2, mark4Point, derived) {
        if (!capLayer || !markerLayer || !numberLayer) return;
        var points = computeCapMarkerPoints(mark1, mark2, mark4Point, derived);
        if (!points) return;
        addPerpendicularFromLine(capLayer, markerLayer, numberLayer, mark1, mark4Point, points.mark11, "a", "a-11");
        addPerpendicularFromLine(capLayer, markerLayer, numberLayer, mark2, mark4Point, points.mark12, "b", "12-b");
    }

    function addPerpendicularFromLine(capLayer, markerLayer, numberLayer, lineStart, lineEnd, targetPoint, labelChar, pathName) {
        if (!capLayer || !markerLayer || !numberLayer || !lineStart || !lineEnd || !targetPoint) return;
        var projection = computePerpendicularProjection(lineStart, lineEnd, targetPoint);
        if (!projection || !projection.projectPoint) return;
        drawDashedLine(capLayer, projection.projectPoint, targetPoint, { name: pathName || "" });
        placeMarker(markerLayer, numberLayer, projection.projectPoint, labelChar || "");
    }

    function addCapArcs(layerHandles, mark1, mark2, mark4Point, derived) {
        if (!layerHandles || !mark1 || !mark2 || !mark4Point || !derived) return;
        var capLayer = layerHandles.capShaping;
        var sleeveBlockLayer = layerHandles.sleeveBlock;
        if (!capLayer) return;
        var points = computeCapMarkerPoints(mark1, mark2, mark4Point, derived);
        if (!points || !points.mark11 || !points.mark12 || !points.mark14) return;

        var mid1114 = midpoint(points.mark11, points.mark14);

        var baselineDir = normalizeVector([mark2[0] - mark1[0], mark2[1] - mark1[1]]);
        var frontHandleLen = cm(3);
        var arc1StartHandle = [
            mark1[0] + baselineDir[0] * frontHandleLen,
            mark1[1] + baselineDir[1] * frontHandleLen
        ];

        var towards14Dir = normalizeVector([points.mark14[0] - mid1114[0], points.mark14[1] - mid1114[1]]);
        var arc1EndHandle = [
            mid1114[0] + towards14Dir[0] * cm(4),
            mid1114[1] + towards14Dir[1] * cm(4)
        ];

        var frontDirMidTo11 = normalizeVector([points.mark11[0] - mid1114[0], points.mark11[1] - mid1114[1]]);
        var frontSecondStartHandle = [
            mid1114[0] + frontDirMidTo11[0] * cm(3),
            mid1114[1] + frontDirMidTo11[1] * cm(3)
        ];
        var dir4To12 = normalizeVector([mark4Point[0] - points.mark12[0], mark4Point[1] - points.mark12[1]]);
        var frontSecondEndHandle = [
            mark4Point[0] + dir4To12[0] * cm(3.3),
            mark4Point[1] + dir4To12[1] * cm(3.3)
        ];
        var frontNodes = [
            { anchor: mark1, right: arc1StartHandle },
            { anchor: mid1114, left: arc1EndHandle, right: frontSecondStartHandle },
            { anchor: mark4Point, left: frontSecondEndHandle }
        ];
        var frontCapPath = drawBezierChain(capLayer, frontNodes, "Front Cap");
        addNotchOnPath(frontCapPath, mark1, mid1114, arc1StartHandle, arc1EndHandle, derived.fAP || 0, capLayer, "Front Cap Notch");
        if (sleeveBlockLayer) {
            var frontBlockPath = drawBezierChain(sleeveBlockLayer, frontNodes, "Front Cap");
            addNotchOnPath(frontBlockPath, mark1, mid1114, arc1StartHandle, arc1EndHandle, derived.fAP || 0, sleeveBlockLayer, "Front Cap Notch");
        }

        var startDirTo11 = normalizeVector([points.mark11[0] - mid1114[0], points.mark11[1] - mid1114[1]]);
        var arc2EndHandle = [
            mid1114[0] + startDirTo11[0] * cm(3),
            mid1114[1] + startDirTo11[1] * cm(3)
        ];

        var mid1213 = midpoint(points.mark12, points.mark13);
        if (!mid1213) mid1213 = mid1114;
        var intersection1213_24 = lineIntersection(points.mark12, points.mark13, mark2, mark4Point);
        if (intersection1213_24) {
            var baselineForwardDir = normalizeVector([mark2[0] - mark1[0], mark2[1] - mark1[1]]);
            var line1213Dir = normalizeVector([points.mark13[0] - points.mark12[0], points.mark13[1] - points.mark12[1]]);
            if (Math.abs(line1213Dir[1]) < 1e-6) {
                line1213Dir = [0, 1];
            } else if (line1213Dir[1] < 0) {
                line1213Dir = [-line1213Dir[0], -line1213Dir[1]];
            }
            var handleMark4 = [
                mark4Point[0] + baselineForwardDir[0] * cm(6.6),
                mark4Point[1] + baselineForwardDir[1] * cm(6.6)
            ];
            var handleIntersectionUpper = [
                intersection1213_24[0] + line1213Dir[0] * cm(2.95),
                intersection1213_24[1] + line1213Dir[1] * cm(2.95)
            ];

            var handleMark2 = [
                mark2[0] - baselineForwardDir[0] * cm(3.5),
                mark2[1] - baselineForwardDir[1] * cm(3.5)
            ];
            var handleIntersectionLower = [
                intersection1213_24[0] - line1213Dir[0] * cm(1.5),
                intersection1213_24[1] - line1213Dir[1] * cm(1.5)
            ];

            var backNodesCustom = [
                { anchor: mark2, right: handleMark2 },
                { anchor: intersection1213_24, left: handleIntersectionLower, right: handleIntersectionUpper },
                { anchor: mark4Point, left: handleMark4 }
            ];
            var wsBackPath = drawBezierChain(capLayer, backNodesCustom, "Back Cap");
            addBackCapNotches(wsBackPath, backNodesCustom, capLayer, derived);
            if (sleeveBlockLayer) {
                var wsBackBlock = drawBezierChain(sleeveBlockLayer, backNodesCustom, "Back Cap");
                addBackCapNotches(wsBackBlock, backNodesCustom, sleeveBlockLayer, derived);
            }
        } else {
            var baselineBackDir = normalizeVector([mark1[0] - mark2[0], mark1[1] - mark2[1]]);
            var arc2StartHandle = [
                mark2[0] + baselineBackDir[0] * cm(4.3),
                mark2[1] + baselineBackDir[1] * cm(4.3)
            ];
            var dirTo13 = normalizeVector([points.mark13[0] - mid1213[0], points.mark13[1] - mid1213[1]]);
            var arc2EndHandleFallback = [
                mid1213[0] + dirTo13[0] * cm(4),
                mid1213[1] + dirTo13[1] * cm(4)
            ];

            var projectionB = computePerpendicularProjection(mark4Point, mark2, points.mark12);
            var mid12b = projectionB && projectionB.projectPoint ? midpoint(points.mark12, projectionB.projectPoint) : midpoint(points.mark12, mark4Point);
            var dirTo12 = normalizeVector([points.mark12[0] - mid1213[0], points.mark12[1] - mid1213[1]]);
            var dir4To12 = normalizeVector([points.mark12[0] - mark4Point[0], points.mark12[1] - mark4Point[1]]);
            var arc3EndHandle = [
                mark4Point[0] + dir4To12[0] * cm(3.3),
                mark4Point[1] + dir4To12[1] * cm(3.3)
            ];
            var targetPoint = mid12b || midpoint(mid1213, mark4Point);
            var f0 = 0.125;
            var f1 = 0.375;
            var f2 = 0.375;
            var f3 = 0.125;
            var restX = f0 * mid1213[0] + f2 * arc3EndHandle[0] + f3 * mark4Point[0];
            var restY = f0 * mid1213[1] + f2 * arc3EndHandle[1] + f3 * mark4Point[1];
            var vecX = targetPoint[0] - restX - f1 * mid1213[0];
            var vecY = targetPoint[1] - restY - f1 * mid1213[1];
            var startLenPt = (vecX * dirTo12[0] + vecY * dirTo12[1]) / f1;
            if (!isFinite(startLenPt) || startLenPt <= 0) startLenPt = cm(3);
            var arc3StartHandleFallback = [
                mid1213[0] + dirTo12[0] * startLenPt,
                mid1213[1] + dirTo12[1] * startLenPt
            ];

            var backNodesFallback = [
                { anchor: mark2, right: arc2StartHandle },
                { anchor: mid1213, left: arc2EndHandleFallback, right: arc3StartHandleFallback },
                { anchor: mark4Point, left: arc3EndHandle }
            ];
            var backCapPath = drawBezierChain(capLayer, backNodesFallback, "Back Cap");
            addBackCapNotches(backCapPath, backNodesFallback, capLayer, derived);
            if (sleeveBlockLayer) {
                var backBlockPath = drawBezierChain(sleeveBlockLayer, backNodesFallback, "Back Cap");
                addBackCapNotches(backBlockPath, backNodesFallback, sleeveBlockLayer, derived);
            }
        }
    }

    function computePerpendicularProjection(lineStart, lineEnd, targetPoint) {
        if (!lineStart || !lineEnd || !targetPoint) return null;
        var vx = lineEnd[0] - lineStart[0];
        var vy = lineEnd[1] - lineStart[1];
        var lenSq = vx * vx + vy * vy;
        if (!isFinite(lenSq) || lenSq <= 0) return null;
        var wx = targetPoint[0] - lineStart[0];
        var wy = targetPoint[1] - lineStart[1];
        var t = (wx * vx + wy * vy) / lenSq;
        var clampedT = Math.max(0, Math.min(1, t));
        var proj = [lineStart[0] + vx * clampedT, lineStart[1] + vy * clampedT];
        return {
            projectPoint: proj,
            t: clampedT
        };
    }

    function lineIntersection(p1, p2, p3, p4) {
        if (!p1 || !p2 || !p3 || !p4) return null;
        var x1 = p1[0], y1 = p1[1];
        var x2 = p2[0], y2 = p2[1];
        var x3 = p3[0], y3 = p3[1];
        var x4 = p4[0], y4 = p4[1];
        var denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4);
        if (Math.abs(denom) < 1e-6) return null;
        var det1 = x1 * y2 - y1 * x2;
        var det2 = x3 * y4 - y3 * x4;
        var x = (det1 * (x3 - x4) - (x1 - x2) * det2) / denom;
        var y = (det1 * (y3 - y4) - (y1 - y2) * det2) / denom;
        if (!isFinite(x) || !isFinite(y)) return null;
        return [x, y];
    }

    function drawBezierSegment(layer, startPt, endPt, startHandlePt, endHandlePt, name) {
        if (!layer || !startPt || !endPt) return;
        ensureLayerWritable(layer);
        var path = layer.pathItems.add();
        if (name) {
            try { path.name = name; } catch (eName) {}
        }
        path.stroked = true;
        path.strokeWidth = 1;
        path.strokeColor = COL_BLACK;
        path.strokeDashes = [];
        path.strokeCap = StrokeCap.BUTTENDCAP;
        path.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        path.filled = false;
        path.closed = false;
        path.setEntirePath([startPt, endPt]);
        var startPoint = path.pathPoints[0];
        var endPoint = path.pathPoints[path.pathPoints.length - 1];
        startPoint.anchor = startPt;
        startPoint.leftDirection = startPt;
        startPoint.rightDirection = startHandlePt || startPt;
        endPoint.anchor = endPt;
        endPoint.leftDirection = endHandlePt || endPt;
        endPoint.rightDirection = endPt;
        return path;
    }

    function drawBezierChain(layer, nodes, name) {
        if (!layer || !nodes || nodes.length < 2) return null;
        ensureLayerWritable(layer);
        var path = layer.pathItems.add();
        if (name) {
            try { path.name = name; } catch (eName) {}
        }
        path.stroked = true;
        path.strokeWidth = 1;
        path.strokeColor = COL_BLACK;
        path.strokeDashes = [];
        path.strokeCap = StrokeCap.BUTTENDCAP;
        path.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        path.filled = false;
        path.closed = false;
        var anchors = [];
        for (var i = 0; i < nodes.length; i++) {
            anchors.push(nodes[i] && nodes[i].anchor ? nodes[i].anchor : [0, 0]);
        }
        path.setEntirePath(anchors);
        for (var p = 0; p < path.pathPoints.length; p++) {
            var node = nodes[p] || {};
            var anchor = node.anchor || anchors[p];
            var left = node.left || anchor;
            var right = node.right || anchor;
            var pp = path.pathPoints[p];
            pp.anchor = anchor;
            pp.leftDirection = (p === 0) ? anchor : left;
            pp.rightDirection = (p === nodes.length - 1) ? anchor : right;
        }
        return path;
    }

    function midpoint(ptA, ptB) {
        if (!ptA || !ptB) return null;
        return [
            (ptA[0] + ptB[0]) / 2,
            (ptA[1] + ptB[1]) / 2
        ];
    }

    function offsetPointTowards(point, targetPoint, distanceCm) {
        if (!point || !targetPoint) return point;
        var dir = normalizeVector([targetPoint[0] - point[0], targetPoint[1] - point[1]]);
        var distPt = cm(isFinite(distanceCm) ? distanceCm : 0);
        return [
            point[0] + dir[0] * distPt,
            point[1] + dir[1] * distPt
        ];
    }

    function drawSmoothMidpointPath(layer, startPt, midPt, endPt, options) {
        if (!layer || !startPt || !midPt || !endPt) return null;
        ensureLayerWritable(layer);
        var path = layer.pathItems.add();
        if (options && options.name) {
            try { path.name = options.name; } catch (eName) {}
        }
        path.stroked = true;
        path.strokeWidth = (options && options.strokeWidth) || 1;
        path.strokeColor = COL_BLACK;
        path.strokeDashes = [];
        path.strokeCap = StrokeCap.BUTTENDCAP;
        path.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        path.filled = false;
        path.closed = false;
        path.setEntirePath([startPt, midPt, endPt]);
        var points = path.pathPoints;
        if (!points || points.length !== 3) return path;
        var startPoint = points[0];
        startPoint.pointType = PointType.CORNER;
        startPoint.leftDirection = (options && options.startLeft) ? options.startLeft : startPt;
        startPoint.rightDirection = (options && options.startRight) ? options.startRight : startPt;

        var midPoint = points[1];
        midPoint.pointType = PointType.SMOOTH;
        midPoint.leftDirection = (options && options.midLeft) ? options.midLeft : midPt;
        midPoint.rightDirection = (options && options.midRight) ? options.midRight : midPt;

        var endPoint = points[2];
        endPoint.pointType = PointType.CORNER;
        endPoint.leftDirection = (options && options.endLeft) ? options.endLeft : endPt;
        endPoint.rightDirection = (options && options.endRight) ? options.endRight : endPt;
        return path;
    }

    function findHorizontalIntersection(startPt, endPt, targetY) {
        if (!startPt || !endPt || typeof targetY !== "number" || !isFinite(targetY)) return null;
        var deltaY = endPt[1] - startPt[1];
        if (!isFinite(deltaY) || Math.abs(deltaY) < 0.00001) return null;
        var t = (targetY - startPt[1]) / deltaY;
        if (t < 0 || t > 1) return null;
        return [
            startPt[0] + (endPt[0] - startPt[0]) * t,
            targetY
        ];
    }

    function projectHorizontalOntoLine(startPt, endPt, targetY) {
        if (!startPt || !endPt || typeof targetY !== "number" || !isFinite(targetY)) return null;
        var deltaY = endPt[1] - startPt[1];
        if (!isFinite(deltaY) || Math.abs(deltaY) < 0.00001) return null;
        var t = (targetY - startPt[1]) / deltaY;
        return [
            startPt[0] + (endPt[0] - startPt[0]) * t,
            targetY
        ];
    }

    function findBezierHorizontalIntersection(p0, p1, p2, p3, targetY) {
        if (!p0 || !p1 || !p2 || !p3 || typeof targetY !== "number" || !isFinite(targetY)) return null;
        var steps = 60;
        var prevT = 0;
        var prevPoint = evaluateBezierPoint(p0, p1, p2, p3, prevT);
        var prevDiff = prevPoint ? (prevPoint[1] - targetY) : null;
        if (!isFinite(prevDiff)) return null;
        if (Math.abs(prevDiff) < 0.00001) return prevPoint;
        for (var i = 1; i <= steps; i++) {
            var t = i / steps;
            var point = evaluateBezierPoint(p0, p1, p2, p3, t);
            if (!point) return null;
            var diff = point[1] - targetY;
            if (Math.abs(diff) < 0.00001) return point;
            if (diff === 0 || diff * prevDiff < 0) {
                var lowT = prevT;
                var highT = t;
                var lowDiff = prevDiff;
                var highDiff = diff;
                for (var iter = 0; iter < 20; iter++) {
                    var midT = (lowT + highT) / 2;
                    var midPoint = evaluateBezierPoint(p0, p1, p2, p3, midT);
                    if (!midPoint) break;
                    var midDiff = midPoint[1] - targetY;
                    if (Math.abs(midDiff) < 0.00001) return midPoint;
                    if (midDiff * lowDiff < 0) {
                        highT = midT;
                        highDiff = midDiff;
                    } else {
                        lowT = midT;
                        lowDiff = midDiff;
                    }
                }
                return evaluateBezierPoint(p0, p1, p2, p3, (lowT + highT) / 2);
            }
            prevT = t;
            prevDiff = diff;
        }
        return null;
    }

    function normalizeVector(vec) {
        if (!vec) return [1, 0];
        var len = Math.sqrt(vec[0] * vec[0] + vec[1] * vec[1]);
        if (!isFinite(len) || len === 0) return [1, 0];
        return [vec[0] / len, vec[1] / len];
    }

    function perpendicularVector(vec) {
        if (!vec) return [0, 0];
        return [-vec[1], vec[0]];
    }

    function ensureRightFacing(vec) {
        if (!vec) return [1, 0];
        if (vec[0] < 0 || (Math.abs(vec[0]) < 0.00001 && vec[1] < 0)) {
            return [-vec[0], -vec[1]];
        }
        return vec;
    }

    function clampValue(value, minValue, maxValue) {
        if (isFinite(minValue)) value = Math.max(value, minValue);
        if (isFinite(maxValue) && maxValue > 0) value = Math.min(value, maxValue);
        return value;
    }

    function computeHandleLength(segmentLenCm, ratio, minCm, maxCm) {
        if (!isFinite(segmentLenCm) || segmentLenCm <= 0 || !isFinite(ratio)) return 0;
        var lenCm = segmentLenCm * ratio;
        return clampValue(lenCm, minCm, maxCm);
    }

    function addScaledVector(point, direction, lengthPt) {
        if (!point || !direction || !isFinite(lengthPt)) return point;
        return [
            point[0] + direction[0] * lengthPt,
            point[1] + direction[1] * lengthPt
        ];
    }

    function computeHandleForPoint(p0, p1, p3, target) {
        if (!p0 || !p1 || !p3 || !target) return null;
        var t = 0.5;
        var u = 1 - t;
        var f0 = u * u * u;
        var f1 = 3 * u * u * t;
        var f2 = 3 * u * t * t;
        var f3 = t * t * t;
        if (f2 === 0) return null;
        var x = target[0] - (f0 * p0[0] + f1 * p1[0] + f3 * p3[0]);
        var y = target[1] - (f0 * p0[1] + f1 * p1[1] + f3 * p3[1]);
        return [x / f2, y / f2];
    }

    function addNotchOnPath(path, startPt, endPt, startHandle, endHandle, distanceCm, targetLayer, notchName, offsetNormalCm) {
        if (!path || !startPt || !endPt || !targetLayer) return null;
        var distancePt = cm(distanceCm || 0);
        if (!isFinite(distancePt) || distancePt <= 0) return null;
        var totalLen = approximateBezierLength(startPt, startHandle, endHandle, endPt);
        if (!isFinite(totalLen) || totalLen <= 0) return null;
        var clampedDist = Math.min(distancePt, totalLen);
        var t = findParameterForLength(startPt, startHandle, endHandle, endPt, clampedDist, totalLen);
        if (!isFinite(t)) return null;
        var point = evaluateBezierPoint(startPt, startHandle, endHandle, endPt, t);
        var tangent = evaluateBezierTangent(startPt, startHandle, endHandle, endPt, t);
        var norm = normalizeVector([-tangent[1], tangent[0]]);
        var offset = cm(offsetNormalCm || 0);
        var basePoint = [
            point[0] + norm[0] * offset,
            point[1] + norm[1] * offset
        ];
        drawNotchAtPoint(targetLayer, basePoint, norm, notchName);
        return { point: basePoint, normal: norm };
    }

    function drawNotchAtPoint(layer, center, normal, notchName) {
        if (!layer || !center || !normal) return;
        var notchHalf = cm(0.25);
        var notchStart = [center[0] - normal[0] * notchHalf, center[1] - normal[1] * notchHalf];
        var notchEnd = [center[0] + normal[0] * notchHalf, center[1] + normal[1] * notchHalf];
        drawDashedLine(layer, notchStart, notchEnd, { name: notchName || "" });
    }

    function addBackCapNotches(path, nodes, layer, derived) {
        if (!path || !nodes || nodes.length < 2 || !layer || !derived) return;
        var baseDistance = num(derived.bAP || 0);
        if (!isFinite(baseDistance) || baseDistance <= 0) return;
        var distances = [baseDistance, baseDistance + 1];
        var names = ["Back Cap Notch", "Back Cap Notch Upper"];
        for (var i = 0; i < distances.length; i++) {
            placeBackNotchOnChain(path, nodes, layer, distances[i], names[i]);
        }
    }

    function placeBackNotchOnChain(path, nodes, layer, distanceCm, notchName) {
        if (!isFinite(distanceCm) || distanceCm <= 0) return;
        var remaining = distanceCm;
        for (var i = 1; i < nodes.length; i++) {
            var prev = nodes[i - 1] || {};
            var curr = nodes[i] || {};
            var startPt = prev.anchor;
            var endPt = curr.anchor;
            if (!startPt || !endPt) continue;
            var startHandle = prev.right || startPt;
            var endHandle = curr.left || endPt;
            var segLenPt = approximateBezierLength(startPt, startHandle, endHandle, endPt);
            var segLenCm = segLenPt / CM_TO_PT;
            if (!isFinite(segLenCm) || segLenCm <= 0) continue;
            if (remaining <= segLenCm) {
                addNotchOnPath(path, startPt, endPt, startHandle, endHandle, remaining, layer, notchName);
                return;
            }
            remaining -= segLenCm;
        }
    }

    function evaluateBezierPoint(p0, p1, p2, p3, t) {
        var u = 1 - t;
        var tt = t * t;
        var uu = u * u;
        var uuu = uu * u;
        var ttt = tt * t;
        var f0 = uuu;
        var f1 = 3 * uu * t;
        var f2 = 3 * u * tt;
        var f3 = ttt;
        return [
            f0 * p0[0] + f1 * p1[0] + f2 * p2[0] + f3 * p3[0],
            f0 * p0[1] + f1 * p1[1] + f2 * p2[1] + f3 * p3[1]
        ];
    }

    function evaluateBezierTangent(p0, p1, p2, p3, t) {
        var u = 1 - t;
        var tt = t * t;
        var uu = u * u;
        var d0 = -3 * uu;
        var d1 = 3 * uu - 6 * u * t;
        var d2 = 6 * u * t - 3 * tt;
        var d3 = 3 * tt;
        return [
            d0 * p0[0] + d1 * p1[0] + d2 * p2[0] + d3 * p3[0],
            d0 * p0[1] + d1 * p1[1] + d2 * p2[1] + d3 * p3[1]
        ];
    }

    function findParameterForLength(p0, p1, p2, p3, targetLen, totalLen) {
        if (targetLen <= 0) return 0;
        if (targetLen >= totalLen) return 1;
        var low = 0;
        var high = 1;
        for (var iter = 0; iter < 20; iter++) {
            var mid = (low + high) / 2;
            var len = approximateBezierLength(p0, p1, p2, p3, mid);
            if (len < targetLen) {
                low = mid;
            } else {
                high = mid;
            }
        }
        return (low + high) / 2;
    }

    function approximateBezierLength(p0, p1, p2, p3, uptoT) {
        var length = 0;
        var prev = p0;
        var span = uptoT || 1;
        var steps = Math.max(5, Math.round(40 * span));
        for (var i = 1; i <= steps; i++) {
            var t = span * (i / steps);
            var pt = evaluateBezierPoint(p0, p1, p2, p3, t);
            var dx = pt[0] - prev[0];
            var dy = pt[1] - prev[1];
            length += Math.sqrt(dx * dx + dy * dy);
            prev = pt;
        }
        return length;
    }

    function computeContentBounds(derived) {
        derived = derived || {};
        var capLineEase = derived.CapLineEase || 0;
        var baselineLengthCm = (derived.SlW || 0) + 1 + capLineEase;
        if (!isFinite(baselineLengthCm) || baselineLengthCm < 1) baselineLengthCm = 1;
        var baselineLengthPt = cm(baselineLengthCm);

        var arcRadiusCm = calculateMark3RadiusCm(derived);
        if (!isFinite(arcRadiusCm) || arcRadiusCm < 0) arcRadiusCm = 0;
        var mark1 = [0, 0];
        var mark2 = [baselineLengthPt, 0];
        var layoutStub = { arcRadiusCm: arcRadiusCm };

        var arcInfo = buildMark3Arc(mark1, derived, layoutStub, ARC_TARGET_LENGTH_CM);
        var mark4 = arcInfo ? buildMark4Intersection(mark2, arcInfo, derived) : null;

        if ((!mark4 || !mark4.point || !arcInfo) && ARC_MAX_LENGTH_CM > ARC_TARGET_LENGTH_CM) {
            arcInfo = buildMark3Arc(mark1, derived, layoutStub, ARC_MAX_LENGTH_CM);
            mark4 = arcInfo ? buildMark4Intersection(mark2, arcInfo, derived) : null;
        }

        var bounds = createEmptyBounds();
        includePointInBounds(bounds, mark1);
        includePointInBounds(bounds, mark2);
        if (mark4 && arcInfo) {
            rebalanceArcAroundMark4(arcInfo, mark4.angle);
        }
        if (arcInfo) includeArcBounds(bounds, arcInfo);
        var mark4Point = (mark4 && mark4.point) ? mark4.point : null;
        if (mark4Point) includePointInBounds(bounds, mark4Point);

        var capPoints = mark4Point ? computeCapMarkerPoints(mark1, mark2, mark4Point, derived) : null;
        if (capPoints) {
            if (capPoints.mark11) includePointInBounds(bounds, capPoints.mark11);
            if (capPoints.mark12) includePointInBounds(bounds, capPoints.mark12);
            if (capPoints.mark13) includePointInBounds(bounds, capPoints.mark13);
            if (capPoints.mark14) includePointInBounds(bounds, capPoints.mark14);
        }

        var sleeveLengthPt = cm(derived.SlL || 0);
        if (mark4Point && sleeveLengthPt > 0) {
            var mark5 = [mark4Point[0], mark4Point[1] - sleeveLengthPt];
            var mark6 = [mark4Point[0], mark4Point[1] - sleeveLengthPt * 0.6];
            var mark1Down5 = [mark1[0], mark5[1]];
            var mark2Down5 = [mark2[0], mark5[1]];
            var mark1Down6 = [mark1[0], mark6[1]];
            var mark2Down6 = [mark2[0], mark6[1]];

            includePointInBounds(bounds, mark5);
            includePointInBounds(bounds, mark6);
            includePointInBounds(bounds, mark1Down5);
            includePointInBounds(bounds, mark2Down5);
            includePointInBounds(bounds, mark1Down6);
            includePointInBounds(bounds, mark2Down6);

        }

        if (mark4Point && capPoints) {
            if (capPoints.mark11) {
                var projectionInfoA = computePerpendicularProjection(mark1, mark4Point, capPoints.mark11);
                if (projectionInfoA && projectionInfoA.projectPoint) includePointInBounds(bounds, projectionInfoA.projectPoint);
            }
            if (capPoints.mark12) {
                var projectionInfoB = computePerpendicularProjection(mark2, mark4Point, capPoints.mark12);
                if (projectionInfoB && projectionInfoB.projectPoint) includePointInBounds(bounds, projectionInfoB.projectPoint);
            }
        }

        if (!bounds.initialized) {
            bounds.minX = 0;
            bounds.maxX = baselineLengthPt;
            bounds.minY = 0;
            bounds.maxY = 0;
        }

        var leftExtentCm = bounds.minX < 0 ? (-bounds.minX / CM_TO_PT) : 0;
        var rightExtentCm = bounds.maxX > 0 ? (bounds.maxX / CM_TO_PT) : 0;
        var topExtentCm = bounds.maxY > 0 ? (bounds.maxY / CM_TO_PT) : 0;
        var bottomExtentCm = bounds.minY < 0 ? (-bounds.minY / CM_TO_PT) : 0;
        var widthCm = rightExtentCm + leftExtentCm;
        var heightCm = topExtentCm + bottomExtentCm;

        if (!isFinite(widthCm) || widthCm <= 0) widthCm = Math.max(baselineLengthCm, 1);
        if (!isFinite(heightCm) || heightCm <= 0) heightCm = Math.max(derived.SlL || 1, 1);

        return {
            baselineLengthCm: baselineLengthCm,
            arcRadiusCm: arcRadiusCm,
            leftExtentCm: leftExtentCm,
            rightExtentCm: rightExtentCm,
            topExtentCm: topExtentCm,
            bottomExtentCm: bottomExtentCm,
            widthCm: widthCm,
            heightCm: heightCm
        };
    }

    function createEmptyBounds() {
        return {
            minX: 0,
            maxX: 0,
            minY: 0,
            maxY: 0,
            initialized: false
        };
    }

    function includePointInBounds(bounds, point) {
        if (!bounds || !point) return;
        var x = point[0];
        var y = point[1];
        if (!isFinite(x) || !isFinite(y)) return;
        if (!bounds.initialized) {
            bounds.minX = bounds.maxX = x;
            bounds.minY = bounds.maxY = y;
            bounds.initialized = true;
            return;
        }
        if (x < bounds.minX) bounds.minX = x;
        if (x > bounds.maxX) bounds.maxX = x;
        if (y < bounds.minY) bounds.minY = y;
        if (y > bounds.maxY) bounds.maxY = y;
    }

    function includeArcBounds(bounds, arcInfo) {
        if (!bounds || !arcInfo) return;
        includePointInBounds(bounds, arcInfo.startPoint);
        includePointInBounds(bounds, arcInfo.endPoint);
        var specialAngles = [0, Math.PI / 2, -Math.PI / 2];
        for (var i = 0; i < specialAngles.length; i++) {
            var ang = specialAngles[i];
            if (angleWithinRange(ang, arcInfo.startAngle, arcInfo.endAngle)) {
                var pt = [
                    arcInfo.center[0] + arcInfo.radiusPt * Math.cos(ang),
                    arcInfo.center[1] + arcInfo.radiusPt * Math.sin(ang)
                ];
                includePointInBounds(bounds, pt);
            }
        }
    }

    function angleWithinRange(angle, start, end) {
        var tol = 1e-6;
        return (angle + tol) >= start && (angle - tol) <= end;
    }

    function calculateMark3RadiusCm(derived) {
        if (!derived) return 0;
        var fAh = (derived.fAhConstruction !== null && derived.fAhConstruction !== undefined) ? derived.fAhConstruction : derived.fAh;
        var capEase = (derived.CapEaseCm !== null && derived.CapEaseCm !== undefined) ? derived.CapEaseCm : 0;
        var radius = 0.97 * (fAh || 0) + 0.25 * (capEase || 0);
        if (!isFinite(radius) || radius < 0) radius = 0;
        return radius;
    }

    function calculateMark4RadiusCm(derived) {
        if (!derived) return 0;
        var bAh = (derived.bAhConstruction !== null && derived.bAhConstruction !== undefined) ? derived.bAhConstruction : derived.bAh;
        var capEase = (derived.CapEaseCm !== null && derived.CapEaseCm !== undefined) ? derived.CapEaseCm : 0;
        var radius = 0.97 * (bAh || 0) + 0.75 * (capEase || 0);
        if (!isFinite(radius) || radius < 0) radius = 0;
        return radius;
    }

    function buildMark3Arc(centerPt, derived, layoutBounds, targetLengthCm) {
        if (!centerPt || !derived) return null;
        var radiusCm = layoutBounds && typeof layoutBounds.arcRadiusCm === "number"
            ? layoutBounds.arcRadiusCm
            : calculateMark3RadiusCm(derived);
        if (!isFinite(radiusCm) || radiusCm <= 0) return null;
        var radiusPt = cm(radiusCm);

        var desiredLengthCm = targetLengthCm;
        if (!isFinite(desiredLengthCm) || desiredLengthCm <= 0) desiredLengthCm = ARC_TARGET_LENGTH_CM;
        var sweepAngle = desiredLengthCm / radiusCm;
        var maxSweep = ARC_END_ANGLE - ARC_MIN_START_ANGLE; // up to 180 degrees
        sweepAngle = Math.min(Math.max(sweepAngle, ARC_MIN_SWEEP_RAD), maxSweep);
        var startAngle = ARC_END_ANGLE - sweepAngle;
        if (startAngle < ARC_MIN_START_ANGLE) {
            startAngle = ARC_MIN_START_ANGLE;
            sweepAngle = ARC_END_ANGLE - startAngle;
        }
        var endAngle = ARC_END_ANGLE;
        var startPoint = polarPoint(centerPt, radiusPt, startAngle);
        var endPoint = polarPoint(centerPt, radiusPt, endAngle);
        if (!startPoint || !endPoint) return null;

        var handleLength = (4 / 3) * radiusPt * Math.tan(sweepAngle / 4);
        if (!isFinite(handleLength)) handleLength = 0;
        var startTangent = [-Math.sin(startAngle), Math.cos(startAngle)];
        var endTangent = [-Math.sin(endAngle), Math.cos(endAngle)];
        var startHandle = [
            startPoint[0] + startTangent[0] * handleLength,
            startPoint[1] + startTangent[1] * handleLength
        ];
        var endHandle = [
            endPoint[0] - endTangent[0] * handleLength,
            endPoint[1] - endTangent[1] * handleLength
        ];
        var totalLengthCm = sweepAngle * radiusCm;

        return {
            startPoint: startPoint,
            endPoint: endPoint,
            center: centerPt,
            radiusCm: radiusCm,
            radiusPt: radiusPt,
            startAngle: startAngle,
            endAngle: endAngle,
            startHandle: startHandle,
            endHandle: endHandle,
            totalLengthCm: totalLengthCm,
            targetLengthCm: desiredLengthCm
        };
    }

    function rebalanceArcAroundMark4(arcInfo, mark4Angle) {
        if (!arcInfo || !isFinite(mark4Angle)) return;
        var radiusCm = arcInfo.radiusCm;
        if (!isFinite(radiusCm) || radiusCm <= 0) return;
        var ratioUp = ARC_MARK4_UP_RATIO;
        var ratioDown = ARC_MARK4_DOWN_RATIO;
        var ratioSum = ratioUp + ratioDown;
        if (!isFinite(ratioSum) || ratioSum <= 0) return;
        var desiredLengthCm = arcInfo.totalLengthCm || arcInfo.targetLengthCm || ARC_TARGET_LENGTH_CM;
        if (!isFinite(desiredLengthCm) || desiredLengthCm <= 0) desiredLengthCm = ARC_TARGET_LENGTH_CM;
        var sweepAngle = desiredLengthCm / radiusCm;
        var maxSweep = ARC_END_ANGLE - ARC_MIN_START_ANGLE;
        sweepAngle = Math.min(Math.max(sweepAngle, ARC_MIN_SWEEP_RAD), maxSweep);
        var upSweep = sweepAngle * (ratioUp / ratioSum);
        var downSweep = sweepAngle - upSweep;
        var startAngle = mark4Angle - downSweep;
        var endAngle = mark4Angle + upSweep;
        var minAngle = ARC_MIN_START_ANGLE;
        var maxAngleLimit = ARC_END_ANGLE;
        if (startAngle < minAngle) {
            var shiftUp = minAngle - startAngle;
            startAngle += shiftUp;
            endAngle += shiftUp;
        }
        if (endAngle > maxAngleLimit) {
            var shiftDown = endAngle - maxAngleLimit;
            startAngle -= shiftDown;
            endAngle -= shiftDown;
        }
        if (startAngle < minAngle) startAngle = minAngle;
        if (endAngle > maxAngleLimit) endAngle = maxAngleLimit;
        var sweep = endAngle - startAngle;
        if (sweep < ARC_MIN_SWEEP_RAD) {
            var mid = (startAngle + endAngle) / 2;
            startAngle = mid - ARC_MIN_SWEEP_RAD / 2;
            endAngle = mid + ARC_MIN_SWEEP_RAD / 2;
        }
        updateArcGeometry(arcInfo, startAngle, endAngle);
    }

    function updateArcGeometry(arcInfo, startAngle, endAngle) {
        if (!arcInfo || !arcInfo.center || !isFinite(arcInfo.radiusPt)) return;
        arcInfo.startAngle = startAngle;
        arcInfo.endAngle = endAngle;
        var startPoint = polarPoint(arcInfo.center, arcInfo.radiusPt, startAngle);
        var endPoint = polarPoint(arcInfo.center, arcInfo.radiusPt, endAngle);
        arcInfo.startPoint = startPoint;
        arcInfo.endPoint = endPoint;
        var sweepAngle = endAngle - startAngle;
        if (sweepAngle < ARC_MIN_SWEEP_RAD) sweepAngle = ARC_MIN_SWEEP_RAD;
        var handleLength = (4 / 3) * arcInfo.radiusPt * Math.tan(sweepAngle / 4);
        if (!isFinite(handleLength)) handleLength = 0;
        var startTangent = [-Math.sin(startAngle), Math.cos(startAngle)];
        var endTangent = [-Math.sin(endAngle), Math.cos(endAngle)];
        arcInfo.startHandle = [
            startPoint[0] + startTangent[0] * handleLength,
            startPoint[1] + startTangent[1] * handleLength
        ];
        arcInfo.endHandle = [
            endPoint[0] - endTangent[0] * handleLength,
            endPoint[1] - endTangent[1] * handleLength
        ];
        arcInfo.totalLengthCm = sweepAngle * (arcInfo.radiusCm || 0);
    }

    function polarPoint(centerPt, radiusPt, angle) {
        if (!centerPt || !isFinite(radiusPt)) return null;
        return [
            centerPt[0] + radiusPt * Math.cos(angle),
            centerPt[1] + radiusPt * Math.sin(angle)
        ];
    }

    function drawArcPath(targetLayer, arcInfo) {
        if (!targetLayer || !arcInfo || !arcInfo.startPoint || !arcInfo.endPoint) return;
        ensureLayerWritable(targetLayer);
        var path = targetLayer.pathItems.add();
        try { path.name = "Arc"; } catch (eName) {}
        path.stroked = true;
        path.strokeWidth = 1;
        path.strokeColor = COL_BLACK;
        path.strokeDashes = [];
        path.strokeCap = StrokeCap.BUTTENDCAP;
        path.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        path.filled = false;
        path.closed = false;
        path.setEntirePath([arcInfo.startPoint, arcInfo.endPoint]);
        try {
            var startPoint = path.pathPoints[0];
            startPoint.anchor = arcInfo.startPoint;
            startPoint.leftDirection = arcInfo.startPoint;
            startPoint.rightDirection = arcInfo.startHandle || arcInfo.startPoint;

            var endPoint = path.pathPoints[path.pathPoints.length - 1];
            endPoint.anchor = arcInfo.endPoint;
            endPoint.leftDirection = arcInfo.endHandle || arcInfo.endPoint;
            endPoint.rightDirection = arcInfo.endPoint;
        } catch (e) {}
    }

    function buildMark4Intersection(mark2, arcInfo, derived) {
        if (!mark2 || !arcInfo || !derived) return null;
        var radiusCm = calculateMark4RadiusCm(derived);
        if (!isFinite(radiusCm) || radiusCm <= 0) return null;
        var radiusPt = cm(radiusCm);
        var intersections = intersectCircles(arcInfo.center, arcInfo.radiusPt, mark2, radiusPt);
        if (!intersections || intersections.length === 0) return null;

        var targetPoint = selectPointOnArc(intersections, arcInfo);
        if (!targetPoint) {
            targetPoint = selectPreferredArcPoint(intersections, arcInfo);
        }
        if (!targetPoint) return null;

        return {
            point: targetPoint.point,
            radiusCm: radiusCm,
            angle: targetPoint.angle
        };
    }

    function intersectCircles(center1, radius1, center2, radius2) {

        if (!center1 || !center2) return null;
        var dx = center2[0] - center1[0];
        var dy = center2[1] - center1[1];
        var dist = Math.sqrt(dx * dx + dy * dy);
        if (!isFinite(dist) || dist === 0) return null;
        var sum = radius1 + radius2;
        if (dist > sum + 0.01) return null;
        var diff = Math.abs(radius1 - radius2);
        if (dist < diff - 0.01) return null;

        var a = (radius1 * radius1 - radius2 * radius2 + dist * dist) / (2 * dist);
        var hSq = radius1 * radius1 - a * a;
        if (hSq < 0) hSq = 0;
        var h = Math.sqrt(hSq);
        var xm = center1[0] + a * dx / dist;
        var ym = center1[1] + a * dy / dist;
        if (h === 0) return [[xm, ym]];
        var rx = -dy * (h / dist);
        var ry = dx * (h / dist);
        return [
            [xm + rx, ym + ry],
            [xm - rx, ym - ry]
        ];
    }

    function selectPointOnArc(points, arcInfo) {
        if (!points || !arcInfo || !arcInfo.center) return null;
        var best = null;
        for (var i = 0; i < points.length; i++) {
            var pt = points[i];
            var angle = Math.atan2(pt[1] - arcInfo.center[1], pt[0] - arcInfo.center[0]);
            if (isAngleOnArc(angle, arcInfo.startAngle, arcInfo.endAngle)) {
                if (!best || angle > best.angle) {
                    best = { point: pt, angle: angle };
                }
            }
        }
        return best;
    }

    function selectPreferredArcPoint(points, arcInfo) {
        if (!points || !arcInfo || !arcInfo.center) return null;
        var best = null;
        for (var i = 0; i < points.length; i++) {
            var pt = points[i];
            var angle = Math.atan2(pt[1] - arcInfo.center[1], pt[0] - arcInfo.center[0]);
            if (!best || angle > best.angle) {
                best = { point: pt, angle: angle };
            }
        }
        return best;
    }

    function isAngleOnArc(angle, startAngle, endAngle) {
        var tol = 1e-6;
        if (startAngle > endAngle) return false;
        return (angle + tol) >= startAngle && (angle - tol) <= endAngle;
    }

    function placeMarker(markerLayer, numberLayer, point, label) {
        if (!markerLayer || !point) return;
        ensureLayerWritable(markerLayer);
        var radius = cm(MARKER_RADIUS_CM);
        var circle = markerLayer.pathItems.ellipse(point[1] + radius, point[0] - radius, radius * 2, radius * 2);
        circle.stroked = false;
        circle.filled = true;
        circle.fillColor = COL_BLACK;

        if (numberLayer !== undefined && numberLayer !== null && label !== undefined && label !== null) {
            ensureLayerWritable(numberLayer);
            var tf = numberLayer.textFrames.add();
            tf.contents = String(label);
            try { tf.textRange.characterAttributes.size = 10; } catch (eSize) {}
            try { tf.textRange.characterAttributes.fillColor = COL_WHITE; } catch (eFill) {}
            try { tf.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (eJust) {}
            centerTextFrame(tf, point);
        }
    }

    function centerTextFrame(tf, anchor) {
        if (!tf || !anchor) return;
        try {
            var bounds = tf.visibleBounds;
            var cx = (bounds[0] + bounds[2]) / 2;
            var cy = (bounds[1] + bounds[3]) / 2;
            tf.translate(anchor[0] - cx, anchor[1] - cy);
        } catch (e) {
            try {
                tf.left = anchor[0] - tf.width / 2;
                tf.top = anchor[1] + tf.height / 2;
            } catch (ignored) {}
        }
    }

    /* ----------------------------
        Utilities
    ---------------------------- */
    function cm(v) { return (v || 0) * CM_TO_PT; }

    function num(val) {
        var parsed = parseFloat(val);
        return isNaN(parsed) ? 0 : parsed;
    }

    function formatNumber(val) {
        var parsed = parseFloat(val);
        if (isNaN(parsed)) return "";
        var rounded = Math.round(parsed * 100) / 100;
        var str = rounded.toFixed(2);
        str = str.replace(/\.00$/, "");
        str = str.replace(/(\.\d)0$/, "$1");
        return str;
    }

    function computeDerived(data) {
        var derivedValues = {};
        var ahh = num(data.AhH);
        derivedValues.AhHConstruction = ahh + num(data.AhHEase);

        var fAh = num(data.fAh);
        var bAh = num(data.bAh);
        derivedValues.fAhConstruction = fAh + num(data.fAhEase);
        derivedValues.bAhConstruction = bAh + num(data.bAhEase);
        var manualAhC = (data.AhC !== null && data.AhC !== undefined && data.AhC !== "") ? num(data.AhC) : (fAh + bAh);
        derivedValues.AhC = manualAhC;
        derivedValues.AhCConstruction = manualAhC + num(data.AhCEase);

        var AL = num(data.AL || 0);
        derivedValues.SlL = AL + num(data.ALEase);
        var upper = num(data.upAC || 0);
        derivedValues.SlW = upper + num(data.upACEase);
        var wrist = num(data.WrC || 0);
        derivedValues.HeW = wrist + num(data.WrCEase);
        derivedValues.CapLineEase = num(data.CapLineEase);

        derivedValues.CapEasePct = num(data.CapEasePct);
        derivedValues.CapEasePctConstruction = derivedValues.CapEasePct + num(data.CapEasePctEase);
        derivedValues.CapEaseCm = derivedValues.AhCConstruction * derivedValues.CapEasePctConstruction / 100;
        derivedValues.CapC = derivedValues.AhCConstruction + derivedValues.CapEaseCm + num(data.CapCEase);
        derivedValues.fAP = num(data.fAP || 0);
        derivedValues.bAP = num(data.bAP || 0);
        return derivedValues;
    }

    function getMeasurementRowDefinitions() {
        return [
            {
                number: 1,
                main: { type: "input", key: "AhH", label: "AhH" },
                ease: { type: "input", key: "AhHEase", label: "Ease", defaultValue: 0 },
                construction: { type: "derived", key: "AhHConstruction", label: "AhH" }
            },
            {
                number: 2,
                main: { type: "input", key: "fAh", label: "fAh" },
                ease: { type: "input", key: "fAhEase", label: "Ease", defaultValue: 0 },
                construction: { type: "derived", key: "fAhConstruction", label: "fAh" }
            },
            {
                number: 3,
                main: { type: "input", key: "bAh", label: "bAh" },
                ease: { type: "input", key: "bAhEase", label: "Ease", defaultValue: 0 },
                construction: { type: "derived", key: "bAhConstruction", label: "bAh" }
            },
            {
                number: 4,
                main: null,
                ease: null,
                construction: { type: "derived", key: "AhCConstruction", label: "AhC", noNumber: true }
            },
            {
                number: 5,
                main: { type: "input", key: "AL", label: "AL", defaultValue: 60 },
                ease: { type: "input", key: "ALEase", label: "Ease", defaultValue: 0 },
                construction: { type: "derived", key: "SlL", label: "SlL" }
            },
            {
                number: 6,
                main: { type: "input", key: "upAC", label: "upAC", defaultValue: 28 },
                ease: { type: "input", key: "upACEase", label: "Ease", defaultValue: 9 },
                construction: { type: "derived", key: "SlW", label: "SlW" }
            },
            {
                number: 7,
                main: { type: "input", key: "WrC", label: "WrC", defaultValue: 16 },
                ease: { type: "input", key: "WrCEase", label: "Ease", defaultValue: 6 },
                construction: { type: "derived", key: "HeW", label: "HeW" }
            },
            {
                number: 8,
                main: { type: "input", key: "CapEasePct", label: "Cap Ease (%)", defaultValue: 3 },
                ease: { type: "input", key: "CapEasePctEase", label: "Ease (%)", defaultValue: 0, noNumber: true },
                construction: { type: "derived", key: "CapEaseCm", label: "Cap Ease (cm)" }
            },
            {
                number: 9,
                main: null,
                ease: null,
                construction: { type: "derived", key: "CapC", label: "CapC" }
            },
            {
                number: 10,
                main: { type: "input", key: "CapLineEase", label: "Cap Line Ease", defaultValue: 0 },
                ease: null,
                construction: null
            },
            {
                number: 11,
                main: { type: "input", key: "fAP", label: "fAP", defaultValue: 4.5 },
                ease: null,
                construction: null
            },
            {
                number: 12,
                main: { type: "input", key: "bAP", label: "bAP", defaultValue: 8.6 },
                ease: null,
                construction: null
            }
        ];
    }

    function registerMeasurementLabel(key, label) {
        if (!key) return;
        measurementLabelMap[key] = label || key;
    }

    function resolveMeasurementLabel(key) {
        if (!key) return "";
        return measurementLabelMap.hasOwnProperty(key) ? measurementLabelMap[key] : key;
    }

    function captureSummaryValue(source, key) {
        if (!source || !key) return null;
        var value = source[key];
        if (value === undefined || value === null || value === "") return null;
        return value;
    }

    function buildSummaryLabel(rowNumber, fieldDef) {
        if (!fieldDef || !fieldDef.label) return "";
        var prefix = (!fieldDef.noNumber && typeof rowNumber === "number") ? (rowNumber + ". ") : "";
        return prefix + fieldDef.label;
    }

    function updateMeasurementSummaryData(data, derived) {
        measurementResults = {};
        easeResults = {};
        finalResults = {};
        measurementLabelMap = {};
        var rows = getMeasurementRowDefinitions();
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var key = (row.main && row.main.key) || (row.construction && row.construction.key) || (row.ease && row.ease.key);
            if (!key) continue;
            var labelSource = row.main || row.construction || row.ease;
            registerMeasurementLabel(key, buildSummaryLabel(row.number, labelSource));
            measurementResults[key] = captureSummaryValue(data, row.main ? row.main.key : null);
            if (measurementResults[key] === null && row.construction) {
                measurementResults[key] = captureSummaryValue(derived, row.construction.key);
            }
            easeResults[key] = captureSummaryValue(data, row.ease ? row.ease.key : null);
            finalResults[key] = row.construction ? captureSummaryValue(derived, row.construction.key) : measurementResults[key];
        }
    }

    function formatReferenceValue(value) {
        if (value === undefined || value === null || value === "") return "-";
        var parsed = parseFloat(value);
        if (!isNaN(parsed)) return formatNumber(parsed);
        return String(value);
    }

    function collectMeasurementRows() {
        var rows = [];
        for (var key in measurementResults) {
            if (!measurementResults.hasOwnProperty(key)) continue;
            rows.push({
                id: resolveMeasurementLabel(key),
                rawId: key,
                meas: formatReferenceValue(measurementResults[key]),
                ease: formatReferenceValue(easeResults[key]),
                finalValue: formatReferenceValue(finalResults[key])
            });
        }
        rows.sort(function (a, b) {
            if (a.id === b.id) {
                if (a.rawId < b.rawId) return -1;
                if (a.rawId > b.rawId) return 1;
                return 0;
            }
            return (a.id < b.id) ? -1 : 1;
        });
        return rows;
    }

    function closeMeasurementPalette() {
        try {
            if (measurementPalette && typeof measurementPalette.close === "function") {
                measurementPalette.close();
            }
        } catch (eLocalClose) {}
        try {
            var globalPalette = $.global[measurementPaletteGlobalKey];
            if (globalPalette && globalPalette.window && typeof globalPalette.window.close === "function") {
                globalPalette.window.close();
            }
            delete $.global[measurementPaletteGlobalKey];
        } catch (eGlobalClose) {}
        measurementPalette = null;
    }

    function buildPaletteStaticText(group, label, width, justification) {
        var st = group.add("statictext", undefined, label);
        if (width !== undefined) {
            st.preferredSize.width = width;
            st.minimumSize.width = width;
        }
        if (justification) {
            try { st.justify = justification; } catch (eJust) {}
        }
        return st;
    }

    function showMeasurementPaletteWindow(patternLabel) {
        closeMeasurementPalette();
        var palette = new Window("palette", "Measurement Summary");
        palette.orientation = "column";
        palette.alignChildren = ["fill", "top"];
        palette.spacing = 8;
        palette.margins = 12;
        palette.preferredSize.width = 560;

        var metaGroup = palette.add("group");
        metaGroup.alignment = "fill";
        metaGroup.add("statictext", undefined, "Pattern: " + (patternLabel || "-"));

        var headerGroup = palette.add("group");
        headerGroup.alignment = "fill";
        headerGroup.spacing = 6;
        var columnWidths = [200, 120, 120, 120];
        buildPaletteStaticText(headerGroup, "Measurement", columnWidths[0], "left");
        buildPaletteStaticText(headerGroup, "Measure", columnWidths[1], "right");
        buildPaletteStaticText(headerGroup, "Ease", columnWidths[2], "right");
        buildPaletteStaticText(headerGroup, "Final", columnWidths[3], "right");

        var rows = collectMeasurementRows();
        if (!rows.length) {
            var emptyRow = palette.add("group");
            emptyRow.alignment = "fill";
            emptyRow.add("statictext", undefined, "No measurements available.");
        } else {
            for (var i = 0; i < rows.length; i++) {
                var row = rows[i];
                var rowGroup = palette.add("group");
                rowGroup.alignment = "fill";
                rowGroup.spacing = 6;
                buildPaletteStaticText(rowGroup, row.id, columnWidths[0], "left");
                buildPaletteStaticText(rowGroup, row.meas, columnWidths[1], "right");
                buildPaletteStaticText(rowGroup, row.ease, columnWidths[2], "right");
                buildPaletteStaticText(rowGroup, row.finalValue, columnWidths[3], "right");
            }
        }

        palette.onClose = function () {
            measurementPalette = null;
            try { delete $.global[measurementPaletteGlobalKey]; } catch (eDel) {}
        };
        palette.show();
        measurementPalette = palette;
        $.global[measurementPaletteGlobalKey] = { window: palette };
    }

    /* ----------------------------
        Measurement Dialog
    ---------------------------- */
    function showMeasurementDialog(data) {
        var dlg = new Window("dialog", "Wide Basic Sleeve - Default size, 38");
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];
        dlg.margins = 18;

        var info = dlg.add("statictext", undefined, "Values in Centimetres");
        info.justification = "center";

        var columnConfigs = [
            { title: "Main Measurements", labelWidth: 130, inputWidth: 70 },
            { title: "Ease", labelWidth: 65, inputWidth: 70 },
            { title: "Construction Measurements", labelWidth: 135, inputWidth: 65 }
        ];

        var header = dlg.add("group");
        header.orientation = "row";
        header.alignChildren = ["left", "center"];
        header.spacing = 0;
        header.margins = [0, 0, 0, 6];
        for (var h = 0; h < columnConfigs.length; h++) {
            addHeaderCell(header, columnConfigs[h].title, getColumnTotalWidth(columnConfigs[h]));
        }

        var rowsContainer = dlg.add("group");
        rowsContainer.orientation = "column";
        rowsContainer.alignChildren = ["fill", "top"];
        rowsContainer.spacing = 6;

        var measurementRows = getMeasurementRowDefinitions();

        var inputs = {};
        var derivedFields = {};

        for (var i = 0; i < measurementRows.length; i++) {
            buildMeasurementRow(rowsContainer, measurementRows[i], columnConfigs);
        }

        var summaryCheckbox = dlg.add("checkbox", undefined, "Show Measurement Summary");
        summaryCheckbox.value = !!data.showMeasurementSummary;
        summaryCheckbox.alignment = ["left", "center"];

        var buttonBar = dlg.add("group");
        buttonBar.alignment = "right";
        buttonBar.spacing = 10;
        var okBtn = buttonBar.add("button", undefined, "OK", { name: "ok" });
        var cancelBtn = buttonBar.add("button", undefined, "Cancel", { name: "cancel" });

        cancelBtn.onClick = function () {
            dlg.close(0);
        };

        okBtn.onClick = function () {
            for (var key in inputs) {
                if (!inputs.hasOwnProperty(key)) continue;
                var field = inputs[key];
                var val = parseFloat(field.text);
                if (!isNaN(val)) {
                    data[key] = val;
                } else if (field.text === "" && (key === "AhH" || key === "fAh" || key === "bAh")) {
                    data[key] = null;
                }
            }
            data.showMeasurementSummary = summaryCheckbox.value ? true : false;
            dlg.close(1);
        };

        updateDerivedFields();
        var dialogResult = dlg.show();
        return dialogResult === 1 ? data : null;

        function addHeaderCell(parent, label, width) {
            var cell = parent.add("group");
            cell.orientation = "row";
            cell.alignChildren = ["fill", "center"];
            cell.margins = [0, 0, 0, 0];
            if (width) {
                cell.minimumSize.width = width;
                cell.maximumSize.width = width;
            }
            var text = cell.add("statictext", undefined, label);
            text.alignment = ["fill", "center"];
            text.justification = "center";
            if (width) {
                text.preferredSize.width = width;
                text.maximumSize.width = width;
            }
            text.textWrap = true;
            try { text.graphics.font = ScriptUI.newFont("Arial", "bold", 11); } catch (eFontBold) {}
        }

        function buildMeasurementRow(parent, rowDef, colConfigs) {
            var rowGroup = parent.add("group");
            rowGroup.orientation = "row";
            rowGroup.alignChildren = ["left", "top"];
            rowGroup.spacing = 12;
            rowGroup.margins = [0, 4, 0, 4];

            var mainDef = cloneWithNumber(rowDef.main, rowDef.number);
            var easeDef = cloneWithNumber(rowDef.ease, rowDef.number);
            var consDef = cloneWithNumber(rowDef.construction, rowDef.number);

            createCell(mainDef, rowGroup, colConfigs[0]);
            createCell(easeDef, rowGroup, colConfigs[1]);
            createCell(consDef, rowGroup, colConfigs[2]);
        }

        function cloneWithNumber(def, number) {
            if (!def) return null;
            var clone = {};
            for (var key in def) {
                if (def.hasOwnProperty(key)) clone[key] = def[key];
            }
            if (clone.label) {
                clone.displayLabel = clone.label;
            }
            return clone;
        }

        function getColumnTotalWidth(config) {
            if (!config) return 0;
            var labelWidth = config.labelWidth || 0;
            var inputWidth = config.inputWidth || 0;
            var gutter = typeof config.gutter === "number" ? config.gutter : 10;
            return labelWidth + inputWidth + gutter;
        }

        function createCell(def, parent, config) {
            var cell = parent.add("group");
            cell.orientation = "row";
            cell.alignChildren = ["left", "top"];
            cell.spacing = 6;
            cell.margins = [0, 1, 0, 1];
            var totalWidth = getColumnTotalWidth(config);
            if (totalWidth > 0) {
                cell.minimumSize.width = totalWidth;
                cell.maximumSize.width = totalWidth;
            }
            var labelWidth = (config && config.labelWidth) ? config.labelWidth : Math.max(0, totalWidth - 80);
            var inputWidth = (config && config.inputWidth) ? config.inputWidth : Math.max(60, totalWidth - labelWidth - 10);

            var labelContainer = cell.add("group");
            labelContainer.orientation = "row";
            labelContainer.alignChildren = ["left", "center"];
            labelContainer.margins = [0, 0, 0, 0];
            labelContainer.alignment = ["fill", "center"];
            if (labelWidth > 0) {
                labelContainer.minimumSize.width = labelWidth;
                labelContainer.maximumSize.width = labelWidth;
                labelContainer.preferredSize.width = labelWidth;
            }

            var controlContainer = cell.add("group");
            controlContainer.orientation = "row";
            controlContainer.alignChildren = ["left", "center"];
            controlContainer.margins = [0, 0, 0, 0];
            if (inputWidth > 0) {
                controlContainer.minimumSize.width = inputWidth;
                controlContainer.maximumSize.width = inputWidth;
                controlContainer.preferredSize.width = inputWidth;
            }

            if (!def) return cell;

            var labelText = def.displayLabel || def.label;
            if (labelText) {
                var label = labelContainer.add("statictext", undefined, labelText + ":", { multiline: true });
                label.alignment = ["left", "center"];
                label.justification = "left";
                if (labelWidth > 0) {
                    label.maximumSize.width = labelWidth;
                    label.preferredSize.width = labelWidth;
                }
            }

            if (def.type === "static") {
                var staticText = controlContainer.add("statictext", undefined, def.value || "-");
                staticText.alignment = ["left", "center"];
                return cell;
            }

            if (def.type === "input") {
                var initial = data[def.key];
                if ((initial === null || initial === undefined || initial === "") && typeof def.defaultValue !== "undefined") {
                    initial = def.defaultValue;
                    data[def.key] = initial;
                }
                var input = createEditBox(controlContainer, formatNumber(initial), false, inputWidth);
                inputs[def.key] = input;
                (function (edit, key) {
                    edit.onChanging = function () {
                        var val = parseFloat(edit.text);
                        if (!isNaN(val)) data[key] = val;
                        updateDerivedFields();
                    };
                })(input, def.key);
                return cell;
            }

            if (def.type === "derived") {
                var derivedBox = createEditBox(controlContainer, "", true, inputWidth);
                registerDerivedField(def.key, derivedBox);
                return cell;
            }

            return cell;
        }

        function createEditBox(parent, initialValue, readOnly, width) {
            var edit = parent.add("edittext", undefined, initialValue || "");
            edit.characters = 6;
            var desiredWidth = width && width > 0 ? width : 70;
            edit.preferredSize.width = desiredWidth;
            edit.minimumSize.width = desiredWidth;
            edit.maximumSize.width = desiredWidth;
            edit.justification = "right";
            if (readOnly) {
                edit.enabled = false;
                try {
                    edit.graphics.backgroundColor = edit.graphics.newBrush(edit.graphics.BrushType.SOLID_COLOR, [0.9, 0.9, 0.9, 1]);
                } catch (eBg) {}
            }
            return edit;
        }

        function registerDerivedField(key, field) {
            if (!derivedFields[key]) derivedFields[key] = [];
            derivedFields[key].push(field);
        }

        function updateDerivedFields() {
            var d = computeDerived(data);
            for (var key in derivedFields) {
                if (!derivedFields.hasOwnProperty(key)) continue;
                var value = d[key];
                var formatted = formatNumber(value);
                var targets = derivedFields[key];
                for (var i = 0; i < targets.length; i++) {
                    targets[i].text = formatted;
                }
            }
        }
    }

})();
