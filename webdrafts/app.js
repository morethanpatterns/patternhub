const NS = "http://www.w3.org/2000/svg";
const INKSCAPE_NS = "http://www.inkscape.org/namespaces/inkscape";
const MM_PER_IN = 25.4;
const PAGE_WIDTH_MM = 600;
const PAGE_HEIGHT_MM = 600;
const PAGE_MARGIN_MM = 25;
const PADDING_MM = 100;
const TOP_PADDING_MM = 10;
const MARKER_RADIUS_MM = 2.6;
const LABEL_FONT_SIZE_MM = 4;
const BACK_GAP_CM = 10;
const BACK_OFFSET_IN = BACK_GAP_CM / 2.54;
const SHARE_URL = "https://morethanpatterns.github.io/patternhub/webdrafts/index.html";
const BUST_CUP_OFFSETS = {
  A: 0.875,
  B: 1.25,
  C: 1.5,
  D: 1.75,
};
const CM_TO_MM = 10;
const ALDRICH_HANDLE_RATIO_CAP = 1.6;
const ALDRICH_COLORS = Object.freeze({
  primary: "#111111",
  guide: "#000000",
});

const baseAldrichDefaults = {
  bust: 88,
  waist: 68,
  hip: 94,
  bustEase: 5,
  waistEase: 3,
  napeToWaist: 41,
  shoulder: 12.25,
  backWidth: 34.4,
  waistToHip: 20.6,
  armscyeDepth: 21,
  chest: 32.4,
  neckSize: 37,
  frontWaistDartBackOff: 2.5,
  frontSideWaistDart: 0,
  backSideWaistDart: 0,
};
const ALDRICH_DEFAULT_FRONT_NECK_DART = computeAldrichFrontNeckDart(baseAldrichDefaults.bust);
const ALDRICH_DEFAULT_DARTS = computeAldrichWaistDarts(
  baseAldrichDefaults.bust,
  baseAldrichDefaults.waist,
  baseAldrichDefaults.bustEase,
  baseAldrichDefaults.waistEase
);
const ALDRICH_DEFAULTS = Object.freeze({
  ...baseAldrichDefaults,
  frontNeckDart: ALDRICH_DEFAULT_FRONT_NECK_DART,
  frontWaistDart: ALDRICH_DEFAULT_DARTS.front,
  backWaistDart: ALDRICH_DEFAULT_DARTS.back,
  frontSideWaistDart: ALDRICH_DEFAULT_DARTS.frontSide,
  backSideWaistDart: ALDRICH_DEFAULT_DARTS.backSide,
  bustWaistDiff: Math.abs(
    computeAldrichWaistDiff(
      baseAldrichDefaults.bust,
      baseAldrichDefaults.waist,
      baseAldrichDefaults.bustEase,
      baseAldrichDefaults.waistEase
    )
  ),
});

function createSvgRoot(w = PAGE_WIDTH_MM, h = PAGE_HEIGHT_MM) {
  const svg = document.createElementNS(NS, "svg");
  svg.setAttribute("xmlns", NS);
  svg.setAttribute("xmlns:inkscape", INKSCAPE_NS);
  svg.setAttribute("version", "1.1");
  svg.setAttribute("width", w + "mm");
  svg.setAttribute("height", h + "mm");
  svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
  return svg;
}

function layer(parent, name, opts = {}) {
  const g = document.createElementNS(NS, "g");
  const id = opts.id || slugifyId(name, opts.prefix);
  if (id) g.setAttribute("id", id);
  g.setAttribute("data-layer", name);
  if (opts.asLayer) {
    g.setAttributeNS(INKSCAPE_NS, "inkscape:groupmode", "layer");
    g.setAttributeNS(INKSCAPE_NS, "inkscape:label", name);
  }
  parent.appendChild(g);
  return g;
}

function path(d, attrs = {}) {
  const p = document.createElementNS(NS, "path");
  p.setAttribute("d", d);
  Object.entries(attrs).forEach(([k, v]) => {
    if (v != null) p.setAttribute(k, v);
  });
  return p;
}

function textNode(x, y, str, attrs = {}) {
  const t = document.createElementNS(NS, "text");
  t.setAttribute("x", x);
  t.setAttribute("y", y);
  t.setAttribute("font-size", attrs["font-size"] || LABEL_FONT_SIZE_MM);
  t.textContent = str;
  Object.entries(attrs).forEach(([k, v]) => {
    if (k !== "font-size") t.setAttribute(k, v);
  });
  return t;
}

function createBounds() {
  return {
    minX: Infinity,
    minY: Infinity,
    maxX: -Infinity,
    maxY: -Infinity,
    include(x, y) {
      if (!Number.isFinite(x) || !Number.isFinite(y)) return;
      if (x < this.minX) this.minX = x;
      if (y < this.minY) this.minY = y;
      if (x > this.maxX) this.maxX = x;
      if (y > this.maxY) this.maxY = y;
    },
  };
}

function fitSvgToBounds(svg, bounds) {
  if (!svg || bounds.minX === Infinity) return;
  const width = Math.max(1, bounds.maxX - bounds.minX);
  const height = Math.max(1, bounds.maxY - bounds.minY);
  const viewMinX = bounds.minX - PADDING_MM;
  const viewMinY = bounds.minY - TOP_PADDING_MM;
  const viewWidth = width + PADDING_MM * 2;
  const viewHeight = height + TOP_PADDING_MM + PADDING_MM;
  svg.setAttribute("viewBox", `${viewMinX} ${viewMinY} ${viewWidth} ${viewHeight}`);
  svg.setAttribute("width", viewWidth + "mm");
  svg.setAttribute("height", viewHeight + "mm");
}

function generateArmstrong(params) {
  const svg = createSvgRoot();
  svg.appendChild(
    Object.assign(document.createElementNS(NS, "metadata"), {
      textContent: JSON.stringify({
        tool: "ArmstrongBodiceWeb",
        units: "mm",
        source: "Armstrong/Bodice/armstrong_bodice_draft_v1.jsx",
      }),
    })
  );

  const layers = createArmstrongLayerStack(svg);
  const bounds = createBounds();
  const origin = {
    x: PAGE_WIDTH_MM - PAGE_MARGIN_MM,
    y: PAGE_MARGIN_MM,
  };

  const points = {};
  const defaultMarkerStyle = {
    radius: 2,
    fontSize: 3,
    offsetY: 0,
  };

  function toSvgCoords(pt) {
    const mapped = {
      x: origin.x - pt.x * MM_PER_IN,
      y: origin.y + pt.y * MM_PER_IN,
    };
    bounds.include(mapped.x, mapped.y);
    return mapped;
  }

  // replace references to use wrapper

  function pickLayer(targetLayer, opts = {}) {
    const stroke = (opts.color || "").toString().toLowerCase();
    const isBlack =
      !stroke ||
      stroke === "black" ||
      stroke === "#000" ||
      stroke === "#000000" ||
      stroke === "#111" ||
      stroke === "#111111";
    if (isBlack) {
      if (targetLayer === layers.front && layers.foundationFront) return layers.foundationFront;
      if (targetLayer === layers.back && layers.foundationBack) return layers.foundationBack;
    }
    return targetLayer;
  }

  function drawLine(targetLayer, start, end, style = "solid", opts = {}) {
    const s = toSvgCoords(start);
    const e = toSvgCoords(end);
    const actualLayer = pickLayer(targetLayer, opts);
    const attrs = {
      fill: "none",
      stroke: opts.color || "#111",
      "stroke-width": opts.width || 0.45,
      "stroke-linecap": "butt",
    };
    if (style === "dashed") {
      attrs["stroke-dasharray"] = opts.dash || "6 4";
    }
    if (opts.name) {
      attrs["data-name"] = opts.name;
    }
    const line = path(`M ${s.x} ${s.y} L ${e.x} ${e.y}`, attrs);
    actualLayer.appendChild(line);
      return line;
    }

  function drawCurve(targetLayer, start, control1, control2, end, opts = {}) {
    const s = toSvgCoords(start);
    const c1 = toSvgCoords(control1);
    const c2 = toSvgCoords(control2);
    const e = toSvgCoords(end);
    const actualLayer = pickLayer(targetLayer, opts);
    const attrs = {
      fill: "none",
      stroke: opts.color || "#2563eb",
      "stroke-width": opts.width || 0.45,
    };
    if (opts.dashed) {
      attrs["stroke-dasharray"] = opts.dash || "6 4";
    }
    if (opts.name) {
      attrs["data-name"] = opts.name;
    }
    const curve = path(`M ${s.x} ${s.y} C ${c1.x} ${c1.y} ${c2.x} ${c2.y} ${e.x} ${e.y}`, attrs);
    actualLayer.appendChild(curve);
      return curve;
    }

function drawBezierArcSegment(targetLayer, startPoint, controlPoint, endPoint, opts = {}) {
    const start = { x: startPoint.x, y: startPoint.y };
    const control = controlPoint ? { x: controlPoint.x, y: controlPoint.y } : null;
    const end = { x: endPoint.x, y: endPoint.y };
    const dxChord = end.x - start.x;
    const dyChord = end.y - start.y;
    const chordLength = Math.hypot(dxChord, dyChord);
    const tangentUnit = chordLength
      ? { x: dxChord / chordLength, y: dyChord / chordLength }
      : { x: 1, y: 0 };
    const normalUnit = { x: -tangentUnit.y, y: tangentUnit.x };

    const applyRatio = (base, ratio) => {
      const tangentComp = ratio && Number.isFinite(ratio.tangent) ? ratio.tangent : 0;
      const normalComp = ratio && Number.isFinite(ratio.normal) ? ratio.normal : 0;
      return {
        x: base.x + (tangentUnit.x * tangentComp + normalUnit.x * normalComp) * chordLength,
        y: base.y + (tangentUnit.y * tangentComp + normalUnit.y * normalComp) * chordLength,
      };
    };

    const buildControl = (base, fallback, ratio, offset, absolute) => {
      if (absolute) return { x: absolute.x, y: absolute.y };
      if (ratio && chordLength) return applyRatio(base, ratio);
      if (fallback) return { x: fallback.x, y: fallback.y };
      if (offset) return movePoint(base, offset.x || 0, offset.y || 0);
      return { x: base.x, y: base.y };
    };

    let ctrl1Fallback = null;
    let ctrl2Fallback = null;
    if (control) {
      ctrl1Fallback = {
        x: start.x + (control.x - start.x) * (2 / 3),
        y: start.y + (control.y - start.y) * (2 / 3),
      };
      ctrl2Fallback = {
        x: end.x + (control.x - end.x) * (2 / 3),
        y: end.y + (control.y - end.y) * (2 / 3),
      };
    }

    let ctrl1 = buildControl(
      start,
      ctrl1Fallback,
      opts.controlStartRatio,
      opts.controlStartOffset,
      opts.controlStartAbsolute
    );
    if (opts.controlStartOffset) {
      ctrl1 = movePoint(ctrl1, opts.controlStartOffset.x || 0, opts.controlStartOffset.y || 0);
    }
    let ctrl2 = buildControl(
      end,
      ctrl2Fallback,
      opts.controlEndRatio,
      opts.controlEndOffset,
      opts.controlEndAbsolute
    );
    if (opts.controlEndOffset) {
      ctrl2 = movePoint(ctrl2, opts.controlEndOffset.x || 0, opts.controlEndOffset.y || 0);
    }

    const s = toSvgCoords(start);
    const c1 = toSvgCoords(ctrl1);
    const c2 = toSvgCoords(ctrl2);
    const e = toSvgCoords(end);
    const actualLayer = pickLayer(targetLayer, opts);
    const attrs = {
      fill: "none",
      stroke: opts.color || "#2563eb",
      "stroke-width": opts.width || 0.45,
    };
    if (opts.dashed) {
      attrs["stroke-dasharray"] = opts.dash || "6 4";
    }
    if (opts.name) {
      attrs["data-name"] = opts.name;
    }
    const curve = path(`M ${s.x} ${s.y} C ${c1.x} ${c1.y} ${c2.x} ${c2.y} ${e.x} ${e.y}`, attrs);
    actualLayer.appendChild(curve);
    return curve;
  }

  function drawInwardArc(targetLayer, startPoint, endPoint, depth = 0.5, options = {}) {
    const start = { x: startPoint.x, y: startPoint.y };
    const end = { x: endPoint.x, y: endPoint.y };
    const dx = end.x - start.x;
    const dy = end.y - start.y;
    const length = Math.hypot(dx, dy);
    if (!length) return null;
    const mid = {
      x: (start.x + end.x) * 0.5,
      y: (start.y + end.y) * 0.5,
    };
    const nx = -dy / length;
    const ny = dx / length;
    const offset = Number.isFinite(depth) && depth !== 0 ? depth : 0.5;
    const inward = {
      x: mid.x + nx * offset,
      y: mid.y + ny * offset,
    };
    const outward = {
      x: mid.x - nx * offset,
      y: mid.y - ny * offset,
    };
    const controlPoint = options.forceOutward ? outward : inward;
    return drawBezierArcSegment(targetLayer, start, controlPoint, end, options);
  }

  function markPoint(pointName, coords, label = pointName) {
    points[pointName] = coords;
    const svgCoords = toSvgCoords(coords);
    const circle = document.createElementNS(NS, "circle");
    circle.setAttribute("cx", svgCoords.x);
    circle.setAttribute("cy", svgCoords.y);
    circle.setAttribute("r", MARKER_RADIUS_MM);
    circle.setAttribute("fill", "#0f172a");
    circle.setAttribute("stroke", "#0f172a");
    circle.setAttribute("stroke-width", 0.2);
    layers.markers.appendChild(circle);
    bounds.include(svgCoords.x - MARKER_RADIUS_MM, svgCoords.y - MARKER_RADIUS_MM);
    bounds.include(svgCoords.x + MARKER_RADIUS_MM, svgCoords.y + MARKER_RADIUS_MM);

    const labelNode = textNode(svgCoords.x, svgCoords.y + 0.4, label, {
      fill: "#fff",
      "font-size": 2.6,
      "text-anchor": "middle",
      "dominant-baseline": "middle",
      "font-weight": "600",
    });
    layers.numbers.appendChild(labelNode);
    bounds.include(svgCoords.x, svgCoords.y);
    return coords;
  }

  // --- Draft steps mirrored from ExtendScript ---
  const pointA = markPoint("A", { x: 0, y: 0 });

  const fullLengthPlusEighth = params.fullLength + 0.125;
  const pointB = markPoint("B", { x: 0, y: fullLengthPlusEighth });
  drawLine(layers.front, pointA, pointB, "solid", { name: "Full Length" });

  const pointC = markPoint("C", {
    x: pointA.x + (params.acrossShoulder - 0.125),
    y: pointA.y,
  });
  drawLine(layers.front, pointA, pointC, "solid", { name: "Front Across Shoulder" });

  const pointD = markPoint("D", {
    x: pointB.x,
    y: pointB.y - params.centreFrontLength,
  });
  drawLine(layers.front, pointB, pointD, "solid", { name: "B-D", color: "#2563eb" });

  const dLeftPoint = { x: pointD.x + 4, y: pointD.y };
  drawLine(layers.foundationFront, pointD, dLeftPoint, "dashed", { name: "Front Neck Guide 1" });

  const pointE = markPoint("E", {
    x: pointB.x + params.bustArc + 0.25,
    y: pointB.y,
  });
  drawLine(layers.front, pointB, pointE, "solid", { name: "Bust Arc" });

  const eUpPoint = { x: pointE.x, y: pointE.y - 11 };
  drawLine(layers.front, pointE, eUpPoint, "solid", { name: "Side Guide Line" });

  const cGuidePoint = { x: pointC.x, y: pointC.y + 4 };
  drawLine(layers.front, pointC, cGuidePoint, "dashed", {
    name: "Shoulder Slope Guide",
    color: "#111",
  });

  const horizontalSeparation = Math.abs(pointC.x - pointB.x);
  const shoulderSlopeLength = params.shoulderSlope;
  const verticalSpanSquared =
    shoulderSlopeLength * shoulderSlopeLength - horizontalSeparation * horizontalSeparation;
  const verticalSpan = verticalSpanSquared > 0 ? Math.sqrt(verticalSpanSquared) : 0;
  let gY = pointB.y - verticalSpan;
  if (gY < pointC.y) gY = pointC.y;
  if (gY > cGuidePoint.y) gY = cGuidePoint.y;
  const pointG = markPoint("G", { x: pointC.x, y: gY });
  drawLine(layers.foundationFront, pointB, pointG, "solid", {
    name: "Shoulder Slope",
  });

  // Point H along shoulder slope (bust depth)
  const gbVectorX = pointB.x - pointG.x;
  const gbVectorY = pointB.y - pointG.y;
  const gbLength = Math.hypot(gbVectorX, gbVectorY) || 1;
  const pointH = markPoint("H", {
    x: pointG.x + (gbVectorX / gbLength) * params.bustDepth,
    y: pointG.y + (gbVectorY / gbLength) * params.bustDepth,
  });

  // Point I on AC line with shoulder length
  const acVectorX = pointC.x - pointA.x;
  const acVectorY = pointC.y - pointA.y;
  const wX = pointA.x - pointG.x;
  const wY = pointA.y - pointG.y;
  const aQuad = acVectorX * acVectorX + acVectorY * acVectorY;
  const bQuad = 2 * (wX * acVectorX + wY * acVectorY);
  const cQuad = wX * wX + wY * wY - params.shoulderLength * params.shoulderLength;
  let discriminant = bQuad * bQuad - 4 * aQuad * cQuad;
  if (discriminant < 0) discriminant = 0;
  let t = 0;
  if (aQuad !== 0) {
    const sqrtDisc = Math.sqrt(discriminant);
    const t1 = (-bQuad + sqrtDisc) / (2 * aQuad);
    const t2 = (-bQuad - sqrtDisc) / (2 * aQuad);
    const clamp01 = (val) => val >= 0 && val <= 1;
    if (clamp01(t1) && clamp01(t2)) {
      t = Math.abs(t1 - 1) < Math.abs(t2 - 1) ? t1 : t2;
    } else if (clamp01(t1)) {
      t = t1;
    } else if (clamp01(t2)) {
      t = t2;
    } else {
      t = Math.max(0, Math.min(1, t1));
    }
  }
  const pointI = markPoint("I", {
    x: pointA.x + acVectorX * t,
    y: pointA.y + acVectorY * t,
  });
  drawLine(layers.front, pointG, pointI, "solid", { name: "Shoulder Length", color: "#2563eb" });

  const diStartControl = movePoint(pointD, 1.8, 0);
  const diEndControl = movePoint(pointI, 0, 1.8);
  drawCurve(layers.front, pointD, diStartControl, diEndControl, pointI, {
    name: "D-I Arc",
    color: "#2563eb",
  });

  const giVecX = pointI.x - pointG.x;
  const giVecY = pointI.y - pointG.y;
  const giLength = Math.hypot(giVecX, giVecY) || 1;
  const perpX = -giVecY / giLength;
  const perpY = giVecX / giLength;
  const tDrop = perpY !== 0 ? (pointD.y - pointI.y) / perpY : 0;
  const iDropPoint = {
    x: pointI.x + perpX * tDrop,
    y: pointI.y + perpY * tDrop,
  };
  drawLine(layers.foundationFront, pointI, iDropPoint, "dashed", {
    name: "Front Neck Guide 2",
  });

  const pointJ = markPoint("J", { x: pointB.x, y: pointH.y });
  const pointL = markPoint("L", {
    x: pointB.x,
    y: pointD.y + (pointJ.y - pointD.y) * 0.5,
  });

  const pointK = markPoint("K", {
    x: pointJ.x + (params.bustSpan + 0.25),
    y: pointJ.y,
  });
  const kDropPoint = { x: pointK.x, y: pointK.y + 0.625 };
  drawLine(layers.foundationFront, pointK, kDropPoint, "dashed", {
    name: "Dart Reduction",
  });
  drawLine(layers.foundationFront, pointJ, pointK, "solid", { name: "Bust Span" });

  const pointM = markPoint("M", {
    x: pointL.x + (params.acrossChest + 0.25),
    y: pointL.y,
  });
  drawLine(layers.foundationFront, pointL, pointM, "solid", { name: "Across Chest" });
  drawLine(
    layers.foundationFront,
    { x: pointM.x, y: pointM.y - 2 },
    { x: pointM.x, y: pointM.y + 1 },
    "dashed",
    { name: "Front Armhole Guide Line" }
  );

  const pointFTop = { x: pointB.x + params.dartPlacement, y: pointB.y };
  const pointF = markPoint("F", { x: pointFTop.x, y: pointFTop.y + 0.1875 });
  drawLine(layers.foundationFront, pointFTop, pointF, "solid", { name: "Dart Placement Guide" });
  drawLine(layers.front, pointB, pointF, "solid", { name: "Front Waist Line 2", color: "#2563eb" });
  drawLine(layers.front, kDropPoint, pointF, "solid", { name: "K-F", color: "#2563eb" });

  const newStrapPlusEighth = params.newStrap + 0.125;
  const dxSide = pointE.x - pointI.x;
  let dySquared = newStrapPlusEighth * newStrapPlusEighth - dxSide * dxSide;
  if (dySquared < 0) dySquared = 0;
  const dySide = Math.sqrt(dySquared);
  const candidateNY1 = pointI.y + dySide;
  const candidateNY2 = pointI.y - dySide;
  const sideUpperY = pointE.y;
  const sideLowerY = pointE.y - 11;
  const clampToSide = (y) => y >= Math.min(sideLowerY, sideUpperY) && y <= Math.max(sideLowerY, sideUpperY);
  let selectedNY = clampToSide(candidateNY1) ? candidateNY1 : candidateNY2;
  if (clampToSide(candidateNY1) && clampToSide(candidateNY2)) {
    selectedNY = Math.abs(candidateNY1 - sideUpperY) < Math.abs(candidateNY2 - sideUpperY) ? candidateNY1 : candidateNY2;
  } else if (!clampToSide(selectedNY)) {
    selectedNY = Math.max(Math.min(selectedNY, sideUpperY), sideLowerY);
  }
  const pointN = markPoint("N", { x: pointE.x, y: selectedNY });
  drawLine(layers.foundationFront, pointI, pointN, "solid", { name: "New Strap" });

  const pointO = markPoint("O", { x: pointE.x, y: pointN.y - params.sideLength });
  drawLine(layers.foundationFront, pointN, pointO, "solid", { name: "Provisional Side Length" });
  drawInwardArc(layers.front, pointG, pointO, 0, {
    name: "G-O Arc",
    color: "#2563eb",
    controlStartRatio: { tangent: 0.5295, normal: 0.494 },
    controlStartOffset: { x: -0.15, y: 0 },
    controlEndRatio: { tangent: -0.1495, normal: 0.3325 },
  });

  const bustCupOffset = resolveBustCupOffset(params.bustCup);
  const pointPBase = { x: pointN.x + bustCupOffset, y: pointN.y };
  let pointPPosition = null;
  const onVecX = pointN.x - pointO.x;
  const onVecY = pointN.y - pointO.y;
  const onDistance = Math.hypot(onVecX, onVecY);
  if (params.sideLength > 0 && onDistance > 0) {
    const r0 = params.sideLength;
    const r1 = bustCupOffset;
    if (onDistance <= r0 + r1 && onDistance >= Math.abs(r0 - r1)) {
      const aInt = (r0 * r0 - r1 * r1 + onDistance * onDistance) / (2 * onDistance);
      const hSq = Math.max(0, r0 * r0 - aInt * aInt);
      const hInt = Math.sqrt(hSq);
      const baseX = pointO.x + (aInt * (pointN.x - pointO.x)) / onDistance;
      const baseY = pointO.y + (aInt * (pointN.y - pointO.y)) / onDistance;
      const offsetX = (-(pointN.y - pointO.y) * hInt) / onDistance;
      const offsetY = ((pointN.x - pointO.x) * hInt) / onDistance;
      const candidate1 = { x: baseX + offsetX, y: baseY + offsetY };
      const candidate2 = { x: baseX - offsetX, y: baseY - offsetY };
      const dist1 =
        (candidate1.x - pointPBase.x) ** 2 + (candidate1.y - pointPBase.y) ** 2;
      const dist2 =
        (candidate2.x - pointPBase.x) ** 2 + (candidate2.y - pointPBase.y) ** 2;
      pointPPosition = dist1 <= dist2 ? candidate1 : candidate2;
    }
  }
  if (!pointPPosition) {
    const opVecX = pointPBase.x - pointO.x;
    const opVecY = pointPBase.y - pointO.y;
    const opLength = Math.hypot(opVecX, opVecY);
    if (opLength > 0 && params.sideLength > 0) {
      const scale = params.sideLength / opLength;
      pointPPosition = {
        x: pointO.x + opVecX * scale,
        y: pointO.y + opVecY * scale,
      };
    } else {
      pointPPosition = { x: pointPBase.x, y: pointPBase.y };
    }
  }
  drawLine(layers.front, pointO, pointPPosition, "solid", { name: "Side Length OP", color: "#2563eb" });
  const pointP = markPoint("P", pointPPosition);

  const waistArcPlusQuarter = params.waistArc + 0.25;
  const bToFValue = Math.abs(pointFTop.x - pointB.x);
  const qDistance = waistArcPlusQuarter - bToFValue;
  const pfVecX = pointF.x - pointP.x;
  const pfVecY = pointF.y - pointP.y;
  const pfLength = Math.hypot(pfVecX, pfVecY);
  let pointQPosition = { x: pointP.x, y: pointP.y };
  if (pfLength > 0) {
    const scaleToQ = qDistance / pfLength;
    pointQPosition = {
      x: pointP.x + pfVecX * scaleToQ,
      y: pointP.y + pfVecY * scaleToQ,
    };
  }
  let pointQTarget = { x: pointQPosition.x, y: pointQPosition.y };
  const kfVecX = pointF.x - kDropPoint.x;
  const kfVecY = pointF.y - kDropPoint.y;
  const kfLength = Math.hypot(kfVecX, kfVecY);
  const kqVecX = pointQTarget.x - kDropPoint.x;
  const kqVecY = pointQTarget.y - kDropPoint.y;
  const kqLength = Math.hypot(kqVecX, kqVecY);
  if (kfLength > 0 && kqLength > 0) {
    const kqScale = kfLength / kqLength;
    pointQTarget = {
      x: kDropPoint.x + kqVecX * kqScale,
      y: kDropPoint.y + kqVecY * kqScale,
    };
  }
  drawLine(layers.front, kDropPoint, pointQTarget, "solid", { name: "KQ Equivalent", color: "#2563eb" });
  const pointQ = markPoint("Q", pointQTarget);
  drawLine(layers.front, pointP, pointQ, "solid", { name: "Front Waist Line 1", color: "#2563eb" });

  // --- Back Draft ---
  const backFullLength = firstNumber(params.fullLengthBack, params.fullLength);
  const backDartPlacement = firstNumber(params.dartPlacementBack, params.dartPlacement, 0);
  const backWaistArcValue = firstNumber(params.waistArcBack, params.waistArc, 0);
  const dartIntake = 1.5;
  const backJOffset = backWaistArcValue + 1.5 + 0.25;
  const centreBackLength = firstNumber(params.centreFrontLengthBack, params.centreFrontLength, backFullLength);
  const backAcrossShoulder = firstNumber(params.acrossShoulderBack, params.acrossShoulder);
  const backBustArc = firstNumber(params.bustArcBack, params.bustArc);
  const backShoulderSlope = firstNumber(params.shoulderSlopeBack, params.shoulderSlope) + 0.125;
  const shoulderLengthPlusHalf = firstNumber(params.shoulderLengthBack, params.shoulderLength, 0) + 0.5;
  const backSideLength = firstNumber(params.sideLengthBack, params.sideLength, 0);
  const backAcrossChest = firstNumber(params.acrossChestBack, params.acrossChest);
  const backNeckPlusEighth = firstNumber(params.backNeck, 0) + 0.125;

  const backOrigin = {
    x: pointA.x - BACK_OFFSET_IN,
    y: pointB.y - (backFullLength + 0.125),
  };
  const backCoord = (dx, dy) => ({
    x: backOrigin.x + dx,
    y: backOrigin.y + dy,
  });

  const backPointA = markPoint("BA", backCoord(0, 0), "A");
  const backPointB = markPoint("BB", backCoord(0, backFullLength), "B");
  drawLine(layers.foundationBack, backPointA, backPointB, "solid", { name: "Back Full Length" });

  const backPointJ = markPoint("BJ", backCoord(-backJOffset, backFullLength), "J");
  const backPointL = markPoint("BL", backCoord(-(backDartPlacement + dartIntake / 2), backFullLength), "L");
  const backPointM = markPoint("BM", backCoord(-backJOffset, backFullLength + 0.1875), "M");
  const backPointD = markPoint("BD", backCoord(0, backFullLength - centreBackLength), "D");
  drawLine(layers.back, backPointD, backPointB, "solid", { name: "Centre Back", color: "#2563eb" });

  const backPointC = markPoint("BC", backCoord(-backAcrossShoulder, 0), "C");
  drawLine(layers.foundationBack, backPointA, backPointC, "solid", { name: "Back Across Shoulder" });
  drawLine(
    layers.foundationBack,
    backPointD,
    { x: backPointD.x - 4, y: backPointD.y },
    "dashed",
    { name: "Back Neck Guide" }
  );
  drawLine(
    layers.foundationBack,
    backPointC,
    { x: backPointC.x, y: backPointC.y + 6 },
    "dashed",
    { name: "Back Shoulder Slope Guide" }
  );

  const backPointE = markPoint("BE", backCoord(-(backBustArc + 0.75), backFullLength), "E");
  drawLine(layers.foundationBack, backPointB, backPointE, "solid", { name: "Back Arc" });
  const backEUp = backCoord(-(backBustArc + 0.75), backFullLength - 10);
  drawLine(layers.foundationBack, backPointE, backEUp, "solid", { name: "Back Side Guide Line" });

  const backPointF = markPoint("BF", backCoord(-backNeckPlusEighth, 0), "F");
  drawBezierArcSegment(layers.back, backPointD, backPointD, backPointF, {
    name: "Back Neck Curve",
    color: "#2563eb",
    controlStartRatio: { tangent: 0.5164137931, normal: -0.1390344828 },
    controlEndRatio: { tangent: -0.1807448276, normal: -0.191337931 },
  });

  const dxSlope = backPointC.x - backPointB.x;
  const baseDy = backPointC.y - backPointB.y;
  const aSlope = 1;
  const bSlope = 2 * baseDy;
  const cSlope = baseDy * baseDy + dxSlope * dxSlope - backShoulderSlope * backShoulderSlope;
  const discriminantSlope = bSlope * bSlope - 4 * aSlope * cSlope;
  let tSlope = 0;
  if (discriminantSlope >= 0) {
    const sqrtDisc = Math.sqrt(discriminantSlope);
    const t1 = (-bSlope + sqrtDisc) / (2 * aSlope);
    const t2 = (-bSlope - sqrtDisc) / (2 * aSlope);
  const withinGuide = (val) => val >= 0 && val <= 6;
    if (withinGuide(t1) && withinGuide(t2)) {
      tSlope = Math.max(t1, t2);
    } else if (withinGuide(t1)) {
      tSlope = t1;
    } else if (withinGuide(t2)) {
      tSlope = t2;
    } else {
      tSlope = Math.max(0, Math.min(6, t1));
    }
  }
  const backPointG = markPoint("BG", backCoord(-backAcrossShoulder, tSlope), "G");
  drawLine(layers.foundationBack, backPointB, backPointG, "solid", { name: "Back Shoulder Slope" });

  const backFGVecX = backPointG.x - backPointF.x;
  const backFGVecY = backPointG.y - backPointF.y;
  const backFGLen = Math.hypot(backFGVecX, backFGVecY) || 1;
  const backPointH = markPoint(
    "BH",
    {
      x: backPointF.x + (backFGVecX * shoulderLengthPlusHalf) / backFGLen + 0.3,
      y: backPointF.y + (backFGVecY * shoulderLengthPlusHalf) / backFGLen,
    },
    "H"
  );
  drawLine(layers.foundationBack, backPointF, backPointH, "solid", { name: "Back Shoulder Length" });

  const backPointP = markPoint("BP", {
    x: (backPointF.x + backPointH.x) / 2,
    y: (backPointF.y + backPointH.y) / 2,
  });
  const fhVecX = backPointH.x - backPointF.x;
  const fhVecY = backPointH.y - backPointF.y;
  const fhLength = Math.hypot(fhVecX, fhVecY) || 1;
  const fhUnitX = fhVecX / fhLength;
  const fhUnitY = fhVecY / fhLength;
  const backPointRBase = {
    x: backPointP.x - fhUnitX * 0.25,
    y: backPointP.y - fhUnitY * 0.25,
  };
  const backPointABase = {
    x: backPointP.x + fhUnitX * 0.25,
    y: backPointP.y + fhUnitY * 0.25,
  };

  const backPointMtoE = backPointE.x - backPointM.x;
  let backPointN = {
    x: backPointE.x,
    y: backPointM.y - Math.sqrt(Math.max(0, backSideLength * backSideLength - backPointMtoE * backPointMtoE)),
  };
  if (backPointN.y < backEUp.y) backPointN.y = backEUp.y;
  backPointN = markPoint("BN", backPointN, "N");
  drawLine(layers.back, backPointM, backPointN, "solid", { name: "Back Side Length", color: "#2563eb" });
  drawBezierArcSegment(layers.back, backPointH, backPointH, backPointN, {
    name: "Back Armhole Curve",
    color: "#2563eb",
    controlStartRatio: { tangent: 0.5836719092, normal: -0.3423131657 },
    controlEndRatio: { tangent: -0.1240008345, normal: -0.3063856393 },
  });

  const backMNLength = Math.hypot(backPointN.x - backPointM.x, backPointN.y - backPointM.y);
  const backPointO = markPoint(
    "BO",
    {
      x: backPointL.x,
      y: backPointL.y - Math.max(0, backMNLength - 1),
    },
    "O"
  );
  drawLine(layers.foundationBack, backPointL, backPointO, "dashed", { name: "Back Dart Leg Guide" });

  const backDLength = Math.hypot(backPointB.x - backPointD.x, backPointB.y - backPointD.y);
  const backPointS = markPoint("BS", backCoord(0, backFullLength - centreBackLength + backDLength / 4), "S");
  const backPointT = markPoint("BT", backCoord(-(backAcrossChest + 0.25), backPointS.y), "T");
  drawLine(layers.foundationBack, backPointS, backPointT, "solid", { name: "Across Back" });
  drawLine(
    layers.foundationBack,
    backCoord(-(backAcrossChest + 0.25), backPointT.y - 1),
    backCoord(-(backAcrossChest + 0.25), backPointT.y + 4),
    "dashed",
    { name: "Back Armhole Guideline" }
  );

  const backPointIBase = backCoord(-backDartPlacement, backFullLength);
  const backPointKBase = backCoord(-(backDartPlacement + dartIntake), backFullLength);
  const extensionLength = 0.125;
  const backPointI = markPoint("BI", extendPoint(backPointO, backPointIBase, extensionLength), "I");
  const backPointK = markPoint("BK", extendPoint(backPointO, backPointKBase, extensionLength), "K");

  drawLine(layers.back, backPointB, backPointI, "solid", { name: "Back Waist Line 1", color: "#2563eb" });
  drawLine(layers.back, backPointK, backPointM, "solid", { name: "Back Waist Line 2", color: "#2563eb" });
  drawLine(layers.foundationBack, backPointP, backPointO, "dashed", { name: "Back Dart Center Guide" });

  drawLine(layers.back, backPointO, backPointI, "solid", { name: "Back Waist Dart Left Leg", color: "#2563eb" });
  drawLine(layers.back, backPointO, backPointK, "solid", { name: "Back Waist Dart Right Leg", color: "#2563eb" });

  const backPOVecX = backPointO.x - backPointP.x;
  const backPOVecY = backPointO.y - backPointP.y;
  const backPOLen = Math.hypot(backPOVecX, backPOVecY) || 1;
  const backPointQ = markPoint(
    "BQ",
    {
      x: backPointP.x + (backPOVecX / backPOLen) * 3,
      y: backPointP.y + (backPOVecY / backPOLen) * 3,
    },
    "Q"
  );

  const dartExtension = 0.125;
  const backPointR = markPoint("BR", extendPoint(backPointQ, backPointRBase, dartExtension), "R");
  const backPointa = markPoint("Ba", extendPoint(backPointQ, backPointABase, dartExtension), "a");
  drawLine(layers.back, backPointQ, backPointR, "solid", {
    name: "Back Shoulder Dart Left Leg",
    color: "#2563eb",
  });
  drawLine(layers.back, backPointQ, backPointa, "solid", {
    name: "Back Shoulder Dart Right Leg",
    color: "#2563eb",
  });

  drawLine(layers.back, backPointa, backPointH, "solid", { name: "Back Shoulder Line 2", color: "#2563eb" });
  drawLine(layers.back, backPointF, backPointR, "solid", { name: "Back Shoulder Line 1", color: "#2563eb" });
  // ------------------------------------------------

  applyLayerVisibility(layers, params);
  fitSvgToBounds(svg, bounds);

  return svg;
}

function generateAldrich(params) {
  const svg = createSvgRoot();
  svg.appendChild(
    Object.assign(document.createElementNS(NS, "metadata"), {
      textContent: JSON.stringify({
        tool: "AldrichCloseFittingBodiceWeb",
        units: "mm",
        source: "Aldrich/Bodice/aldrich_close_fitting_bodice_with_waist_shaping_v1.jsx",
      }),
    })
  );

  const layers = createAldrichLayerStack(svg);
  const bounds = createBounds();
  const origin = {
    x: PAGE_MARGIN_MM,
    y: PAGE_MARGIN_MM,
  };

  const points = {};

  function toSvgCoords(pt) {
    const mapped = {
      x: origin.x + pt.x * CM_TO_MM,
      y: origin.y + pt.y * CM_TO_MM,
    };
    bounds.include(mapped.x, mapped.y);
    return mapped;
  }

  function applyArtHandle(point, handle, signX = 1, signY = 1) {
    const artY = -point.y;
    const nextArtY = artY + signY * handle.y;
    return {
      x: point.x + signX * handle.x,
      y: -nextArtY,
    };
  }

  function drawSegment(layerNode, start, end, opts = {}) {
    if (!layerNode || !start || !end) return null;
    const s = toSvgCoords(start);
    const e = toSvgCoords(end);
    const attrs = {
      fill: "none",
      stroke: opts.color || "#111",
      "stroke-width": opts.width || 0.45,
      "stroke-linecap": "butt",
    };
    if (opts.dashed) {
      attrs["stroke-dasharray"] = opts.dash || "6 4";
    }
    const pathEl = path(`M ${s.x} ${s.y} L ${e.x} ${e.y}`, attrs);
    if (opts.name) pathEl.setAttribute("data-name", opts.name);
    layerNode.appendChild(pathEl);
    return pathEl;
  }

  const defaultMarkerStyle = {
    radius: MARKER_RADIUS_MM * 1.05,
    fontSize: LABEL_FONT_SIZE_MM * 0.8,
    offsetY: 0,
  };

  function drawCurveSegment(layerNode, start, control1, control2, end, opts = {}) {
    if (!layerNode || !start || !end || !control1 || !control2) return null;
    const s = toSvgCoords(start);
    const c1 = toSvgCoords(control1);
    const c2 = toSvgCoords(control2);
    const e = toSvgCoords(end);
    const attrs = {
      fill: "none",
      stroke: opts.color || ALDRICH_COLORS.primary,
      "stroke-width": opts.width || 0.45,
    };
    if (opts.dashed) {
      attrs["stroke-dasharray"] = opts.dash || "6 4";
    }
    const pathEl = path(`M ${s.x} ${s.y} C ${c1.x} ${c1.y} ${c2.x} ${c2.y} ${e.x} ${e.y}`, attrs);
    if (opts.name) pathEl.setAttribute("data-name", opts.name);
    layerNode.appendChild(pathEl);
    return pathEl;
  }

  function drawPolylineSegment(layerNode, coordList = [], opts = {}) {
    if (!layerNode || !coordList.length) return null;
    const pointsList = coordList.map((coord) => toSvgCoords(coord));
    const d = pointsList
      .map((pt, idx) => `${idx === 0 ? "M" : "L"} ${pt.x} ${pt.y}`)
      .join(" ");
    const attrs = {
      fill: "none",
      stroke: opts.color || "#111",
      "stroke-width": opts.width || 0.45,
    };
    if (opts.dashed) {
      attrs["stroke-dasharray"] = opts.dash || "6 4";
    }
    if (opts.name) {
      attrs["data-name"] = opts.name;
    }
    const pathEl = path(d, attrs);
    layerNode.appendChild(pathEl);
    return pathEl;
  }

  function addMarker(id, coords, style = {}) {
    if (!layers.markers || !layers.numbers || !coords) return;
    const anchor = toSvgCoords(coords);
    const radius =
      typeof style.radius === "number" && Number.isFinite(style.radius) ? style.radius : MARKER_RADIUS_MM;
    const fontSize =
      typeof style.fontSize === "number" && Number.isFinite(style.fontSize) ? style.fontSize : LABEL_FONT_SIZE_MM;
    const offsetY =
      typeof style.offsetY === "number" && Number.isFinite(style.offsetY)
        ? style.offsetY
        : LABEL_FONT_SIZE_MM * 0.35;
    const circle = document.createElementNS(NS, "circle");
    circle.setAttribute("cx", anchor.x);
    circle.setAttribute("cy", anchor.y);
    circle.setAttribute("r", radius);
    circle.setAttribute("fill", "#111");
    circle.setAttribute("stroke", "#111");
    circle.setAttribute("stroke-width", 0.25);
    layers.markers.appendChild(circle);

    const label = textNode(anchor.x, anchor.y + offsetY, String(id), {
      "font-size": fontSize,
      fill: "#fff",
      "text-anchor": "middle",
      "dominant-baseline": "middle",
    });
    layers.numbers.appendChild(label);
  }

  function registerPoint(id, coords, opts = {}) {
    points[id] = coords;
    if (!opts.skipMarker) {
      addMarker(id, coords, opts.markerStyle || defaultMarkerStyle);
    }
    return coords;
  }

  function registerLetterPoint(id, coords, label) {
    points[id] = coords;
    if (layers.letters && coords) {
      const anchor = toSvgCoords(coords);
      const letterNode = textNode(anchor.x, anchor.y, String(label || id), {
        "font-size": LABEL_FONT_SIZE_MM,
        fill: "#111",
        "text-anchor": "middle",
      });
      layers.letters.appendChild(letterNode);
    }
    return coords;
  }

  const depth01 = 1.5;
  const depth12 = params.armscyeDepth + 0.5;
  const vertical02 = depth01 + depth12;
  const halfBust = params.bust / 2;
  const width23 = halfBust + params.bustEase;
  const napeToWaist = params.napeToWaist;

  registerPoint("0", { x: 0, y: 0 });
  registerPoint("1", { x: 0, y: depth01 });
  registerPoint("2", { x: 0, y: vertical02 });
  registerPoint("3", { x: width23, y: vertical02 });
  registerPoint("4", { x: width23, y: 0 });
  registerPoint("5", { x: 0, y: depth01 + napeToWaist });
  registerPoint("6", { x: width23, y: depth01 + napeToWaist });
  const pointC = registerLetterPoint("c", { x: points["6"].x, y: points["6"].y + 1 }, "c");

  registerPoint("9", { x: params.neckSize / 5 - 0.2, y: points["0"].y });
  registerPoint("10", {
    x: points["1"].x,
    y: points["1"].y + params.armscyeDepth / 5 - 0.7,
  });
  registerPoint("20", {
    x: points["4"].x - (params.neckSize / 5 - 0.7),
    y: points["4"].y,
  });
  registerPoint("21", {
    x: points["4"].x,
    y: points["4"].y + (params.neckSize / 5 - 0.2),
  });
  registerPoint("22", {
    x: points["3"].x - params.chest / 2 - params.frontNeckDart / 2,
    y: points["3"].y,
  });

  const distanceB = computeAldrichPointBDistance(params.bust);
  if (distanceB > 0) {
    const diag = distanceB / Math.SQRT2;
    const pointB = registerLetterPoint(
      "b",
      {
        x: points["22"].x - diag,
        y: points["22"].y - diag,
      },
      "b"
    );
    drawSegment(layers.foundation, points["22"], pointB, {
      dashed: true,
      color: ALDRICH_COLORS.guide,
      name: "Front Armhole Guideline",
    });
  }

  registerPoint("23", midpoint(points["3"], points["22"]));
  registerPoint("24", { x: points["23"].x, y: points["5"].y }, { skipMarker: true });
  registerPoint("26", { x: points["23"].x, y: points["23"].y + 2.5 });
  registerPoint("27", {
    x: points["20"].x - params.frontNeckDart,
    y: points["20"].y,
  });

  registerPoint("31", {
    x: points["22"].x,
    y: points["22"].y + ((points["21"].y - points["3"].y) / 3),
  });
  drawSegment(layers.foundation, points["5"], pointC, {
    color: ALDRICH_COLORS.primary,
    name: "Waist Drop Guide",
  });
  drawSegment(layers.foundation, points["2"], points["3"], {
    dashed: true,
    color: ALDRICH_COLORS.guide,
    name: "Bust Line",
  });
  drawSegment(layers.foundation, points["5"], points["6"], {
    dashed: true,
    color: ALDRICH_COLORS.guide,
    name: "Waistline",
  });
  drawSegment(layers.back, points["1"], points["5"], {
    color: ALDRICH_COLORS.primary,
    name: "Centre Back (CB)",
  });
  drawSegment(layers.foundation, points["4"], pointC, {
    color: ALDRICH_COLORS.primary,
    name: "CF Line",
  });
  drawSegment(layers.front, points["21"], pointC, {
    color: ALDRICH_COLORS.primary,
    name: "Centre Front (CF)",
  });
  drawSegment(layers.foundation, points["0"], points["4"], {
    color: ALDRICH_COLORS.primary,
    name: "0 - 4",
  });
  drawSegment(layers.foundation, points["0"], points["5"], {
    color: ALDRICH_COLORS.primary,
    name: "Foundation Centre Back",
  });

  const shoulderDeltaY = Math.abs(points["10"].y - points["9"].y);
  const shoulderPlusOne = params.shoulder + 1;
  let shoulderHorizontal = 0;
  if (shoulderPlusOne > shoulderDeltaY) {
    const squared = Math.pow(shoulderPlusOne, 2) - Math.pow(shoulderDeltaY, 2);
    shoulderHorizontal = Math.sqrt(Math.max(squared, 0));
  }
  registerPoint("11", {
    x: points["9"].x + shoulderHorizontal,
    y: points["10"].y,
  });
  registerPoint("12", midpoint(points["9"], points["11"]));
  const guide12Down = { x: points["12"].x, y: points["12"].y + 5 };
  drawSegment(layers.foundation, points["12"], guide12Down, {
    dashed: true,
    color: ALDRICH_COLORS.guide,
    name: "Back Shoulder Guide",
  });
  registerPoint("13", { x: guide12Down.x - 1, y: guide12Down.y });
  drawSegment(layers.foundation, guide12Down, points["13"], {
    dashed: true,
    color: ALDRICH_COLORS.guide,
  });
  const shoulderSegmentLength = distanceBetween(points["9"], points["11"]);
  const shoulderDartHalfWidth = 0.5;
  if (shoulderSegmentLength > 0.0001) {
    const shoulderUnitX = (points["11"].x - points["9"].x) / shoulderSegmentLength;
    const shoulderUnitY = (points["11"].y - points["9"].y) / shoulderSegmentLength;
    const shoulderDartLeft = {
      x: points["12"].x - shoulderUnitX * shoulderDartHalfWidth,
      y: points["12"].y - shoulderUnitY * shoulderDartHalfWidth,
    };
    const shoulderDartRight = {
      x: points["12"].x + shoulderUnitX * shoulderDartHalfWidth,
      y: points["12"].y + shoulderUnitY * shoulderDartHalfWidth,
    };
    const leftLegLength = distanceBetween(shoulderDartLeft, points["13"]);
    const rightVector = {
      x: shoulderDartRight.x - points["13"].x,
      y: shoulderDartRight.y - points["13"].y,
    };
    const rightVectorLength = Math.hypot(rightVector.x, rightVector.y);
    let shoulderDartRightAdjusted = shoulderDartRight;
    if (rightVectorLength > 0.0001 && leftLegLength > 0.0001) {
      const scale = leftLegLength / rightVectorLength;
      shoulderDartRightAdjusted = {
        x: points["13"].x + rightVector.x * scale,
        y: points["13"].y + rightVector.y * scale,
      };
    }
    drawPolylineSegment(layers.back, [shoulderDartLeft, points["13"], shoulderDartRightAdjusted], {
      color: ALDRICH_COLORS.primary,
      name: "Back Shoulder Dart",
    });
    drawSegment(layers.back, points["9"], shoulderDartLeft, {
      color: ALDRICH_COLORS.primary,
      name: "Back Shoulder Line",
    });
    drawSegment(layers.back, shoulderDartRightAdjusted, points["11"], {
      color: ALDRICH_COLORS.primary,
      name: "Back Shoulder Dart Right to 11",
    });
  } else {
    drawSegment(layers.back, points["9"], points["11"], {
      color: ALDRICH_COLORS.primary,
      name: "Back Shoulder Line",
    });
  }

  const backNeckChord = distanceBetween(points["1"], points["9"]);
  const handle1 = clampHandleToChordCm(4.3, 0, backNeckChord);
  const handle9 = clampHandleToChordCm(-1, -1, backNeckChord);
  const backNeckControl1 = applyArtHandle(points["1"], handle1, 1, 1);
  const backNeckControl2 = applyArtHandle(points["9"], handle9, 1, 1);
  drawCurveSegment(
    layers.back,
    points["1"],
    backNeckControl1,
    backNeckControl2,
    points["9"],
    { name: "Back Neck Curve", color: ALDRICH_COLORS.primary }
  );

  const frontNeckChord = distanceBetween(points["20"], points["21"]);
  const handle20 = clampHandleToChordCm(0, -4, frontNeckChord);
  const handle21 = clampHandleToChordCm(-4, 0, frontNeckChord);
  const frontNeckControl1 = applyArtHandle(points["20"], handle20, 1, 1);
  const frontNeckControl2 = applyArtHandle(points["21"], handle21, 1, 1);
  drawCurveSegment(
    layers.front,
    points["20"],
    frontNeckControl1,
    frontNeckControl2,
    points["21"],
    { name: "Front Neck Curve", color: ALDRICH_COLORS.primary }
  );

  drawSegment(layers.front, points["20"], points["26"], {
    color: ALDRICH_COLORS.primary,
    name: "Front Neck Dart Left Leg",
  });
  drawSegment(layers.front, points["27"], points["26"], {
    color: ALDRICH_COLORS.primary,
    name: "Front Neck Dart Right Leg",
  });

  registerPoint("28", { x: points["11"].x, y: points["11"].y + 1.5 });
  registerPoint("29", { x: points["28"].x + 10, y: points["28"].y });
  drawSegment(layers.foundation, points["11"], points["28"], {
    dashed: true,
    color: ALDRICH_COLORS.guide,
    name: "Back Shoulder Drop",
  });
  drawSegment(layers.foundation, points["28"], points["29"], {
    dashed: true,
    color: ALDRICH_COLORS.guide,
    name: "Shoulder Balance Line",
  });
  const frontShoulderDeltaY = Math.abs(points["28"].y - points["27"].y);
  const frontShoulderHorizontalSq = Math.pow(params.shoulder, 2) - Math.pow(frontShoulderDeltaY, 2);
  let frontShoulderHorizontal = 0;
  if (frontShoulderHorizontalSq > 0) {
    frontShoulderHorizontal = Math.sqrt(frontShoulderHorizontalSq);
  }
  let candidateX30 = points["27"].x - frontShoulderHorizontal;
  const minFrontX = Math.min(points["28"].x, points["29"].x);
  const maxFrontX = Math.max(points["28"].x, points["29"].x);
  if (candidateX30 < minFrontX) candidateX30 = minFrontX;
  if (candidateX30 > maxFrontX) candidateX30 = maxFrontX;
  registerPoint("30", { x: candidateX30, y: points["28"].y });
  drawSegment(layers.front, points["27"], points["30"], {
    color: ALDRICH_COLORS.primary,
    name: "Front Shoulder Line",
  });

  registerPoint("14", {
    x: points["2"].x + params.backWidth / 2 + 0.5,
    y: points["2"].y,
  });
  const pointADistance = computeAldrichPointADistance(params.bust);
  if (pointADistance > 0) {
    const diagComponent = pointADistance / Math.SQRT2;
    const pointA = registerLetterPoint(
      "a",
      {
        x: points["14"].x + diagComponent,
        y: points["14"].y - diagComponent,
      },
      "a"
    );
    drawSegment(layers.foundation, points["14"], pointA, {
      dashed: true,
      color: ALDRICH_COLORS.guide,
      name: "Back Armhole Guideline",
    });
  }
  registerPoint("15", { x: points["14"].x, y: points["10"].y });
  registerPoint("16", midpoint(points["14"], points["15"]));
  registerPoint("17", midpoint(points["2"], points["14"]));
  registerPoint("32", midpoint(points["14"], points["22"]));
  drawSegment(layers.foundation, points["14"], points["15"], {
    dashed: true,
    color: ALDRICH_COLORS.guide,
    name: "Back Width Line",
  });

  const backArmholeChord = distanceBetween(points["11"], points["32"]);
  const backHandle11 = clampHandleToChordCm(3.2, 9.79, backArmholeChord);
  const backHandle32 = clampHandleToChordCm(6.15, 0, backArmholeChord);
  const backArmControl1 = applyArtHandle(points["11"], backHandle11, -1, -1);
  const backArmControl2 = applyArtHandle(points["32"], backHandle32, -1, -1);
  drawCurveSegment(
    layers.back,
    points["11"],
    backArmControl1,
    backArmControl2,
    points["32"],
    { name: "Back Armhole Curve", color: ALDRICH_COLORS.primary }
  );

  const frontArmholeChord = distanceBetween(points["30"], points["32"]);
  const frontHandle30 = clampHandleToChordCm(7, 11.25, frontArmholeChord);
  const frontHandle32 = clampHandleToChordCm(6.47, 0, frontArmholeChord);
  const frontArmControl1 = applyArtHandle(points["30"], frontHandle30, 1, -1);
  const frontArmControl2 = applyArtHandle(points["32"], frontHandle32, 1, -1);
  drawCurveSegment(
    layers.front,
    points["30"],
    frontArmControl1,
    frontArmControl2,
    points["32"],
    { name: "Front Armhole Curve", color: ALDRICH_COLORS.primary }
  );

  drawSegment(layers.foundation, points["22"], points["31"], {
    dashed: true,
    color: ALDRICH_COLORS.guide,
    name: "Chest Line",
  });

  const waistLineY = points["5"].y;
  registerPoint("33", { x: points["23"].x, y: waistLineY });

  const line5 = points["5"];
  if (line5 && pointC) {
    const intersectionSpecs = [
      {
        id: "d",
        x: points["17"] ? points["17"].x : null,
        top: points["17"],
        bottom: points["17"] ? { x: points["17"].x, y: pointC.y } : null,
      },
      {
        id: "e",
        x: points["23"] ? points["23"].x : null,
        top: points["23"],
        bottom: points["23"] ? { x: points["23"].x, y: pointC.y } : null,
      },
      {
        id: "f",
        x: points["32"] ? points["32"].x : null,
        top: points["32"],
        bottom: points["32"] ? { x: points["32"].x, y: pointC.y } : null,
      },
    ];
    intersectionSpecs.forEach((spec) => {
      if (!spec.top || !spec.bottom || !isFiniteNumber(spec.x)) return;
      const intersectionPoint = intersectLineWithVertical(line5, pointC, spec.x);
      if (!intersectionPoint) return;
      if (!valueBetween(intersectionPoint.y, spec.top.y, spec.bottom.y, 0.001)) return;
      registerLetterPoint(spec.id, intersectionPoint, spec.id);
    });
  }

  const frontWaistDartWidth = params.frontWaistDart;
  if (points["e"] && Number.isFinite(frontWaistDartWidth) && frontWaistDartWidth > 0) {
    const half = frontWaistDartWidth / 2;
    const dartBaseLeft = { x: points["e"].x - half, y: points["e"].y };
    const dartBaseRight = { x: points["e"].x + half, y: points["e"].y };
    const dartBackOff = Number.isFinite(params.frontWaistDartBackOff) ? params.frontWaistDartBackOff : 2.5;
    const dartApex = { x: points["e"].x, y: points["26"].y + dartBackOff };
    drawSegment(layers.front, dartBaseLeft, dartApex, {
      color: ALDRICH_COLORS.primary,
      name: "Front Waist Dart Left Leg",
    });
    drawSegment(layers.front, dartBaseRight, dartApex, {
      color: ALDRICH_COLORS.primary,
      name: "Front Waist Dart Right Leg",
    });
    drawSegment(layers.foundation, dartApex, points["e"], {
      dashed: true,
      color: ALDRICH_COLORS.guide,
      name: "Front Waist Dart Bisector",
    });
  }

  const backWaistDartWidth = params.backWaistDart;
  if (points["d"] && points["17"] && Number.isFinite(backWaistDartWidth) && backWaistDartWidth > 0) {
    const half = backWaistDartWidth / 2;
    const backLeftBase = { x: points["d"].x - half, y: points["d"].y };
    const backRightBase = { x: points["d"].x + half, y: points["d"].y };
    const backApex = { x: points["17"].x, y: points["17"].y };
    drawSegment(layers.back, backLeftBase, backApex, {
      color: ALDRICH_COLORS.primary,
      name: "Back Waist Dart Left Leg",
    });
    drawSegment(layers.back, backRightBase, backApex, {
      color: ALDRICH_COLORS.primary,
      name: "Back Waist Dart Right Leg",
    });
    drawSegment(layers.foundation, points["17"], points["d"], {
      dashed: true,
      color: ALDRICH_COLORS.guide,
      name: "Back Waist Dart Bisector",
    });
  }

  const waistLineStart = points["5"];
  const waistLineEnd = pointC;
  const waistLineMinX =
    waistLineStart && waistLineEnd ? Math.min(waistLineStart.x, waistLineEnd.x) : null;
  const waistLineMaxX =
    waistLineStart && waistLineEnd ? Math.max(waistLineStart.x, waistLineEnd.x) : null;

  const clampWaistX = (x) => {
    if (waistLineMinX == null || waistLineMaxX == null || !Number.isFinite(x)) return x;
    const minX = Math.min(waistLineMinX, waistLineMaxX);
    const maxX = Math.max(waistLineMinX, waistLineMaxX);
    if (x < minX) return minX;
    if (x > maxX) return maxX;
    return x;
  };

  const waistLinePointAtX = (x) => {
    if (!waistLineStart || !waistLineEnd) return null;
    const deltaX = waistLineEnd.x - waistLineStart.x;
    if (Math.abs(deltaX) < 0.0001) {
      return { x, y: waistLineStart.y };
    }
    const slope = (waistLineEnd.y - waistLineStart.y) / deltaX;
    return {
      x,
      y: waistLineStart.y + slope * (x - waistLineStart.x),
    };
  };

  const waistAxisPoint =
    points["f"] ||
    (points["32"] && Number.isFinite(points["32"].x)
      ? waistLinePointAtX(clampWaistX(points["32"].x))
      : waistLineStart && waistLineEnd
      ? waistLinePointAtX(waistLineStart.x)
      : null);

  if (waistAxisPoint && points["32"]) {
    drawSegment(layers.foundation, points["32"], waistAxisPoint, {
      dashed: true,
      color: ALDRICH_COLORS.guide,
      name: "Side",
    });
    const sideApex = { x: points["32"].x, y: points["32"].y };

    const backSideWidth = params.backSideWaistDart;
    if (Number.isFinite(backSideWidth) && backSideWidth > 0) {
      const baseX = clampWaistX(waistAxisPoint.x - backSideWidth);
      const sideLeftBase = waistLinePointAtX(baseX) || { x: baseX, y: waistAxisPoint.y };
      drawSegment(layers.back, sideLeftBase, sideApex, {
        color: ALDRICH_COLORS.primary,
        name: "Back Side Waist Dart",
      });
      if (points["5"]) {
        drawSegment(layers.back, points["5"], sideLeftBase, {
          color: ALDRICH_COLORS.primary,
          name: "Back Waist Line",
        });
      }
    }

    const frontSideWidth = params.frontSideWaistDart;
    if (Number.isFinite(frontSideWidth) && frontSideWidth > 0) {
      const baseX = clampWaistX(waistAxisPoint.x + frontSideWidth);
      const sideRightBase = waistLinePointAtX(baseX) || { x: baseX, y: waistAxisPoint.y };
      drawSegment(layers.front, sideRightBase, sideApex, {
        color: ALDRICH_COLORS.primary,
        name: "Front Side Waist Dart",
      });
      if (pointC) {
        drawSegment(layers.front, pointC, sideRightBase, {
          color: ALDRICH_COLORS.primary,
          name: "Front Waist Line",
        });
      }
    }
  }

  drawSegment(layers.foundation, points["10"], points["1"], {
    color: ALDRICH_COLORS.primary,
    name: "Back Blade",
  });
  const backSquareEnd = { x: points["10"].x + width23 / 2, y: points["10"].y };
  drawSegment(layers.foundation, points["10"], backSquareEnd, {
    dashed: true,
    color: ALDRICH_COLORS.guide,
    name: "Back Blade Guide",
  });

  applyLayerVisibility(layers, params);
  fitSvgToBounds(svg, bounds);
  return svg;
}

function movePoint(pt, dx = 0, dy = 0) {
  return { x: pt.x + dx, y: pt.y + dy };
}

function extendPoint(origin, target, extra = 0) {
  const vecX = target.x - origin.x;
  const vecY = target.y - origin.y;
  const length = Math.hypot(vecX, vecY);
  if (!length) return { x: target.x, y: target.y };
  const scale = (length + extra) / length;
  return {
    x: origin.x + vecX * scale,
    y: origin.y + vecY * scale,
  };
}

function createArmstrongLayerStack(svg) {
  const foundation = layer(svg, "Foundation", { asLayer: true });
  const foundationFront = layer(foundation, "Front Guides", { asLayer: true, prefix: "foundation" });
  const foundationBack = layer(foundation, "Back Guides", { asLayer: true, prefix: "foundation" });

  const front = layer(svg, "Front Bodice", { asLayer: true });
  const back = layer(svg, "Back Bodice", { asLayer: true });
  const labelsParent = layer(svg, "Labels & Markers", { asLayer: true });
  const labels = layer(labelsParent, "Labels", { asLayer: true, prefix: "labels" });
  const markers = layer(labelsParent, "Markers", { asLayer: true, prefix: "labels" });
  const numbers = layer(labelsParent, "Numbers", { asLayer: true, prefix: "labels" });

  return {
    foundation,
    foundationFront,
    foundationBack,
    front,
    back,
    labelsParent,
    labels,
    markers,
    numbers,
  };
}

function createAldrichLayerStack(svg) {
  const foundation = layer(svg, "Aldrich Guides", { asLayer: true, prefix: "aldrich" });
  const front = layer(svg, "Aldrich Front Bodice", { asLayer: true, prefix: "aldrich" });
  const back = layer(svg, "Aldrich Back Bodice", { asLayer: true, prefix: "aldrich" });
  const darts = layer(svg, "Aldrich Darts", { asLayer: true, prefix: "aldrich" });
  const labelsParent = layer(svg, "Aldrich Labels & Markers", { asLayer: true, prefix: "aldrich" });
  const markers = layer(labelsParent, "Markers", { asLayer: true, prefix: "aldrich" });
  const numbers = layer(labelsParent, "Numbers", { asLayer: true, prefix: "aldrich" });
  const letters = layer(labelsParent, "Letters", { asLayer: true, prefix: "aldrich" });
  return {
    foundation,
    front,
    back,
    darts,
    labelsParent,
    markers,
    numbers,
    letters,
  };
}

function readArmstrongParams() {
  const bustSpan = getNumber("bustSpan", 4.0625);
  const dartPlacement = getNumber("dartPlacement", 3.4375);
  return {
    fullLength: getNumber("fullLength", 18),
    acrossShoulder: getNumber("acrossShoulder", 7.9375),
    centreFrontLength: getNumber("centreFrontLength", 14.875),
    bustArc: getNumber("bustArc", 10.375),
    shoulderSlope: getNumber("shoulderSlope", 18.125),
    bustDepth: getNumber("bustDepth", 9.6875),
    shoulderLength: getNumber("shoulderLength", 5.375),
    bustSpan,
    acrossChest: getNumber("acrossChest", 6.9375),
    dartPlacement,
    newStrap: getNumber("newStrap", 18.1875),
    sideLength: getNumber("sideLength", 8.5),
    waistArc: getNumber("waistArc", 7.375),
    bustCup: getText("bustCup", "B Cup"),
    fullLengthBack: getNumber("fullLengthBack", 17.875),
    acrossShoulderBack: getNumber("acrossShoulderBack", 8.1875),
    centreFrontLengthBack: getNumber("centreFrontLengthBack", 17),
    bustArcBack: getNumber("bustArcBack", 9),
    shoulderSlopeBack: getNumber("shoulderSlopeBack", 17.375),
    shoulderLengthBack: getNumber("shoulderLengthBack", 5.375),
    bustSpanBack: bustSpan,
    acrossChestBack: getNumber("acrossChestBack", 7.1875),
    dartPlacementBack: dartPlacement,
    sideLengthBack: getNumber("sideLengthBack", 8.5),
    waistArcBack: getNumber("waistArcBack", 7),
    backNeck: getNumber("backNeck", 3.125),
    showGuides: getCheckbox("showGuides", true),
    showMarkers: getCheckbox("showMarkers", true),
  };
}

function readAldrichParams() {
  return {
    bust: getNumber("aldrichBust", ALDRICH_DEFAULTS.bust),
    bustEase: getNumber("aldrichBustEase", ALDRICH_DEFAULTS.bustEase),
    waist: getNumber("aldrichWaist", ALDRICH_DEFAULTS.waist),
    waistEase: getNumber("aldrichWaistEase", ALDRICH_DEFAULTS.waistEase),
    hip: getNumber("aldrichHip", ALDRICH_DEFAULTS.hip),
    waistToHip: getNumber("aldrichWaistToHip", ALDRICH_DEFAULTS.waistToHip),
    napeToWaist: getNumber("aldrichNapeToWaist", ALDRICH_DEFAULTS.napeToWaist),
    shoulder: getNumber("aldrichShoulder", ALDRICH_DEFAULTS.shoulder),
    backWidth: getNumber("aldrichBackWidth", ALDRICH_DEFAULTS.backWidth),
    chest: getNumber("aldrichChest", ALDRICH_DEFAULTS.chest),
    armscyeDepth: getNumber("aldrichArmscyeDepth", ALDRICH_DEFAULTS.armscyeDepth),
    neckSize: getNumber("aldrichNeckSize", ALDRICH_DEFAULTS.neckSize),
    frontNeckDart: getNumber("aldrichFrontNeckDart", ALDRICH_DEFAULTS.frontNeckDart),
    frontWaistDart: getNumber("aldrichFrontWaistDart", ALDRICH_DEFAULTS.frontWaistDart),
    backWaistDart: getNumber("aldrichBackWaistDart", ALDRICH_DEFAULTS.backWaistDart),
    frontSideWaistDart: getNumber("aldrichFrontSideWaistDart", ALDRICH_DEFAULTS.frontSideWaistDart),
    backSideWaistDart: getNumber("aldrichBackSideWaistDart", ALDRICH_DEFAULTS.backSideWaistDart),
    frontWaistDartBackOff: getNumber(
      "aldrichFrontWaistDartBackOff",
      ALDRICH_DEFAULTS.frontWaistDartBackOff
    ),
    bustWaistDiff: getNumber("aldrichBustWaistDiff", ALDRICH_DEFAULTS.bustWaistDiff),
    closeWaistShaping: getCheckbox("aldrichCloseWaist", true),
    reducedDarting: getCheckbox("aldrichReducedDarting", false),
    showGuides: getCheckbox("aldrichShowGuides", true),
    showMarkers: getCheckbox("aldrichShowMarkers", true),
  };
}

function setInputNumber(id, value, precision = 2) {
  const input = document.getElementById(id);
  if (!input || !Number.isFinite(value)) return;
  const factor = Math.pow(10, precision);
  const rounded = Math.round(value * factor) / factor;
  input.value = rounded;
}

function getAldrichReductionFactor() {
  const reducedToggle = document.getElementById("aldrichReducedDarting");
  return reducedToggle && reducedToggle.checked ? 0.75 : 1;

}

function applyAldrichDartDistribution(diff, options = {}) {
  const reducedToggle = document.getElementById("aldrichReducedDarting");
  const useReduced = reducedToggle && reducedToggle.checked;
  const darts = computeAldrichWaistDartsFromDiff(diff, { reduced: useReduced });
  const { front = 0, back = 0, frontSide = 0, backSide = 0 } = darts;
  const force = options.force === true;
  const skipDiffField = options.skipDiffField === true;
  const resetDartFlags = options.resetDartFlags === true || options.resetManual === true;
  const resetDiffFlag = options.resetDiffFlag === true || options.resetManual === true;

  if (force || !aldrichAutoState.frontWaistDartEdited) {
    setInputNumber("aldrichFrontWaistDart", front);
  }
  if (force || !aldrichAutoState.backWaistDartEdited) {
    setInputNumber("aldrichBackWaistDart", back);
  }
  if (force || !aldrichAutoState.frontSideWaistDartEdited) {
    setInputNumber("aldrichFrontSideWaistDart", frontSide);
  }
  if (force || !aldrichAutoState.backSideWaistDartEdited) {
    setInputNumber("aldrichBackSideWaistDart", backSide);
  }
  if (!skipDiffField && (force || !aldrichAutoState.bustWaistDiffEdited) && Number.isFinite(diff)) {
    setInputNumber("aldrichBustWaistDiff", Math.abs(diff));
  }

  if (resetDartFlags) {
    aldrichAutoState.frontWaistDartEdited = false;
    aldrichAutoState.backWaistDartEdited = false;
    aldrichAutoState.frontSideWaistDartEdited = false;
    aldrichAutoState.backSideWaistDartEdited = false;
  }
  if (resetDiffFlag) {
    aldrichAutoState.bustWaistDiffEdited = false;
  }
}

function updateAldrichDerivedFields(options = {}) {
  const bust = getNumber("aldrichBust", ALDRICH_DEFAULTS.bust);
  const waist = getNumber("aldrichWaist", ALDRICH_DEFAULTS.waist);
  const bustEase = getNumber("aldrichBustEase", ALDRICH_DEFAULTS.bustEase);
  const waistEase = getNumber("aldrichWaistEase", ALDRICH_DEFAULTS.waistEase);
  const force = options.force === true;
  const resetManual = options.resetManual === true;
  if (resetManual) {
    aldrichAutoState.frontNeckDartEdited = false;
    aldrichAutoState.bustWaistDiffEdited = false;
  }
  const frontNeck = computeAldrichFrontNeckDart(bust);
  if (Number.isFinite(frontNeck) && (force || !aldrichAutoState.frontNeckDartEdited)) {
    setInputNumber("aldrichFrontNeckDart", frontNeck);
  }
  const diff = computeAldrichWaistDiff(bust, waist, bustEase, waistEase);
  const manualDiffLocked = aldrichAutoState.bustWaistDiffEdited && !resetManual;
  if (Number.isFinite(diff) && !manualDiffLocked) {
    applyAldrichDartDistribution(diff, {
      force,
      resetDartFlags: resetManual,
      resetDiffFlag: resetManual,
    });
  }
}

function initAldrichAutoFields() {
  if (aldrichAutoInitialized) return;
  aldrichAutoInitialized = true;
  const baseFieldIds = [
    "aldrichBust",
    "aldrichBustEase",
    "aldrichWaist",
    "aldrichWaistEase",
    "aldrichChest",
    "aldrichBackWidth",
    "aldrichHip",
    "aldrichWaistToHip",
    "aldrichNapeToWaist",
    "aldrichShoulder",
    "aldrichArmscyeDepth",
    "aldrichNeckSize",
  ];
  baseFieldIds.forEach((id) => {
    const input = document.getElementById(id);
    if (!input) return;
    ["input", "change"].forEach((evt) => {
      input.addEventListener(evt, () => updateAldrichDerivedFields());
    });
  });

  const frontNeckInput = document.getElementById("aldrichFrontNeckDart");
  if (frontNeckInput) {
    frontNeckInput.addEventListener("input", () => {
      aldrichAutoState.frontNeckDartEdited = true;
    });
  }

  [
    ["aldrichFrontWaistDart", "frontWaistDartEdited"],
    ["aldrichBackWaistDart", "backWaistDartEdited"],
    ["aldrichFrontSideWaistDart", "frontSideWaistDartEdited"],
    ["aldrichBackSideWaistDart", "backSideWaistDartEdited"],
  ].forEach(([id, flag]) => {
    const input = document.getElementById(id);
    if (!input) return;
    input.addEventListener("input", () => {
      aldrichAutoState[flag] = true;
      aldrichAutoState.bustWaistDiffEdited = false;
    });
  });

  const bustWaistInput = document.getElementById("aldrichBustWaistDiff");
  if (bustWaistInput) {
    bustWaistInput.addEventListener("input", () => {
      aldrichAutoState.bustWaistDiffEdited = true;
    });
    bustWaistInput.addEventListener("change", () => {
      const diffVal = parseFloat(bustWaistInput.value);
      if (Number.isFinite(diffVal)) {
        applyAldrichDartDistribution(diffVal, {
          force: true,
          resetDartFlags: true,
          skipDiffField: true,
        });
        aldrichAutoState.bustWaistDiffEdited = true;
        scheduleRegen();
      }
    });
  }

  const closeToggle = document.getElementById("aldrichCloseWaist");
  const reducedToggle = document.getElementById("aldrichReducedDarting");
  if (closeToggle && reducedToggle) {
    closeToggle.addEventListener("change", () => {
      if (closeToggle.checked) {
        reducedToggle.checked = false;
      } else if (!reducedToggle.checked) {
        closeToggle.checked = true;
      }
      updateAldrichDerivedFields({ force: true, resetManual: true });
      scheduleRegen();
    });
    reducedToggle.addEventListener("change", () => {
      if (reducedToggle.checked) {
        closeToggle.checked = false;
      } else if (!closeToggle.checked) {
        closeToggle.checked = true;
      }
      updateAldrichDerivedFields({ force: true, resetManual: true });
      scheduleRegen();
    });
  }

  updateAldrichDerivedFields({ force: true, resetManual: true });
}

function getNumber(id, fallback) {
  const input = document.getElementById(id);
  if (!input) return fallback;
  const value = parseFloat(input.value);
  return Number.isFinite(value) ? value : fallback;
}

function getCheckbox(id, fallback) {
  const input = document.getElementById(id);
  if (!input) return fallback;
  return input.checked;
}

function downloadSVG(svg, filename = "armstrong_bodice.svg") {
  const blob = new Blob([new XMLSerializer().serializeToString(svg)], {
    type: "image/svg+xml;charset=utf-8",
  });
  const url = URL.createObjectURL(blob);
  const link = Object.assign(document.createElement("a"), {
    href: url,
    download: filename,
  });
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

let preview = null;
let currentSvg = null;
let regenTimer = null;
let patternSelect = null;
let armstrongControls = null;
let patternPlaceholder = null;
let appInitialized = false;
let aldrichAutoInitialized = false;
const aldrichAutoState = {
  frontNeckDartEdited: false,
  frontWaistDartEdited: false,
  backWaistDartEdited: false,
  frontSideWaistDartEdited: false,
  backSideWaistDartEdited: false,
  bustWaistDiffEdited: false,
};
const PATTERN_CONFIGS = {
  armstrong: {
    title: "Armstrong's Bodice",
    elementId: "armstrongControls",
    readParams: readArmstrongParams,
    generate: generateArmstrong,
    downloadId: "download",
    shareId: "share",
    filename: "armstrong_bodice.svg",
  },
  aldrich: {
    title: "Aldrich's Close Fitting Bodice",
    elementId: "aldrichControls",
    readParams: readAldrichParams,
    generate: generateAldrich,
    downloadId: "downloadAldrich",
    shareId: "shareAldrich",
    filename: "aldrich_close_fitting_bodice.svg",
  },
};

function regen() {
  if (regenTimer) {
    clearTimeout(regenTimer);
    regenTimer = null;
  }
  if (!preview) return;
  preview.innerHTML = "";
  const selectedPattern = patternSelect ? patternSelect.value : "armstrong";
  const config = PATTERN_CONFIGS[selectedPattern];
  if (!config) {
    const label = patternSelect
      ? patternSelect.options[patternSelect.selectedIndex]?.text || "This draft"
      : "This draft";
    showPreviewMessage(`${label} is coming soon.`);
    currentSvg = null;
    return;
  }
  if (selectedPattern === "aldrich") {
    updateAldrichDerivedFields();
  }
  currentSvg = config.generate(config.readParams());
  preview.appendChild(currentSvg);
  if (typeof window !== "undefined" && window.localStorage) {
    try {
      window.localStorage.setItem("patternhub:lastPattern", selectedPattern);
    } catch (errStorePattern) {
      console.warn("Unable to persist pattern selection:", errStorePattern);
    }
  }
}

function ensurePatternSelection(patternKey) {
  if (!patternSelect || patternSelect.value === patternKey) return;
  patternSelect.value = patternKey;
  patternSelect.dispatchEvent(new Event("change", { bubbles: true }));
}

const scheduleRegen = () => {
  clearTimeout(regenTimer);
  regenTimer = setTimeout(regen, 200);
};

function initApp() {
  if (appInitialized) return;
  appInitialized = true;

  preview = document.getElementById("preview");
  patternSelect = document.getElementById("patternSelect");
  let initialPattern = "armstrong";
  if (typeof window !== "undefined" && window.localStorage) {
    try {
      const storedPattern = window.localStorage.getItem("patternhub:lastPattern");
      if (storedPattern && PATTERN_CONFIGS[storedPattern]) {
        initialPattern = storedPattern;
      }
    } catch (errLoadPattern) {
      console.warn("Unable to read pattern selection from storage:", errLoadPattern);
    }
  }
  if (patternSelect && patternSelect.value !== initialPattern && PATTERN_CONFIGS[initialPattern]) {
    patternSelect.value = initialPattern;
  }
  armstrongControls = document.getElementById("armstrongControls");
  patternPlaceholder = document.getElementById("patternPlaceholder");
  enhancePatternSelect();
  enhanceBustCupSelect();
  initAldrichAutoFields();
  Object.entries(PATTERN_CONFIGS).forEach(([key, config]) => {
    if (!config) return;
    const downloadButton = config.downloadId ? document.getElementById(config.downloadId) : null;
    if (downloadButton) {
      downloadButton.addEventListener("click", () => {
        ensurePatternSelection(key);
        if (!currentSvg) regen();
        if (currentSvg) downloadSVG(currentSvg, config.filename || `${key}.svg`);
      });
    }

    const shareButton = config.shareId ? document.getElementById(config.shareId) : null;
    if (shareButton) {
      const defaultShareLabel = shareButton.textContent;
      shareButton.addEventListener("click", async () => {
        ensurePatternSelection(key);
        const title = config.title || "Pattern Draft";
        try {
          if (navigator.share) {
            await navigator.share({
              title,
              url: SHARE_URL || window.location.href,
            });
          } else if (navigator.clipboard) {
            await navigator.clipboard.writeText(SHARE_URL || window.location.href);
            shareButton.textContent = "Link Copied!";
            setTimeout(() => (shareButton.textContent = defaultShareLabel), 1500);
          } else {
            window.open(SHARE_URL || window.location.href, "_blank");
          }
        } catch (err) {
          console.error("Share failed:", err);
        }
      });
    }
  });

  document
    .querySelectorAll(".controls input, .controls select")
    .forEach((el) => {
      if (el.id === "download" || el.id === "patternSelect") {
        return;
      }
      const type = el.type || el.tagName;
      if (type === "checkbox" || el.tagName === "SELECT") {
        el.addEventListener("change", scheduleRegen);
      } else if (type === "number" || type === "text") {
        el.addEventListener("input", scheduleRegen);
      }
    });

  if (patternSelect) {
    patternSelect.addEventListener("change", () => {
      updatePatternVisibility();
      regen();
    });
  }

  updatePatternVisibility();
  regen();
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initApp);
} else {
  initApp();
}

function enhancePatternSelect() {
  const select = document.getElementById("patternSelect");
  if (!select || select.dataset.enhanced === "true") return;
  const label = select.closest(".pattern-picker label");
  if (!label) return;

  const dropdown = document.createElement("div");
  dropdown.className = "pattern-dropdown";
  dropdown.dataset.patternDropdown = "true";

  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "pattern-dropdown__toggle";
  toggle.setAttribute("aria-haspopup", "listbox");
  toggle.setAttribute("aria-expanded", "false");
  dropdown.appendChild(toggle);

  const menu = document.createElement("div");
  menu.className = "pattern-dropdown__menu";
  menu.setAttribute("role", "listbox");
  dropdown.appendChild(menu);

  label.insertBefore(dropdown, select);

  const optionButtons = [];

  const createOptionButton = (option) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "pattern-dropdown__option";
    btn.setAttribute("role", "option");
    btn.dataset.value = option.value;
    btn.textContent = option.textContent;
    if (option.disabled) {
      btn.disabled = true;
      btn.setAttribute("aria-disabled", "true");
    }
    btn.addEventListener("click", () => {
      if (btn.disabled) return;
      if (select.value !== option.value) {
        select.value = option.value;
        select.dispatchEvent(new Event("change", { bubbles: true }));
      }
      closeMenu();
    });
    optionButtons.push(btn);
    return btn;
  };

  const appendOption = (container, option) => {
    container.appendChild(createOptionButton(option));
  };

  const defaultGroup = document.createElement("div");
  defaultGroup.className = "pattern-dropdown__group";

  Array.from(select.children).forEach((child) => {
    if (child.tagName === "OPTGROUP") {
      const group = document.createElement("div");
      group.className = "pattern-dropdown__group";
      if (child.label) {
        const heading = document.createElement("p");
        heading.className = "pattern-dropdown__group-title";
        heading.textContent = child.label;
        group.appendChild(heading);
      }
      Array.from(child.children).forEach((option) => appendOption(group, option));
      menu.appendChild(group);
    } else if (child.tagName === "OPTION") {
      appendOption(defaultGroup, child);
    }
  });

  if (defaultGroup.children.length) {
    menu.insertBefore(defaultGroup, menu.firstChild);
  }

  const syncSelection = () => {
    const selectedOption = select.options[select.selectedIndex];
    toggle.textContent = selectedOption ? selectedOption.textContent : "Select a draft";
    optionButtons.forEach((btn) => {
      const isSelected = btn.dataset.value === select.value;
      btn.classList.toggle("is-selected", isSelected);
      btn.setAttribute("aria-selected", isSelected ? "true" : "false");
    });
  };

  const closeMenu = () => {
    dropdown.classList.remove("is-open");
    toggle.setAttribute("aria-expanded", "false");
  };

  const openMenu = () => {
    dropdown.classList.add("is-open");
    toggle.setAttribute("aria-expanded", "true");
  };

  toggle.addEventListener("click", () => {
    if (dropdown.classList.contains("is-open")) {
      closeMenu();
    } else {
      openMenu();
    }
  });

  document.addEventListener("click", (event) => {
    if (!dropdown.contains(event.target)) {
      closeMenu();
    }
  });

  dropdown.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeMenu();
      toggle.focus();
    }
  });

  dropdown.addEventListener("focusout", (event) => {
    if (!dropdown.contains(event.relatedTarget)) {
      closeMenu();
    }
  });

  select.addEventListener("change", syncSelection);

  syncSelection();
  select.dataset.enhanced = "true";
  select.classList.add("pattern-picker__native");
}

function enhanceBustCupSelect() {
  const select = document.getElementById("bustCup");
  if (!select || select.dataset.enhanced === "true") return;
  const label = select.parentElement;
  if (!label) return;

  const dropdown = document.createElement("div");
  dropdown.className = "mini-dropdown";

  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "mini-dropdown__toggle";
  toggle.setAttribute("aria-haspopup", "listbox");
  toggle.setAttribute("aria-expanded", "false");
  dropdown.appendChild(toggle);

  const menu = document.createElement("div");
  menu.className = "mini-dropdown__menu";
  menu.setAttribute("role", "listbox");
  dropdown.appendChild(menu);

  label.insertBefore(dropdown, select);
  select.classList.add("mini-dropdown__native");

  const buttons = [];

  const sync = () => {
    const selectedOption = select.options[select.selectedIndex];
    toggle.textContent = selectedOption ? selectedOption.textContent : "Select";
    buttons.forEach((btn) => {
      const isSelected = btn.dataset.value === select.value;
      btn.classList.toggle("is-selected", isSelected);
      btn.setAttribute("aria-selected", isSelected ? "true" : "false");
    });
  };

  Array.from(select.options).forEach((option) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "mini-dropdown__option";
    btn.dataset.value = option.value;
    btn.textContent = option.textContent;
    if (option.disabled) {
      btn.disabled = true;
      btn.setAttribute("aria-disabled", "true");
    }
    btn.addEventListener("click", () => {
      if (btn.disabled) return;
      if (select.value !== option.value) {
        select.value = option.value;
        select.dispatchEvent(new Event("change", { bubbles: true }));
      }
      closeMenu();
    });
    buttons.push(btn);
    menu.appendChild(btn);
  });

  const closeMenu = () => {
    dropdown.classList.remove("is-open");
    toggle.setAttribute("aria-expanded", "false");
  };

  const openMenu = () => {
    dropdown.classList.add("is-open");
    toggle.setAttribute("aria-expanded", "true");
  };

  toggle.addEventListener("click", () => {
    if (dropdown.classList.contains("is-open")) {
      closeMenu();
    } else {
      openMenu();
    }
  });

  document.addEventListener("click", (event) => {
    if (!dropdown.contains(event.target)) {
      closeMenu();
    }
  });

  dropdown.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeMenu();
      toggle.focus();
    }
  });

  dropdown.addEventListener("focusout", (event) => {
    if (!dropdown.contains(event.relatedTarget)) {
      closeMenu();
    }
  });

  select.addEventListener("change", sync);
  sync();

  select.dataset.enhanced = "true";
}

function updatePatternVisibility() {
  const selectedValue = patternSelect ? patternSelect.value : "armstrong";
  let hasVisibleSection = false;

  Object.entries(PATTERN_CONFIGS).forEach(([key, config]) => {
    const section = config.elementId ? document.getElementById(config.elementId) : null;
    if (!section) return;
    if (key === selectedValue) {
      section.hidden = false;
      hasVisibleSection = true;
    } else {
      section.hidden = true;
    }
  });

  if (!patternPlaceholder) {
    return;
  }

  if (hasVisibleSection) {
    patternPlaceholder.hidden = true;
    return;
  }

  patternPlaceholder.hidden = false;
  const label =
    patternSelect?.options[patternSelect.selectedIndex]?.text || "This draft";
  const placeholderMessage = patternPlaceholder.querySelector("p");
  if (placeholderMessage) {
    placeholderMessage.textContent = `${label} is coming soon.`;
  }
  if (preview) {
    preview.innerHTML = "";
    showPreviewMessage(`${label} is coming soon.`);
  }
}

function resolveBustCupOffset(cup) {
  if (!cup) return BUST_CUP_OFFSETS.B;
  let normalized = String(cup).toUpperCase();
  if (normalized.includes("CUP")) normalized = normalized.replace("CUP", "");
  normalized = normalized.replace(/[^A-Z]/g, "").trim();
  const key = normalized.charAt(0);
  return BUST_CUP_OFFSETS[key] || BUST_CUP_OFFSETS.B;
}

function getText(id, fallback) {
  const input = document.getElementById(id);
  if (!input) return fallback;
  const value = input.value;
  return value !== undefined && value !== null && value !== "" ? value : fallback;
}

function slugifyId(name, prefix = "") {
  if (!name) return prefix || "layer";
  const base = name
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  if (!base) return prefix || "layer";
  return prefix ? `${prefix}-${base}` : base;
}

function firstNumber(...values) {
  for (const value of values) {
    if (Number.isFinite(value)) return value;
  }
  return 0;
}

function applyLayerVisibility(layers, params) {
  const showGuides = params.showGuides !== false;
  const showMarkers = params.showMarkers !== false;
  const guideDisplay = showGuides ? null : "none";
  [layers.foundation, layers.foundationFront, layers.foundationBack].forEach((layer) => {
    if (!layer) return;
    if (guideDisplay) {
      layer.setAttribute("display", guideDisplay);
    } else {
      layer.removeAttribute("display");
    }
  });

  const markerDisplay = showMarkers ? null : "none";
  [layers.labelsParent].forEach((layer) => {
    if (!layer) return;
    if (markerDisplay) {
      layer.setAttribute("display", markerDisplay);
    } else {
      layer.removeAttribute("display");
    }
  });
}

function showPreviewMessage(text) {
  if (!preview) return;
  const msg = document.createElement("p");
  msg.textContent = text;
  msg.style.color = "#475569";
  msg.style.fontStyle = "italic";
  msg.style.margin = "24px";
  preview.appendChild(msg);
}

function computeAldrichFrontNeckDart(bust) {
  if (!Number.isFinite(bust)) return 7;
  const baseBustLow = 88;
  const baseBustHigh = 110;
  const baseLowValue = 7;
  const baseHighValue = 10;
  if (bust < baseBustLow) {
    const diffLow = ((baseBustLow - bust) / 4) * 0.6;
    return baseLowValue - diffLow;
  }
  if (bust <= 104) {
    const diffMidLow = ((bust - baseBustLow) / 4) * 0.6;
    return baseLowValue + diffMidLow;
  }
  if (bust <= baseBustHigh) {
    const diffMidHigh = ((baseBustHigh - bust) / 6) * 0.6;
    return baseHighValue - diffMidHigh;
  }
  const diffHigh = ((bust - baseBustHigh) / 6) * 0.6;
  return baseHighValue + diffHigh;
}

function computeAldrichWaistDiff(bust, waist, bustEase, waistEase) {
  if (!Number.isFinite(bust) || !Number.isFinite(waist)) return NaN;
  let bustHalf = bust / 2;
  let waistHalf = waist / 2;
  if (Number.isFinite(bustEase)) bustHalf += bustEase;
  if (Number.isFinite(waistEase)) waistHalf += waistEase;
  return bustHalf - waistHalf;
}

function computeAldrichWaistDartsFromDiff(diff, options = {}) {
  if (!Number.isFinite(diff)) {
    return { front: 0, back: 0, frontSide: 0, backSide: 0 };
  }
  const reduced = options.reduced === true;
  let waistDifference = Math.abs(diff);
  if (reduced) {
    waistDifference *= 0.75;
  }
  const subtractBase = reduced ? 5 : 6;
  let x = (waistDifference - subtractBase) / 4;
  if (!Number.isFinite(x) || x < 0) {
    x = 0;
  }
  if (reduced) {
    return {
      backSide: x,
      frontSide: x + 1,
      back: x + 1.5,
      front: x + 2.5,
    };
  }
  return {
    backSide: x,
    frontSide: x + 1,
    back: x + 2,
    front: x + 3,
  };
}

function computeAldrichWaistDarts(bust, waist, bustEase, waistEase, options = {}) {
  const diff = computeAldrichWaistDiff(bust, waist, bustEase, waistEase);
  return computeAldrichWaistDartsFromDiff(diff, options);
}

function computeAldrichPointADistance(bust) {
  if (!Number.isFinite(bust)) return 2.5;
  if (bust <= 80) return 2.25;
  if (bust >= 96 && bust <= 106) return 3;
  if (bust > 106 && bust <= 128) return 3.5;
  if (bust > 80 && bust <= 99) return 2.5;
  return 3.5;
}

function computeAldrichPointBDistance(bust) {
  if (!Number.isFinite(bust)) return 2;
  if (bust <= 80) return 1.75;
  if (bust >= 96 && bust <= 106) return 2.5;
  if (bust > 106 && bust <= 128) return 3;
  if (bust > 80 && bust <= 99) return 2.0;
  return 3;
}

function clampHandleToChordCm(baseXcm, baseYcm, chordCm, ratio = ALDRICH_HANDLE_RATIO_CAP) {
  const baseLen = Math.hypot(baseXcm, baseYcm);
  if (!Number.isFinite(chordCm) || chordCm <= 0) {
    return { x: 0, y: 0 };
  }
  if (!Number.isFinite(baseLen) || baseLen < 0.0001) {
    return { x: baseXcm, y: baseYcm };
  }
  const maxLen = chordCm * ratio;
  if (maxLen >= baseLen) {
    return { x: baseXcm, y: baseYcm };
  }
  const scale = maxLen / baseLen;
  return {
    x: baseXcm * scale,
    y: baseYcm * scale,
  };
}

function distanceBetween(a, b) {
  if (!a || !b) return 0;
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  return Math.hypot(dx, dy);
}

function midpoint(a, b) {
  return {
    x: (a.x + b.x) / 2,
    y: (a.y + b.y) / 2,
  };
}

function isFiniteNumber(val) {
  return typeof val === "number" && Number.isFinite(val);
}

function intersectLineWithVertical(lineStart, lineEnd, xConst) {
  if (!lineStart || !lineEnd || !isFiniteNumber(xConst)) {
    return null;
  }
  const dx = lineEnd.x - lineStart.x;
  if (Math.abs(dx) < 0.000001) {
    return null;
  }
  const t = (xConst - lineStart.x) / dx;
  if (t < -0.0001 || t > 1.0001) {
    return null;
  }
  return {
    x: xConst,
    y: lineStart.y + t * (lineEnd.y - lineStart.y),
  };
}

function valueBetween(val, a, b, tolerance = 0.0001) {
  const minVal = Math.min(a, b) - tolerance;
  const maxVal = Math.max(a, b) + tolerance;
  return val >= minVal && val <= maxVal;
}
