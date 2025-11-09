//@target illustrator

;(function () {
  if (typeof app === 'undefined' || !app.documents) {
    alert('This script must be run from Adobe Illustrator.')
    return
  }

  var CM_TO_PT = 28.346456693
  var DASH_PATTERN = [25, 12]
  var LABEL_FONT_PT = 12
  var LABEL_OFFSET_CM = 0.5
  var LABEL_INSIDE_OFFSET_CM = 1
  var MARKER_RADIUS_CM = 0.45
  var LINE_STROKE_PT = 1
  var HANDLE_RATIO_CAP = 1.6
  var measurementPalette = null
  var DRAFT_GROUP_NAME = 'Close Fitting Bodice Draft'
  var MARKERS_GROUP_NAME = 'Close Fitting Bodice Markers'
  var NUMBERS_GROUP_NAME = 'Close Fitting Bodice Numbers'
  var LABELS_GROUP_NAME = 'Close Fitting Bodice Labels'
  var BACK_GROUP_NAME = 'Close Fitting Back Draft'

  function createRGBColor(r, g, b) {
    var color = new RGBColor()
    color.red = r
    color.green = g
    color.blue = b
    return color
  }

  function blackColor() {
    return createRGBColor(0, 0, 0)
  }

  function whiteColor() {
    return createRGBColor(255, 255, 255)
  }

  function norm(value) {
    return (value || '').toString().toLowerCase()
  }

  function cm(val) {
    return val * CM_TO_PT
  }

  function ptToCm(val) {
    return val / CM_TO_PT
  }

  function formatCm(val) {
    if (isNaN(val)) {
      return ''
    }
    var rounded = Math.round(val * 100) / 100
    return rounded.toFixed(2) + ' cm'
  }

  function formatNumber(val) {
    if (isNaN(val)) {
      return ''
    }
    var rounded = Math.round(val * 100) / 100
    return rounded.toFixed(2)
  }

  function computeFrontNeckDart(bust) {
    if (!isFinite(bust)) {
      return 7
    }
    var baseBustLow = 88
    var baseBustHigh = 110
    var baseLowValue = 7
    var baseHighValue = 10
    if (bust < baseBustLow) {
      var diffLow = ((baseBustLow - bust) / 4) * 0.6
      return baseLowValue - diffLow
    }
    if (bust <= 104) {
      var diffMidLow = ((bust - baseBustLow) / 4) * 0.6
      return baseLowValue + diffMidLow
    }
    if (bust <= baseBustHigh) {
      var diffMidHigh = ((baseBustHigh - bust) / 6) * 0.6
      return baseHighValue - diffMidHigh
    }
    var diffHigh = ((bust - baseBustHigh) / 6) * 0.6
    return baseHighValue + diffHigh
  }

  function computeWaistDiff(bust, waist, bustEase, waistEase) {
    if (isNaN(bust) || isNaN(waist)) return NaN
    var bustHalf = bust / 2
    var waistHalf = waist / 2
    if (!isNaN(bustEase)) bustHalf += bustEase
    if (!isNaN(waistEase)) waistHalf += waistEase
    return bustHalf - waistHalf
  }

  function computeWaistDartsFromDiff(diff) {
    if (!isFinite(diff)) {
      return { front: 0, back: 0, side: 0 }
    }
    var targetDiff = Math.abs(diff)
    var x = (targetDiff - 6) / 4
    return {
      front: x + 3,
      back: x + 2,
      side: 2 * x + 1,
    }
  }

  function computeWaistDarts(bust, waist, bustEase, waistEase) {
    var diff = computeWaistDiff(bust, waist, bustEase, waistEase)
    return computeWaistDartsFromDiff(diff)
  }

  function computePointADistance(bust) {
    // The provided bands overlap slightly, so prioritise the 96-106 range for the 3 cm step.
    if (!isFinite(bust)) {
      return 2.5
    }
    if (bust <= 80) {
      return 2.25
    }
    if (bust >= 96 && bust <= 106) {
      return 3
    }
    if (bust > 106 && bust <= 128) {
      return 3.5
    }
    if (bust > 80 && bust <= 99) {
      return 2.5
    }
    return 3.5
  }

  function computePointBDistance(bust) {
    if (!isFinite(bust)) {
      return 2
    }
    if (bust <= 80) {
      return 1.75
    }
    if (bust >= 96 && bust <= 106) {
      return 2.5
    }
    if (bust > 106 && bust <= 128) {
      return 3
    }
    if (bust > 80 && bust <= 99) {
      return 2.0
    }
    return 3
  }

  function distanceBetween(a, b) {
    var dx = b.x - a.x
    var dy = b.y - a.y
    return Math.sqrt(dx * dx + dy * dy)
  }

  function clampHandleToChord(baseXcm, baseYcm, chordCm, ratio) {
    var baseLen = Math.sqrt(baseXcm * baseXcm + baseYcm * baseYcm)
    if (!isFiniteNumber(chordCm) || chordCm <= 0) {
      return { x: 0, y: 0 }
    }
    if (!isFiniteNumber(baseLen) || baseLen < 0.0001) {
      return { x: cm(baseXcm), y: cm(baseYcm) }
    }
    var maxLen = chordCm * ratio
    if (maxLen >= baseLen) {
      return { x: cm(baseXcm), y: cm(baseYcm) }
    }
    var scale = maxLen / baseLen
    return {
      x: cm(baseXcm * scale),
      y: cm(baseYcm * scale),
    }
  }

  function midpoint(a, b) {
    return {
      x: (a.x + b.x) / 2,
      y: (a.y + b.y) / 2,
    }
  }

  function isFiniteNumber(val) {
    return typeof val === 'number' && isFinite(val)
  }

  function intersectLineWithVertical(lineStart, lineEnd, xConst) {
    if (!lineStart || !lineEnd || !isFiniteNumber(xConst)) {
      return null
    }
    var dx = lineEnd.x - lineStart.x
    if (Math.abs(dx) < 0.000001) {
      return null
    }
    var t = (xConst - lineStart.x) / dx
    if (t < -0.0001 || t > 1.0001) {
      return null
    }
    var y = lineStart.y + t * (lineEnd.y - lineStart.y)
    return {
      x: xConst,
      y: y,
    }
  }

  function valueBetween(val, a, b, tolerance) {
    var tol = tolerance || 0.0001
    var minVal = Math.min(a, b) - tol
    var maxVal = Math.max(a, b) + tol
    return val >= minVal && val <= maxVal
  }

  function getOrCreateDocument() {
    var doc
    if (app.documents.length === 0) {
      doc = app.documents.add()
    } else {
      var current = app.activeDocument
      if (documentHasContent(current)) {
        doc = app.documents.add()
      } else {
        doc = current
      }
    }
    return doc
  }

  function documentHasContent(doc) {
    if (!doc) {
      return false
    }
    try {
      return doc.pageItems.length > 0
    } catch (err) {
      return false
    }
  }

  function showMeasurementDialog(defaults) {
    var dlg = new Window('dialog', 'Measurement Panel, Default Size = 12')
    dlg.orientation = 'column'
    dlg.alignChildren = 'fill'
    dlg.margins = 16
    dlg.spacing = 12

    var panel = dlg.add('panel', undefined, '')
    panel.orientation = 'column'
    panel.alignChildren = 'fill'
    panel.margins = 14
    panel.spacing = 8

    function addField(label, defaultValue) {
      var group = panel.add('group')
      group.orientation = 'row'
      group.alignChildren = ['left', 'center']
      var staticText = group.add('statictext', undefined, label + ':')
      staticText.preferredSize.width = 180
      var field = group.add(
        'edittext',
        undefined,
        defaultValue !== undefined ? String(defaultValue) : '',
      )
      field.characters = 8
      return field
    }

    var fields = {
      bust: addField('1. Bust (cm)', formatNumber(defaults.bust)),
      bustEase: addField('Bust Ease (cm)', formatNumber(defaults.bustEase)),
      waist: addField('2. Waist (cm)', formatNumber(defaults.waist)),
      waistEase: addField('Waist Ease (cm)', formatNumber(defaults.waistEase)),
      hip: addField('3. Hip (cm)', formatNumber(defaults.hip)),
      napeToWaist: addField(
        '4. Nape to Waist (cm)',
        formatNumber(defaults.napeToWaist),
      ),
      shoulder: addField('5. Shoulder (cm)', formatNumber(defaults.shoulder)),
      backWidth: addField(
        '6. Back Width (cm)',
        formatNumber(defaults.backWidth),
      ),
      waistToHip: addField(
        '7. Waist to Hip (cm)',
        formatNumber(defaults.waistToHip),
      ),
      armscyeDepth: addField(
        '8. Armscye Depth (cm)',
        formatNumber(defaults.armscyeDepth),
      ),
      chest: addField('9. Chest (cm)', formatNumber(defaults.chest)),
      neckSize: addField('10. Neck Size (cm)', formatNumber(defaults.neckSize)),
      frontNeckDart: addField(
        'Front Neck Dart (cm)',
        formatNumber(defaults.frontNeckDart),
      ),
      frontWaistDart: addField(
        'Front Waist Dart (cm)',
        formatNumber(defaults.frontWaistDart),
      ),
      backWaistDart: addField(
        'Back Waist Dart (cm)',
        formatNumber(defaults.backWaistDart),
      ),
      sideWaistDart: addField(
        'Side Waist Darts (cm)',
        formatNumber(defaults.sideWaistDart),
      ),
      frontWaistDartBackOff: addField(
        'Dart Apex Offset (cm)',
        formatNumber(
      isFiniteNumber(defaults.frontWaistDartBackOff)
            ? defaults.frontWaistDartBackOff
            : 2.5,
        ),
      ),
      bustWaistDiff: addField(
        'Bust-Waist Difference (cm)',
        formatNumber(defaults.bustWaistDiff),
      ),
    }

    var toggleGroup = panel.add('group')
    toggleGroup.orientation = 'column'
    toggleGroup.alignChildren = 'left'
    toggleGroup.margins = [0, 8, 0, 0]

    var closeWaistToggle = toggleGroup.add(
      'checkbox',
      undefined,
      'Close Waist Shaping',
    )
    closeWaistToggle.value = defaults.closeWaistShaping !== false

    var reducedDartToggle = toggleGroup.add(
      'checkbox',
      undefined,
      'Reduced Darting (75% of Dart)',
    )
    reducedDartToggle.value = !!defaults.reducedDarting
    if (reducedDartToggle.value) {
      closeWaistToggle.value = false
    } else if (!closeWaistToggle.value) {
      closeWaistToggle.value = true
    }

    function getDartReductionFactor() {
      return reducedDartToggle.value ? 0.75 : 1
    }

    function refreshWaistDistribution() {
      var bustValue = getFieldNumber(fields.bust)
      var waistValue = getFieldNumber(fields.waist)
      var bustEaseValue = getFieldNumber(fields.bustEase)
      var waistEaseValue = getFieldNumber(fields.waistEase)
      if (isNaN(bustValue) || isNaN(waistValue)) {
        return
      }
      var diff = computeWaistDiff(
        bustValue,
        waistValue,
        bustEaseValue,
        waistEaseValue,
      )
      applyDartDistribution(diff, { force: true, resetManualDarts: true })
      waistDiffManuallyEdited = false
    }

    var frontNeckDartManuallyEdited = false

    function markFrontNeckDartManual() {
      frontNeckDartManuallyEdited = true
    }

    function getFieldNumber(field) {
      var raw = field.text || ''
      var normalized = raw.replace(',', '.').replace(/[^0-9\.\-]/g, '')
      return parseFloat(normalized)
    }

    function updateFrontNeckDart() {
      if (frontNeckDartManuallyEdited) {
        return
      }
      var bustValue = getFieldNumber(fields.bust)
      var chestValue = getFieldNumber(fields.chest)
      var backWidthValue = getFieldNumber(fields.backWidth)
      if (isNaN(bustValue) || isNaN(chestValue) || isNaN(backWidthValue)) {
        return
      }
      fields.frontNeckDart.text = formatNumber(
        computeFrontNeckDart(bustValue, chestValue, backWidthValue),
      )
    }

    var waistDartManuallyEdited = {
      front: false,
      back: false,
      side: false,
    }
    var waistDiffManuallyEdited = false

    function markWaistDartManual(key) {
      waistDartManuallyEdited[key] = true
      waistDiffManuallyEdited = false
    }

    function anyWaistDartManuallyEdited() {
      return (
        waistDartManuallyEdited.front ||
        waistDartManuallyEdited.back ||
        waistDartManuallyEdited.side
      )
    }

    function updateBustWaistDiffFromDarts() {
      var frontVal = getFieldNumber(fields.frontWaistDart)
      var backVal = getFieldNumber(fields.backWaistDart)
      var sideVal = getFieldNumber(fields.sideWaistDart)
      if (isNaN(frontVal) || isNaN(backVal) || isNaN(sideVal)) return
      var diffSum = frontVal + backVal + sideVal
      fields.bustWaistDiff.text = formatNumber(Math.abs(diffSum))
    }

    function applyDartDistribution(diff, options) {
      options = options || {}
      var darts = computeWaistDartsFromDiff(diff)
      var factor = getDartReductionFactor()
      var adjustedFront = darts.front * factor
      var adjustedBack = darts.back * factor
      var adjustedSide = darts.side * factor
      var distributedTotal = adjustedFront + adjustedBack + adjustedSide
      if (options.force || !waistDartManuallyEdited.front)
        fields.frontWaistDart.text = formatNumber(adjustedFront)
      if (options.force || !waistDartManuallyEdited.back)
        fields.backWaistDart.text = formatNumber(adjustedBack)
      if (options.force || !waistDartManuallyEdited.side)
        fields.sideWaistDart.text = formatNumber(adjustedSide)
      if (!options.skipDiffField)
        fields.bustWaistDiff.text = formatNumber(Math.abs(distributedTotal))
      if (options.resetManualDarts) {
        waistDartManuallyEdited.front = false
        waistDartManuallyEdited.back = false
        waistDartManuallyEdited.side = false
      }
    }

    function updateWaistDarts() {
      var bustValue = getFieldNumber(fields.bust)
      var waistValue = getFieldNumber(fields.waist)
      var bustEaseValue = getFieldNumber(fields.bustEase)
      var waistEaseValue = getFieldNumber(fields.waistEase)
      if (isNaN(bustValue) || isNaN(waistValue)) {
        return
      }
      var diff = computeWaistDiff(
        bustValue,
        waistValue,
        bustEaseValue,
        waistEaseValue,
      )
      if (!waistDiffManuallyEdited && !anyWaistDartManuallyEdited()) {
        applyDartDistribution(diff, { force: true, resetManualDarts: true })
        waistDiffManuallyEdited = false
      } else if (!waistDiffManuallyEdited) {
        updateBustWaistDiffFromDarts()
      }
    }

    function onBustFieldChange() {
      updateFrontNeckDart()
      updateWaistDarts()
    }

    function onWaistFieldChange() {
      updateWaistDarts()
    }

    function onWaistEaseFieldChange() {
      updateWaistDarts()
    }

    function onBustEaseFieldChange() {
      updateWaistDarts()
    }

    closeWaistToggle.onClick = function () {
      if (closeWaistToggle.value) {
        reducedDartToggle.value = false
      } else if (!reducedDartToggle.value) {
        closeWaistToggle.value = true
      }
      refreshWaistDistribution()
    }

    reducedDartToggle.onClick = function () {
      if (reducedDartToggle.value) {
        closeWaistToggle.value = false
      } else if (!closeWaistToggle.value) {
        closeWaistToggle.value = true
      }
      refreshWaistDistribution()
    }

    var summaryToggle = toggleGroup.add(
      'checkbox',
      undefined,
      'Show measurement summary',
    )
    summaryToggle.value = defaults.showMeasurementPalette === true

    fields.bust.onChanging = onBustFieldChange
    fields.bust.onChange = onBustFieldChange
    fields.bustEase.onChanging = onBustEaseFieldChange
    fields.bustEase.onChange = onBustEaseFieldChange
    fields.waist.onChanging = onWaistFieldChange
    fields.waist.onChange = onWaistFieldChange
    fields.waistEase.onChanging = onWaistEaseFieldChange
    fields.waistEase.onChange = onWaistEaseFieldChange
    fields.chest.onChanging = updateFrontNeckDart
    fields.chest.onChange = updateFrontNeckDart
    fields.backWidth.onChanging = updateFrontNeckDart
    fields.backWidth.onChange = updateFrontNeckDart
    fields.frontNeckDart.onChanging = markFrontNeckDartManual
    fields.frontNeckDart.onChange = markFrontNeckDartManual
    fields.frontWaistDart.onChanging = function () {
      markWaistDartManual('front')
    }
    fields.frontWaistDart.onChange = function () {
      markWaistDartManual('front')
      updateBustWaistDiffFromDarts()
    }
    fields.backWaistDart.onChanging = function () {
      markWaistDartManual('back')
    }
    fields.backWaistDart.onChange = function () {
      markWaistDartManual('back')
      updateBustWaistDiffFromDarts()
    }
    fields.sideWaistDart.onChanging = function () {
      markWaistDartManual('side')
    }
    fields.sideWaistDart.onChange = function () {
      markWaistDartManual('side')
      updateBustWaistDiffFromDarts()
    }
    fields.bustWaistDiff.onChanging = function () {
      waistDiffManuallyEdited = true
    }
    fields.bustWaistDiff.onChange = function () {
      var diffVal = getFieldNumber(fields.bustWaistDiff)
      if (isNaN(diffVal)) return
      waistDiffManuallyEdited = true
      applyDartDistribution(diffVal, { force: true, resetManualDarts: true })
    }
    updateFrontNeckDart()
    updateWaistDarts()

    var buttonRow = dlg.add('group')
    buttonRow.orientation = 'row'
    buttonRow.alignment = 'right'
    var okBtn = buttonRow.add('button', undefined, 'OK', { name: 'ok' })
    var cancelBtn = buttonRow.add('button', undefined, 'Cancel', {
      name: 'cancel',
    })

    var result = null

    function readValue(field, key) {
      var value = getFieldNumber(field)
      if (isNaN(value)) {
        throw new Error('Please enter a numeric value for ' + key + '.')
      }
      return value
    }

    okBtn.onClick = function () {
      try {
        result = {
          bust: readValue(fields.bust, 'Bust'),
          waist: readValue(fields.waist, 'Waist'),
          hip: readValue(fields.hip, 'Hip'),
          waistEase: readValue(fields.waistEase, 'Waist Ease'),
          bustEase: readValue(fields.bustEase, 'Bust Ease'),
          napeToWaist: readValue(fields.napeToWaist, 'Nape to Waist'),
          shoulder: readValue(fields.shoulder, 'Shoulder'),
          backWidth: readValue(fields.backWidth, 'Back Width'),
          waistToHip: readValue(fields.waistToHip, 'Waist to Hip'),
          frontNeckDart: readValue(fields.frontNeckDart, 'Front Neck Dart'),
          frontWaistDart: readValue(fields.frontWaistDart, 'Front Waist Dart'),
          backWaistDart: readValue(fields.backWaistDart, 'Back Waist Dart'),
          sideWaistDart: readValue(fields.sideWaistDart, 'Side Waist Darts'),
          frontWaistDartBackOff: readValue(
            fields.frontWaistDartBackOff,
            'Dart Apex Offset',
          ),
          bustWaistDiff: readValue(
            fields.bustWaistDiff,
            'Bust-Waist Difference',
          ),
          armscyeDepth: readValue(fields.armscyeDepth, 'Armscye Depth'),
          chest: readValue(fields.chest, 'Chest'),
          neckSize: readValue(fields.neckSize, 'Neck Size'),
          closeWaistShaping: !!closeWaistToggle.value,
          reducedDarting: !!reducedDartToggle.value,
          showMeasurementPalette: !!summaryToggle.value,
        }
        dlg.close(1)
      } catch (validationErr) {
        alert(validationErr.message)
      }
    }

    cancelBtn.onClick = function () {
      dlg.close(0)
    }

    var response = dlg.show()
    if (response !== 1) {
      return null
    }
    return result
  }

  function closeMeasurementPalette() {
    try {
      if (measurementPalette && typeof measurementPalette.close === 'function') {
        measurementPalette.close()
      }
    } catch (errClosePalette) {}
    try {
      var globalScope = $.global
      if (
        globalScope &&
        globalScope.aldrichMeasurementPalette &&
        globalScope.aldrichMeasurementPalette.window &&
        typeof globalScope.aldrichMeasurementPalette.window.close === 'function'
      ) {
        globalScope.aldrichMeasurementPalette.window.close()
      }
      if (globalScope && globalScope.aldrichMeasurementPalette) {
        delete globalScope.aldrichMeasurementPalette
      }
    } catch (errGlobalClose) {}
    measurementPalette = null
  }

  function getMeasurementSummaryDefinitions() {
    return [
      { label: '1. Bust', key: 'bust' },
      { label: 'Bust Ease', key: 'bustEase' },
      { label: '2. Waist', key: 'waist' },
      { label: 'Waist Ease', key: 'waistEase' },
      { label: '3. Hip', key: 'hip' },
      { label: '4. Nape to Waist', key: 'napeToWaist' },
      { label: '5. Shoulder', key: 'shoulder' },
      { label: '6. Back Width', key: 'backWidth' },
      { label: '7. Waist to Hip', key: 'waistToHip' },
      { label: '8. Armscye Depth', key: 'armscyeDepth' },
      { label: '9. Chest', key: 'chest' },
      { label: '10. Neck Size', key: 'neckSize' },
      { label: 'Front Neck Dart', key: 'frontNeckDart' },
      { label: 'Front Waist Dart', key: 'frontWaistDart' },
      { label: 'Back Waist Dart', key: 'backWaistDart' },
      { label: 'Side Waist Darts', key: 'sideWaistDart' },
      { label: 'Dart Apex Offset', key: 'frontWaistDartBackOff' },
      { label: 'Bust-Waist Difference', key: 'bustWaistDiff' },
      { label: 'Close Waist Shaping', key: 'closeWaistShaping', type: 'boolean' },
      { label: 'Reduced Darting', key: 'reducedDarting', type: 'boolean' },
    ]
  }

  function formatSummaryValue(def, values) {
    if (!def || !values) {
      return '-'
    }
    var val = values[def.key]
    if (def.type === 'boolean') {
      return val ? 'Yes' : 'No'
    }
    if (isFiniteNumber(val)) {
      return formatCm(val)
    }
    return '-'
  }

  function showMeasurementPalette(values) {
    if (!values) return
    closeMeasurementPalette()
    var palette = new Window('palette', 'Measurement Summary')
    palette.orientation = 'column'
    palette.alignChildren = ['fill', 'top']
    palette.spacing = 10
    palette.margins = 16
    palette.preferredSize.width = 360

    var header = palette.add('statictext', undefined, 'Measurements (cm)')
    header.alignment = 'fill'

    var divider = palette.add('panel', undefined, '')
    divider.alignment = 'fill'

    var defs = getMeasurementSummaryDefinitions()
    var listGroup = palette.add('group')
    listGroup.orientation = 'column'
    listGroup.alignChildren = ['fill', 'top']
    listGroup.spacing = 4

    for (var i = 0; i < defs.length; i++) {
      var def = defs[i]
      var row = listGroup.add('group')
      row.orientation = 'row'
      row.alignChildren = ['left', 'center']
      row.spacing = 8
      row.margins = [0, 0, 0, 0]
      var label = row.add('statictext', undefined, def.label)
      label.preferredSize = [200, 18]
      var valueText = row.add(
        'statictext',
        undefined,
        formatSummaryValue(def, values),
      )
      valueText.preferredSize = [120, 18]
      valueText.justify = 'right'
    }

    palette.onClose = function () {
      measurementPalette = null
      try {
        if ($.global && $.global.aldrichMeasurementPalette) {
          delete $.global.aldrichMeasurementPalette
        }
      } catch (errDeleteGlobal) {}
    }

    palette.show()
    measurementPalette = palette
    try {
      $.global.aldrichMeasurementPalette = { window: palette }
    } catch (errStorePalette) {}
  }

  function ensureLayerStructure(doc) {
    function ensureLayer(name) {
      for (var i = 0; i < doc.layers.length; i++) {
        var lyr = doc.layers[i]
        if (norm(lyr.name) === norm(name)) {
          unlockLayer(lyr)
          return lyr
        }
      }
      var newLayer = doc.layers.add()
      newLayer.name = name
      unlockLayer(newLayer)
      return newLayer
    }

    function ensureSubLayer(parent, name) {
      if (!parent) {
        return null
      }
      for (var i = 0; i < parent.layers.length; i++) {
        var sub = parent.layers[i]
        if (norm(sub.name) === norm(name)) {
          unlockLayer(sub)
          return sub
        }
      }
      var created = parent.layers.add()
      created.name = name
      unlockLayer(created)
      return created
    }

    function unlockLayer(layer) {
      try {
        layer.locked = false
      } catch (eLock) {}
      try {
        layer.visible = true
      } catch (eVis) {}
    }

    function removeLayerByName(name) {
      var goal = norm(name)
      for (var i = doc.layers.length - 1; i >= 0; i--) {
        var lyr = doc.layers[i]
        if (norm(lyr.name) === goal) {
          try {
            lyr.remove()
          } catch (eRemove) {}
        }
      }
    }

    function removeGroupByName(layer, name) {
      if (!layer) {
        return
      }
      var goal = norm(name)
      var groups = layer.groupItems
      for (var i = groups.length - 1; i >= 0; i--) {
        var group = groups[i]
        if (norm(group.name) === goal) {
          try {
            group.remove()
          } catch (eRemoveGroup) {}
        }
      }
    }

    var basicFrame = ensureLayer('Foundation')
    var backLayer = ensureLayer('Back')
    var frontLayer = ensureLayer('Front')
    var numbersLabelsLayer = ensureLayer('Numbers & Markers')

    var subLayers = {
      // labels: ensureSubLayer(numbersLabelsLayer, 'Labels'), // Labels layer intentionally disabled.
      numbers: ensureSubLayer(numbersLabelsLayer, 'Numbers'),
      markers: ensureSubLayer(numbersLabelsLayer, 'Markers'),
    }
    var dartsLayer = ensureLayer('Darts')
    try {
      if (frontLayer && dartsLayer) {
        dartsLayer.move(frontLayer, ElementPlacement.PLACEBEFORE)
      }
    } catch (eDartsReorder) {}

    removeLayerByName('Layer 1')
    removeGroupByName(basicFrame, DRAFT_GROUP_NAME)
    removeGroupByName(backLayer, BACK_GROUP_NAME)
    removeGroupByName(frontLayer, 'Front Waist Darts')
    removeGroupByName(backLayer, 'Back Waist Darts')
    removeGroupByName(subLayers.numbers, NUMBERS_GROUP_NAME)
    // removeGroupByName(subLayers.labels, LABELS_GROUP_NAME)

    // Labels were removed; only numbers and markers remain under this layer.

    return {
      basicFrame: basicFrame,
      back: backLayer,
      front: frontLayer,
      numbersAndMarkers: {
        layer: numbersLabelsLayer,
        numbers: subLayers.numbers,
        markers: subLayers.markers,
      },
      darts: dartsLayer,
    }
  }

  function resetLayerContents(layer) {
    if (!layer) {
      return null
    }
    try {
      layer.locked = false
    } catch (eUnlockLayer) {}
    try {
      layer.visible = true
    } catch (eVisLayer) {}
    function removeCollection(collection) {
      if (!collection) return
      for (var i = collection.length - 1; i >= 0; i--) {
        try {
          collection[i].remove()
        } catch (eRemoveItem) {}
      }
    }
    removeCollection(layer.pageItems)
    removeCollection(layer.pathItems)
    removeCollection(layer.groupItems)
    removeCollection(layer.compoundPathItems)
    removeCollection(layer.textFrames)
    return layer
  }

  function toArt(pt, originX, originY) {
    return [originX + cm(pt.x), originY - cm(pt.y)]
  }

  function centerTextFrame(tf, anchor) {
    try {
      var bounds = tf.visibleBounds
      var cx = (bounds[0] + bounds[2]) * 0.5
      var cy = (bounds[1] + bounds[3]) * 0.5
      tf.translate(anchor[0] - cx, anchor[1] - cy)
    } catch (errCenter) {
      try {
        tf.left = anchor[0] - tf.width / 2
        tf.top = anchor[1] + tf.height / 2
      } catch (errPosition) {}
    }
  }

  function createLineLabel(labelLayer, originX, originY, start, end, text) {
    if (!labelLayer || !text) {
      return null
    }
    var startArt = toArt(start, originX, originY)
    var endArt = toArt(end, originX, originY)
    var dx = end.x - start.x
    var dy = end.y - start.y
    var length = Math.sqrt(dx * dx + dy * dy)
    if (length === 0) {
      return null
    }
    var midpointPt = midpoint(start, end)
    var nx = -dy / length
    var ny = dx / length
    if (ny < 0) {
      nx = -nx
      ny = -ny
    }
    var anchor = {
      x: midpointPt.x + nx * LABEL_OFFSET_CM,
      y: midpointPt.y + ny * LABEL_OFFSET_CM,
    }
    var anchorArt = toArt(anchor, originX, originY)
    var tf = labelLayer.textFrames.add()
    tf.contents = text
    try {
      tf.textRange.characterAttributes.size = LABEL_FONT_PT
      tf.textRange.characterAttributes.fillColor = blackColor()
      tf.textRange.paragraphAttributes.justification = Justification.CENTER
    } catch (eLabelAttr) {}
    tf.position = anchorArt
    centerTextFrame(tf, anchorArt)
    var dxArt = endArt[0] - startArt[0]
    var dyArt = endArt[1] - startArt[1]
    var angleRad = Math.atan2(dyArt, dxArt)
    if (angleRad < -Math.PI / 2) {
      angleRad += Math.PI
    } else if (angleRad > Math.PI / 2) {
      angleRad -= Math.PI
    }
    var angleDeg = (angleRad * 180) / Math.PI
    try {
      tf.rotate(angleDeg, true, true, true, true, Transformation.CENTER)
    } catch (errRotate) {
      try {
        tf.rotate(angleDeg)
      } catch (eRotateFallback) {}
    }
    return tf
  }

  function drawLine(
    lineGroup,
    labelLayer,
    originX,
    originY,
    start,
    end,
    options,
  ) {
    if (!lineGroup || !start || !end) {
      return {
        path: null,
        label: null,
      }
    }
    var opts = options || {}
    var path = lineGroup.pathItems.add()
    path.stroked = true
    path.strokeWidth = LINE_STROKE_PT
    try {
      path.strokeColor = blackColor()
    } catch (eStroke) {}
    path.filled = false
    try {
      path.closed = false
    } catch (eClosed) {}
    path.strokeDashes = opts.dashed ? DASH_PATTERN : []
    path.setEntirePath([
      toArt(start, originX, originY),
      toArt(end, originX, originY),
    ])
    if (opts.name) {
      try {
        path.name = opts.name
      } catch (eName) {}
    }
    var label = null
    if (opts.labelText) {
      label = createLineLabel(
        labelLayer,
        originX,
        originY,
        start,
        end,
        opts.labelText,
      )
    }
    return {
      path: path,
      label: label,
    }
  }

  function drawCurve(lineGroup, originX, originY, start, end, options) {
    if (!lineGroup || !start || !end) {
      return null
    }
    var startArt = toArt(start, originX, originY)
    var endArt = toArt(end, originX, originY)
    var path = lineGroup.pathItems.add()
    path.stroked = true
    path.strokeWidth = LINE_STROKE_PT
    try {
      path.strokeColor = blackColor()
    } catch (eCurveStroke) {}
    path.filled = false
    try {
      path.closed = false
    } catch (eCurveClosed) {}
    path.strokeDashes = []
    path.setEntirePath([startArt, endArt])
    try {
      var pts = path.pathPoints
      if (pts.length >= 2) {
        var dx = endArt[0] - startArt[0]
        var dy = endArt[1] - startArt[1]
        var handleScale = 0.4
        pts[0].rightDirection = [
          pts[0].anchor[0] + dx * handleScale,
          pts[0].anchor[1] + dy * handleScale,
        ]
        pts[1].leftDirection = [
          pts[1].anchor[0] - dx * handleScale,
          pts[1].anchor[1] - dy * handleScale,
        ]
      }
    } catch (eHandles) {}
    if (options && options.name) {
      try {
        path.name = options.name
      } catch (eCurveName) {}
    }
    return path
  }

  function createMarker(
    markersGroup,
    numbersGroup,
    originX,
    originY,
    position,
    label,
  ) {
    if (!markersGroup || !numbersGroup) {
      return {
        circle: null,
        text: null,
      }
    }
    if (!position || !isFinite(position.x) || !isFinite(position.y)) {
      return {
        circle: null,
        text: null,
      }
    }
    var anchor = toArt(position, originX, originY)
    var radiusPts = cm(MARKER_RADIUS_CM)
    var circle = markersGroup.pathItems.ellipse(
      anchor[1] + radiusPts,
      anchor[0] - radiusPts,
      radiusPts * 2,
      radiusPts * 2,
    )
    circle.stroked = true
    circle.strokeWidth = LINE_STROKE_PT
    try {
      circle.strokeColor = blackColor()
    } catch (eCircleStroke) {}
    circle.filled = true
    try {
      circle.fillColor = blackColor()
    } catch (eCircleFill) {}
    circle.strokeDashes = []
    var tf = null
    try {
      tf = numbersGroup.textFrames.add()
    } catch (eAddMarkerText) {
      if (doc && doc.textFrames) {
        try {
          tf = doc.textFrames.add()
          if (tf && tf.move) {
            try {
              tf.move(numbersGroup, ElementPlacement.PLACEATEND)
            } catch (eMoveMarkerText) {}
          }
        } catch (eDocAddText) {}
      }
    }
    if (tf) {
      tf.contents = label
      try {
        tf.textRange.characterAttributes.size = LABEL_FONT_PT
        tf.textRange.characterAttributes.fillColor = whiteColor()
        tf.textRange.paragraphAttributes.justification = Justification.CENTER
      } catch (eMarkerText) {}
      try {
        tf.position = anchor
      } catch (eMarkerPos) {
        try {
          tf.left = anchor[0]
          tf.top = anchor[1]
        } catch (eMarkerPosFallback) {}
      }
      centerTextFrame(tf, anchor)
      try {
        tf.zOrder(ZOrderMethod.BRINGTOFRONT)
      } catch (eMarkerZ) {}
    }
    return {
      circle: circle,
      text: tf,
    }
  }

  function getItemBounds(item) {
    if (!item) {
      return null
    }
    var bounds = null
    try {
      bounds = item.visibleBounds
    } catch (errVisible) {}
    if (!bounds) {
      try {
        bounds = item.geometricBounds
      } catch (errGeo) {}
    }
    return bounds
  }

  function unionBounds(existing, candidate) {
    if (!candidate) {
      return existing
    }
    if (!existing) {
      return candidate.slice(0)
    }
    return [
      Math.min(existing[0], candidate[0]),
      Math.max(existing[1], candidate[1]),
      Math.max(existing[2], candidate[2]),
      Math.min(existing[3], candidate[3]),
    ]
  }

  function cropArtboardAroundItems(doc, artboardIndex, items, marginCm) {
    if (!doc || artboardIndex < 0 || artboardIndex >= doc.artboards.length) {
      return
    }
    var bounds = null
    for (var i = 0; i < items.length; i++) {
      var itemBounds = getItemBounds(items[i])
      bounds = unionBounds(bounds, itemBounds)
    }
    if (!bounds) {
      return
    }
    var margin = cm(marginCm)
    var rect = [
      bounds[0] - margin,
      bounds[1] + margin,
      bounds[2] + margin,
      bounds[3] - margin,
    ]
    try {
      doc.artboards[artboardIndex].artboardRect = rect
    } catch (errCrop) {}
  }

  var doc = getOrCreateDocument()
  if (!doc) {
    return
  }

  try {
    doc.rulerUnits = RulerUnits.Centimeters
  } catch (errUnits) {}

  var defaultMeasurements = {
    bust: 88,
    waist: 68,
    hip: 94,
    bustEase: 5,
    waistEase: 3,
    napeToWaist: 41,
    shoulder: 12.25,
    backWidth: 34.4,
    waistToHip: 20.6,
    frontNeckDart: 0,
    frontWaistDart: 0,
    frontWaistDartBackOff: 2.5,
    backWaistDart: 0,
    sideWaistDart: 0,
    bustWaistDiff: 0,
    armscyeDepth: 21,
    chest: 32.4,
    neckSize: 37,
    closeWaistShaping: true,
    reducedDarting: false,
    showMeasurementPalette: false,
  }
  defaultMeasurements.frontNeckDart = computeFrontNeckDart(
    defaultMeasurements.bust,
    defaultMeasurements.chest,
    defaultMeasurements.backWidth,
  )
  defaultMeasurements.bustWaistDiff = Math.abs(
    computeWaistDiff(
      defaultMeasurements.bust,
      defaultMeasurements.waist,
      defaultMeasurements.bustEase,
      defaultMeasurements.waistEase,
    ),
  )
  var defaultWaistDarts = computeWaistDarts(
    defaultMeasurements.bust,
    defaultMeasurements.waist,
    defaultMeasurements.bustEase,
    defaultMeasurements.waistEase,
  )
  defaultMeasurements.frontWaistDart = defaultWaistDarts.front
  defaultMeasurements.backWaistDart = defaultWaistDarts.back
  defaultMeasurements.sideWaistDart = defaultWaistDarts.side

  var measurements = showMeasurementDialog(defaultMeasurements)
  if (!measurements) {
    closeMeasurementPalette()
    return
  }
  if (measurements.showMeasurementPalette) {
    showMeasurementPalette(measurements)
  } else {
    closeMeasurementPalette()
  }

  var layers = ensureLayerStructure(doc)
  var artboardIndex = doc.artboards.getActiveArtboardIndex()
  var artboard = doc.artboards[artboardIndex]
  var artboardRect = artboard.artboardRect
  var originX = artboardRect[0]
  var originY = artboardRect[1]

  var draftGroup = resetLayerContents(layers.basicFrame)
  var backDraftGroup = resetLayerContents(layers.back)
  var frontDraftGroup = resetLayerContents(layers.front)
  var markersGroup = resetLayerContents(layers.numbersAndMarkers.markers)
  var numbersGroup = resetLayerContents(layers.numbersAndMarkers.numbers)
  var labelsGroup = null; // Labels layer removed; keep null for createLineLabel guards.
  try {
    if (markersGroup && markersGroup.zOrder) {
      markersGroup.zOrder(ZOrderMethod.SENDTOBACK)
    }
  } catch (eMarkersOrder) {}
  try {
    if (numbersGroup && numbersGroup.zOrder) {
      numbersGroup.zOrder(ZOrderMethod.BRINGTOFRONT)
    }
  } catch (eNumbersOrder) {}

  var items = []
  var points = {}

  function registerPoint(id, coords) {
    points[id] = coords
    var marker = createMarker(
      markersGroup,
      numbersGroup,
      originX,
      originY,
      coords,
      String(id),
    )
    if (marker.circle) items.push(marker.circle)
    if (marker.text) items.push(marker.text)
    return coords
  }

  function registerLetterPoint(id, coords, label) {
    points[id] = coords
    if (!labelsGroup) {
      return coords
    }
    var text = (label || id).toString().toLowerCase()
    var anchor = toArt(coords, originX, originY)
    var tf = labelsGroup.textFrames.add()
    tf.contents = text
    try {
      tf.textRange.characterAttributes.size = LABEL_FONT_PT
      tf.textRange.characterAttributes.fillColor = blackColor()
      tf.textRange.paragraphAttributes.justification = Justification.CENTER
    } catch (errLetterAttr) {}
    tf.position = anchor
    centerTextFrame(tf, anchor)
    items.push(tf)
    return coords
  }

  function buildLineOptions(name, ptA, ptB) {
    return {
      name: name,
      labelText: name + ' ' + formatCm(distanceBetween(ptA, ptB)),
    }
  }

  function addLine(idA, idB, options) {
    var result = drawLine(
      draftGroup,
      labelsGroup,
      originX,
      originY,
      points[idA],
      points[idB],
      options,
    )
    if (result.path) items.push(result.path)
    if (result.label) items.push(result.label)
    return result
  }

  var depth01 = 1.5
  var depth12 = measurements.armscyeDepth + 0.5
  var vertical02 = depth01 + depth12
  var halfBust = measurements.bust / 2
  var width23 = halfBust + measurements.bustEase
  var napeToWaist = measurements.napeToWaist
  registerPoint('0', { x: 0, y: 0 })
  registerPoint('1', { x: 0, y: depth01 })
  registerPoint('2', { x: 0, y: vertical02 })
  registerPoint('3', { x: width23, y: points['2'].y })
  registerPoint('4', { x: width23, y: 0 })
  registerPoint('5', { x: 0, y: points['1'].y + napeToWaist })
  registerPoint('6', { x: width23, y: points['5'].y })
  registerLetterPoint('c', { x: points['6'].x, y: points['6'].y + 1 }, 'c')
  var guide5Options = buildLineOptions(
    'Waist Drop Guide',
    points['5'],
    points['c'],
  )
  guide5Options.dashed = false
  guide5Options.labelText = null
  var guide5toC = drawLine(
    layers.basicFrame,
    labelsGroup,
    originX,
    originY,
    points['5'],
    points['c'],
    guide5Options,
  )
  if (guide5toC.path) items.push(guide5toC.path)
  if (guide5toC.label) items.push(guide5toC.label)
  var point5 = points['5']
  registerPoint('9', { x: measurements.neckSize / 5 - 0.2, y: points['0'].y })
  registerPoint('10', {
    x: points['1'].x,
    y: points['1'].y + measurements.armscyeDepth / 5 - 0.7,
  })
  registerPoint('20', {
    x: points['4'].x - (measurements.neckSize / 5 - 0.7),
    y: points['4'].y,
  })
  registerPoint('21', {
    x: points['4'].x,
    y: points['4'].y + (measurements.neckSize / 5 - 0.2),
  })
  registerPoint('22', {
    x: points['3'].x - measurements.chest / 2 - measurements.frontNeckDart / 2,
    y: points['3'].y,
  })
  var distance22toB = computePointBDistance(measurements.bust)
  if (distance22toB > 0) {
    var diagComponentB = distance22toB / Math.SQRT2
    var pointB = {
      x: points['22'].x - diagComponentB,
      y: points['22'].y - diagComponentB,
    }
    registerLetterPoint('b', pointB, 'b')
    var guide22Options = buildLineOptions(
      'Front Armhole Guideline',
      points['22'],
      points['b'],
    )
    guide22Options.dashed = true
    guide22Options.labelText = null
    var guide22toB = drawLine(
      layers.basicFrame,
      labelsGroup,
      originX,
      originY,
      points['22'],
      points['b'],
      guide22Options,
    )
    if (guide22toB.path) items.push(guide22toB.path)
    if (guide22toB.label) items.push(guide22toB.label)
  }
  registerPoint('23', midpoint(points['3'], points['22']))
  var waistLineY = points['5'].y
  registerPoint('24', { x: points['23'].x, y: waistLineY })
  registerPoint('26', { x: points['23'].x, y: points['23'].y + 2.5 })
  registerPoint('27', {
    x: points['20'].x - measurements.frontNeckDart,
    y: points['20'].y,
  })
  var dart20Options = buildLineOptions(
    'Front Neck Dart Left Leg',
    points['20'],
    points['26'],
  )
  dart20Options.labelText = null
  var dart20to26 = drawLine(
    frontDraftGroup,
    labelsGroup,
    originX,
    originY,
    points['20'],
    points['26'],
    dart20Options,
  )
  if (dart20to26.path) items.push(dart20to26.path)
  if (dart20to26.label) items.push(dart20to26.label)
  var dart27Options = buildLineOptions(
    'Front Neck Dart Right Leg',
    points['27'],
    points['26'],
  )
  dart27Options.labelText = null
  var dart27to26 = drawLine(
    frontDraftGroup,
    labelsGroup,
    originX,
    originY,
    points['27'],
    points['26'],
    dart27Options,
  )
  if (dart27to26.path) items.push(dart27to26.path)
  if (dart27to26.label) items.push(dart27to26.label)
  var thirdOf3To21 = (points['21'].y - points['3'].y) / 3
  registerPoint('31', { x: points['22'].x, y: points['22'].y + thirdOf3To21 })
  var guide2231Options = buildLineOptions(
    'Chest Line',
    points['22'],
    points['31'],
  )
  guide2231Options.dashed = true
  guide2231Options.labelText = null
  var guide22to31 = drawLine(
    layers.basicFrame,
    labelsGroup,
    originX,
    originY,
    points['22'],
    points['31'],
    guide2231Options,
  )
  if (guide22to31.path) items.push(guide22to31.path)
  if (guide22to31.label) items.push(guide22to31.label)

  var bustLineOptions = buildLineOptions('Bust Line', points['2'], points['3'])
  bustLineOptions.dashed = true
  addLine('2', '3', bustLineOptions)
  var cfOptions = buildLineOptions(
    'Centre Front (CF)',
    points['21'],
    points['c'],
  )
  cfOptions.labelText =
    'Centre Front (CF) ' + formatCm(distanceBetween(points['21'], points['c']))
  var cfLine = drawLine(
    frontDraftGroup,
    labelsGroup,
    originX,
    originY,
    points['21'],
    points['c'],
    cfOptions,
  )
  if (cfLine.path) items.push(cfLine.path)
  if (cfLine.label) items.push(cfLine.label)
  var baseLineOptions = buildLineOptions('0 - 4', points['4'], points['0'])
  baseLineOptions.labelText = null
  addLine('4', '0', baseLineOptions)
  var foundationCbOptions = buildLineOptions(
    'Foundation Centre Back',
    points['0'],
    point5,
  )
  foundationCbOptions.labelText = null
  addLine('0', '5', foundationCbOptions)
  var waistLineOptions = buildLineOptions('Waistline', points['5'], points['6'])
  waistLineOptions.dashed = true
  addLine('5', '6', waistLineOptions)
  var frontEdgeOptions = buildLineOptions('CF Line', points['4'], points['c'])
  frontEdgeOptions.labelText = null
  addLine('4', 'c', frontEdgeOptions)
  var cbOptions = buildLineOptions('Centre Back (CB)', points['1'], point5)
  cbOptions.labelText =
    'Centre Back (CB) ' + formatCm(distanceBetween(points['1'], point5))
  var cbLine = drawLine(
    backDraftGroup,
    labelsGroup,
    originX,
    originY,
    points['1'],
    point5,
    cbOptions,
  )
  if (cbLine.path) items.push(cbLine.path)
  if (cbLine.label) items.push(cbLine.label)
  if (cfLine.label) {
    try {
      var cfMidpoint = midpoint(points['3'], points['4'])
      var cfAnchor = toArt(
        { x: cfMidpoint.x - LABEL_INSIDE_OFFSET_CM, y: cfMidpoint.y },
        originX,
        originY,
      )
      cfLine.label.position = cfAnchor
      centerTextFrame(cfLine.label, cfAnchor)
      cfLine.label.rotate(180, true, true, true, true, Transformation.CENTER)
    } catch (eAdjustCfLabel) {}
  }
  if (cbLine.label) {
    try {
      var cbMidpoint = midpoint(points['0'], point5)
      var cbAnchor = toArt(
        { x: cbMidpoint.x + LABEL_INSIDE_OFFSET_CM, y: cbMidpoint.y },
        originX,
        originY,
      )
      cbLine.label.position = cbAnchor
      centerTextFrame(cbLine.label, cbAnchor)
      cbLine.label.rotate(180, true, true, true, true, Transformation.CENTER)
    } catch (eAdjustCbLabel) {}
  }

  var shoulderPlusOne = measurements.shoulder + 1
  var shoulderDeltaY = Math.abs(points['10'].y - points['9'].y)
  var shoulderHorizontal = 0
  if (shoulderPlusOne > shoulderDeltaY) {
    var shoulderHorizontalSq =
      Math.pow(shoulderPlusOne, 2) - Math.pow(shoulderDeltaY, 2)
    shoulderHorizontal = Math.sqrt(Math.max(shoulderHorizontalSq, 0))
  }
  registerPoint('11', {
    x: points['9'].x + shoulderHorizontal,
    y: points['10'].y,
  })
  var shoulderLine = drawLine(
    backDraftGroup,
    labelsGroup,
    originX,
    originY,
    points['9'],
    points['11'],
    buildLineOptions('Back Shoulder Line', points['9'], points['11']),
  )
  if (shoulderLine.path) items.push(shoulderLine.path)
  if (shoulderLine.label) items.push(shoulderLine.label)

  registerPoint('12', midpoint(points['9'], points['11']))
  var guide12Down = { x: points['12'].x, y: points['12'].y + 5 }
  var guide12Vertical = drawLine(
    backDraftGroup,
    labelsGroup,
    originX,
    originY,
    points['12'],
    guide12Down,
    { dashed: true },
  )
  registerPoint('13', { x: guide12Down.x - 1, y: guide12Down.y })
  var guide12Horizontal = drawLine(
    backDraftGroup,
    labelsGroup,
    originX,
    originY,
    guide12Down,
    points['13'],
    { dashed: true },
  )
  var backShoulderDartGuide = null
  if (guide12Vertical.path && guide12Horizontal.path) {
    var pointsVertical = guide12Vertical.path.pathPoints
    var pointsHorizontal = guide12Horizontal.path.pathPoints
    if (pointsVertical.length === 2 && pointsHorizontal.length === 2) {
      var combinedPath = layers.basicFrame.pathItems.add()
      combinedPath.stroked = true
      combinedPath.strokeWidth = LINE_STROKE_PT
      combinedPath.strokeDashes = DASH_PATTERN
      try {
        combinedPath.strokeColor = blackColor()
      } catch (eCombineColor) {}
      combinedPath.filled = false
      try {
        combinedPath.closed = false
      } catch (eCombinedClosed) {}
      var firstSegment = [
        toArt(points['12'], originX, originY),
        toArt(guide12Down, originX, originY),
        toArt(points['13'], originX, originY),
      ]
      combinedPath.setEntirePath(firstSegment)
      try {
        combinedPath.name = 'Back Shoulder Dart Guide'
      } catch (eNameCombined) {}
      backShoulderDartGuide = combinedPath
      try {
        guide12Vertical.path.remove()
      } catch (eRemoveVertical) {}
      try {
        guide12Horizontal.path.remove()
      } catch (eRemoveHorizontal) {}
    }
  }
  if (backShoulderDartGuide) {
    items.push(backShoulderDartGuide)
  } else {
    if (guide12Vertical.path) items.push(guide12Vertical.path)
    if (guide12Horizontal.path) items.push(guide12Horizontal.path)
  }
  var shoulderDartHalfWidth = 0.5
  if (points['9'] && points['11'] && points['12'] && points['13']) {
    var shoulderSegmentLength = distanceBetween(points['9'], points['11'])
    if (shoulderSegmentLength > 0.0001) {
      var shoulderUnitX =
        (points['11'].x - points['9'].x) / shoulderSegmentLength
      var shoulderUnitY =
        (points['11'].y - points['9'].y) / shoulderSegmentLength
      var shoulderDartLeft = {
        x: points['12'].x - shoulderUnitX * shoulderDartHalfWidth,
        y: points['12'].y - shoulderUnitY * shoulderDartHalfWidth,
      }
      var shoulderDartRight = {
        x: points['12'].x + shoulderUnitX * shoulderDartHalfWidth,
        y: points['12'].y + shoulderUnitY * shoulderDartHalfWidth,
      }
      var leftLegLength = distanceBetween(shoulderDartLeft, points['13'])
      var rightVector = {
        x: shoulderDartRight.x - points['13'].x,
        y: shoulderDartRight.y - points['13'].y,
      }
      var rightVectorLength = Math.sqrt(
        rightVector.x * rightVector.x + rightVector.y * rightVector.y,
      )
      var shoulderDartRightAdjusted = shoulderDartRight
      if (rightVectorLength > 0.0001 && leftLegLength > 0.0001) {
        var scale = leftLegLength / rightVectorLength
        shoulderDartRightAdjusted = {
          x: points['13'].x + rightVector.x * scale,
          y: points['13'].y + rightVector.y * scale,
        }
      }
      var shoulderDartPath = backDraftGroup.pathItems.add()
      shoulderDartPath.stroked = true
      shoulderDartPath.strokeWidth = LINE_STROKE_PT
      try {
        shoulderDartPath.strokeColor = blackColor()
      } catch (eShoulderDartColor) {}
      shoulderDartPath.filled = false
      try {
        shoulderDartPath.closed = false
      } catch (eShoulderDartClosed) {}
      shoulderDartPath.strokeDashes = []
      shoulderDartPath.setEntirePath([
        toArt(shoulderDartLeft, originX, originY),
        toArt(points['13'], originX, originY),
        toArt(shoulderDartRightAdjusted, originX, originY),
      ])
      try {
        shoulderDartPath.name = 'Back Shoulder Dart'
      } catch (eShoulderDartName) {}
      items.push(shoulderDartPath)

      if (shoulderLine && shoulderLine.path) {
        shoulderLine.path.setEntirePath([
          toArt(points['9'], originX, originY),
          toArt(shoulderDartLeft, originX, originY),
        ])
      }
      if (shoulderLine && shoulderLine.label) {
        try {
          shoulderLine.label.remove()
        } catch (eShoulderLineLabel) {}
        shoulderLine.label = null
      }
      var shoulderRightConnectOptions = buildLineOptions(
        'Back Shoulder Dart Right to 11',
        shoulderDartRightAdjusted,
        points['11'],
      )
      shoulderRightConnectOptions.labelText = null
      var shoulderRightConnectLine = drawLine(
        backDraftGroup,
        labelsGroup,
        originX,
        originY,
        shoulderDartRightAdjusted,
        points['11'],
        shoulderRightConnectOptions,
      )
      if (shoulderRightConnectLine.path)
        items.push(shoulderRightConnectLine.path)
      if (shoulderRightConnectLine.label)
        items.push(shoulderRightConnectLine.label)
    }
  }
  var curve19 = drawCurve(
    backDraftGroup,
    originX,
    originY,
    points['1'],
    points['9'],
  )
  if (curve19) {
    items.push(curve19)
    try {
      var curvePoints = curve19.pathPoints
      if (curvePoints.length >= 2) {
        var chord19Cm = distanceBetween(points['1'], points['9'])
        var handle1Vector = clampHandleToChord(
          4.3,
          0,
          chord19Cm,
          HANDLE_RATIO_CAP,
        )
        curvePoints[0].rightDirection = [
          curvePoints[0].anchor[0] + handle1Vector.x,
          curvePoints[0].anchor[1] + handle1Vector.y,
        ]
        var handle9Vector = clampHandleToChord(
          -1,
          -1,
          chord19Cm,
          HANDLE_RATIO_CAP,
        )
        curvePoints[1].leftDirection = [
          curvePoints[1].anchor[0] + handle9Vector.x,
          curvePoints[1].anchor[1] + handle9Vector.y,
        ]
      }
    } catch (eCurveHandle) {}
    try {
      curve19.name = 'Back Neck Curve'
      var backNeckLabel = createLineLabel(
        labelsGroup,
        originX,
        originY,
        points['1'],
        points['9'],
        'Back Neck Curve ' +
          formatCm(distanceBetween(points['1'], points['9'])),
      )
      if (backNeckLabel) {
        items.push(backNeckLabel)
      }
    } catch (eNameCurve19) {}
  }
  var curve20to21 = drawCurve(
    frontDraftGroup,
    originX,
    originY,
    points['20'],
    points['21'],
  )
  if (curve20to21) {
    items.push(curve20to21)
    try {
      curve20to21.name = 'Front Neck Curve'
    } catch (eNameCurve2021) {}
    var frontNeckLabel = createLineLabel(
      labelsGroup,
      originX,
      originY,
      points['20'],
      points['21'],
      'Front Neck Curve ' +
        formatCm(distanceBetween(points['20'], points['21'])),
    )
    if (frontNeckLabel) items.push(frontNeckLabel)
  }
  try {
    var curve2021Points = curve20to21.pathPoints
    if (curve2021Points.length >= 2) {
      var chord2021Cm = distanceBetween(points['20'], points['21'])
      var handle20Vector = clampHandleToChord(
        0,
        -4,
        chord2021Cm,
        HANDLE_RATIO_CAP,
      )
      curve2021Points[0].rightDirection = [
        curve2021Points[0].anchor[0] + handle20Vector.x,
        curve2021Points[0].anchor[1] + handle20Vector.y,
      ]
      var handle21Vector = clampHandleToChord(
        -4,
        0,
        chord2021Cm,
        HANDLE_RATIO_CAP,
      )
      curve2021Points[1].leftDirection = [
        curve2021Points[1].anchor[0] + handle21Vector.x,
        curve2021Points[1].anchor[1] + handle21Vector.y,
      ]
    }
  } catch (eCurve2021) {}
  registerPoint('28', { x: points['11'].x, y: points['11'].y + 1.5 })
  var guide1128Options = buildLineOptions(
    'Back Shoulder Drop',
    points['11'],
    points['28'],
  )
  guide1128Options.dashed = true
  guide1128Options.labelText = null
  var guide11to28 = drawLine(
    layers.basicFrame,
    labelsGroup,
    originX,
    originY,
    points['11'],
    points['28'],
    guide1128Options,
  )
  if (guide11to28.path) items.push(guide11to28.path)
  if (guide11to28.label) items.push(guide11to28.label)
  registerPoint('29', { x: points['28'].x + 10, y: points['28'].y })
  var guide2829Options = buildLineOptions(
    'Shoulder Balance Line',
    points['28'],
    points['29'],
  )
  guide2829Options.dashed = true
  guide2829Options.labelText = null
  var guide28to29 = drawLine(
    layers.basicFrame,
    labelsGroup,
    originX,
    originY,
    points['28'],
    points['29'],
    guide2829Options,
  )
  if (guide28to29.path) items.push(guide28to29.path)
  if (guide28to29.label) items.push(guide28to29.label)
  var frontShoulderDeltaY = Math.abs(points['28'].y - points['27'].y)
  var frontShoulderHorizontalSq =
    Math.pow(measurements.shoulder, 2) - Math.pow(frontShoulderDeltaY, 2)
  var frontShoulderHorizontal =
    frontShoulderHorizontalSq > 0 ? Math.sqrt(frontShoulderHorizontalSq) : 0
  var candidateX30 = points['27'].x - frontShoulderHorizontal
  var minFrontX = Math.min(points['28'].x, points['29'].x)
  var maxFrontX = Math.max(points['28'].x, points['29'].x)
  if (candidateX30 < minFrontX) candidateX30 = minFrontX
  if (candidateX30 > maxFrontX) candidateX30 = maxFrontX
  registerPoint('30', { x: candidateX30, y: points['28'].y })
  var frontShoulderLine = drawLine(
    frontDraftGroup,
    labelsGroup,
    originX,
    originY,
    points['27'],
    points['30'],
    buildLineOptions('Front Shoulder Line', points['27'], points['30']),
  )
  if (frontShoulderLine.path) items.push(frontShoulderLine.path)
  if (frontShoulderLine.label) items.push(frontShoulderLine.label)

  registerPoint('14', {
    x: points['2'].x + measurements.backWidth / 2 + 0.5,
    y: points['2'].y,
  })
  var distance14toA = computePointADistance(measurements.bust)
  if (distance14toA > 0) {
    var diagonalComponent = distance14toA / Math.SQRT2
    var pointA = {
      x: points['14'].x + diagonalComponent,
      y: points['14'].y - diagonalComponent,
    }
    registerLetterPoint('a', pointA, 'a')
    var guide14aOptions = buildLineOptions(
      'Back Armhole Guideline',
      points['14'],
      points['a'],
    )
    guide14aOptions.dashed = true
    guide14aOptions.labelText = null
    var guide14toA = drawLine(
      layers.basicFrame,
      labelsGroup,
      originX,
      originY,
      points['14'],
      points['a'],
      guide14aOptions,
    )
    if (guide14toA.path) items.push(guide14toA.path)
    if (guide14toA.label) items.push(guide14toA.label)
  }
  registerPoint('15', { x: points['14'].x, y: points['10'].y })
  var guide14Options = buildLineOptions(
    'Back Width Line',
    points['14'],
    points['15'],
  )
  guide14Options.dashed = true
  guide14Options.labelText = null
  var guide14Vertical = drawLine(
    layers.basicFrame,
    labelsGroup,
    originX,
    originY,
    points['14'],
    points['15'],
    guide14Options,
  )
  if (guide14Vertical.path) items.push(guide14Vertical.path)
  if (guide14Vertical.label) items.push(guide14Vertical.label)
  registerPoint('16', midpoint(points['14'], points['15']))
  registerPoint('17', midpoint(points['2'], points['14']))
  registerPoint('32', midpoint(points['14'], points['22']))
  var curve11to32 = drawCurve(
    backDraftGroup,
    originX,
    originY,
    points['11'],
    points['32'],
  )
  if (curve11to32) {
    items.push(curve11to32)
    try {
      var curve1132Points = curve11to32.pathPoints
      if (curve1132Points.length >= 2) {
        var backChordCm = distanceBetween(points['11'], points['32'])
        var handle11Vector = clampHandleToChord(
          3.2,
          9.79,
          backChordCm,
          HANDLE_RATIO_CAP,
        )
        curve1132Points[0].rightDirection = [
          curve1132Points[0].anchor[0] - handle11Vector.x,
          curve1132Points[0].anchor[1] - handle11Vector.y,
        ]
        var handle32Vector = clampHandleToChord(
          6.15,
          0,
          backChordCm,
          HANDLE_RATIO_CAP,
        )
        curve1132Points[1].leftDirection = [
          curve1132Points[1].anchor[0] - handle32Vector.x,
          curve1132Points[1].anchor[1] - handle32Vector.y,
        ]
      }
    } catch (eCurve1132) {}
    try {
      curve11to32.name = 'Back Armhole Curve'
    } catch (eNameCurve1132) {}
    try {
      var backArmholeLength = curve11to32.length
      if (isFiniteNumber(backArmholeLength)) {
        var backArmholeLabel = createLineLabel(
          labelsGroup,
          originX,
          originY,
          points['11'],
          points['32'],
          'Back Armhole Curve ' + formatCm(ptToCm(backArmholeLength)),
        )
        if (backArmholeLabel) {
          items.push(backArmholeLabel)
        }
      }
    } catch (eBackArmholeLabel) {}
  }
  var curve30to32 = drawCurve(
    frontDraftGroup,
    originX,
    originY,
    points['30'],
    points['32'],
  )
  if (curve30to32) {
    items.push(curve30to32)
    try {
      var curve3032Points = curve30to32.pathPoints
      if (curve3032Points.length >= 2) {
        var frontChordCm = distanceBetween(points['30'], points['32'])
        var handle30Vector = clampHandleToChord(
          7,
          11.25,
          frontChordCm,
          HANDLE_RATIO_CAP,
        )
        curve3032Points[0].rightDirection = [
          curve3032Points[0].anchor[0] + handle30Vector.x,
          curve3032Points[0].anchor[1] - handle30Vector.y,
        ]
        var handle32RightVector = clampHandleToChord(
          6.47,
          0,
          frontChordCm,
          HANDLE_RATIO_CAP,
        )
        curve3032Points[1].leftDirection = [
          curve3032Points[1].anchor[0] + handle32RightVector.x,
          curve3032Points[1].anchor[1] - handle32RightVector.y,
        ]
      }
    } catch (eCurve3032) {}
    try {
      curve30to32.name = 'Front Armhole Curve'
    } catch (eNameCurve3032) {}
    try {
      var frontArmholeLength = curve30to32.length
      if (isFiniteNumber(frontArmholeLength)) {
        var frontArmholeLabel = createLineLabel(
          labelsGroup,
          originX,
          originY,
          points['30'],
          points['32'],
          'Front Armhole Curve ' + formatCm(ptToCm(frontArmholeLength)),
        )
        if (frontArmholeLabel) {
          items.push(frontArmholeLabel)
        }
      }
    } catch (eFrontArmholeLabel) {}
  }
  registerPoint('33', { x: points['32'].x, y: waistLineY })
  registerPoint('18', { x: points['17'].x, y: points['5'].y })

  var line5 = points['5']
  var pointC = points['c']
  if (line5 && pointC) {
    var intersectionSpecs = [
      {
        id: 'd',
        x: points['17'] ? points['17'].x : null,
        top: points['17'],
        bottom: { x: points['17'] ? points['17'].x : null, y: pointC.y },
      },
      {
        id: 'e',
        x: points['23'] ? points['23'].x : null,
        top: points['23'],
        bottom: { x: points['23'] ? points['23'].x : null, y: pointC.y },
      },
      {
        id: 'f',
        x: points['32'] ? points['32'].x : null,
        top: points['32'],
        bottom: { x: points['32'] ? points['32'].x : null, y: pointC.y },
      },
    ]
    for (var interIdx = 0; interIdx < intersectionSpecs.length; interIdx++) {
      var spec = intersectionSpecs[interIdx]
      if (!spec.top || !spec.bottom || !isFiniteNumber(spec.x)) {
        continue
      }
      var intersectionPoint = intersectLineWithVertical(line5, pointC, spec.x)
      if (!intersectionPoint) {
        continue
      }
      if (
        !valueBetween(intersectionPoint.y, spec.top.y, spec.bottom.y, 0.001)
      ) {
        continue
      }
      registerLetterPoint(spec.id, intersectionPoint, spec.id)
    }
  }

  if (points['d']) {
    var backWaistDartGroup = backDraftGroup.groupItems.add()
    backWaistDartGroup.name = 'Back Waist Darts'
    var trimmedBackDartOptions = buildLineOptions(
      'Back Waist Dart Bisector',
      points['17'],
      points['d'],
    )
    trimmedBackDartOptions.dashed = true
    trimmedBackDartOptions.labelText = null
    var trimmedBackDart = drawLine(
      backWaistDartGroup,
      labelsGroup,
      originX,
      originY,
      points['17'],
      points['d'],
      trimmedBackDartOptions,
    )
    if (trimmedBackDart.path) items.push(trimmedBackDart.path)
    if (trimmedBackDart.label) items.push(trimmedBackDart.label)
  }
  if (points['e']) {
    var frontWaistDartGroup = frontDraftGroup.groupItems.add()
    frontWaistDartGroup.name = 'Front Waist Darts'
    var frontWaistDartWidth = measurements.frontWaistDart
    if (
      isFiniteNumber(frontWaistDartWidth) &&
      frontWaistDartWidth > 0.0001 &&
      points['26']
    ) {
      var frontDartHalf = frontWaistDartWidth / 2
      var dartBaseLeft = { x: points['e'].x - frontDartHalf, y: points['e'].y }
      var dartBaseRight = { x: points['e'].x + frontDartHalf, y: points['e'].y }
      var frontBackOff = isFiniteNumber(measurements.frontWaistDartBackOff)
        ? measurements.frontWaistDartBackOff
        : 2.5
      var frontDartApex = { x: points['e'].x, y: points['26'].y + frontBackOff }
      var frontLeftDartOptions = buildLineOptions(
        'Front Waist Dart Left Leg',
        dartBaseLeft,
        frontDartApex,
      )
      frontLeftDartOptions.labelText = null
      var frontRightDartOptions = buildLineOptions(
        'Front Waist Dart Right Leg',
        dartBaseRight,
        frontDartApex,
      )
      frontRightDartOptions.labelText = null
      var frontLeftDart = drawLine(
        frontWaistDartGroup,
        labelsGroup,
        originX,
        originY,
        dartBaseLeft,
        frontDartApex,
        frontLeftDartOptions,
      )
      var frontRightDart = drawLine(
        frontWaistDartGroup,
        labelsGroup,
        originX,
        originY,
        dartBaseRight,
        frontDartApex,
        frontRightDartOptions,
      )
      if (frontLeftDart.path) items.push(frontLeftDart.path)
      if (frontRightDart.path) items.push(frontRightDart.path)
      if (frontLeftDart.label) items.push(frontLeftDart.label)
      if (frontRightDart.label) items.push(frontRightDart.label)
      var trimmedFrontDartOptions = buildLineOptions(
        'Front Waist Dart Bisector',
        frontDartApex,
        points['e'],
      )
      trimmedFrontDartOptions.dashed = true
      trimmedFrontDartOptions.labelText = null
      var trimmedFrontDart = drawLine(
        frontWaistDartGroup,
        labelsGroup,
        originX,
        originY,
        frontDartApex,
        points['e'],
        trimmedFrontDartOptions,
      )
      if (trimmedFrontDart.path) items.push(trimmedFrontDart.path)
      if (trimmedFrontDart.label) items.push(trimmedFrontDart.label)
    }
  }
  if (points['f']) {
    var trimmedPrincessOptions = buildLineOptions(
      'Side',
      points['32'],
      points['f'],
    )
    trimmedPrincessOptions.dashed = true
    trimmedPrincessOptions.labelText = null
    var trimmedPrincessLine = drawLine(
      layers.basicFrame,
      labelsGroup,
      originX,
      originY,
      points['32'],
      points['f'],
      trimmedPrincessOptions,
    )
    if (trimmedPrincessLine.path) items.push(trimmedPrincessLine.path)
    if (trimmedPrincessLine.label) items.push(trimmedPrincessLine.label)
    if (points['32']) {
      var sideWaistDartWidth = measurements.sideWaistDart
      if (isFiniteNumber(sideWaistDartWidth) && sideWaistDartWidth > 0.0001) {
        var sideOffset = (sideWaistDartWidth - 1) / 2
        if (!isFiniteNumber(sideOffset) || sideOffset < 0) {
          sideOffset = sideWaistDartWidth / 2
        }
        var sideLeftBase = { x: points['f'].x - sideOffset, y: points['f'].y }
        var sideRightBase = { x: points['f'].x + sideOffset, y: points['f'].y }
        var sideApex = { x: points['32'].x, y: points['32'].y }
        var sideLeftOptions = buildLineOptions(
          'Back Side Waist Dart',
          sideLeftBase,
          sideApex,
        )
        sideLeftOptions.labelText = null
        var sideRightOptions = buildLineOptions(
          'Front Side Waist Dart',
          sideRightBase,
          sideApex,
        )
        sideRightOptions.labelText = null
        var sideLeftDart = drawLine(
          backDraftGroup,
          labelsGroup,
          originX,
          originY,
          sideLeftBase,
          sideApex,
          sideLeftOptions,
        )
        var sideRightDart = drawLine(
          frontDraftGroup,
          labelsGroup,
          originX,
          originY,
          sideRightBase,
          sideApex,
          sideRightOptions,
        )
        var frontSideGuideOptions = buildLineOptions(
          'Front Waist Line',
          points['c'],
          sideRightBase,
        )
        frontSideGuideOptions.labelText = null
        var frontSideGuide = drawLine(
          frontDraftGroup,
          labelsGroup,
          originX,
          originY,
          points['c'],
          sideRightBase,
          frontSideGuideOptions,
        )
        if (frontSideGuide && frontSideGuide.path) {
          frontSideGuide.path.strokeDashes = []
          items.push(frontSideGuide.path)
          if (frontSideGuide.label) items.push(frontSideGuide.label)
        }
        var backSideGuideOptions = buildLineOptions(
          'Back Waist Line',
          point5,
          sideLeftBase,
        )
        backSideGuideOptions.labelText = null
        var backSideGuide = drawLine(
          backDraftGroup,
          labelsGroup,
          originX,
          originY,
          point5,
          sideLeftBase,
          backSideGuideOptions,
        )
        if (backSideGuide && backSideGuide.path) {
          backSideGuide.path.strokeDashes = []
          items.push(backSideGuide.path)
          if (backSideGuide.label) items.push(backSideGuide.label)
        }
        if (sideLeftDart.path) items.push(sideLeftDart.path)
        if (sideRightDart.path) items.push(sideRightDart.path)
        if (sideLeftDart.label) items.push(sideLeftDart.label)
        if (sideRightDart.label) items.push(sideRightDart.label)
      }
    }
  }
  if (points['d'] && points['17']) {
    var backWaistDartWidth = measurements.backWaistDart
    if (isFiniteNumber(backWaistDartWidth) && backWaistDartWidth > 0.0001) {
      var backHalf = backWaistDartWidth / 2
      var backLeftBase = { x: points['d'].x - backHalf, y: points['d'].y }
      var backRightBase = { x: points['d'].x + backHalf, y: points['d'].y }
      var backApex = { x: points['17'].x, y: points['17'].y }
      var backLeftOptions = buildLineOptions(
        'Back Waist Dart Left Leg',
        backLeftBase,
        backApex,
      )
      backLeftOptions.labelText = null
      var backRightOptions = buildLineOptions(
        'Back Waist Dart Right Leg',
        backRightBase,
        backApex,
      )
      backRightOptions.labelText = null
      var backLeftDart = drawLine(
        backWaistDartGroup,
        labelsGroup,
        originX,
        originY,
        backLeftBase,
        backApex,
        backLeftOptions,
      )
      var backRightDart = drawLine(
        backWaistDartGroup,
        labelsGroup,
        originX,
        originY,
        backRightBase,
        backApex,
        backRightOptions,
      )
      if (backLeftDart.path) items.push(backLeftDart.path)
      if (backRightDart.path) items.push(backRightDart.path)
      if (backLeftDart.label) items.push(backLeftDart.label)
      if (backRightDart.label) items.push(backRightDart.label)
    }
  }

  var backBladeOptions = buildLineOptions(
    'Back Blade',
    points['10'],
    points['1'],
  )
  backBladeOptions.labelText = null
  var backBladeLine = drawLine(
    layers.basicFrame,
    labelsGroup,
    originX,
    originY,
    points['10'],
    points['1'],
    backBladeOptions,
  )
  if (backBladeLine.path) items.push(backBladeLine.path)
  if (backBladeLine.label) items.push(backBladeLine.label)
  var backSquareEnd = {
    x: points['10'].x + width23 / 2,
    y: points['10'].y,
  }
  var backBladeGuideOptions = buildLineOptions(
    'Back Blade Guide',
    points['10'],
    backSquareEnd,
  )
  backBladeGuideOptions.dashed = true
  backBladeGuideOptions.labelText = null
  var dashedBackLine = drawLine(
    layers.basicFrame,
    labelsGroup,
    originX,
    originY,
    points['10'],
    backSquareEnd,
    backBladeGuideOptions,
  )
  if (dashedBackLine.path) items.push(dashedBackLine.path)
  if (dashedBackLine.label) items.push(dashedBackLine.label)

  try {
    if (layers.darts) {
      layers.darts.remove()
    }
  } catch (errRemoveDartsLayer) {}

  if (items.length) {
    cropArtboardAroundItems(doc, artboardIndex, items, 10)
  }

  try {
    doc.fitArtboardToWindow(artboardIndex)
  } catch (errFit) {}

  try {
    doc.artboards.setActiveArtboardIndex(artboardIndex)
    app.executeMenuCommand('fitartboardinwindow')
  } catch (errMenuFitArtboard) {}

  try {
    app.executeMenuCommand('fitall')
  } catch (errMenuFitAll) {}

  doc.selection = null
})()
