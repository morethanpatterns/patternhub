(function() {
    var doc = ensureDocument();
    if (!doc) return;

    function cm(val) {
        return val * 28.3464566929;
    }

    function ptToCm(val) {
        return val / 28.3464566929;
    }

    function ensureDocument() {
        var widthPt = cm(120);
        var heightPt = cm(120);

        function addDocument() {
            try {
                return app.documents.add(DocumentColorSpace.RGB, widthPt, heightPt);
            } catch (errAddDoc) {
                try {
                    return app.documents.add();
                } catch (errFallbackAdd) {
                    return null;
                }
            }
        }

        if (app.documents.length === 0) {
            return addDocument();
        }

        var active = null;
        try {
            active = app.activeDocument;
        } catch (errActiveDoc) {
            active = null;
        }
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
    var latestDerived = null;
    var lastRecommendedOptimalBalanceText = null;
    var measurementPalette = null;
    var selectedProfile = null;
    var defaults = {
        AhD: 20.1,
        NeG: 6.5,
        MoL: 75,
        BL: 41.6,
        BLBal: 0,
        HiD: 20,
        FitIndex: 3,
        BG: 16.5,
        AG: 9.3,
        BrG: 18.2,
        BrC: 88,
        WaC: 68,
        HiC: 97,
        ShG: 12.2,
        BrD: 28.1,
        FL: 45.3,
        BackShoulderEase: 0.7,
        OptimalBalance: 3.5,
        ShA: 20,
        ShoulderDifference: 2,
        showMeasurementPalette: false
    };
    var shouldShowMeasurementPalette = defaults.showMeasurementPalette !== false;
    var measurementLabelMap = {};

    function registerMeasurementLabel(key, label) {
        if (!key) return;
        measurementLabelMap[key] = label || key;
    }

    function resolveMeasurementLabel(key) {
        if (!key) return '';
        if (measurementLabelMap.hasOwnProperty(key)) return measurementLabelMap[key];
        return key;
    }
    var FIT_PROFILES = [{
        name: 'Fit 0',
        ease: {
            AhD: 0.25,
            BrC: 0,
            WaC: 0,
            HiC: 0,
            BG: 0,
            AG: 0,
            BrG: 0,
            ShG: 0
        },
        notes: {
            AhD: '0 - 0.5 cm',
            BrC: '0 cm',
            WaC: '0 cm',
            HiC: '0 cm'
        }
    }, {
        name: 'Fit 1',
        ease: {
            AhD: 0.45,
            BrC: 2,
            WaC: 1,
            HiC: 1,
            BG: 0.1,
            AG: 0.3,
            BrG: 0.6,
            ShG: 0.1
        },
        notes: {
            AhD: '0.2 - 0.7 cm',
            BrC: '2 cm',
            WaC: '0 - 2 cm',
            HiC: '0 - 2 cm'
        }
    }, {
        name: 'Fit 2',
        ease: {
            AhD: 0.75,
            BrC: 4,
            WaC: 3,
            HiC: 3,
            BG: 0.3,
            AG: 0.9,
            BrG: 0.8,
            ShG: 0.2
        },
        notes: {
            AhD: '0.5 - 1.0 cm',
            BrC: '4 cm',
            WaC: '2 - 4 cm',
            HiC: '2 - 4 cm'
        }
    }, {
        name: 'Fit 3',
        ease: {
            AhD: 1.3,
            BrC: 6,
            WaC: 5,
            HiC: 5,
            BG: 0.5,
            AG: 1.5,
            BrG: 1.0,
            ShG: 0.3
        },
        notes: {
            BrC: '6 cm',
            WaC: '4 - 6 cm',
            HiC: '4 - 6 cm'
        }
    }, {
        name: 'Fit 4',
        ease: {
            AhD: 1.7,
            BrC: 8,
            WaC: 6,
            HiC: 6,
            BG: 0.8,
            AG: 2.0,
            BrG: 1.2,
            ShG: 0.4
        },
        notes: {
            BrC: '8 cm',
            WaC: '4 - 8 cm',
            HiC: '4 - 8 cm'
        }
    }, {
        name: 'Fit 5',
        ease: {
            AhD: 2.1,
            BrC: 10,
            WaC: 10,
            HiC: 7,
            BG: 1.1,
            AG: 2.5,
            BrG: 1.4,
            ShG: 0.5
        },
        notes: {
            BrC: '10 cm',
            WaC: '8 - 12 cm',
            HiC: '6 - 8 cm'
        }
    }, {
        name: 'Fit 6',
        ease: {
            AhD: 2.5,
            BrC: 12,
            WaC: 12,
            HiC: 8,
            BG: 1.4,
            AG: 3.0,
            BrG: 1.6,
            ShG: 0.6
        },
        notes: {
            BrC: '12 cm',
            WaC: '8 - 16 cm',
            HiC: '6 - 10 cm'
        }
    }];
    var MEASURE_ROWS = [{
        id: 'AhD',
        label: '1. AhD',
        defaultValue: defaults.AhD
    }, {
        id: 'BrC',
        label: '2. BrC',
        defaultValue: defaults.BrC
    }, {
        id: 'WaC',
        label: '3. WaC',
        defaultValue: defaults.WaC
    }, {
        id: 'HiC',
        label: '4. HiC',
        defaultValue: defaults.HiC
    }, {
        id: 'BG',
        label: '5. BG',
        defaultValue: defaults.BG
    }, {
        id: 'AG',
        label: '6. AG',
        defaultValue: defaults.AG
    }, {
        id: 'BrG',
        label: '7. BrG',
        defaultValue: defaults.BrG
    }, {
        id: 'ShG',
        label: '8. ShG',
        defaultValue: defaults.ShG
    }];
    for (var measIdx = 0; measIdx < MEASURE_ROWS.length; measIdx++) {
        registerMeasurementLabel(MEASURE_ROWS[measIdx].id, MEASURE_ROWS[measIdx].label);
    }
    var BL_NUMBER = MEASURE_ROWS.length + 1;
    var BL_ROW_DEF = {
        id: 'BL',
        label: BL_NUMBER + '. BL',
        defaultValue: defaults.BL,
        finalLabel: BL_NUMBER + '. BL (final)'
    };
    registerMeasurementLabel(BL_ROW_DEF.id, BL_ROW_DEF.label);
    var FL_NUMBER = BL_NUMBER + 1;
    var SECONDARY_ROWS = [{
        id: 'NeG',
        labelBase: 'NeG',
        finalLabelBase: 'NeG (final)',
        defaultValue: defaults.NeG,
        easeOptions: {
            enabled: false
        },
        finalOptions: {
            enabled: false
        }
    }, {
        id: 'MoL',
        labelBase: 'MoL',
        finalLabelBase: 'MoL (final)',
        defaultValue: defaults.MoL,
        easeOptions: {
            enabled: false
        },
        finalOptions: {
            enabled: false
        }
    }, {
        id: 'HiD',
        labelBase: 'HiD',
        finalLabelBase: 'HiD (final)',
        defaultValue: defaults.HiD,
        easeOptions: {
            enabled: false
        },
        finalOptions: {
            enabled: false
        }
    }, {
        id: 'ShA',
        labelBase: 'ShA (deg)',
        finalLabelBase: 'ShA (deg)',
        defaultValue: defaults.ShA,
        easeOptions: {
            enabled: false
        },
        finalOptions: {
            enabled: false
        }
    }, {
        id: 'BrD',
        labelBase: 'BrD',
        finalLabelBase: 'BrD (final)',
        defaultValue: defaults.BrD,
        easeOptions: {
            enabled: false
        },
        finalOptions: {
            enabled: false
        }
    }];

    var SECONDARY_START_NUMBER = FL_NUMBER + 1;
    for (var sIdx = 0; sIdx < SECONDARY_ROWS.length; sIdx++) {
        var secondary = SECONDARY_ROWS[sIdx];
        var secondaryNumber = SECONDARY_START_NUMBER + sIdx;
        var baseLabel = secondary.labelBase || secondary.id;
        secondary.label = secondaryNumber + '. ' + baseLabel;
        var finalBase = secondary.finalLabelBase;
        if (finalBase) {
            secondary.finalLabel = secondaryNumber + '. ' + finalBase;
        } else if (secondary.finalLabel) {
            secondary.finalLabel = secondaryNumber + '. ' + secondary.finalLabel;
        } else {
            secondary.finalLabel = secondaryNumber + '. ' + baseLabel + ' (final)';
        }
    }
    for (var secondaryIdx = 0; secondaryIdx < SECONDARY_ROWS.length; secondaryIdx++) {
        registerMeasurementLabel(SECONDARY_ROWS[secondaryIdx].id, SECONDARY_ROWS[secondaryIdx].label);
    }

    function fitNames() {
        var names = [];
        for (var i = 0; i < FIT_PROFILES.length; i++) names.push(FIT_PROFILES[i].name);
        return names;
    }
    var dlg = new Window('dialog', 'Measurement Panel (Default Size: 38, Fit 3)');
    dlg.orientation = 'column';
    dlg.alignChildren = 'left';
    dlg.spacing = 12;
    var mainColumns = dlg.add('group');
    mainColumns.orientation = 'row';
    mainColumns.alignChildren = 'top';
    mainColumns.alignment = 'fill';
    mainColumns.spacing = 12;
    var measurementPanel = mainColumns.add('panel', undefined, 'Measurement Panel (Default Size: 38, Fit 3)');
    measurementPanel.orientation = 'column';
    measurementPanel.alignChildren = 'fill';
    measurementPanel.margins = 12;
    measurementPanel.spacing = 12;
    measurementPanel.alignment = 'fill';
    var optionsPanel = mainColumns.add('panel', undefined, 'Options');

    function addFieldRow(panel, label, defaultValue, options) {
        options = options || {};
        var row = panel.add('group');
        row.alignChildren = ['left', 'center'];
        var caption = label;
        if (options.note) caption += ' (' + options.note + ')';
        var st = row.add('statictext', undefined, caption + ':');
        if (options.labelWidth !== undefined) {
            st.minimumSize.width = options.labelWidth;
            st.preferredSize.width = options.labelWidth;
        }
        if (options.helpTip) {
            try {
                st.helpTip = options.helpTip;
            } catch (eHelpLabel) {}
        }
        if (options.note) {
            try {
                st.graphics.font = ScriptUI.newFont('dialog', 'italic', 8);
            } catch (eFont) {}
        }
        if (options && options.type === 'checkbox') {
            var checkbox = row.add('checkbox', undefined, '');
            checkbox.value = !!defaultValue;
            return checkbox;
        } else {
            var field = row.add('edittext', undefined, (defaultValue !== null && defaultValue !== undefined && defaultValue !== '') ? String(defaultValue) : '');
            field.characters = options.characters || 8;
            if (options.helpTip) {
                try {
                    field.helpTip = options.helpTip;
                } catch (eHelpField) {}
            }
            if (options.enabled === false) {
                try {
                    field.enabled = false;
                } catch (eOff) {}
            }
            return field;
        }
    }
    var columnWidths = {
        measurement: 210,
        ease: 180,
        construction: 210
    };
    var columnLabelWidths = {
        measurement: 120,
        ease: 120,
        construction: 140
    };
    var optionsLabelWidth = 220;
    var fitRow = measurementPanel.add('group');
    fitRow.orientation = 'row';
    fitRow.alignChildren = ['left', 'center'];
    fitRow.alignment = 'fill';
    fitRow.spacing = 6;
    fitRow.add('statictext', undefined, 'Fit Category:');
    var fitDropdown = fitRow.add('dropdownlist', undefined, fitNames());
    fitDropdown.selection = defaults.FitIndex;
    fitDropdown.preferredSize.width = 140;
    var headerRow = measurementPanel.add('group');
    headerRow.orientation = 'row';
    headerRow.alignChildren = ['left', 'center'];
    headerRow.spacing = 12;
    headerRow.alignment = 'fill';

    function addHeaderCell(label, width) {
        var header = headerRow.add('statictext', undefined, label);
        header.preferredSize.width = width;
        return header;
    }
    addHeaderCell('Main Measurement', columnWidths.measurement);
    addHeaderCell('Ease', columnWidths.ease);
    addHeaderCell('Construction Measurement', columnWidths.construction);
    var rowsContainer = measurementPanel.add('group');
    rowsContainer.orientation = 'column';
    rowsContainer.alignChildren = 'fill';
    rowsContainer.spacing = 6;
    rowsContainer.alignment = 'fill';

    function addDivider(target) {
        var dividerGroup = target.add('group');
        dividerGroup.alignment = 'fill';
        dividerGroup.margins = 0;
        var divider = dividerGroup.add('panel');
        divider.alignment = 'fill';
        divider.margins = 0;
        divider.minimumSize.height = 1;
        divider.maximumSize.height = 1;
    }

    function collapseGroup(group) {
        if (!group) return;
        try {
            group.visible = false;
        } catch (eVisibility) {}
        try {
            group.minimumSize.height = 0;
            group.maximumSize.height = 0;
        } catch (eHeight) {}
        try {
            group.margins = 0;
        } catch (eMargins) {}
        try {
            group.spacing = 0;
        } catch (eSpacing) {}
    }

    function configureDerivedField(field, labelWidth) {
        if (!field) return;
        var hostGroup = field.parent;
        if (!hostGroup) return;
        try {
            hostGroup.margins = 0;
        } catch (eMargin) {}
        try {
            hostGroup.spacing = 6;
        } catch (eSp) {}
        try {
            hostGroup.alignment = ['fill', 'top'];
        } catch (eAlign) {}
        if (labelWidth !== undefined) {
            for (var i = 0; i < hostGroup.children.length; i++) {
                if (hostGroup.children[i].type === 'statictext') {
                    try {
                        hostGroup.children[i].minimumSize.width = labelWidth;
                        hostGroup.children[i].preferredSize.width = labelWidth;
                    } catch (eWidth) {}
                    break;
                }
            }
        }
    }

    function addMeasurementRow(config) {
        config = config || {};
        var rowWrapper = rowsContainer.add('group');
        rowWrapper.orientation = 'column';
        rowWrapper.alignChildren = 'fill';
        rowWrapper.spacing = 6;
        var row = rowWrapper.add('group');
        row.orientation = 'row';
        row.alignChildren = ['left', 'center'];
        row.spacing = 12;
        row.alignment = 'fill';

        function addCell(width) {
            var cell = row.add('group');
            cell.orientation = 'column';
            cell.alignChildren = ['left', 'top'];
            cell.spacing = 4;
            cell.alignment = 'top';
            cell.preferredSize.width = width;
            cell.maximumSize.width = width;
            return cell;
        }
        var measCell = addCell(columnWidths.measurement);
        var easeColumnWidth = columnWidths.ease;
        if (config.noEaseColumn === true) easeColumnWidth = 0;
        var easeCell = addCell(easeColumnWidth);
        if (config.noEaseColumn === true) {
            try {
                easeCell.visible = false;
            } catch (eHideEase) {}
        }
        var consCell = addCell(columnWidths.construction);
        function prepareOptions(source, labelWidth) {
            var prepared = {};
            if (source) {
                for (var optKey in source) {
                    if (source.hasOwnProperty(optKey)) prepared[optKey] = source[optKey];
                }
            }
            if (labelWidth !== undefined && prepared.labelWidth === undefined) prepared.labelWidth = labelWidth;
            return prepared;
        }
        var measField = addFieldRow(measCell, config.label, config.defaultValue, prepareOptions(config.measOptions, columnLabelWidths.measurement));
        var easeField = null;
        if (config.noEaseColumn === true) {
            easeField = null;
        } else {
            easeField = addFieldRow(easeCell, config.easeLabel || config.label, config.easeDefault, prepareOptions(config.easeOptions, columnLabelWidths.ease));
        }
        var finalField = addFieldRow(consCell, config.finalLabel || config.label, config.finalDefault, prepareOptions(config.finalOptions, columnLabelWidths.construction));
        addDivider(rowWrapper);
        return {
            measField: measField,
            easeField: easeField,
            finalField: finalField,
            measCell: measCell,
            easeCell: easeCell,
            consCell: consCell
        };
    }
    var rowRefs = {};
    var secondaryRowRefs = {};
    var brwField = null;
    var wawField = null;
    var hiwField = null;
    var brgPlusField = null;
    var bShSField = null;

    function createRowForDefinition(def) {
        var rawEaseNote = (def.id === 'BL') ? '- / +' : (FIT_PROFILES[defaults.FitIndex].notes[def.id] || '');
        var easeNote = normalizeNote(rawEaseNote);
        var easeLabel = def.easeLabel;
        if (!easeLabel) {
            if (def.id === 'BL') easeLabel = 'BL Balance';
            else if (def.id === 'ShG') easeLabel = 'ShG Ease';
            else easeLabel = def.label;
        }
        var easeDefault = def.hasOwnProperty('easeDefault') ? def.easeDefault : ((def.id === 'BL') ? defaults.BLBal : '');
        var easeOptions = {};
        if (def.easeOptions) {
            for (var easeKey in def.easeOptions) {
                if (def.easeOptions.hasOwnProperty(easeKey)) easeOptions[easeKey] = def.easeOptions[easeKey];
            }
        }
        if (easeNote && easeOptions.note === undefined) easeOptions.note = easeNote;
        var finalLabel = def.finalLabel;
        if (!finalLabel) {
            if (def.id === 'AhD') finalLabel = '1. AhD+';
            else if (def.id === 'BG') finalLabel = '5. BG+';
            else if (def.id === 'AG') finalLabel = '6. AG+';
            else if (def.id === 'ShG') finalLabel = '8. fShS';
            else if (def.id === 'BL') finalLabel = '9. BL (final)';
            else finalLabel = def.label;
        }
        var finalDefault = def.hasOwnProperty('finalDefault') ? def.finalDefault : '';
        var finalOptions = {
            enabled: false
        };
        if (def.finalOptions) {
            finalOptions = {};
            for (var finalKey in def.finalOptions) {
                if (def.finalOptions.hasOwnProperty(finalKey)) finalOptions[finalKey] = def.finalOptions[finalKey];
            }
        }
        var measOptions = {};
        if (def.measOptions) {
            for (var measKey in def.measOptions) {
                if (def.measOptions.hasOwnProperty(measKey)) measOptions[measKey] = def.measOptions[measKey];
            }
        }
        var row = addMeasurementRow({
            label: def.label,
            defaultValue: def.defaultValue,
            measOptions: measOptions,
            easeLabel: easeLabel,
            easeDefault: easeDefault,
            easeOptions: easeOptions,
            finalLabel: finalLabel,
            finalDefault: finalDefault,
            finalOptions: finalOptions,
            noEaseColumn: def.noEaseColumn === true
        });
        var labelNumberPrefix = '';
        if (typeof def.label === 'string') {
            var numberMatch = /^(\d+)\.\s*/.exec(def.label);
            if (numberMatch) labelNumberPrefix = numberMatch[1] + '. ';
        }
        if (def.id === 'BrC') {
            var brwLabel = labelNumberPrefix ? labelNumberPrefix + 'BrW' : 'BrW';
            brwField = addFieldRow(row.consCell, brwLabel, '', {
                enabled: false,
                labelWidth: columnLabelWidths.construction
            });
            try {
                row.finalField.parent.visible = false;
            } catch (eHideBrC) {}
            collapseGroup(row.finalField.parent);
            configureDerivedField(brwField, columnLabelWidths.construction);
        }
        if (def.id === 'WaC') {
            var wawLabel = labelNumberPrefix ? labelNumberPrefix + 'WaW' : 'WaW';
            wawField = addFieldRow(row.consCell, wawLabel, '', {
                enabled: false,
                labelWidth: columnLabelWidths.construction
            });
            try {
                row.finalField.parent.visible = false;
            } catch (eHideWaW) {}
            collapseGroup(row.finalField.parent);
            configureDerivedField(wawField, columnLabelWidths.construction);
        }
        if (def.id === 'HiC') {
            var hiwLabel = labelNumberPrefix ? labelNumberPrefix + 'HiW' : 'HiW';
            hiwField = addFieldRow(row.consCell, hiwLabel, '', {
                enabled: false,
                labelWidth: columnLabelWidths.construction
            });
            try {
                row.finalField.parent.visible = false;
            } catch (eHideHiW) {}
            collapseGroup(row.finalField.parent);
            configureDerivedField(hiwField, columnLabelWidths.construction);
        }
        if (def.id === 'BrG') {
            var brgPlusLabel = labelNumberPrefix ? labelNumberPrefix + 'BrG+' : 'BrG+';
            brgPlusField = addFieldRow(row.consCell, brgPlusLabel, '', {
                enabled: false,
                labelWidth: columnLabelWidths.construction
            });
            try {
                row.finalField.parent.visible = false;
            } catch (eHideBrG) {}
            collapseGroup(row.finalField.parent);
            configureDerivedField(brgPlusField, columnLabelWidths.construction);
        }
        if (def.id === 'ShG') {
            var bShSLabel = labelNumberPrefix ? labelNumberPrefix + 'bShS' : 'bShS';
            bShSField = addFieldRow(row.consCell, bShSLabel, '', {
                enabled: false,
                labelWidth: columnLabelWidths.construction
            });
            configureDerivedField(bShSField, columnLabelWidths.construction);
        }
        rowRefs[def.id] = row;
        return row;
    }
    for (var i = 0; i < MEASURE_ROWS.length; i++) {
        createRowForDefinition(MEASURE_ROWS[i]);
    }
    createRowForDefinition(BL_ROW_DEF);
    var flEaseNote = normalizeNote('- / +');
    var flLabel = FL_NUMBER + '. FL';
    var flFinalLabel = FL_NUMBER + '. FL (final)';
    registerMeasurementLabel('FL', flLabel);
    var flRow = addMeasurementRow({
        label: flLabel,
        defaultValue: defaults.FL,
        easeLabel: 'FL Balance',
        easeDefault: 0,
        easeOptions: flEaseNote ? {
            note: flEaseNote
        } : {},
        finalLabel: flFinalLabel,
        finalDefault: '',
        finalOptions: {
            enabled: false
        }
    });
    rowRefs.FL = flRow;
    var flMeasureField = flRow.measField;
    var flBalanceField = flRow.easeField;
    var flFinalField = flRow.finalField;
    var optimalBalanceField = addFieldRow(flRow.consCell, 'Optimal Balance', defaults.OptimalBalance, {
        labelWidth: columnLabelWidths.construction
    });
    var individualBalanceField = addFieldRow(flRow.consCell, 'Individual Balance', '', {
        enabled: false,
        labelWidth: columnLabelWidths.construction
    });
    var finalBalanceField = addFieldRow(flRow.consCell, 'Final Balance', '', {
        enabled: false,
        labelWidth: columnLabelWidths.construction
    });
    for (var j = 0; j < SECONDARY_ROWS.length; j++) {
        var secondaryDef = SECONDARY_ROWS[j];
        secondaryRowRefs[secondaryDef.id] = createRowForDefinition(secondaryDef);
    }
    optionsPanel.orientation = 'column';
    optionsPanel.alignChildren = 'left';
    optionsPanel.margins = 10;
    optionsPanel.spacing = 6;
    optionsPanel.alignment = ['left', 'top'];
    var paletteCheckbox = optionsPanel.add('checkbox', undefined, 'Show Measurement Summary');
    paletteCheckbox.value = defaults.showMeasurementPalette !== false;
    try {
        paletteCheckbox.helpTip = 'Opens a palette listing measurement, ease, and final values instead of creating a reference artboard.';
    } catch (ePaletteTip) {}
    var shoulderDiffLabelNumber = SECONDARY_START_NUMBER + SECONDARY_ROWS.length;
    var shoulderDifferenceLabel = shoulderDiffLabelNumber + '. Shoulder Difference (deg)';
    var shoulderDifferenceField = addFieldRow(optionsPanel, shoulderDifferenceLabel, defaults.ShoulderDifference, {
        labelWidth: optionsLabelWidth,
        helpTip: 'page 159, number 24'
    });
    registerMeasurementLabel('ShoulderDifference', shoulderDifferenceLabel);
    try {
        shoulderDifferenceField.parent.children[0].helpTip = 'page 159, number 24';
    } catch (eShoulderTip) {}

    function parseField(field, fallback) {
        var v = parseFloat(field.text);
        return isNaN(v) ? fallback : v;
    }

    function formatValue(val) {
        return (Math.round(val * 100) / 100).toFixed(2);
    }

    function trimString(str) {
        if (str === undefined || str === null) return '';
        return String(str).replace(/^\s+|\s+$/g, '');
    }

    function normalizeNote(note) {
        if (note === undefined || note === null) return '';
        var clean = String(note);
        clean = clean.replace(/\u2012|\u2013|\u2014|\u2015|\u2212/g, '-');
        clean = clean.replace(/\s*-\s*/g, ' - ');
        clean = clean.replace(/\s+/g, ' ');
        clean = clean.replace(/\s+cm/gi, ' cm');
        return clean.replace(/^\s+|\s+$/g, '');
    }

    function renumberLabel(text, number) {
        if (!text || typeof text !== 'string') return text;
        var remainder = text.replace(/^\d+\.\s*/, '');
        return number + '. ' + remainder;
    }

    function calculateRecommendedOptimalBalance(brc) {
        if (isNaN(brc)) return null;
        if (brc < 80) return 3.5;
        if (brc <= 89) return 3.5;
        if (brc <= 99) return 4;
        if (brc <= 109) return ((brc - 100) / 10) + 4.5;
        if (brc <= 119) return ((brc - 100) / 10) + 5;
        if (brc <= 129) return ((brc - 100) / 10) + 5.5;
        if (brc <= 150) return ((brc - 100) / 10) + 6;
        return ((brc - 100) / 10) + 6;
    }

    function updateConstruction() {
        for (var key in rowRefs) {
            if (!rowRefs.hasOwnProperty(key)) continue;
            var ref = rowRefs[key];
            var meas = parseField(ref.measField, 0);
            var ease = (ref.easeField) ? parseField(ref.easeField, 0) : 0;
            ref.finalField.text = formatValue(meas + ease);
        }
        var brMeasure = parseField(rowRefs['BrC'].measField, defaults.BrC);
        var brFinal = parseField(rowRefs['BrC'].finalField, defaults.BrC);
        var waFinal = parseField(rowRefs['WaC'].finalField, defaults.WaC);
        var hiFinal = parseField(rowRefs['HiC'].finalField, defaults.HiC);
        var shA = parseField(rowRefs['ShA'].measField, defaults.ShA);
        var shoulderDifference = parseField(shoulderDifferenceField, defaults.ShoulderDifference);
        var frontShoulderAngle = shA + shoulderDifference;
        var backShoulderAngle = shA - shoulderDifference;
        var brw = brFinal / 2;
        var waw = waFinal / 2;
        var hiw = hiFinal / 2;
        var brgPlus = parseField(rowRefs['BrG'].finalField, defaults.BrG);
        var fShS = parseField(rowRefs['ShG'].finalField, defaults.ShG);
        var bShS = fShS + defaults.BackShoulderEase;
        var flMeasure = parseField(rowRefs['FL'].measField || flMeasureField, defaults.FL);
        var flBalance = parseField(flBalanceField, 0);
        var flFinal = flMeasure + flBalance;
        var BLMeas = parseField(rowRefs['BL'].measField, defaults.BL);
        var BLBal = rowRefs['BL'].easeField ? parseField(rowRefs['BL'].easeField, defaults.BLBal) : 0;
        var BLFinal = parseField(rowRefs['BL'].finalField, BLMeas + BLBal);
        var individualBalance = flMeasure - BLMeas;
        var recommendedOptimal = calculateRecommendedOptimalBalance(brMeasure);
        if (recommendedOptimal !== null) {
            var recommendedOptimalText = formatValue(recommendedOptimal);
            var currentOptimalText = trimString(optimalBalanceField.text);
            if (currentOptimalText === '' || currentOptimalText === lastRecommendedOptimalBalanceText || lastRecommendedOptimalBalanceText === null) {
                optimalBalanceField.text = recommendedOptimalText;
                currentOptimalText = recommendedOptimalText;
            }
            lastRecommendedOptimalBalanceText = recommendedOptimalText;
        } else {
            lastRecommendedOptimalBalanceText = null;
        }
        var optimalBalance = parseField(optimalBalanceField, defaults.OptimalBalance);
        var finalBalance = flFinal - BLFinal;
        var balanceDelta = Math.abs(finalBalance - optimalBalance);
        var balanceWithinTolerance = balanceDelta <= 1;
        if (brwField) brwField.text = formatValue(brw);
        if (wawField) wawField.text = formatValue(waw);
        if (hiwField) hiwField.text = formatValue(hiw);
        if (brgPlusField) brgPlusField.text = formatValue(brgPlus);
        if (bShSField) bShSField.text = formatValue(bShS);
        flFinalField.text = formatValue(flFinal);
        individualBalanceField.text = formatValue(individualBalance);
        finalBalanceField.text = formatValue(finalBalance);
        latestDerived = {
            brw: brw,
            waw: waw,
            hiw: hiw,
            brgPlus: brgPlus,
            fShS: fShS,
            bShS: bShS,
            flMeasure: flMeasure,
            flBalance: flBalance,
            flFinal: flFinal,
            blMeasure: BLMeas,
            blBalance: BLBal,
            blFinal: BLFinal,
            individualBalance: individualBalance,
            optimalBalance: optimalBalance,
            finalBalance: finalBalance,
            balanceWithinTolerance: balanceWithinTolerance,
            balanceDelta: balanceDelta,
            shA: shA,
            shoulderDifference: shoulderDifference,
            frontShoulderAngle: frontShoulderAngle,
            backShoulderAngle: backShoulderAngle
        };
    }

    function updateEaseFromFit() {
        var idx = fitDropdown.selection ? fitDropdown.selection.index : defaults.FitIndex;
        if (idx < 0 || idx >= FIT_PROFILES.length) idx = defaults.FitIndex;
        var profile = FIT_PROFILES[idx];
        for (var key in rowRefs) {
            if (!rowRefs.hasOwnProperty(key)) continue;
            if (!profile.ease.hasOwnProperty(key)) continue;
            if (!rowRefs[key].easeField) continue;
            var easeVal = profile.ease[key];
            if (easeVal === undefined) easeVal = 0;
            rowRefs[key].easeField.text = formatValue(easeVal);
            var noteStr = profile.notes[key] || '';
            var st = rowRefs[key].easeField.parent.children[0];
            if (st && st.text !== undefined) {
                var labelText = String(st.text);
                var colonIdx = labelText.indexOf(':');
                var beforeColon = colonIdx >= 0 ? labelText.substring(0, colonIdx) : labelText;
                var parenIdx = beforeColon.indexOf('(');
                if (parenIdx >= 0) beforeColon = beforeColon.substring(0, parenIdx);
                var cleanBase = trimString(beforeColon);
                st.text = cleanBase + (noteStr ? ' (' + noteStr + '):' : ':');
            }
        }
        updateConstruction();
    }
    for (var keyInit in rowRefs) {
        if (!rowRefs.hasOwnProperty(keyInit)) continue;
        rowRefs[keyInit].measField.onChanging = updateConstruction;
        if (rowRefs[keyInit].easeField) rowRefs[keyInit].easeField.onChanging = updateConstruction;
    }
    for (var keyInitSecondary in secondaryRowRefs) {
        if (!secondaryRowRefs.hasOwnProperty(keyInitSecondary)) continue;
        var secondaryRef = secondaryRowRefs[keyInitSecondary];
        secondaryRef.measField.onChanging = updateConstruction;
        if (secondaryRef.easeField) secondaryRef.easeField.onChanging = updateConstruction;
    }
    shoulderDifferenceField.onChanging = updateConstruction;
    fitDropdown.onChange = updateEaseFromFit;
    var buttonRow = dlg.add('group');
    buttonRow.alignment = 'right';
    buttonRow.add('button', undefined, 'OK', {
        name: 'ok'
    });
    buttonRow.add('button', undefined, 'Cancel', {
        name: 'cancel'
    });
    updateEaseFromFit();
    if (dlg.show() !== 1) return;
    selectedProfile = FIT_PROFILES[fitDropdown.selection ? fitDropdown.selection.index : defaults.FitIndex];
    shouldShowMeasurementPalette = paletteCheckbox.value ? true : false;
    var measurementResults = {};
    var easeResults = {};
    var finalResults = {};
    for (var keyRes in rowRefs) {
        if (!rowRefs.hasOwnProperty(keyRes)) continue;
        measurementResults[keyRes] = parseField(rowRefs[keyRes].measField, 0);
        var easeField = rowRefs[keyRes].easeField;
        easeResults[keyRes] = easeField ? parseField(easeField, 0) : 0;
        finalResults[keyRes] = parseField(rowRefs[keyRes].finalField, 0);
    }
    for (var keyResSecondary in secondaryRowRefs) {
        if (!secondaryRowRefs.hasOwnProperty(keyResSecondary)) continue;
        var secRefRes = secondaryRowRefs[keyResSecondary];
        measurementResults[keyResSecondary] = parseField(secRefRes.measField, 0);
        var secEaseField = secRefRes.easeField;
        easeResults[keyResSecondary] = secEaseField ? parseField(secEaseField, 0) : 0;
        finalResults[keyResSecondary] = parseField(secRefRes.finalField, 0);
    }
    var shoulderDifferenceValue = parseField(shoulderDifferenceField, defaults.ShoulderDifference);
    measurementResults.ShoulderDifference = shoulderDifferenceValue;
    easeResults.ShoulderDifference = 0;
    finalResults.ShoulderDifference = shoulderDifferenceValue;
    var NeG = parseField(secondaryRowRefs['NeG'].measField, defaults.NeG);
    var MoL = parseField(secondaryRowRefs['MoL'].measField, defaults.MoL);
    var BL = parseField(rowRefs['BL'].measField, defaults.BL);
    var BLBal = parseField(rowRefs['BL'].easeField, defaults.BLBal);
    var BLFinal = parseField(rowRefs['BL'].finalField, BL + BLBal);
    var HiD = parseField(secondaryRowRefs['HiD'].measField, defaults.HiD);
    var AhDPlus = finalResults['AhD'];
    var ShA = parseField(secondaryRowRefs['ShA'].measField, defaults.ShA);
    var fShSValue = parseField(rowRefs['ShG'].finalField, defaults.ShG);
    var shoulderDifference = shoulderDifferenceValue;
    var frontShoulderAngle = ShA + shoulderDifference;
    var backShoulderAngle = ShA - shoulderDifference;
    var bShS = fShSValue + defaults.BackShoulderEase;

    function makeRGB(r, g, b) {
        var c = new RGBColor();
        c.red = r;
        c.green = g;
        c.blue = b;
        return c;
    }

    function norm(name) {
        return (name && name.toString) ? name.toString().replace(/^\s+|\s+$/g, '').toLowerCase() : '';
    }

    function removeLayerIfNamed(target) {
        var goal = norm(target);
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var layer = doc.layers[i];
            if (norm(layer.name) === goal) {
                try {
                    layer.locked = false;
                } catch (e1) {}
                try {
                    layer.visible = true;
                } catch (e2) {}
                try {
                    layer.remove();
                } catch (e3) {}
            }
        }
    }
    removeLayerIfNamed('Layer 1');
    removeLayerIfNamed('layer');
    removeLayerIfNamed('<layer>');
    removeLayerIfNamed('&lt;layer&gt;');

    function ensureArtboardSize(widthCm, heightCm) {
        var idx = doc.artboards.getActiveArtboardIndex();
        var ab = doc.artboards[idx];
        var rect = ab.artboardRect;
        var desiredW = cm(widthCm);
        var desiredH = cm(heightCm);
        var currentW = rect[2] - rect[0];
        var currentH = rect[1] - rect[3];
        if (Math.abs(currentW - desiredW) > 0.5 || Math.abs(currentH - desiredH) > 0.5) {
            ab.artboardRect = [rect[0], rect[1], rect[0] + desiredW, rect[1] - desiredH];
        }
        return ab.artboardRect;
    }
    var artboardRect = ensureArtboardSize(100, 100);
    var originX = artboardRect[0] + cm(30);
    var originY = artboardRect[1] - cm(30);

    function normalizePoint(pt) {
        var x = 0,
            y = 0;
        if (pt != null) {
            if (typeof pt.x === 'number') x = pt.x;
            else if (typeof pt.length === 'number' && pt.length > 0 && typeof pt[0] === 'number') x = pt[0];
            if (typeof pt.y === 'number') y = pt.y;
            else if (typeof pt.length === 'number' && pt.length > 1 && typeof pt[1] === 'number') y = pt[1];
        }
        if (isNaN(x)) x = 0;
        if (isNaN(y)) y = 0;
        return [x, y];
    }

    function toArt(pt) {
        var coords = normalizePoint(pt);
        return [originX + cm(coords[0]), originY - cm(coords[1])];
    }

    function findLayerByName(name) {
        var goal = norm(name);
        for (var i = 0; i < doc.layers.length; i++) {
            if (norm(doc.layers[i].name) === goal) return doc.layers[i];
        }
        return null;
    }

    function ensureLayer(name) {
        var layer = findLayerByName(name);
        if (layer) return layer;
        layer = doc.layers.add();
        layer.name = name;
        return layer;
    }

    function ensureSubLayer(parent, name) {
        if (!parent || !parent.layers) return ensureLayer(name);
        var goal = norm(name);
        for (var i = 0; i < parent.layers.length; i++) {
            if (norm(parent.layers[i].name) === goal) return parent.layers[i];
        }
        var sub = parent.layers.add();
        sub.name = name;
        return sub;
    }

    function clearLayerContents(layer) {
        if (!layer) return;
        try {
            var items = layer.pageItems;
            for (var i = items.length - 1; i >= 0; i--) {
                try {
                    items[i].remove();
                } catch (eRemoveItem) {}
            }
        } catch (eLayerItems) {}
        try {
            if (layer.layers) {
                for (var s = layer.layers.length - 1; s >= 0; s--) {
                    try {
                        layer.layers[s].remove();
                    } catch (eRemoveSubLayer) {}
                }
            }
        } catch (eLayerSub) {}
    }

    function duplicateItemsToLayer(names, destinationLayer) {
        if (!destinationLayer || !names || !names.length) return;
        var lookup = {};
        for (var i = 0; i < names.length; i++) {
            lookup[names[i]] = true;
        }
        try {
            var items = doc.pageItems;
            for (var j = 0; j < items.length; j++) {
                var item = items[j];
                if (!item) continue;
                var itemName = '';
                try {
                    itemName = item.name;
                } catch (eNameRead) {}
                if (!itemName) continue;
                if (lookup[itemName]) {
                    try {
                        item.duplicate(destinationLayer, ElementPlacement.PLACEATEND);
                    } catch (eDuplicateItem) {}
                }
            }
        } catch (eDocItems) {}
    }

    function findLayerPathItemsByName(layer, name) {
        var matches = [];
        if (!layer || !name) return matches;
        try {
            var paths = layer.pathItems;
            for (var i = 0; i < paths.length; i++) {
                var candidate = paths[i];
                if (!candidate) continue;
                var candidateName = '';
                try {
                    candidateName = candidate.name || '';
                } catch (eNameRead) {}
                if (candidateName === name) matches.push(candidate);
            }
        } catch (eLayerPaths) {}
        return matches;
    }

    function ensureCentreBackOnLayer(layer, backPoints) {
        if (!layer || !backPoints || !backPoints.length) return;
        var cbPoints = [];
        for (var i = 0; i < backPoints.length; i++) {
            if (backPoints[i]) cbPoints.push(backPoints[i]);
        }
        if (!cbPoints.length) return;
        var cbPaths = findLayerPathItemsByName(layer, 'Centre Back (CB)');
        var cbPath = cbPaths.length ? cbPaths[0] : layer.pathItems.add();
        cbPath.closed = false;
        try {
            cbPath.name = 'Centre Back (CB)';
        } catch (eCBName) {}
        var artPoints = [];
        for (var j = 0; j < cbPoints.length; j++) {
            artPoints.push(toArt(cbPoints[j]));
        }
        cbPath.setEntirePath(artPoints);
        try {
            cbPath.stroked = true;
            cbPath.strokeWidth = 1;
            cbPath.strokeColor = FRAME_COLOR;
        } catch (eCBStroke) {}
        try {
            cbPath.filled = false;
        } catch (eCBFill) {}
        for (var c = 1; c < cbPaths.length; c++) {
            try {
                cbPaths[c].remove();
            } catch (eRemoveCB) {}
        }
    }

    function fromArt(pt) {
        var coords = normalizePoint(pt);
        return {
            x: (coords[0] - originX) / cm(1),
            y: (originY - coords[1]) / cm(1)
        };
    }

    function resetGroup(parent, name) {
        if (!parent) return null;
        var groups = parent.groupItems;
        for (var i = groups.length - 1; i >= 0; i--) {
            if (groups[i].name === name) {
                try {
                    groups[i].remove();
                } catch (eRem) {}
            }
        }
        var group = parent.groupItems.add();
        group.name = name;
        return group;
    }

    function removeGroupsByName(parent, names) {
        if (!parent || !names || !names.length) return;
        try {
            var groups = parent.groupItems;
            for (var i = groups.length - 1; i >= 0; i--) {
                var grp = groups[i];
                if (!grp) continue;
                for (var n = 0; n < names.length; n++) {
                    if (grp.name === names[n]) {
                        try {
                            grp.remove();
                        } catch (eRemGroup) {}
                        break;
                    }
                }
            }
        } catch (eParentGroups) {}
    }

    function clearGroupDeep(group) {
        if (!group) return;
        try {
            for (var i = group.pageItems.length - 1; i >= 0; i--) {
                try {
                    group.pageItems[i].remove();
                } catch (ePI) {}
            }
        } catch (ePage) {}
        try {
            if (group.compoundPathItems) group.compoundPathItems.removeAll();
        } catch (eCP) {}
        try {
            if (group.pathItems) group.pathItems.removeAll();
        } catch (ePath) {}
        try {
            if (group.textFrames) group.textFrames.removeAll();
        } catch (eText) {}
        try {
            if (group.groupItems) group.groupItems.removeAll();
        } catch (eGroup) {}
        try {
            if (group.layers) {
                var subLayers = group.layers;
                for (var j = subLayers.length - 1; j >= 0; j--) {
                    try {
                        subLayers[j].remove();
                    } catch (eSubLayer) {}
                }
            }
        } catch (eLayer) {}
    }

    function removeLayerByName(docRef, targetName) {
        if (!docRef || !docRef.layers || !targetName) return;
        var targetNorm = norm(targetName);
        for (var i = docRef.layers.length - 1; i >= 0; i--) {
            var layer = docRef.layers[i];
            if (norm(layer.name) !== targetNorm) continue;
            try {
                layer.locked = false;
            } catch (eUnlockLayer) {}
            try {
                layer.visible = true;
            } catch (eRevealLayer) {}
            try {
                layer.remove();
            } catch (eRemoveLayer) {}
        }
    }

    function removeArtboardByName(docRef, targetName) {
        if (!docRef || !docRef.artboards || !targetName) return;
        var targetNorm = norm(targetName);
        for (var i = docRef.artboards.length - 1; i >= 0; i--) {
            var artboard = docRef.artboards[i];
            if (norm(artboard.name) !== targetNorm) continue;
            try {
                if (docRef.artboards.length > 1) {
                    var newIndex = (i === 0) ? 1 : 0;
                    docRef.artboards.setActiveArtboardIndex(newIndex);
                }
            } catch (eSelectArt) {}
            try {
                artboard.remove();
            } catch (eRemoveArt) {}
        }
    }

    function removePageItemsByName(container, names) {
        if (!container || !names || !names.length) return;
        try {
            var items = container.pageItems;
            for (var i = items.length - 1; i >= 0; i--) {
                var item = items[i];
                if (!item) continue;
                for (var n = 0; n < names.length; n++) {
                    if (item.name === names[n]) {
                        try {
                            item.remove();
                        } catch (eRemItem) {}
                        break;
                    }
                }
            }
        } catch (eItems) {}
    }
    var basicFrameLayer = ensureLayer('Basic Frame');
    try {
        basicFrameLayer.locked = false;
    } catch (eBF1) {}
    try {
        basicFrameLayer.visible = true;
    } catch (eBF2) {}
    try {
        basicFrameLayer.color = LAYER_COLOR;
    } catch (eColor) {}
    var labelsLayerContainer = ensureLayer('Labels, Markers & Numbers');
    try {
        labelsLayerContainer.locked = false;
    } catch (eLbl1) {}
    try {
        labelsLayerContainer.visible = true;
    } catch (eLbl2) {}
    var casualLinesLayer = ensureLayer('Casual Bodice Lines');
    try {
        casualLinesLayer.locked = false;
    } catch (eCasLock) {}
    try {
        casualLinesLayer.visible = true;
    } catch (eCasVis) {}
    for (var iLayer = doc.layers.length - 1; iLayer >= 0; iLayer--) {
        var lyr = doc.layers[iLayer];
        var nameNorm = norm(lyr.name);
        if (nameNorm !== norm('Basic Frame') && nameNorm !== norm('Labels, Markers & Numbers') && nameNorm !== norm('Casual Bodice Lines')) {
            try {
                lyr.locked = false;
            } catch (eLk) {}
            try {
                lyr.visible = true;
            } catch (eVs) {}
            try {
                lyr.remove();
            } catch (eRm) {}
        }
    }
    basicFrameLayer = ensureLayer('Basic Frame');
    labelsLayerContainer = ensureLayer('Labels, Markers & Numbers');
    casualLinesLayer = ensureLayer('Casual Bodice Lines');
    clearGroupDeep(basicFrameLayer);
    var linesLayer = basicFrameLayer;
    clearGroupDeep(casualLinesLayer);
    var markersLayer = ensureSubLayer(labelsLayerContainer, 'Markers');
    var numbersLayer = ensureSubLayer(labelsLayerContainer, 'Numbers');
    var labelsLayer = ensureSubLayer(labelsLayerContainer, 'Labels');
    clearGroupDeep(markersLayer);
    clearGroupDeep(numbersLayer);
    clearGroupDeep(labelsLayer);
    var markersGroup = markersLayer;
    var numbersGroup = numbersLayer;
    var labelsGroup = labelsLayer;
    try {
        basicFrameLayer.zOrder(ZOrderMethod.SENDTOBACK);
    } catch (eOrderBasic) {}
    try {
        casualLinesLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
    } catch (eOrderCasual) {}
    try {
        labelsLayerContainer.zOrder(ZOrderMethod.BRINGTOFRONT);
    } catch (eOrderLabels) {}
    removeLayerByName(doc, 'Measurement Reference');
    removeArtboardByName(doc, 'Measurement Reference');
    var FRAME_COLOR = makeRGB(0, 0, 0);
    var NUMBER_FILL_COLOR = makeRGB(255, 255, 255);
    var LABEL_FONT_SIZE_PT = 13;
    var NUMBER_FONT_SIZE_PT = 9;
    var MARKER_RADIUS_CM = 0.25;
    var FRAME_LINE_LENGTH_CM = 40;
    var LABEL_RIGHT_OFFSET_CM = 5;
    var LABEL_VERTICAL_OFFSET_CM = 0.5;
    var LINE_LABEL_HORIZONTAL_SHIFT_CM = ptToCm(200);
    var CONNECTOR_BLUE_COLOR = makeRGB(0, 0, 0);
    var LAYER_COLOR = makeRGB(0, 0, 0);
    var HIP_LEFT_OFFSET_CM = 2;
    var DASH_PATTERN = [25, 12];
    var LABEL_SMALL_OFFSET_CM = 0.5;
    var FRONT_ARM_LABEL_NORMAL_OFFSET_CM = -0.5;
    var MIN_HANDLE_LENGTH_CM = 0.5;
    var labelAlignmentRefs = {};

    function drawFrameLine(pA, pB, name, color, targetGroup) {
        var target = targetGroup || linesLayer;
        if (!target) return null;
        var path = null;
        if (name) {
            try {
                var existing = target.pathItems;
                for (var idx = existing.length - 1; idx >= 0; idx--) {
                    var candidate = existing[idx];
                    if (!candidate) continue;
                    var candidateName = '';
                    try {
                        candidateName = candidate.name;
                    } catch (eNameRead) {}
                    if (candidateName === name) {
                        if (!path) {
                            path = candidate;
                        } else {
                            try {
                                candidate.remove();
                            } catch (eDupRem) {}
                        }
                    }
                }
            } catch (eExisting) {}
        }
        if (!path) {
            path = target.pathItems.add();
            if (name) {
                try {
                    path.name = name;
                } catch (eNameAssign) {}
            }
        }
        path.setEntirePath([toArt(pA), toArt(pB)]);
        path.stroked = true;
        path.strokeWidth = 1;
        path.strokeColor = color || FRAME_COLOR;
        path.filled = false;
        return path;
    }

    function drawCurveBetween(pA, pB, options) {
        options = options || {};
        var color = options.color || FRAME_COLOR;
        var name = options.name;
        var bulgeCm = options.bulgeCm;
        if (bulgeCm === undefined || bulgeCm === null || isNaN(bulgeCm)) bulgeCm = 1;
        var startArt = toArt(pA);
        var endArt = toArt(pB);
        var container = linesLayer;
        if (!container) return null;
        var path = container.pathItems.add();
        path.setEntirePath([startArt, endArt]);
        path.stroked = true;
        path.strokeWidth = 1;
        path.strokeColor = color;
        path.filled = false;
        try {
            path.closed = false;
        } catch (eClosed) {}
        var startPoint = path.pathPoints[0];
        var endPoint = path.pathPoints[1];
        var dx = endArt[0] - startArt[0];
        var dy = endArt[1] - startArt[1];
        var len = Math.sqrt((dx * dx) + (dy * dy));
        var nx = 0;
        var ny = 0;
        if (len > 0) {
            nx = -dy / len;
            ny = dx / len;
        }
        var bulge = cm(bulgeCm);
        var flattenStart = options.flattenStart === true;
        var flattenEnd = options.flattenEnd === true;
        startPoint.pointType = flattenStart ? PointType.CORNER : PointType.SMOOTH;
        endPoint.pointType = flattenEnd ? PointType.CORNER : PointType.SMOOTH;
        startPoint.leftDirection = startPoint.anchor;
        endPoint.rightDirection = endPoint.anchor;
        if (flattenStart) {
            startPoint.rightDirection = startPoint.anchor;
        } else {
            startPoint.rightDirection = [startArt[0] + (dx / 3) + (nx * bulge), startArt[1] + (dy / 3) + (ny * bulge)];
        }
        if (flattenEnd) {
            endPoint.leftDirection = endPoint.anchor;
        } else {
            endPoint.leftDirection = [endArt[0] - (dx / 3) + (nx * bulge), endArt[1] - (dy / 3) + (ny * bulge)];
        }
        if (name) {
            try {
                path.name = name;
            } catch (eCurveName) {}
        }
        return path;
    }

    function setHorizontalHandle(pathItem, pointIndex, side, lengthCm, directionSign) {
        try {
            if (!pathItem) return;
            var pts = pathItem.pathPoints;
            if (!pts || pointIndex < 0 || pointIndex >= pts.length) return;
            var pt = pts[pointIndex];
            var effectiveLength = (!isNaN(lengthCm) && lengthCm > 0) ? lengthCm : MIN_HANDLE_LENGTH_CM;
            var delta = cm(effectiveLength) * (directionSign >= 0 ? 1 : -1);
            pt.pointType = PointType.SMOOTH;
            if (side === 'left') {
                pt.leftDirection = [pt.anchor[0] + delta, pt.anchor[1]];
                pt.rightDirection = pt.anchor;
            } else {
                pt.rightDirection = [pt.anchor[0] + delta, pt.anchor[1]];
                pt.leftDirection = pt.anchor;
            }
        } catch (eHandle) {}
    }

    function isVerticalLine(pA, pB) {
        return Math.abs(pA.x - pB.x) < 0.0001 && Math.abs(pA.y - pB.y) > 0.0001;
    }

    function isHorizontalLine(pA, pB) {
        return Math.abs(pA.y - pB.y) < 0.0001 && Math.abs(pA.x - pB.x) > 0.0001;
    }

    function midpoint(pA, pB) {
        return {
            x: (pA.x + pB.x) / 2,
            y: (pA.y + pB.y) / 2
        };
    }

    function offsetAlongNormal(pA, pB, offsetCm) {
        var mid = midpoint(pA, pB);
        if (!offsetCm) return mid;
        var dx = pB.x - pA.x;
        var dy = pB.y - pA.y;
        var len = Math.sqrt((dx * dx) + (dy * dy));
        if (len < 0.0001) return mid;
        var nx = dy / len;
        var ny = -dx / len;
        return {
            x: mid.x + nx * offsetCm,
            y: mid.y + ny * offsetCm
        };
    }

    function rotateVerticalLabel(tf, anchor) {
        if (!tf) return;
        var rotationDegrees = -90;
        var extraRotation = 180;
        var rotated = false;
        try {
            tf.rotate(rotationDegrees);
            rotated = true;
        } catch (eRotFrame) {}
        if (!rotated) {
            try {
                tf.textRange.characterAttributes.rotation = rotationDegrees;
                rotated = true;
            } catch (eRot) {}
        }
        if (rotated) {
            var extraApplied = false;
            try {
                tf.rotate(extraRotation);
                extraApplied = true;
            } catch (eExtraFrame) {}
            if (!extraApplied) {
                try {
                    tf.textRange.characterAttributes.rotation = rotationDegrees + extraRotation;
                    extraApplied = true;
                } catch (eExtraChar) {}
            }
            try {
                centerTextFrame(tf, anchor);
            } catch (eCenter) {}
        }
    }

    function computeLabelPoint(pA, pB, text) {
        var vertical = isVerticalLine(pA, pB);
        var horizontal = isHorizontalLine(pA, pB);
        var basePoint;
        if (text === 'Front Arm Line') {
            basePoint = offsetAlongNormal(pA, pB, FRONT_ARM_LABEL_NORMAL_OFFSET_CM);
        } else if (vertical) {
            basePoint = {
                x: pA.x - LABEL_RIGHT_OFFSET_CM,
                y: (pA.y + pB.y) / 2
            };
        } else if (horizontal) {
            basePoint = lineLabelAnchor(pA, pB);
        } else {
            basePoint = midpoint(pA, pB);
        }
        var smallOffset = LABEL_SMALL_OFFSET_CM;
        switch (text) {
            case 'Back Arm Line':
                labelAlignmentRefs.backArmLineY = basePoint.y;
                basePoint = {
                    x: pA.x + smallOffset,
                    y: basePoint.y
                };
                break;
            case 'Front Arm Line':
                if (labelAlignmentRefs.backArmLineY !== undefined) basePoint.y = labelAlignmentRefs.backArmLineY;
                break;
            case 'Back Side Line 1':
            case 'Back Side Line 2':
                if (labelAlignmentRefs.backArmLineY !== undefined) basePoint.y = labelAlignmentRefs.backArmLineY;
                labelAlignmentRefs.backSideLineY = basePoint.y;
                basePoint = {
                    x: pA.x - smallOffset,
                    y: basePoint.y
                };
                break;
            case 'Front Side Line':
                if (labelAlignmentRefs.backArmLineY !== undefined) basePoint.y = labelAlignmentRefs.backArmLineY;
                basePoint = {
                    x: pA.x + smallOffset,
                    y: basePoint.y
                };
                break;
            case 'Front Dart Line':
                if (labelAlignmentRefs.backArmLineY !== undefined) basePoint.y = labelAlignmentRefs.backArmLineY;
                basePoint = {
                    x: pA.x - smallOffset,
                    y: basePoint.y
                };
                break;
            case 'Centre Front (CF)':
                basePoint = {
                    x: pA.x + smallOffset,
                    y: (pA.y + pB.y) / 2
                };
                break;
            case 'Centre Back (CB)':
                basePoint = {
                    x: pA.x - smallOffset,
                    y: (pA.y + pB.y) / 2
                };
                break;
        }
        return basePoint;
    }

    function formatReferenceValue(value) {
        if (value === undefined || value === null) return '-';
        var num = parseFloat(value);
        if (!isNaN(num)) return formatValue(num);
        if (value === true) return 'Yes';
        if (value === false) return 'No';
        return String(value);
    }

    function collectMeasurementRows() {
        var rows = [];
        var seen = {};
        for (var key in measurementResults) {
            if (!measurementResults.hasOwnProperty(key)) continue;
            if (seen[key]) continue;
            seen[key] = true;
            rows.push({
                id: resolveMeasurementLabel(key),
                rawId: key,
                meas: formatReferenceValue(measurementResults[key]),
                ease: formatReferenceValue(easeResults[key]),
                finalValue: formatReferenceValue(finalResults[key])
            });
        }
        rows.sort(function(a, b) {
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
            if (measurementPalette && typeof measurementPalette.close === 'function') {
                measurementPalette.close();
            }
        } catch (eCloseLocal) {}
        try {
            var globalPalette = $.global.hofenbitzerMeasurementPalette;
            if (globalPalette && globalPalette.window && typeof globalPalette.window.close === 'function') {
                globalPalette.window.close();
            }
            delete $.global.hofenbitzerMeasurementPalette;
        } catch (eCloseGlobal) {}
        measurementPalette = null;
    }

    function buildPaletteStaticText(group, label, width, justification) {
        var st = group.add('statictext', undefined, label);
        if (width !== undefined) {
            st.preferredSize.width = width;
            st.minimumSize.width = width;
        }
        if (justification) {
            try {
                st.justify = justification;
            } catch (eJustSet) {}
        }
        return st;
    }

    function showMeasurementPaletteWindow(profileName) {
        closeMeasurementPalette();
        var palette = new Window('palette', 'Measurement Summary');
        palette.orientation = 'column';
        palette.alignChildren = ['fill', 'top'];
        palette.spacing = 8;
        palette.margins = 12;
        palette.preferredSize.width = 560;

        var metaGroup = palette.add('group');
        metaGroup.alignment = 'fill';
        metaGroup.add('statictext', undefined, 'Fit Profile: ' + (profileName || '-'));

        var headerGroup = palette.add('group');
        headerGroup.alignment = 'fill';
        headerGroup.spacing = 6;
        var columnWidths = [200, 120, 120, 120];
        buildPaletteStaticText(headerGroup, 'Measurement', columnWidths[0], 'left');
        buildPaletteStaticText(headerGroup, 'Measure', columnWidths[1], 'right');
        buildPaletteStaticText(headerGroup, 'Ease', columnWidths[2], 'right');
        buildPaletteStaticText(headerGroup, 'Final', columnWidths[3], 'right');

        var rows = collectMeasurementRows();
        if (!rows.length) {
            var emptyRow = palette.add('group');
            emptyRow.alignment = 'fill';
            emptyRow.add('statictext', undefined, 'No measurements available.');
        } else {
            for (var i = 0; i < rows.length; i++) {
                var row = rows[i];
                var rowGroup = palette.add('group');
                rowGroup.alignment = 'fill';
                rowGroup.spacing = 6;
                buildPaletteStaticText(rowGroup, row.id, columnWidths[0], 'left');
                buildPaletteStaticText(rowGroup, row.meas, columnWidths[1], 'right');
                buildPaletteStaticText(rowGroup, row.ease, columnWidths[2], 'right');
                buildPaletteStaticText(rowGroup, row.finalValue, columnWidths[3], 'right');
            }
        }

        palette.onClose = function() {
            measurementPalette = null;
            try {
                delete $.global.hofenbitzerMeasurementPalette;
            } catch (eDeleteGlobal) {}
        };

        measurementPalette = palette;
        $.global.hofenbitzerMeasurementPalette = {
            window: palette
        };
        try {
            palette.show();
        } catch (eShow) {}
    }

    function getItemBounds(item) {
        if (!item) return null;
        var bounds = null;
        try {
            bounds = item.visibleBounds;
        } catch (eVisibleBounds) {}
        if (!bounds) {
            try {
                bounds = item.geometricBounds;
            } catch (eGeoBounds) {}
        }
        return bounds;
    }

    function unionBounds(current, candidate) {
        if (!candidate) return current;
        if (!current) return [candidate[0], candidate[1], candidate[2], candidate[3]];
        if (candidate[0] < current[0]) current[0] = candidate[0];
        if (candidate[1] > current[1]) current[1] = candidate[1];
        if (candidate[2] > current[2]) current[2] = candidate[2];
        if (candidate[3] < current[3]) current[3] = candidate[3];
        return current;
    }

    function cropArtboardAroundItems(artboardIndex, items, marginCm) {
        if (artboardIndex < 0 || artboardIndex >= doc.artboards.length) return;
        if (!items || !items.length) return;
        var union = null;
        for (var i = 0; i < items.length; i++) {
            var bounds = getItemBounds(items[i]);
            if (!bounds) continue;
            union = unionBounds(union, bounds);
        }
        if (!union) return;
        var marginPts = cm(marginCm);
        var newRect = [union[0] - marginPts, union[1] + marginPts, union[2] + marginPts, union[3] - marginPts];
        try {
            doc.artboards[artboardIndex].artboardRect = newRect;
        } catch (eCrop) {}
    }

    function drawLineWithLabel(pA, pB, text, color, targetGroup) {
        var path = drawFrameLine(pA, pB, text, color, targetGroup);
        if (!text) return path;
        var labelPoint = computeLabelPoint(pA, pB, text);
        var tf = drawCenteredLabel(labelPoint, text);
        if (tf && isVerticalLine(pA, pB)) {
            rotateVerticalLabel(tf, toArt(labelPoint));
        }
        return path;
    }

    function centerTextFrame(tf, anchor) {
        if (!tf) return;
        try {
            var bounds = tf.visibleBounds;
            tf.translate(anchor[0] - (bounds[0] + bounds[2]) / 2, anchor[1] - (bounds[1] + bounds[3]) / 2);
        } catch (eBounds) {
            try {
                tf.left = anchor[0] - (tf.width / 2);
                tf.top = anchor[1] + (tf.height / 2);
            } catch (eFallback) {}
        }
    }

    function placeFrameMarker(pt, label) {
        var anchor = toArt(pt);
        var radius = cm(MARKER_RADIUS_CM);
        var circle = markersGroup.pathItems.ellipse(anchor[1] + radius, anchor[0] - radius, radius * 2, radius * 2);
        if (label != null) {
            try {
                circle.name = 'Marker ' + label;
            } catch (eMkName) {}
        }
        circle.stroked = true;
        circle.strokeWidth = 1;
        circle.strokeColor = FRAME_COLOR;
        circle.filled = true;
        circle.fillColor = FRAME_COLOR;
        if (label != null) {
            var tf = numbersGroup.textFrames.add();
            tf.contents = String(label);
            try {
                tf.name = 'Marker ' + label + ' Number';
            } catch (eNumName) {}
            tf.textRange.characterAttributes.size = NUMBER_FONT_SIZE_PT;
            tf.textRange.characterAttributes.fillColor = NUMBER_FILL_COLOR;
            try {
                tf.textRange.characterAttributes.textFont = app.textFonts.getByName('ArialMT');
            } catch (eFont) {}
            try {
                tf.textRange.paragraphAttributes.justification = Justification.CENTER;
            } catch (eJust) {}
            try {
                tf.textRange.characterAttributes.baselineShift = 0;
            } catch (eShift) {}
            centerTextFrame(tf, anchor);
            try {
                circle.zOrder(ZOrderMethod.SENDTOBACK);
            } catch (eZ) {}
        }
        return circle;
    }

    function drawCenteredLabel(midPoint, text) {
        var anchor = toArt(midPoint);
        var tf = labelsGroup.textFrames.add();
        tf.contents = text;
        tf.textRange.characterAttributes.size = LABEL_FONT_SIZE_PT;
        try {
            tf.textRange.characterAttributes.textFont = app.textFonts.getByName('ArialMT');
        } catch (eFont) {}
        tf.textRange.characterAttributes.fillColor = makeRGB(0, 0, 0);
        centerTextFrame(tf, anchor);
        return tf;
    }

    function lineLabelAnchor(startPt, endPt) {
        var midX = (startPt.x + endPt.x) / 2;
        var hemX = startPt.x - LABEL_RIGHT_OFFSET_CM;
        var baseX = (midX + hemX) / 2 + LINE_LABEL_HORIZONTAL_SHIFT_CM;
        return {
            x: baseX,
            y: startPt.y - LABEL_VERTICAL_OFFSET_CM
        };
    }

    function extendLineToY(pA, pB, targetY) {
        var dy = pB.y - pA.y;
        var dx = pB.x - pA.x;
        if (Math.abs(dy) < 0.0001) {
            return {
                x: pB.x,
                y: targetY
            };
        }
        var t = (targetY - pA.y) / dy;
        return {
            x: pA.x + dx * t,
            y: targetY
        };
    }

    function distanceBetween(pA, pB) {
        if (!pA || !pB) return null;
        var dx = pB.x - pA.x;
        var dy = pB.y - pA.y;
        return Math.sqrt((dx * dx) + (dy * dy));
    }

    function pointOnSegment(pt, segA, segB, tolerance) {
        if (!pt || !segA || !segB) return false;
        var tol = (tolerance !== undefined) ? tolerance : 0.01;
        var minX = Math.min(segA.x, segB.x) - tol;
        var maxX = Math.max(segA.x, segB.x) + tol;
        var minY = Math.min(segA.y, segB.y) - tol;
        var maxY = Math.max(segA.y, segB.y) + tol;
        if (pt.x < minX || pt.x > maxX || pt.y < minY || pt.y > maxY) return false;
        var cross = Math.abs((segB.x - segA.x) * (pt.y - segA.y) - (segB.y - segA.y) * (pt.x - segA.x));
        var segLen = distanceBetween(segA, segB);
        if (!segLen || segLen < 1e-6) return false;
        return (cross / segLen) <= tol;
    }

    function lineIntersection(p1, p2, p3, p4) {
        if (!p1 || !p2 || !p3 || !p4) return null;
        var x1 = p1.x,
            y1 = p1.y;
        var x2 = p2.x,
            y2 = p2.y;
        var x3 = p3.x,
            y3 = p3.y;
        var x4 = p4.x,
            y4 = p4.y;
        var denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4);
        if (Math.abs(denom) < 1e-6) return null;
        var det1 = (x1 * y2) - (y1 * x2);
        var det2 = (x3 * y4) - (y3 * x4);
        var x = (det1 * (x3 - x4) - (x1 - x2) * det2) / denom;
        var y = (det1 * (y3 - y4) - (y1 - y2) * det2) / denom;
        var candidate = {
            x: x,
            y: y
        };
        if (!pointOnSegment(candidate, p1, p2) || !pointOnSegment(candidate, p3, p4)) return null;
        return candidate;
    }

    function trimBackWaistLine(path, waistStart, waistEnd, sideStart, sideEnd, point7, waistLineY) {
        if (!path || !waistStart || !waistEnd || !point7 || waistLineY === undefined || waistLineY === null) return;
        var usableSideStart = sideStart || waistStart;
        var usableSideEnd = sideEnd || waistEnd;
        var trimmedStart = lineIntersection(waistStart, waistEnd, usableSideStart, usableSideEnd);
        if (!trimmedStart) trimmedStart = waistStart;
        var trimmedEnd = {
            x: point7.x,
            y: waistLineY
        };
        path.setEntirePath([toArt(trimmedStart), toArt(trimmedEnd)]);
    }

    function setPathBetween(path, startPt, endPt) {
        if (!path || !startPt || !endPt) return;
        path.setEntirePath([toArt(startPt), toArt(endPt)]);
    }

    function trimHorizontalPathRightOf(path, cutAnchor, toleranceCm) {
        if (!path || !cutAnchor) return;
        var cutPoint = fromArt(cutAnchor);
        var tol = (toleranceCm !== undefined) ? toleranceCm : 0.05;
        var pts = path.pathPoints;
        if (!pts || pts.length < 2) return;
        var start = fromArt(pts[0].anchor);
        var end = fromArt(pts[pts.length - 1].anchor);
        var horizontal = Math.abs(start.y - end.y) <= tol;
        if (!horizontal) return;
        var left = start.x <= end.x ? start : end;
        var right = start.x > end.x ? start : end;
        if (cutPoint.x <= left.x + tol) {
            try {
                path.remove();
            } catch (eRemovePath) {}
            return;
        }
        var newRight = right;
        if (cutPoint.x < right.x - tol) newRight = {
            x: cutPoint.x,
            y: right.y
        };
        var ordered = left === start;
        var artLeft = toArt(left);
        var artNewRight = toArt(newRight);
        if (ordered) path.setEntirePath([artLeft, artNewRight]);
        else path.setEntirePath([artNewRight, artLeft]);
    }

    var rightBoundaryCm = (artboardRect[2] - originX) / cm(1);
    var topBoundaryCm = (originY - artboardRect[1]) / cm(1);
    var point1 = {
        x: rightBoundaryCm - 10,
        y: topBoundaryCm + 10
    };
    placeFrameMarker([point1.x, point1.y], 1);
    var point1aOffset = NeG + 0.5;
    if (isNaN(point1aOffset)) point1aOffset = NeG;
    var point1a = {
        x: point1.x - point1aOffset,
        y: point1.y
    };
    var point1aExtension = {
        x: point1a.x - 10,
        y: point1a.y
    };
    var line1To1a = drawFrameLine([point1.x, point1.y], [point1a.x, point1a.y], '1 - 1a');
    try {
        line1To1a.strokeDashes = DASH_PATTERN;
    } catch (eDash1a) {}
    var line1aExtensionPath = drawFrameLine([point1a.x, point1a.y], [point1aExtension.x, point1aExtension.y], '1a - Extension');
    try {
        line1aExtensionPath.strokeDashes = DASH_PATTERN;
    } catch (eDash1aExt) {}
    placeFrameMarker([point1a.x, point1a.y], '1a');
    var dist12 = (NeG / 3) + 1;
    var point2 = {
        x: point1.x,
        y: point1.y + dist12
    };
    placeFrameMarker([point2.x, point2.y], 2);
    var curveBulgeCm = -Math.max(0.5, (NeG + 0.5) / 3);
    var backNeckCurve = drawCurveBetween(point1a, point2, {
        name: 'Back Neck Curve',
        bulgeCm: curveBulgeCm
    });
    if (backNeckCurve) {
        var backHandleLength = Math.max(MIN_HANDLE_LENGTH_CM, NeG / 2);
        setHorizontalHandle(backNeckCurve, 1, 'left', backHandleLength, -1);
    }
    var frontShoulderEnd = null;
    var backShoulderEnd = null;
    var point16 = null;
    var point3 = {
        x: point2.x,
        y: point2.y + MoL
    };
    placeFrameMarker([point3.x, point3.y], 3);
    var point4 = {
        x: point2.x,
        y: point2.y + AhDPlus
    };
    placeFrameMarker([point4.x, point4.y], 4);
    var point5 = {
        x: point2.x,
        y: point2.y + BLFinal
    };
    placeFrameMarker([point5.x, point5.y], 5);
    var point6 = {
        x: point5.x,
        y: point5.y + HiD
    };
    placeFrameMarker([point6.x, point6.y], 6);
    var hemLineStart = {
        x: point3.x,
        y: point3.y
    };
    var bustLineStart = {
        x: point4.x,
        y: point4.y
    };
    var waistLineStart = {
        x: point5.x,
        y: point5.y
    };
    var hipLineStart = {
        x: point6.x,
        y: point6.y
    };

    var point6a = {
        x: point6.x - HIP_LEFT_OFFSET_CM,
        y: point6.y
    };
    placeFrameMarker([point6a.x, point6a.y], '6a');

    var point9 = extendLineToY(point2, point6a, point4.y);
    placeFrameMarker([point9.x, point9.y], 9);

    var point7 = extendLineToY(point2, point6a, point5.y);
    placeFrameMarker([point7.x, point7.y], 7);

    var point8 = extendLineToY(point2, point6a, point3.y);

    var backArmBase = point9;
    var AGPlus = finalResults['AG'];
    if (isNaN(AGPlus)) AGPlus = parseField(rowRefs['AG'].finalField, 0);
    if (isNaN(AGPlus)) AGPlus = 0;
    var BGPlus = finalResults['BG'];
    if (isNaN(BGPlus)) BGPlus = parseField(rowRefs['BG'].finalField, 0);
    if (isNaN(BGPlus)) BGPlus = 0;
    var BrGPlus = finalResults['BrG'];
    if (isNaN(BrGPlus)) BrGPlus = parseField(rowRefs['BrG'].finalField, 0);
    if (isNaN(BrGPlus)) BrGPlus = 0;
    var point10 = {
        x: backArmBase.x - BGPlus,
        y: bustLineStart.y
    };
    var point11 = {
        x: point10.x - (AGPlus * 2 / 3),
        y: point10.y
    };
    var point12 = {
        x: point11.x - 15,
        y: point11.y
    };
    var point13 = {
        x: point12.x - (AGPlus / 3),
        y: point12.y
    };
    var point14 = {
        x: point13.x - BrGPlus,
        y: point13.y
    };
    placeFrameMarker([point10.x, point10.y], 10);
    placeFrameMarker([point11.x, point11.y], 11);
    placeFrameMarker([point12.x, point12.y], 12);
    placeFrameMarker([point13.x, point13.y], 13);
    var point13aOffset = AGPlus / 4;
    if (isNaN(point13aOffset)) point13aOffset = 0;
    var point13a = {
        x: point13.x,
        y: point13.y - point13aOffset
    };
    placeFrameMarker([point13a.x, point13a.y], '13a');
    placeFrameMarker([point14.x, point14.y], 14);
    if (bShS > 0) {
        var frontShoulderRadians = frontShoulderAngle * Math.PI / 180;
        frontShoulderEnd = {
            x: point1a.x - Math.cos(frontShoulderRadians) * bShS,
            y: point1a.y + Math.sin(frontShoulderRadians) * bShS
        };
        var backShoulderRadians = backShoulderAngle * Math.PI / 180;
        var dirX = -Math.cos(backShoulderRadians);
        var dirY = Math.sin(backShoulderRadians);
        var lineLength = bShS;
        var tIntersect = null;
        if (Math.abs(dirX) > 0.0001) {
            var candidateT = (point10.x - point1a.x) / dirX;
            if (candidateT >= 0) tIntersect = candidateT;
        } else if (Math.abs(point10.x - point1a.x) < 0.0001) {
            tIntersect = 0;
        }
        var effectiveLength = lineLength;
        if (tIntersect !== null && tIntersect > effectiveLength) effectiveLength = tIntersect;
        backShoulderEnd = {
            x: point1a.x + dirX * effectiveLength,
            y: point1a.y + dirY * effectiveLength
        };
        if (tIntersect !== null) {
            point16 = {
                x: point10.x,
                y: point1a.y + dirY * tIntersect
            };
        }
        drawFrameLine([point1a.x, point1a.y], [backShoulderEnd.x, backShoulderEnd.y], 'Back Shoulder Line');
    }
    if (point16) {
        placeFrameMarker([point16.x, point16.y], '16');
    }
    var point17 = null;
    var point17a = null;
    var point18 = null;
    if (point16) {
        var midpoint1610 = {
            x: (point16.x + point10.x) / 2,
            y: (point16.y + point10.y) / 2
        };
        point17 = {
            x: midpoint1610.x - 1,
            y: midpoint1610.y
        };
        placeFrameMarker([point17.x, point17.y], '17');
        var shoulderBladeEnd = {
            x: point2.x,
            y: point17.y
        };
        var shoulderBladeLine = drawLineWithLabel({
            x: point17.x,
            y: point17.y
        }, shoulderBladeEnd, 'Shoulder Blade Line', FRAME_COLOR);
        try {
            shoulderBladeLine.strokeDashes = DASH_PATTERN;
        } catch (eDashShoulder) {}
        var midpoint1710 = {
            x: (point17.x + point10.x) / 2,
            y: (point17.y + point10.y) / 2
        };
        point17a = {
            x: midpoint1710.x - 1.5,
            y: midpoint1710.y
        };
        placeFrameMarker([point17a.x, point17a.y], '17a');
        point18 = {
            x: point13.x,
            y: point17a.y
        };
        var line17a18 = drawFrameLine([point17a.x, point17a.y], [point18.x, point18.y], '17a - 18');
        try {
            line17a18.strokeDashes = DASH_PATTERN;
        } catch (eDash17a18) {}
        placeFrameMarker([point18.x, point18.y], '18');
    }
    var topLineY = point1.y;
    var bustLineY = bustLineStart.y;
    var waistLineY = waistLineStart.y;
    var hipLineY = hipLineStart.y;
    var hemLineY = hemLineStart.y;
    var hemEnd = {
        x: point14.x,
        y: hemLineY
    };
    var bustEnd = {
        x: point14.x,
        y: bustLineY
    };
    var waistEnd = {
        x: point14.x,
        y: waistLineY
    };
    var hipEnd = {
        x: point14.x,
        y: hipLineY
    };

    var point25 = null;
    var backDiagVec = null;
    if (point11 && point6a && point2) {
        backDiagVec = {
            x: point6a.x - point2.x,
            y: point6a.y - point2.y
        };
        if (Math.abs(backDiagVec.x) > 0.0001 || Math.abs(backDiagVec.y) > 0.0001) {
            var refPoint25 = {
                x: point11.x + backDiagVec.x,
                y: point11.y + backDiagVec.y
            };
            point25 = extendLineToY(point11, refPoint25, hemLineY);
        } else {
            backDiagVec = null;
        }
    }

    var line1125 = null;
    if (point25) {
        line1125 = drawFrameLine([point11.x, point11.y], [point25.x, point25.y], 'Back Side Straightening Line', FRAME_COLOR, casualLinesLayer);
        if (line1125) {
            try {
                line1125.strokeDashes = DASH_PATTERN;
            } catch (eDash1125) {}
        }
        placeFrameMarker([point25.x, point25.y], '25');
    }

    var point26 = null;
    if (point11) {
        point26 = {
            x: point11.x,
            y: hipLineY
        };
        placeFrameMarker([point26.x, point26.y], '26');
    }

    var point27 = {
        x: point12.x,
        y: hipLineY
    };
    placeFrameMarker([point27.x, point27.y], '27');

    var point28 = null;
    if (backDiagVec) {
        var refPoint28 = {
            x: point11.x + backDiagVec.x,
            y: point11.y + backDiagVec.y
        };
        point28 = extendLineToY(point11, refPoint28, hipLineY);
    } else if (point11) {
        point28 = {
            x: point11.x,
            y: hipLineY
        };
    }
    if (point28) {
        placeFrameMarker([point28.x, point28.y], '28');
    }

    var point19a = {
        x: point14.x,
        y: hipLineY
    };
    placeFrameMarker([point19a.x, point19a.y], '19a');

    var point19 = {
        x: point14.x,
        y: waistLineY
    };
    placeFrameMarker([point19.x, point19.y], '19');
    var hipSpanBack = (point6a && point26) ? distanceBetween(point6a, point26) : null;
    var hipSpanFront = (point19a && point27) ? distanceBetween(point19a, point27) : null;
    var hiG = null;
    var hipShortage = null;
    var halfHipShortage = null;
    var halfHipShortageMagnitude = null;
    if (hipSpanBack !== null && hipSpanFront !== null) {
        hiG = hipSpanBack + hipSpanFront;
    }
    var halfHiW = (latestDerived && typeof latestDerived.hiw == 'number') ? latestDerived.hiw : null;
    if (hiG !== null && halfHiW !== null) {
        hipShortage = hiG - halfHiW;
        halfHipShortage = hipShortage / 2;
        halfHipShortageMagnitude = Math.abs(halfHipShortage);
    }
    if (hiG !== null) {
        registerMeasurementLabel('HiG', 'Hip Sum (HiG)');
        measurementResults.HiG = hiG;
        easeResults.HiG = 0;
        finalResults.HiG = hiG;
    }
    if (hipShortage !== null) {
        registerMeasurementLabel('HipShortage', 'Hip Shortage');
        measurementResults.HipShortage = hipShortage;
        easeResults.HipShortage = 0;
        finalResults.HipShortage = hipShortage;
    }
    if (latestDerived) {
        latestDerived.hiG = hiG;
        latestDerived.hipShortage = hipShortage;
        latestDerived.hipSpanBack = hipSpanBack;
        latestDerived.hipSpanFront = hipSpanFront;
    }

    var point29 = null;
    if (point27 && halfHipShortageMagnitude !== null) {
        point29 = {
            x: point27.x + halfHipShortageMagnitude,
            y: hipLineY
        };
        placeFrameMarker([point29.x, point29.y], '29');
    }
    var point30 = null;
    if (point28 && halfHipShortageMagnitude !== null) {
        point30 = {
            x: point28.x - halfHipShortageMagnitude,
            y: hipLineY
        };
        placeFrameMarker([point30.x, point30.y], '30');
    }
    var hipLinePoint6a = null;
    var hemConnectorEnd = null;
    var waistConnectorEnd = null;

    if (point30) {
        var point30Hem = extendLineToY(point11, point30, hemLineY);
        var line1130 = null;
        if (point30Hem) {
            line1130 = drawFrameLine([point11.x, point11.y], [point30Hem.x, point30Hem.y], 'Back Side Line 1', FRAME_COLOR, casualLinesLayer);
            if (line1130) {
                try {
                    line1130.strokeDashes = [];
                } catch (eSolid1130) {}
            }
        }

        var point30Waist = (typeof waistLineY === 'number') ? extendLineToY(point11, point30, waistLineY) : null;

        function perpendicularFootOnBackDiagonal(sourcePoint) {
            if (!sourcePoint || !point2 || !point6a) return null;
            var diagX = point6a.x - point2.x;
            var diagY = point6a.y - point2.y;
            var diagLenSq = (diagX * diagX) + (diagY * diagY);
            if (diagLenSq < 1e-6) return null;
            var tDiag = ((sourcePoint.x - point2.x) * diagX + (sourcePoint.y - point2.y) * diagY) / diagLenSq;
            return {
                x: point2.x + diagX * tDiag,
                y: point2.y + diagY * tDiag
            };
        }

        hipLinePoint6a = perpendicularFootOnBackDiagonal(point30);

        if (!hipLinePoint6a && point6a) {
            hipLinePoint6a = {
                x: point6a.x,
                y: hipLineY
            };
        } else if (!hipLinePoint6a && point30) {
            hipLinePoint6a = {
                x: point30.x,
                y: hipLineY
            };
        }

        if (hipLinePoint6a) {
            var line306a = drawFrameLine([point30.x, point30.y], [hipLinePoint6a.x, hipLinePoint6a.y], 'Back Hip Line', FRAME_COLOR, casualLinesLayer);
            if (line306a) {
                try {
                    line306a.strokeDashes = DASH_PATTERN;
                } catch (eSolid306a) {}
            }
        }

        if (point30Hem) {
            hemConnectorEnd = perpendicularFootOnBackDiagonal(point30Hem);
            if (hemConnectorEnd) {
                var hemConnector = drawFrameLine([point30Hem.x, point30Hem.y], [hemConnectorEnd.x, hemConnectorEnd.y], 'Back Hem Line', FRAME_COLOR, casualLinesLayer);
                if (hemConnector) {
                    try {
                        hemConnector.strokeDashes = [];
                    } catch (eSolidHemConnector) {}
                }
            }
        }

        if (point30Waist) {
            waistConnectorEnd = perpendicularFootOnBackDiagonal(point30Waist);
            if (waistConnectorEnd) {
                var waistConnector = drawFrameLine([point30Waist.x, point30Waist.y], [waistConnectorEnd.x, waistConnectorEnd.y], 'Back Waist Line', FRAME_COLOR, casualLinesLayer);
                if (waistConnector) {
                    try {
                        waistConnector.strokeDashes = DASH_PATTERN.slice(0);
                    } catch (eSolidWaistConnector) {}
                }
            }
        }
    }

    if (hemConnectorEnd) {
        point8 = hemConnectorEnd;
    }

    if (point8) {
        placeFrameMarker([point8.x, point8.y], 8);
        var centreBackWithoutSeam = drawFrameLine([point2.x, point2.y], [point8.x, point8.y], 'Centre Back (CB)');
        if (centreBackWithoutSeam && point7) {
            centreBackWithoutSeam.setEntirePath([toArt(point2), toArt(point7), toArt(point8)]);
        }
    }

    var point29Hem = null;
    if (point29) {
        point29Hem = extendLineToY(point12, point29, hemLineY);
        drawFrameLine([point12.x, point12.y], [point29Hem.x, point29Hem.y], 'New Front Side Line', FRAME_COLOR, casualLinesLayer);
    }

    var frontLengthFinal = finalResults['FL'];
    if (isNaN(frontLengthFinal)) frontLengthFinal = defaults.FL;
    var point20Offset = frontLengthFinal - 1;
    if (isNaN(point20Offset) || point20Offset < 0) point20Offset = 0;
    var point20 = {
        x: point19.x,
        y: point19.y - point20Offset
    };
    placeFrameMarker([point20.x, point20.y], '20');
    var point20GuideLength = 20;
    var point20Guide = {
        x: point20.x + point20GuideLength,
        y: point20.y
    };
    var line20Guide = drawFrameLine([point20.x, point20.y], [point20Guide.x, point20Guide.y], '20 Guideline', FRAME_COLOR);
    try {
        line20Guide.strokeDashes = DASH_PATTERN;
    } catch (eDash20) {}
    var point20aOffset = NeG;
    if (isNaN(point20aOffset)) point20aOffset = 0;
    var point20a = {
        x: point20.x + point20aOffset,
        y: point20.y
    };
    placeFrameMarker([point20a.x, point20a.y], '20a');
    var point23Offset = NeG + 0.5;
    if (isNaN(point23Offset)) point23Offset = NeG;
    if (isNaN(point23Offset)) point23Offset = 0;
    var point23 = {
        x: point20.x,
        y: point20.y + point23Offset
    };
    placeFrameMarker([point23.x, point23.y], '23');
    var frontNeckBulge = Math.max(MIN_HANDLE_LENGTH_CM, NeG / 2);
    var frontNeckCurve = drawCurveBetween(point20a, point23, {
        name: 'Front Neck Curve',
        bulgeCm: frontNeckBulge
    });
    if (frontNeckCurve) {
        var frontHandleLength = Math.max(MIN_HANDLE_LENGTH_CM, NeG / 2);
        setHorizontalHandle(frontNeckCurve, 1, 'left', frontHandleLength, 1);
    }
    var point24 = null;
    if (fShSValue > 0) {
        var frontShoulderAngleSafe = isNaN(frontShoulderAngle) ? 0 : frontShoulderAngle;
        var frontShoulderRadiansFront = frontShoulderAngleSafe * Math.PI / 180;
        point24 = {
            x: point20a.x + Math.cos(frontShoulderRadiansFront) * fShSValue,
            y: point20a.y + Math.sin(frontShoulderRadiansFront) * fShSValue
        };
        drawFrameLine([point20a.x, point20a.y], [point24.x, point24.y], 'Front Shoulder Line');
        placeFrameMarker([point24.x, point24.y], '24');
    }
    var BrDValue = finalResults['BrD'];
    if (isNaN(BrDValue)) BrDValue = parseField(rowRefs['BrD'].measField, defaults.BrD);
    if (isNaN(BrDValue)) BrDValue = defaults.BrD;
    var point21Offset = BrDValue - 1;
    if (isNaN(point21Offset) || point21Offset < 0) point21Offset = 0;
    var point21 = {
        x: point20.x,
        y: point20.y + point21Offset
    };
    placeFrameMarker([point21.x, point21.y], '21');
    var dartOffset = (BrGPlus / 2) - 0.3;
    if (isNaN(dartOffset)) dartOffset = 0;
    if (dartOffset < 0) dartOffset = 0;
    var point22 = {
        x: point21.x + dartOffset,
        y: point21.y
    };
    var line21to22 = drawFrameLine([point21.x, point21.y], [point22.x, point22.y], 'Bust Distance');
    try {
        line21to22.strokeDashes = DASH_PATTERN;
    } catch (eDash21) {}
    placeFrameMarker([point22.x, point22.y], '22');
    var frontDartTopY = point20.y;
    if (point24) {
        var shoulderDX = point24.x - point20a.x;
        if (Math.abs(shoulderDX) > 0.0001) {
            var tIntersectDart = (point22.x - point20a.x) / shoulderDX;
            if (tIntersectDart >= 0 && tIntersectDart <= 1) {
                frontDartTopY = point20a.y + (point24.y - point20a.y) * tIntersectDart;
            }
        } else if (Math.abs(point22.x - point20a.x) < 0.0001) {
            var minShoulderY = Math.min(point20a.y, point24.y);
            var maxShoulderY = Math.max(point20a.y, point24.y);
            if (frontDartTopY < minShoulderY) frontDartTopY = minShoulderY;
            if (frontDartTopY > maxShoulderY) frontDartTopY = maxShoulderY;
        }
    }
    if (isNaN(frontDartTopY)) frontDartTopY = point20.y;
    var frontDartTop = {
        x: point22.x,
        y: frontDartTopY
    };
    var frontDartBottom = {
        x: point22.x,
        y: hemLineY
    };
    drawLineWithLabel({
        x: point2.x,
        y: topLineY
    }, {
        x: point2.x,
        y: hemLineY
    }, 'Centre Back (CB)', FRAME_COLOR);
    var backArmLineTopY = point16 ? point16.y : topLineY;
    var backArmLineStart = {
        x: point10.x,
        y: backArmLineTopY
    };
    var backArmLineEnd = {
        x: point10.x,
        y: hipLineY
    };
    var backArmLinePath = drawLineWithLabel(backArmLineStart, backArmLineEnd, 'Back Arm Line', FRAME_COLOR);
    if (backArmLinePath) {
        try {
            backArmLinePath.strokeDashes = DASH_PATTERN;
        } catch (eDashBackArm) {}
    }
    drawLineWithLabel({
        x: point11.x,
        y: point10.y
    }, {
        x: point11.x,
        y: hemLineY
    }, 'Back Side Line 2', FRAME_COLOR);
    var frontSideLinePath = drawLineWithLabel({
        x: point12.x,
        y: point12.y
    }, {
        x: point12.x,
        y: hemLineY
    }, 'Front Side Line', FRAME_COLOR);
    if (frontSideLinePath) {
        try {
            frontSideLinePath.strokeDashes = DASH_PATTERN;
        } catch (eDashFrontSide) {}
    }
    var frontArmLineTopY = topLineY + 8;
    if (frontArmLineTopY > waistLineY) frontArmLineTopY = waistLineY;
    var frontArmLineStart = {
        x: point13.x,
        y: frontArmLineTopY
    };
    var frontArmLineEnd = {
        x: point13.x,
        y: waistLineY
    };
    var frontArmLinePath = drawLineWithLabel(frontArmLineStart, frontArmLineEnd, 'Front Arm Line', FRAME_COLOR);
    if (frontArmLinePath) {
        try {
            frontArmLinePath.strokeDashes = DASH_PATTERN;
        } catch (eDashFrontArm) {}
    }
    var frontDartLinePath = drawLineWithLabel(frontDartTop, frontDartBottom, 'Front Dart Line', FRAME_COLOR);
    try {
        frontDartLinePath.strokeDashes = DASH_PATTERN;
    } catch (eDashDart) {}
    drawLineWithLabel({
        x: point14.x,
        y: point20.y
    }, {
        x: point14.x,
        y: hemLineY
    }, 'Centre Front (CF)', FRAME_COLOR);
    var cfBottom = {
        x: point14.x,
        y: hemLineY
    };
    var frontSideBottom = point29Hem ? {
        x: point29Hem.x,
        y: point29Hem.y
    } : {
        x: point12.x,
        y: hemLineY
    };
    drawFrameLine(cfBottom, frontSideBottom, 'Front Hem Line', FRAME_COLOR, casualLinesLayer);
    var hemLinePath = drawLineWithLabel(hemLineStart, hemEnd, 'Hem Line', FRAME_COLOR);
    if (hemLinePath) {
        try {
            hemLinePath.strokeDashes = DASH_PATTERN;
        } catch (eDashHem) {}
    }
    var baseArtboardIndex = doc.artboards.getActiveArtboardIndex();
    if (shouldShowMeasurementPalette) {
        showMeasurementPaletteWindow(selectedProfile ? selectedProfile.name : '');
    } else {
        closeMeasurementPalette();
    }
    cropArtboardAroundItems(baseArtboardIndex, [linesLayer, casualLinesLayer, markersGroup, numbersGroup, labelsGroup], 10);
    var bustLinePath = drawLineWithLabel(bustLineStart, bustEnd, 'Front Bust Line', FRAME_COLOR);
    try {
        bustLinePath.strokeDashes = DASH_PATTERN;
    } catch (eDashBust) {}
    var waistLinePath = drawLineWithLabel(waistLineStart, waistEnd, 'Front Waist Line', FRAME_COLOR);
    try {
        waistLinePath.strokeDashes = DASH_PATTERN;
    } catch (eDashWaist) {}
    var hipLinePath = drawLineWithLabel(hipLineStart, hipEnd, 'Front Hip Line', FRAME_COLOR);
    try {
        hipLinePath.strokeDashes = DASH_PATTERN;
    } catch (eDashHip) {}
    if (!$.global.guidoBodiceFrame) $.global.guidoBodiceFrame = {};
    $.global.guidoBodiceFrame.originCm = {
        x: 0,
        y: 0
    };
    $.global.guidoBodiceFrame.extentsCm = {
        left: (artboardRect[0] - originX) / cm(1),
        right: rightBoundaryCm,
        top: topBoundaryCm,
        bottom: (artboardRect[3] - originY) / cm(1)
    };
    $.global.guidoBodiceFrame.helpers = {
        drawFrameLine: drawFrameLine,
        drawCurve: drawCurveBetween,
        placeFrameMarker: placeFrameMarker,
        drawLabel: drawCenteredLabel
    };
    $.global.guidoBodiceFrame.points = {
        p1: point1,
        p1a: point1a,
        p1aExtension: point1aExtension,
        p2: point2,
        p3: point3,
        p4: point4,
        p5: point5,
        p6: point6,
        p6a: point6a,
        p7: point7,
        p8: point8,
        p9: point9,
        p10: point10,
        p11: point11,
        p12: point12,
        p13: point13,
        p13a: point13a,
        p14: point14,
        point19: point19,
        point19a: point19a,
        point20: point20,
        point20a: point20a,
        point21: point21,
        point22: point22,
        point23: point23,
        point24: point24,
        point25: point25,
        point26: point26,
        point27: point27,
        point28: point28,
        point29: point29,
        point30: point30,
        frontDartTop: frontDartTop,
        frontDartBottom: frontDartBottom,
        hemEnd: hemEnd,
        bustEnd: bustEnd,
        waistEnd: waistEnd,
        hipEnd: hipEnd,
        frontShoulderEnd: frontShoulderEnd,
        backShoulderEnd: backShoulderEnd,
        p16: point16,
        p17: point17,
        p17a: point17a,
        p18: point18
    };
    $.global.guidoBodiceFrame.activePoints = {
        bust: point9,
        waist: point7,
        hem: point8
    };
    $.global.guidoBodiceFrame.measurements = {

        measurements: measurementResults,

        ease: easeResults,

        constructed: finalResults,

        BL: BL,

        BLBal: BLBal,

        BLFinal: BLFinal,

        NeG: NeG,

        AGPlus: AGPlus,

        BGPlus: BGPlus,

        BrGPlus: BrGPlus,

        BrD: BrDValue,

        centreFrontX: point14.x,

        backArmBase: backArmBase,

        backArmLineX: point10.x,

        MoL: MoL,

        HiD: HiD,


        hiG: hiG,

        hipShortage: hipShortage,

        fitProfile: selectedProfile.name,

        shoulderDifference: shoulderDifference,

        frontShoulderAngle: frontShoulderAngle,

        backShoulderAngle: backShoulderAngle,

        fShS: fShSValue,

        bShS: bShS

    };

    function drawCasualFrontArmhole() {
        if (!point24 || !point12 || !point13a || !point14) return;

        var casualLayer = ensureLayer('Casual Bodice Lines');
        try {
            casualLayer.locked = false;
        } catch (eCasLockFront) {}
        try {
            casualLayer.visible = true;
        } catch (eCasVisFront) {}
        var casualStroke = makeRGB(0, 0, 0);

        removePageItemsByName(casualLayer, ['Front Armhole Curve']);

        function makeVector(a, b) {
            return {
                x: b.x - a.x,
                y: b.y - a.y
            };
        }

        function addVec(pt, vec) {
            return {
                x: pt.x + vec.x,
                y: pt.y + vec.y
            };
        }

        function magnitude(v) {
            return Math.sqrt((v.x * v.x) + (v.y * v.y));
        }

        function normalizeVec(v) {
            var len = magnitude(v);
            if (len < 1e-6) return {
                x: 0,
                y: 0
            };
            return {
                x: v.x / len,
                y: v.y / len
            };
        }

        function scaleVec(v, scalar) {
            return {
                x: v.x * scalar,
                y: v.y * scalar
            };
        }

        function bezierPoint(t, p0, p1, p2, p3) {
            var mt = 1 - t;
            var mt2 = mt * mt;
            var t2 = t * t;
            var a = mt2 * mt;
            var b = 3 * mt2 * t;
            var c = 3 * mt * t2;
            var d = t * t2;
            return {
                x: a * p0.x + b * p1.x + c * p2.x + d * p3.x,
                y: a * p0.y + b * p1.y + c * p2.y + d * p3.y
            };
        }

        function bezierDerivative(t, p0, p1, p2, p3) {
            var mt = 1 - t;
            return {
                x: 3 * mt * mt * (p1.x - p0.x) + 6 * mt * t * (p2.x - p1.x) + 3 * t * t * (p3.x - p2.x),
                y: 3 * mt * mt * (p1.y - p0.y) + 6 * mt * t * (p2.y - p1.y) + 3 * t * t * (p3.y - p2.y)
            };
        }

        var shoulderTip = point24;
        var bustPoint = point12;
        var guidePoint = point13a;
        var shoulderBase = point20a ? point20a : shoulderTip;

        var dir1214 = makeVector(bustPoint, point14);
        var dir1214Len = magnitude(dir1214);
        if (dir1214Len < 1e-6) return;
        var dir1214Norm = normalizeVec(dir1214);

        var shoulderVector = makeVector(shoulderBase, shoulderTip);
        var bustVector = makeVector(shoulderTip, bustPoint);

        var perp = {
            x: -shoulderVector.y,
            y: shoulderVector.x
        };
        if (magnitude(perp) < 1e-6) perp = {
            x: -bustVector.y,
            y: bustVector.x
        };
        if ((perp.x * (guidePoint.x - shoulderTip.x)) + (perp.y * (guidePoint.y - shoulderTip.y)) < 0) {
            perp.x *= -1;
            perp.y *= -1;
        }
        perp = normalizeVec(perp);

        var shoulderHandleLen = Math.max(magnitude(bustVector) * 0.45, 1.2);
        var shoulderHandle = scaleVec(perp, shoulderHandleLen);

        var handleLen = dir1214Len * 0.5;
        if (handleLen < 0.5) handleLen = 0.5;
        var tSolve = 0.5;

        var p0 = shoulderTip;
        var p1 = addVec(shoulderTip, shoulderHandle);
        var p3 = bustPoint;

        for (var iter = 0; iter < 25; iter++) {
            var p2 = addVec(p3, scaleVec(dir1214Norm, handleLen));
            var current = bezierPoint(tSolve, p0, p1, p2, p3);
            var diff = {
                x: current.x - guidePoint.x,
                y: current.y - guidePoint.y
            };
            if (Math.abs(diff.x) < 0.0005 && Math.abs(diff.y) < 0.0005) break;

            var dBdt = bezierDerivative(tSolve, p0, p1, p2, p3);
            var coeff = 3 * (1 - tSolve) * tSolve * tSolve;
            var dBdL = scaleVec(dir1214Norm, coeff);
            var det = (dBdt.x * dBdL.y) - (dBdt.y * dBdL.x);
            if (Math.abs(det) < 1e-6) break;

            var deltaT = (-diff.x * dBdL.y + dBdL.x * diff.y) / det;
            var deltaL = (-dBdt.x * diff.y + dBdt.y * diff.x) / det;
            tSolve += deltaT;
            handleLen += deltaL;
            if (tSolve < 0.05) tSolve = 0.05;
            if (tSolve > 0.95) tSolve = 0.95;
            if (handleLen < 0.05) handleLen = 0.05;
        }

        var shoulderAnchor = toArt(p0);
        var bustAnchor = toArt(p3);
        var handleStartAnchor = toArt(addVec(p0, shoulderHandle));
        var handleEndAnchor = toArt(addVec(p3, scaleVec(dir1214Norm, handleLen)));

        var path = casualLayer.pathItems.add();
        path.name = 'Front Armhole Curve';
        path.stroked = true;
        path.strokeWidth = 1;
        path.strokeColor = casualStroke;
        path.filled = false;
        path.closed = false;
        path.setEntirePath([shoulderAnchor, bustAnchor]);

        var pts = path.pathPoints;
        if (pts.length === 2) {
            var startPt = pts[0];
            startPt.pointType = PointType.SMOOTH;
            startPt.leftDirection = shoulderAnchor;
            startPt.rightDirection = handleStartAnchor;

            var endPt = pts[1];
            endPt.pointType = PointType.SMOOTH;
            endPt.leftDirection = handleEndAnchor;
            endPt.rightDirection = bustAnchor;
        }
    }

    function drawCasualBackArmhole() {
        if (!backShoulderEnd || !point17 || !point17a || !point11 || !point4) return;

        var casualLayer = ensureLayer('Casual Bodice Lines');
        try {
            casualLayer.locked = false;
        } catch (eCasLockBack) {}
        try {
            casualLayer.visible = true;
        } catch (eCasVisBack) {}
        var casualStroke = makeRGB(0, 0, 0);

        removePageItemsByName(casualLayer, ['Back Armhole Curve', 'Back Armhole Curve (Shoulder-17)', 'Back Armhole Curve (17-11)']);

        function makeVector(a, b) {
            return {
                x: b.x - a.x,
                y: b.y - a.y
            };
        }

        function addVec(pt, vec) {
            return {
                x: pt.x + vec.x,
                y: pt.y + vec.y
            };
        }

        function magnitude(v) {
            return Math.sqrt((v.x * v.x) + (v.y * v.y));
        }

        function normalizeVec(v) {
            var len = magnitude(v);
            if (len < 1e-6) return {
                x: 0,
                y: 0
            };
            return {
                x: v.x / len,
                y: v.y / len
            };
        }

        function scaleVec(v, scalar) {
            return {
                x: v.x * scalar,
                y: v.y * scalar
            };
        }

        function bezierPoint(p0, p1, p2, p3, t) {
            var mt = 1 - t;
            var mt2 = mt * mt;
            var t2 = t * t;
            var a = mt2 * mt;
            var b = 3 * mt2 * t;
            var c = 3 * mt * t2;
            var d = t * t2;
            return {
                x: a * p0.x + b * p1.x + c * p2.x + d * p3.x,
                y: a * p0.y + b * p1.y + c * p2.y + d * p3.y
            };
        }

        function bezierDerivative(p0, p1, p2, p3, t) {
            var mt = 1 - t;
            return {
                x: 3 * mt * mt * (p1.x - p0.x) + 6 * mt * t * (p2.x - p1.x) + 3 * t * t * (p3.x - p2.x),
                y: 3 * mt * mt * (p1.y - p0.y) + 6 * mt * t * (p2.y - p1.y) + 3 * t * t * (p3.y - p2.y)
            };
        }

        var startAnchor = backShoulderEnd;
        var midAnchor = point17;
        var guidePoint = point17a;
        var endAnchor = point11;
        var guideLinePoint = point4;

        var startToMid = makeVector(startAnchor, midAnchor);
        var startLen = magnitude(startToMid);
        if (startLen < 1e-6) return;
        var startDir = normalizeVec(startToMid);
        var startHandleLen = Math.min(startLen * 0.35, 5);
        if (startHandleLen < 0.5) startHandleLen = 0.5;
        var startHandle = addVec(startAnchor, scaleVec(startDir, startHandleLen));
        var verticalDrop = guidePoint ? Math.abs(guidePoint.y - midAnchor.y) : 5;
        if (verticalDrop < 0.5) verticalDrop = 0.5;
        var midDown = {
            x: midAnchor.x,
            y: midAnchor.y + verticalDrop
        };
        var midOutgoingHandle = {
            x: guidePoint ? guidePoint.x + 1 : midDown.x,
            y: midDown.y
        };
        var midIncomingHandle = {
            x: midAnchor.x - (midOutgoingHandle.x - midAnchor.x),
            y: midAnchor.y - (midOutgoingHandle.y - midAnchor.y)
        };

        var lineDir = makeVector(endAnchor, guideLinePoint);
        var lineDirLen = magnitude(lineDir);
        if (lineDirLen < 1e-6) lineDir = makeVector(endAnchor, midAnchor);
        lineDir = normalizeVec(lineDir);

        var handleLenEnd = magnitude(makeVector(endAnchor, guidePoint || endAnchor));
        if (handleLenEnd < 0.5) handleLenEnd = 0.5;

        var tSolve = 0.5;
        var p0 = midAnchor;
        var p1 = midOutgoingHandle;
        var p3 = endAnchor;

        for (var iter = 0; iter < 30; iter++) {
            var p2 = addVec(p3, scaleVec(lineDir, handleLenEnd));
            var target = guidePoint || midAnchor;
            var current = bezierPoint(p0, p1, p2, p3, tSolve);
            var diff = {
                x: current.x - target.x,
                y: current.y - target.y
            };
            if (Math.abs(diff.x) < 0.0003 && Math.abs(diff.y) < 0.0003) break;

            var dBdt = bezierDerivative(p0, p1, p2, p3, tSolve);
            var coeff = 3 * (1 - tSolve) * tSolve * tSolve;
            var dBdL = scaleVec(lineDir, coeff);
            var det = (dBdt.x * dBdL.y) - (dBdt.y * dBdL.x);
            if (Math.abs(det) < 1e-8) break;

            var deltaT = (-diff.x * dBdL.y + dBdL.x * diff.y) / det;
            var deltaL = (-dBdt.x * diff.y + dBdt.y * diff.x) / det;
            tSolve += deltaT;
            handleLenEnd += deltaL;
            if (tSolve < 0.05) tSolve = 0.05;
            if (tSolve > 0.95) tSolve = 0.95;
            if (handleLenEnd < 0.05) handleLenEnd = 0.05;
        }

        var endLeftHandle = addVec(endAnchor, scaleVec(lineDir, handleLenEnd));

        var armholePath = casualLayer.pathItems.add();
        armholePath.name = 'Back Armhole Curve';
        armholePath.stroked = true;
        armholePath.strokeWidth = 1;
        armholePath.strokeColor = casualStroke;
        armholePath.filled = false;
        armholePath.closed = false;
        armholePath.setEntirePath([toArt(startAnchor), toArt(midAnchor), toArt(endAnchor)]);

        var ptsBack = armholePath.pathPoints;
        if (ptsBack.length === 3) {
            var startPtBack = ptsBack[0];
            startPtBack.pointType = PointType.SMOOTH;
            startPtBack.leftDirection = toArt(startAnchor);
            startPtBack.rightDirection = toArt(startHandle);

            var midPtBack = ptsBack[1];
            midPtBack.pointType = PointType.SMOOTH;
            midPtBack.leftDirection = toArt(midIncomingHandle);
            midPtBack.rightDirection = toArt(midOutgoingHandle);

            var endPtBack = ptsBack[2];
            endPtBack.pointType = PointType.SMOOTH;
            endPtBack.leftDirection = toArt(endLeftHandle);
            endPtBack.rightDirection = toArt(endAnchor);
        }
    }

    drawCasualFrontArmhole();
    drawCasualBackArmhole();

    var backBodiceLayer = ensureLayer('Back Bodice');
    try {
        backBodiceLayer.locked = false;
    } catch (eBackLock) {}
    try {
        backBodiceLayer.visible = true;
    } catch (eBackVis) {}
    clearLayerContents(backBodiceLayer);
    duplicateItemsToLayer([
        'Back Neck Curve',
        'Back Shoulder Line',
        'Centre Back (CB)',
        'Back Hem Line',
        'Back Hip Line',
        'Back Armhole Curve',
        'Back Side Line 1',
        'Back Waist Line',
        'Shoulder Blade Line',
        'Bust Line'
    ], backBodiceLayer);
    var backBustStart = (point11 && point30 && typeof bustLineY === 'number') ? extendLineToY(point11, point30, bustLineY) : null;
    var backBustEnd = point9 ? {
        x: point9.x,
        y: point9.y
    } : null;
    var bustLineCopies = findLayerPathItemsByName(backBodiceLayer, 'Bust Line');
    for (var bb = 0; bb < bustLineCopies.length; bb++) {
        var bustCopy = bustLineCopies[bb];
        if (!bustCopy) continue;
        if (backBustStart && backBustEnd) setPathBetween(bustCopy, backBustStart, backBustEnd);
        try {
            bustCopy.name = 'Back Bust Line';
        } catch (eRenameBust) {}
    }
    var hipLineCopies = findLayerPathItemsByName(backBodiceLayer, 'Back Hip Line');
    for (var bh = 0; bh < hipLineCopies.length; bh++) {
        var hipCopy = hipLineCopies[bh];
        if (!hipCopy) continue;
        try {
            hipCopy.strokeDashes = DASH_PATTERN.slice(0);
        } catch (eDashHip) {}
    }
    if (point17 && point2) {
        var cbEndPoint = point8 ? {
            x: point8.x,
            y: point8.y
        } : (point7 ? {
            x: point7.x,
            y: point7.y
        } : null);
        var shoulderStart = {
            x: point17.x,
            y: point17.y
        };
        var shoulderEnd = {
            x: point2.x,
            y: point17.y
        };
        var cbStartPoint = {
            x: point2.x,
            y: point2.y
        };
        var cbEnd = cbEndPoint || shoulderEnd;
        var shoulderIntersection = lineIntersection(shoulderStart, shoulderEnd, cbStartPoint, cbEnd);
        if (!shoulderIntersection) shoulderIntersection = shoulderEnd;
        var shoulderCopies = findLayerPathItemsByName(backBodiceLayer, 'Shoulder Blade Line');
        for (var sb = 0; sb < shoulderCopies.length; sb++) {
            var sbCopy = shoulderCopies[sb];
            if (!sbCopy) continue;
            setPathBetween(sbCopy, shoulderStart, shoulderIntersection);
        }
    }
    ensureCentreBackOnLayer(backBodiceLayer, [point2, point7, point8]);
    if (point9 && point11) {
        var backBustLinePath = backBodiceLayer.pathItems.add();
        backBustLinePath.closed = false;
        backBustLinePath.stroked = true;
        backBustLinePath.strokeWidth = 1;
        backBustLinePath.strokeColor = FRAME_COLOR;
        backBustLinePath.filled = false;
        try {
            backBustLinePath.strokeDashes = DASH_PATTERN.slice(0);
        } catch (eDashBackBust) {}
        try {
            backBustLinePath.name = 'Back Bust Line';
        } catch (eNameBackBust) {}
        backBustLinePath.setEntirePath([toArt(point9), toArt(point11)]);
    }
    try {
        backBodiceLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
    } catch (eBackOrder) {}

    var frontBodiceLayer = ensureLayer('Front Bodice');
    try {
        frontBodiceLayer.locked = false;
    } catch (eFrontLock) {}
    try {
        frontBodiceLayer.visible = true;
    } catch (eFrontVis) {}
    clearLayerContents(frontBodiceLayer);
    duplicateItemsToLayer([
        'Front Armhole Curve',
        'Front Neck Curve',
        'Centre Front (CF)',
        'New Front Side Line',
        'Front Hem Line',
        'Front Shoulder Line',
        'Front Hip Line',
        'Front Bust Line',
        'Hem Line',
        'Front Waist Line'
    ], frontBodiceLayer);
    var frontBustLine = findLayerPathItemsByName(frontBodiceLayer, 'Front Bust Line');
    for (var fb = 0; fb < frontBustLine.length; fb++) {
        trimHorizontalPathRightOf(frontBustLine[fb], toArt(point12));
    }
    var frontWaistLine = findLayerPathItemsByName(frontBodiceLayer, 'Front Waist Line');
    for (var fw = 0; fw < frontWaistLine.length; fw++) {
        try {
            frontWaistLine[fw].remove();
        } catch (eRemoveFrontWaist) {}
    }
    var frontWaistEndPoint = extendLineToY(point12, point29 ? point29 : point12, waistLineStart.y);
    if (!frontWaistEndPoint) {
        frontWaistEndPoint = {
            x: point12.x,
            y: waistLineStart.y
        };
    }
    var replacementWaist = frontBodiceLayer.pathItems.add();
    replacementWaist.closed = false;
    replacementWaist.stroked = true;
    replacementWaist.strokeWidth = 1;
    replacementWaist.strokeColor = FRAME_COLOR;
    try {
        replacementWaist.strokeDashes = DASH_PATTERN.slice(0);
    } catch (eDashWaistFront) {}
    replacementWaist.filled = false;
    try {
        replacementWaist.name = 'Front Waist Line';
    } catch (eNameWaistFront) {}
    replacementWaist.setEntirePath([toArt(point19), toArt(frontWaistEndPoint)]);
    var frontHipLine = findLayerPathItemsByName(frontBodiceLayer, 'Front Hip Line');
    for (var fh = 0; fh < frontHipLine.length; fh++) {
        if (point29) trimHorizontalPathRightOf(frontHipLine[fh], toArt(point29));
    }
    var frontCFPaths = findLayerPathItemsByName(frontBodiceLayer, 'Centre Front (CF)');
    var cfEnd = {
        x: point14.x,
        y: hemLineY
    };
    for (var cfIdx = 0; cfIdx < frontCFPaths.length; cfIdx++) {
        var cfPath = frontCFPaths[cfIdx];
        if (!cfPath) continue;
        cfPath.setEntirePath([toArt(point23), toArt(cfEnd)]);
    }
    try {
        frontBodiceLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
    } catch (eFrontOrder) {}

    if (labelsGroup) {
        try {
            labelsGroup.visible = false;
        } catch (eHideLabels) {}
    }

    try {
        app.executeMenuCommand('fitin');
    } catch (eFit) {}
})();
