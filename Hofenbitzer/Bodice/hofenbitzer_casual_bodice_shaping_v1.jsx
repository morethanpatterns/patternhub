(function() {
    function documentHasArtwork(targetDoc) {
        if (!targetDoc) return false;
        try {
            if (targetDoc.pageItems.length > 0) return true;
        } catch (ePageItems) {}
        try {
            if (targetDoc.pathItems.length > 0) return true;
        } catch (ePathItems) {}
        try {
            if (targetDoc.textFrames.length > 0) return true;
        } catch (eTextFrames) {}
        return false;
    }

    function ensureWorkingDocument() {
        if (app.documents.length === 0) return app.documents.add();
        var active = null;
        try {
            active = app.activeDocument;
        } catch (eActiveDoc) {}
        if (!active) return app.documents.add();
        if (documentHasArtwork(active)) return app.documents.add();
        return active;
    }
    var doc = ensureWorkingDocument();

    function cm(val) {
        return val * 28.3464566929;
    }

    function ptToCm(val) {
        return val / 28.3464566929;
    }
    var latestDerived = null;
    var lastRecommendedOptimalBalanceText = null;
    var FORCE_HEADLESS_UI = false;
    function makeStubContainer(kind, label) {
        return {
            type: kind || '',
            text: label || '',
            value: false,
            items: [],
            selection: null,
            children: [],
            preferredSize: {},
            minimumSize: {},
            maximumSize: {},
            add: function(childKind, bounds, childLabel, props) {
                var child = makeStubContainer(childKind, childLabel);
                child.parent = this;
                this.children.push(child);
                if (childKind === 'dropdownlist') child.items = [];
                return child;
            }
        };
    }
    var defaults = {
        AhD: 20.1,
        NeG: 6.5,
        MoL: 75,
        BL: 41.6,
        BLBal: 0,
        HiD: 20,
        BackContour: 2,
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
        BackShoulderDartIntake: 1.5,
        WaistShaping: 1,
        OptimalBalance: 3.5,
        ShA: 20,
        ShoulderDifference: 2,
        showMeasurementPalette: false
    };
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
    var BL_NUMBER = MEASURE_ROWS.length + 1;
    var BL_ROW_DEF = {
        id: 'BL',
        label: BL_NUMBER + '. BL',
        defaultValue: defaults.BL,
        finalLabel: BL_NUMBER + '. BL (final)'
    };
    var FL_NUMBER = BL_NUMBER + 1;
    var SECONDARY_ROWS = [{
        id: 'NeG',
        label: '10. NeG',
        defaultValue: defaults.NeG,
        easeOptions: {
            enabled: false
        },
        finalLabel: '10. NeG (final)',
        finalOptions: {
            enabled: false
        }
    }, {
        id: 'MoL',
        label: '11. MoL',
        defaultValue: defaults.MoL,
        easeOptions: {
            enabled: false
        },
        finalLabel: '11. MoL (final)',
        finalOptions: {
            enabled: false
        }
    }, {
        id: 'HiD',
        label: '12. HiD',
        defaultValue: defaults.HiD,
        easeOptions: {
            enabled: false
        },
        finalLabel: '12. HiD (final)',
        finalOptions: {
            enabled: false
        }
    }, {
        id: 'ShA',
        label: '13. ShA (deg)',
        defaultValue: defaults.ShA,
        easeOptions: {
            enabled: false
        },
        finalLabel: '13. ShA (deg)',
        finalOptions: {
            enabled: false
        }
    }, {
        id: 'BrD',
        label: '14. BrD',
        defaultValue: defaults.BrD,
        easeOptions: {
            enabled: false
        },
        finalLabel: '14. BrD (final)',
        finalOptions: {
            enabled: false
        }
    }];

    function fitNames() {
        var names = [];
        for (var i = 0; i < FIT_PROFILES.length; i++) names.push(FIT_PROFILES[i].name);
        return names;
    }
    var uiBuildFailed = false;
    var dlg = null;
    if (!FORCE_HEADLESS_UI) {
        try {
            dlg = new Window('dialog', 'Measurement Panel (Default Size: 38, Fit 3)');
        } catch (eDialog) {
            uiBuildFailed = true;
            dlg = makeStubContainer('dialog');
        }
    } else {
        dlg = makeStubContainer('dialog');
    }
    var measurementPalette = null;
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
        var fallback = {
            text: (defaultValue !== null && defaultValue !== undefined && defaultValue !== '') ? String(defaultValue) : '',
            characters: options.characters || 8,
            enabled: false,
            onChanging: null,
            parent: null,
            value: (options && options.type === 'checkbox') ? !!defaultValue : undefined
        };
        if (uiBuildFailed || FORCE_HEADLESS_UI || !panel || typeof panel.add !== 'function') {
            uiBuildFailed = true;
            return fallback;
        }
        try {
            var row = panel.add('group');
            row.orientation = 'row';
            row.alignChildren = ['left', 'center'];
            var caption = label;
            if (options.note) caption += ' (' + options.note + ')';
            var st = row.add('statictext', undefined, caption + ':');
            if (options.labelWidth !== undefined) {
                st.minimumSize.width = options.labelWidth;
                st.preferredSize.width = options.labelWidth;
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
                var field = null;
                try {
                    field = row.add('edittext', undefined, (defaultValue !== null && defaultValue !== undefined && defaultValue !== '') ? String(defaultValue) : '');
                } catch (eAddField) {
                    uiBuildFailed = true;
                    return fallback;
                }
                field.characters = options.characters || 8;
                if (options.enabled === false) {
                    try {
                        field.enabled = false;
                    } catch (eOff) {}
                }
                return field;
            }
        } catch (eAddFieldRow) {
            uiBuildFailed = true;
            return fallback;
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
    var fitRow = measurementPanel.add('group');
    fitRow.orientation = 'row';
    fitRow.alignChildren = ['left', 'center'];
    fitRow.alignment = 'fill';
    fitRow.spacing = 6;
    fitRow.add('statictext', undefined, 'Fit Category:');
    var fitDropdown = fitRow.add('dropdownlist', undefined, fitNames());
    try {
        fitDropdown.selection = defaults.FitIndex;
    } catch (eFitSel) {
        fitDropdown.selection = {
            index: defaults.FitIndex
        };
    }
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
    var showSummaryCheckbox = addCheckbox(optionsPanel, 'Show Measurement Summary', defaults.showMeasurementPalette === true, 'Displays a palette listing measurement, ease, and final values');
    var shoulderDiffLabelNumber = 15;
    var shoulderDifferenceField = addFieldRow(optionsPanel, shoulderDiffLabelNumber + '. Shoulder Difference (deg)', defaults.ShoulderDifference, {});
    var backShoulderDartLabelNumber = shoulderDiffLabelNumber + 1;
    var backShoulderDartField = addFieldRow(optionsPanel, backShoulderDartLabelNumber + '. Back shoulder dart intake', defaults.BackShoulderDartIntake, {
        note: '1cm - 1.5cm'
    });
    var waistShapingLabelNumber = backShoulderDartLabelNumber + 1;
    var waistShapingField = addFieldRow(optionsPanel, waistShapingLabelNumber + '. Waist shaping', defaults.WaistShaping, {
        note: '0.5cm - 1.5cm'
    });
    shoulderDifferenceField.helpTip = 'page 159, number 24';
    try {
        shoulderDifferenceField.parent.children[0].helpTip = 'page 159, number 24';
    } catch (eShoulderTip) {}
    backShoulderDartField.helpTip = 'page 162, number 31';
    try {
        backShoulderDartField.parent.children[0].helpTip = 'page 162, number 31';
    } catch (eBackDartTip) {}
    waistShapingField.helpTip = 'page 161, number 30';
    try {
        waistShapingField.parent.children[0].helpTip = 'page 161, number 30';
    } catch (eWaistTip) {}
    function addCheckbox(panel, label, defaultValue, helpTip) {
        try {
            if (uiBuildFailed || FORCE_HEADLESS_UI || !panel || typeof panel.add !== 'function') throw new Error('no panel');
            var group = panel.add('group');
            group.orientation = 'row';
            group.alignChildren = ['left', 'center'];
            var cb = group.add('checkbox', undefined, label);
            cb.value = !!defaultValue;
            if (helpTip) cb.helpTip = helpTip;
            return cb;
        } catch (eAddCheckbox) {
            uiBuildFailed = true;
            return {
                value: !!defaultValue,
                enabled: false,
                onClick: null,
                helpTip: helpTip || ''
            };
        }
    }
    var backContourCheckbox = addCheckbox(optionsPanel, 'Hollow Back Curve / Flat Buttocks', false, 'page 157, number 8. An average of the 2-3cm (2.7cm is used)');

    function parseField(field, fallback) {
        if (!field || field.text === undefined || field.text === null) return fallback;
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

    function ensureBackContourValue() {}

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
                id: key,
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
            var globalPalette = $.global.guidoMeasurementPalette;
            if (globalPalette && globalPalette.window && typeof globalPalette.window.close === 'function') {
                globalPalette.window.close();
            }
            delete $.global.guidoMeasurementPalette;
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
        var rows = collectMeasurementRows();
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
                delete $.global.guidoMeasurementPalette;
            } catch (eDeleteGlobal) {}
        };

        measurementPalette = palette;
        $.global.guidoMeasurementPalette = {
            window: palette
        };
        try {
            palette.show();
        } catch (eShow) {}
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
        var frontShoulderAngle = shA - shoulderDifference;
        var backShoulderAngle = shA + shoulderDifference;
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
        var idx = defaults.FitIndex;
        try {
            if (fitDropdown && fitDropdown.selection !== undefined && fitDropdown.selection !== null) {
                if (typeof fitDropdown.selection.index === 'number') {
                    idx = fitDropdown.selection.index;
                } else if (typeof fitDropdown.selection === 'number') {
                    idx = fitDropdown.selection;
                }
            }
        } catch (eFitRead) {
            idx = defaults.FitIndex;
        }
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
            var st = null;
            try {
                var parentGroup = rowRefs[key].easeField ? rowRefs[key].easeField.parent : null;
                if (parentGroup && parentGroup.children && parentGroup.children.length > 0) {
                    st = parentGroup.children[0];
                }
            } catch (eEaseLabel) {
                st = null;
            }
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
        if (rowRefs[keyInit].measField && typeof rowRefs[keyInit].measField.onChanging !== 'undefined') {
            rowRefs[keyInit].measField.onChanging = updateConstruction;
        }
        if (rowRefs[keyInit].easeField && typeof rowRefs[keyInit].easeField.onChanging !== 'undefined') {
            rowRefs[keyInit].easeField.onChanging = updateConstruction;
        }
    }
    for (var keyInitSecondary in secondaryRowRefs) {
        if (!secondaryRowRefs.hasOwnProperty(keyInitSecondary)) continue;
        var secondaryRef = secondaryRowRefs[keyInitSecondary];
        if (secondaryRef.measField && typeof secondaryRef.measField.onChanging !== 'undefined') {
            secondaryRef.measField.onChanging = updateConstruction;
        }
        if (secondaryRef.easeField && typeof secondaryRef.easeField.onChanging !== 'undefined') {
            secondaryRef.easeField.onChanging = updateConstruction;
        }
    }
    if (shoulderDifferenceField && typeof shoulderDifferenceField.onChanging !== 'undefined') {
        shoulderDifferenceField.onChanging = updateConstruction;
    }
    // Back contour UI removed; nothing to wire up.
    if (backShoulderDartField && typeof backShoulderDartField.onChanging !== 'undefined') backShoulderDartField.onChanging = updateConstruction;
    if (waistShapingField && typeof waistShapingField.onChanging !== 'undefined') waistShapingField.onChanging = updateConstruction;
    // Back contour checkbox removed; nothing to wire up.
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
    var dialogResult = 1;
    if (!uiBuildFailed && dlg && typeof dlg.show === 'function') {
        dialogResult = dlg.show();
    }
        if (dialogResult !== 1) return;
    var selectedProfileIndex = defaults.FitIndex;
    try {
        if (fitDropdown && fitDropdown.selection !== undefined && fitDropdown.selection !== null) {
            if (typeof fitDropdown.selection.index === 'number') {
                selectedProfileIndex = fitDropdown.selection.index;
            } else if (typeof fitDropdown.selection === 'number') {
                selectedProfileIndex = fitDropdown.selection;
            }
        }
    } catch (eFitSelect) {
        selectedProfileIndex = defaults.FitIndex;
    }
    var selectedProfile = FIT_PROFILES[selectedProfileIndex];
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
    var backContourAuto = false;
    var BackContour = defaults.BackContour;
    measurementResults.BackContour = BackContour;
    easeResults.BackContour = 0;
    finalResults.BackContour = BackContour;
    var backShoulderDartIntake = parseField(backShoulderDartField, defaults.BackShoulderDartIntake);
    if (!isFinite(backShoulderDartIntake) || backShoulderDartIntake <= 0) backShoulderDartIntake = defaults.BackShoulderDartIntake;
    if (backShoulderDartIntake < 1) backShoulderDartIntake = 1;
    if (backShoulderDartIntake > 1.5) backShoulderDartIntake = 1.5;
    backShoulderDartField.text = formatValue(backShoulderDartIntake);
    measurementResults.BackShoulderDartIntake = backShoulderDartIntake;
    easeResults.BackShoulderDartIntake = 0;
    finalResults.BackShoulderDartIntake = backShoulderDartIntake;
    var waistShaping = parseField(waistShapingField, defaults.WaistShaping);
    if (!isFinite(waistShaping) || waistShaping <= 0) waistShaping = defaults.WaistShaping;
    if (waistShaping < 0.5) waistShaping = 0.5;
    if (waistShaping > 1.5) waistShaping = 1.5;
    waistShapingField.text = formatValue(waistShaping);
    measurementResults.WaistShaping = waistShaping;
    easeResults.WaistShaping = 0;
    finalResults.WaistShaping = waistShaping;
    closeMeasurementPalette();
    if (showSummaryCheckbox && showSummaryCheckbox.value) {
        var profileName = fitDropdown.selection && fitDropdown.selection.text ? fitDropdown.selection.text : ('Fit ' + (selectedProfileIndex + 1));
        showMeasurementPaletteWindow(profileName);
    }
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
    var frontShoulderAngle = ShA - shoulderDifference;
    var backShoulderAngle = ShA + shoulderDifference;
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

    function artToModel(artPt) {
        if (!artPt || typeof artPt.length !== 'number' || artPt.length < 2) return null;
        var x = (artPt[0] - originX) / cm(1);
        var y = (originY - artPt[1]) / cm(1);
        if (!isFinite(x) || !isFinite(y)) return null;
        return {
            x: x,
            y: y
        };
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

    function ensureSubLayer(parentLayer, name) {
        if (!parentLayer || !parentLayer.layers) return null;
        var goal = norm(name);
        for (var i = 0; i < parentLayer.layers.length; i++) {
            var sub = parentLayer.layers[i];
            if (norm(sub.name) === goal) return sub;
        }
        var newLayer = parentLayer.layers.add();
        newLayer.name = name;
        return newLayer;
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

    function removeGroupByName(parent, name) {
        if (!parent) return;
        var groups = parent.groupItems;
        for (var i = groups.length - 1; i >= 0; i--) {
            if (groups[i].name === name) {
                try {
                    groups[i].remove();
                } catch (eGrpRem) {}
            }
        }
    }

    function removePageItemsByName(parent, name) {
        if (!parent || !name) return;
        var items = parent.pageItems;
        for (var i = items.length - 1; i >= 0; i--) {
            var candidate = items[i];
            if (!candidate) continue;
            var candidateName = '';
            try {
                candidateName = candidate.name;
            } catch (eItemName) {}
            if (candidateName === name) {
                try {
                    candidate.remove();
                } catch (eItemRemove) {}
            }
        }
    }

    function removeGroupsByName(parent, names) {
        if (!parent || !names || !names.length) return;
        var groups = parent.groupItems;
        for (var i = groups.length - 1; i >= 0; i--) {
            var grp = groups[i];
            if (!grp) continue;
            for (var n = 0; n < names.length; n++) {
                if (grp.name === names[n]) {
                    try {
                        grp.remove();
                    } catch (eGrpRem) {}
                    break;
                }
            }
        }
    }

    function clearLayerDeep(layer) {
        if (!layer) return;
        try {
            var items = layer.pageItems;
            for (var i = items.length - 1; i >= 0; i--) {
                try {
                    items[i].remove();
                } catch (eLayerItem) {}
            }
        } catch (eLayerPage) {}
        try {
            if (layer.compoundPathItems) layer.compoundPathItems.removeAll();
        } catch (eLayerCompound) {}
        try {
            if (layer.pathItems) layer.pathItems.removeAll();
        } catch (eLayerPaths) {}
        try {
            if (layer.textFrames) layer.textFrames.removeAll();
        } catch (eLayerText) {}
        try {
            if (layer.groupItems) layer.groupItems.removeAll();
        } catch (eLayerGroups) {}
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
    var labelsLayer = ensureLayer('Labels, Markers & Numbers');
    try {
        labelsLayer.locked = false;
    } catch (eLbl1) {}
    try {
        labelsLayer.visible = true;
    } catch (eLbl2) {}
    var dartsLayer = ensureLayer('Darts & Shaping');
    try {
        dartsLayer.locked = false;
    } catch (eDartsLock) {}
    try {
        dartsLayer.visible = true;
    } catch (eDartsVis) {}
    var dartsSubLayer = ensureSubLayer(dartsLayer, 'Darts');
    if (dartsSubLayer) {
        try {
            dartsSubLayer.locked = false;
        } catch (eDartsSubLock) {}
        try {
            dartsSubLayer.visible = true;
        } catch (eDartsSubVis) {}
    }
    var shapingLayer = ensureSubLayer(dartsLayer, 'Shaping');
    if (shapingLayer) {
        try {
            shapingLayer.locked = false;
        } catch (eShapeLockInit) {}
        try {
            shapingLayer.visible = true;
        } catch (eShapeVisInit) {}
    }
    var casualLayer = ensureLayer('Casual Bodice');
    try {
        casualLayer.locked = false;
    } catch (eCasLockInit) {}
    try {
        casualLayer.visible = true;
    } catch (eCasVisInit) {}
    for (var iLayer = doc.layers.length - 1; iLayer >= 0; iLayer--) {
        var lyr = doc.layers[iLayer];
        var nameNorm = norm(lyr.name);
        if (nameNorm !== norm('Basic Frame') && nameNorm !== norm('Labels, Markers & Numbers') && nameNorm !== norm('Darts & Shaping') && nameNorm !== norm('Casual Bodice') && nameNorm !== norm('Patterns')) {
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
    labelsLayer = ensureLayer('Labels, Markers & Numbers');
    dartsLayer = ensureLayer('Darts & Shaping');
    dartsSubLayer = ensureSubLayer(dartsLayer, 'Darts');
    shapingLayer = ensureSubLayer(dartsLayer, 'Shaping');
    casualLayer = ensureLayer('Casual Bodice');
    var patternsLayer = ensureLayer('Patterns');
    var frontBodiceLayer = patternsLayer ? ensureSubLayer(patternsLayer, 'Front Bodice') : null;
    var backBodiceLayer = patternsLayer ? ensureSubLayer(patternsLayer, 'Back Bodice') : null;
    function unlockLayer(targetLayer) {
        if (!targetLayer) return;
        try {
            targetLayer.locked = false;
        } catch (eUnlockLayer) {}
        try {
            targetLayer.visible = true;
        } catch (eShowLayer) {}
    }
    unlockLayer(basicFrameLayer);
    unlockLayer(labelsLayer);
    unlockLayer(dartsLayer);
    unlockLayer(dartsSubLayer);
    unlockLayer(shapingLayer);
    unlockLayer(casualLayer);
    unlockLayer(patternsLayer);
    unlockLayer(frontBodiceLayer);
    unlockLayer(backBodiceLayer);
    clearLayerDeep(basicFrameLayer);
    var linesLayer = basicFrameLayer;
    removeGroupsByName(basicFrameLayer, ['Markers', 'Numbers']);
    var markersGroup = resetGroup(labelsLayer, 'Markers');
    var numbersGroup = resetGroup(labelsLayer, 'Numbers');
    var labelsGroup = resetGroup(labelsLayer, 'Labels');
    clearGroupDeep(markersGroup);
    clearGroupDeep(numbersGroup);
    clearGroupDeep(labelsGroup);
    var CASUAL_LINE_NAMES = [
        'New Front Side Line',
        'New Centre Back (CB)',
        'New Back Waist Line',
        'New Back Hem Line',
        'New Back Hip Line',
        'New Back Side Line',
        'Provisional Back Side Line',
        'Front Hem Line'
    ];
    for (var idxCasual = 0; idxCasual < CASUAL_LINE_NAMES.length; idxCasual++) {
        var casualName = CASUAL_LINE_NAMES[idxCasual];
        removePageItemsByName(linesLayer, casualName);
        removePageItemsByName(casualLayer, casualName);
    }
    if (dartsSubLayer) {
        clearGroupDeep(dartsSubLayer);
        removeGroupByName(dartsSubLayer, 'Darts');
        removePageItemsByName(dartsSubLayer, 'Back Shoulder Dart (Rotated)');
        removePageItemsByName(dartsSubLayer, 'Back Shoulder Dart');
        removePageItemsByName(dartsSubLayer, 'Front Shoulder Dart');
    }
    var dartsGroup = dartsSubLayer;
    if (shapingLayer) clearLayerDeep(shapingLayer);
    var shapingTarget = shapingLayer;
    var SHAPING_LINE_NAMES = [
        'Upper Front Side Line',
        'Upper Back Side Line',
        'Lower Front Side Curve',
        'Lower Back Side Curve',
        '36 Guide',
        '37 Guide'
    ];
    for (var idxShape = 0; idxShape < SHAPING_LINE_NAMES.length; idxShape++) {
        var shapeName = SHAPING_LINE_NAMES[idxShape];
        removePageItemsByName(linesLayer, shapeName);
        if (shapingLayer) removePageItemsByName(shapingLayer, shapeName);
    }
    if (patternsLayer) {
        try {
            patternsLayer.locked = false;
        } catch (ePatternLockInit) {}
        try {
            patternsLayer.visible = true;
        } catch (ePatternVisInit) {}
    }
    if (frontBodiceLayer) {
        try {
            frontBodiceLayer.locked = false;
        } catch (eFrontPatternLock) {}
        try {
            frontBodiceLayer.visible = true;
        } catch (eFrontPatternVis) {}
        clearLayerDeep(frontBodiceLayer);
    }
    if (backBodiceLayer) {
        try {
            backBodiceLayer.locked = false;
        } catch (eBackPatternLock) {}
        try {
            backBodiceLayer.visible = true;
        } catch (eBackPatternVis) {}
        clearLayerDeep(backBodiceLayer);
    }
    try {
        basicFrameLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
    } catch (eOrder1) {}
    try {
        labelsLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
    } catch (eOrder2) {}
    try {
        casualLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
    } catch (eOrderCasual) {}
    try {
        dartsLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
    } catch (eOrderDarts) {}
    if (dartsSubLayer) {
        try {
            dartsSubLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (eOrderDartsSub) {}
    }
    if (shapingLayer) {
        try {
            shapingLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (eOrderShaping) {}
    }
    if (patternsLayer) {
        try {
            patternsLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (eOrderPatterns) {}
    }
    var FRAME_COLOR = makeRGB(0, 0, 0);
    var NUMBER_FILL_COLOR = makeRGB(255, 255, 255);
    var LABEL_FONT_SIZE_PT = 13;
    var NUMBER_FONT_SIZE_PT = 9;
    var MARKER_RADIUS_CM = 0.25;
    var FRAME_LINE_LENGTH_CM = 40;
    var LABEL_RIGHT_OFFSET_CM = 5;
    var LABEL_VERTICAL_OFFSET_CM = 0.5;
    var LINE_LABEL_HORIZONTAL_SHIFT_CM = ptToCm(200);
    var CONNECTOR_BLUE_COLOR = makeRGB(0, 102, 204);
    var LAYER_COLOR = makeRGB(0, 102, 153);
    var HIP_LEFT_OFFSET_CM = 2;
    var DASH_PATTERN = [25, 12];
    var DART_TRIANGLE_COLOR = makeRGB(0, 0, 0);
    var LABEL_SMALL_OFFSET_CM = 0.5;
    var FRONT_ARM_LABEL_NORMAL_OFFSET_CM = -0.5;
    var MIN_HANDLE_LENGTH_CM = 0.5;
    var labelAlignmentRefs = {};
    var FRONT_BODICE_COPY_SUFFIX = '';

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
        var target = options.target || linesLayer;
        if (!target) return null;
        var path = target.pathItems.add();
        if (name) {
            try {
                path.name = name;
            } catch (eCurveNameAssign) {}
        }
        path.setEntirePath([startArt, endArt]);
        path.stroked = true;
        path.strokeWidth = 1;
        path.strokeColor = color || FRAME_COLOR;
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

    function rotatePoint(pt, pivot, angleRadians) {
        if (!pt || !pivot || isNaN(angleRadians)) return null;
        var cosA = Math.cos(angleRadians);
        var sinA = Math.sin(angleRadians);
        var dx = pt.x - pivot.x;
        var dy = pt.y - pivot.y;
        return {
            x: pivot.x + (dx * cosA) - (dy * sinA),
            y: pivot.y + (dx * sinA) + (dy * cosA)
        };
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
            case 'Back Side Line':
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

    function drawLineWithLabel(pA, pB, text, color) {
        var path = drawFrameLine(pA, pB, text, color);
        if (!text) return path;
        var labelPoint = computeLabelPoint(pA, pB, text);
        var tf = drawCenteredLabel(labelPoint, text);
        if (tf && isVerticalLine(pA, pB)) {
            rotateVerticalLabel(tf, toArt(labelPoint));
        }
        return path;
    }

    function collectPageItemsByName(container, goalName, results) {
        if (!container || !goalName) return;
        if (!results) results = [];
        var cleanGoal = goalName.toString();
        var items = null;
        try {
            items = container.pageItems;
        } catch (ePageItemsRead) {}
        if (items) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (!item) continue;
                var candidateName = '';
                try {
                    candidateName = item.name;
                } catch (eNameRead) {}
                if (candidateName === cleanGoal) results.push(item);
                if (item.typename === 'GroupItem') collectPageItemsByName(item, cleanGoal, results);
            }
        }
    }

    function copyItemsByNameToLayer(sourceLayer, itemName, targetLayer, renameSuffix) {
        if (!sourceLayer || !itemName || !targetLayer) return;
        var matches = [];
        collectPageItemsByName(sourceLayer, itemName, matches);
        if (!matches.length) return;
        for (var i = 0; i < matches.length; i++) {
            var original = matches[i];
            var duplicate = null;
            try {
                duplicate = original.duplicate(targetLayer, ElementPlacement.PLACEATBEGINNING);
            } catch (eDuplicate) {
                duplicate = null;
            }
            if (!duplicate) continue;
            if (renameSuffix) {
                try {
                    duplicate.name = itemName + renameSuffix;
                } catch (eRename) {}
            }
            try {
                duplicate.zOrder(ZOrderMethod.BRINGTOFRONT);
            } catch (eOrderDup) {}
        }
    }

    function copyFrontComponentsToPatternsLayer() {
        if (!frontBodiceLayer) return;
        var specs = [{
            layer: shapingLayer,
            name: 'Lower Front Side Curve'
        }, {
            layer: shapingLayer,
            name: 'Upper Front Side Line'
        }, {
            layer: dartsSubLayer,
            name: 'Front Shoulder Dart'
        }, {
            layer: casualLayer,
            name: 'Front Armhole Curve'
        }, {
            layer: casualLayer,
            name: 'Front Hem Line'
        }, {
            layer: casualLayer,
            name: 'Hip Line'
        }, {
            layer: basicFrameLayer,
            name: 'Waist Line'
        }, {
            layer: basicFrameLayer,
            name: 'Centre Front (CF)'
        }, {
            layer: basicFrameLayer,
            name: 'Front Dart Line'
        }, {
            layer: basicFrameLayer,
            name: 'Front Dart Line Upper'
        }, {
            layer: basicFrameLayer,
            name: 'Front Shoulder Line'
        }, {
            layer: basicFrameLayer,
            name: 'Bust Distance'
        }, {
            layer: basicFrameLayer,
            name: 'Front Neck Curve'
        }];
        for (var i = 0; i < specs.length; i++) {
            var spec = specs[i];
            copyItemsByNameToLayer(spec.layer, spec.name, frontBodiceLayer, FRONT_BODICE_COPY_SUFFIX);
        }
    }

    function trimBackDartLegAtShoulderLine() {
        if (!backBodiceLayer) return;
        var dartLegs = findPathItems(backBodiceLayer, 'Back Shoulder Dart Leg');
        var shoulderLines = findPathItems(backBodiceLayer, 'Back Shoulder Line (Original)');
        if (!dartLegs.length || !shoulderLines.length) return;
        var shoulder = shoulderLines[0];
        if (!shoulder || shoulder.typename !== 'PathItem') return;
        var shpPts = shoulder.pathPoints;
        if (!shpPts || shpPts.length < 2) return;
        var shoulderStart = shpPts[0].anchor;
        var shoulderEnd = shpPts[shpPts.length - 1].anchor;
        if (!shoulderStart || !shoulderEnd) return;
        for (var i = 0; i < dartLegs.length; i++) {
            var leg = dartLegs[i];
            if (!leg || leg.typename !== 'PathItem') continue;
            var pts = leg.pathPoints;
            if (!pts || pts.length < 2) continue;
            var legStart = pts[0].anchor;
            var legEnd = pts[pts.length - 1].anchor;
            if (!legStart || !legEnd) continue;
            var intersect = segmentIntersection(legStart, legEnd, shoulderStart, shoulderEnd);
            if (!intersect) continue;
            leg.setEntirePath([
                [legStart[0], legStart[1]],
                [intersect[0], intersect[1]]
            ]);
            var lPts = leg.pathPoints;
            if (lPts && lPts.length === 2) {
                for (var p = 0; p < lPts.length; p++) {
                    lPts[p].leftDirection = [lPts[p].anchor[0], lPts[p].anchor[1]];
                    lPts[p].rightDirection = [lPts[p].anchor[0], lPts[p].anchor[1]];
                    lPts[p].pointType = PointType.CORNER;
                }
            }
        }
    }

    function cutBackShoulderLineAtDartLeg() {
        if (!backBodiceLayer) return;
        var shoulderLines = findPathItems(backBodiceLayer, 'Back Shoulder Line (Original)');
        var dartLegs = findPathItems(backBodiceLayer, 'Back Shoulder Dart Leg');
        if (!shoulderLines.length || !dartLegs.length) return;
        var shoulder = shoulderLines[0];
        var dartLeg = dartLegs[0];
        if (!shoulder || shoulder.typename !== 'PathItem') return;
        if (!dartLeg || dartLeg.typename !== 'PathItem') return;
        var sPts = shoulder.pathPoints;
        var dPts = dartLeg.pathPoints;
        if (!sPts || sPts.length < 2 || !dPts || dPts.length < 2) return;
        var shoulderStart = sPts[0].anchor;
        var shoulderEnd = sPts[sPts.length - 1].anchor;
        var legStart = dPts[0].anchor;
        var legEnd = dPts[dPts.length - 1].anchor;
        if (!shoulderStart || !shoulderEnd || !legStart || !legEnd) return;

        function pointOnLine(pt, a, b) {
            var vx = b[0] - a[0];
            var vy = b[1] - a[1];
            var wx = pt[0] - a[0];
            var wy = pt[1] - a[1];
            var len2 = vx * vx + vy * vy;
            if (len2 < 1e-6) return false;
            var cross = Math.abs(vx * wy - vy * wx);
            var dist = cross / Math.sqrt(len2);
            return dist < 0.5;
        }

        var intersection = null;
        if (pointOnLine(legStart, shoulderStart, shoulderEnd)) intersection = legStart;
        else if (pointOnLine(legEnd, shoulderStart, shoulderEnd)) intersection = legEnd;
        if (!intersection) intersection = segmentIntersection(shoulderStart, shoulderEnd, legStart, legEnd);
        if (!intersection) return;

        var keepAnchor = null;
        var candidates = [];
        if (shoulderStart[0] > intersection[0]) candidates.push(shoulderStart);
        if (shoulderEnd[0] > intersection[0]) candidates.push(shoulderEnd);
        if (candidates.length) {
            keepAnchor = candidates[0];
            for (var c = 1; c < candidates.length; c++) {
                if (candidates[c][0] > keepAnchor[0]) keepAnchor = candidates[c];
            }
        } else {
            keepAnchor = (shoulderStart[0] > shoulderEnd[0]) ? shoulderStart : shoulderEnd;
        }

        shoulder.setEntirePath([
            [intersection[0], intersection[1]],
            [keepAnchor[0], keepAnchor[1]]
        ]);
        var sPtsNew = shoulder.pathPoints;
        if (sPtsNew && sPtsNew.length === 2) {
            for (var p = 0; p < sPtsNew.length; p++) {
                sPtsNew[p].leftDirection = [sPtsNew[p].anchor[0], sPtsNew[p].anchor[1]];
                sPtsNew[p].rightDirection = [sPtsNew[p].anchor[0], sPtsNew[p].anchor[1]];
                sPtsNew[p].pointType = PointType.CORNER;
            }
        }
    }

    function copyBackComponentsToPatternsLayer() {
        if (!backBodiceLayer) return;
        var specs = [{
            layer: shapingLayer,
            name: 'Upper Back Side Line'
        }, {
            layer: shapingLayer,
            name: 'Lower Back Side Curve'
        }, {
            layer: dartsSubLayer,
            name: 'Back Shoulder Dart'
        }, {
            layer: casualLayer,
            name: 'Back Armhole Curve (17-11)'
        }, {
            layer: casualLayer,
            name: 'New Centre Back (CB)'
        }, {
            layer: casualLayer,
            name: 'New Back Waist Line'
        }, {
            layer: casualLayer,
            name: 'New Back Hem Line'
        }, {
            layer: casualLayer,
            name: 'New Back Hip Line'
        }, {
            layer: basicFrameLayer,
            name: 'Bust Line'
        }, {
            layer: basicFrameLayer,
            name: 'Back Shoulder Dart Leg'
        }, {
            layer: basicFrameLayer,
            name: 'Shoulder Blade Line'
        }, {
            layer: basicFrameLayer,
            name: 'Back Neck Curve'
        }];
        for (var i = 0; i < specs.length; i++) {
            var spec = specs[i];
            copyItemsByNameToLayer(spec.layer, spec.name, backBodiceLayer, FRONT_BODICE_COPY_SUFFIX);
        }
    }

    function drawBackBodiceLine(name, startPoint, endPoint, dashed) {
        if (!backBodiceLayer || !startPoint || !endPoint) return null;
        removePageItemsByName(backBodiceLayer, name);
        var start = toArt(startPoint);
        var end = toArt(endPoint);
        if (!start || !end) return null;
        var path = backBodiceLayer.pathItems.add();
        path.name = name;
        path.stroked = true;
        path.strokeWidth = 1;
        path.strokeColor = FRAME_COLOR;
        path.filled = false;
        path.closed = false;
        path.setEntirePath([start, end]);
        if (dashed) {
            try {
                path.strokeDashes = DASH_PATTERN;
            } catch (eDashBack) {}
        }
        try {
            path.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (eBackOrder) {}
        return path;
    }

    function findPathItems(layer, name) {
        var matches = [];
        collectPageItemsByName(layer, name, matches);
        return matches;
    }

    function segmentIntersection(a1, a2, b1, b2) {
        var x1 = a1[0],
            y1 = a1[1],
            x2 = a2[0],
            y2 = a2[1];
        var x3 = b1[0],
            y3 = b1[1],
            x4 = b2[0],
            y4 = b2[1];
        var denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4);
        if (Math.abs(denom) < 1e-6) return null;
        var px = ((x1 * y2 - y1 * x2) * (x3 - x4) - (x1 - x2) * (x3 * y4 - y3 * x4)) / denom;
        var py = ((x1 * y2 - y1 * x2) * (y3 - y4) - (y1 - y2) * (x3 * y4 - y3 * x4)) / denom;
        function within(p, a, b) {
            return (p >= Math.min(a, b) - 1e-6) && (p <= Math.max(a, b) + 1e-6);
        }
        if (within(px, x1, x2) && within(py, y1, y2) && within(px, x3, x4) && within(py, y3, y4)) {
            return [px, py];
        }
        return null;
    }

    function trimBackShoulderDartLegs() {
        if (!backBodiceLayer) return;
        removePageItemsByName(backBodiceLayer, 'Back Shoulder Dart Left');
        removePageItemsByName(backBodiceLayer, 'Back Shoulder Dart Right');
        removePageItemsByName(backBodiceLayer, 'Back Shoulder Dart Cut 1');
        removePageItemsByName(backBodiceLayer, 'Back Shoulder Dart Cut 2');
        var darts = findPathItems(backBodiceLayer, 'Back Shoulder Dart');
        if (!darts || !darts.length) return;
        if (!point17a || !point34) return;
        var cutA = toArt(point17a);
        var cutB = toArt(point34);
        if (!cutA || !cutB) return;
        for (var i = 0; i < darts.length; i++) {
            var dart = darts[i];
            if (!dart || dart.typename !== 'PathItem') continue;
            var pts = dart.pathPoints;
            if (!pts || pts.length < 3) continue;
            var anchors = [];
            for (var p = 0; p < pts.length; p++) {
                var a = pts[p].anchor;
                if (!a) continue;
                anchors.push([a[0], a[1]]);
            }
            if (anchors.length < 3) continue;

            function nearestIndex(target) {
                var best = -1;
                var bestDist = null;
                for (var idx = 0; idx < anchors.length; idx++) {
                    var dx = anchors[idx][0] - target[0];
                    var dy = anchors[idx][1] - target[1];
                    var d2 = dx * dx + dy * dy;
                    if (bestDist === null || d2 < bestDist) {
                        bestDist = d2;
                        best = idx;
                    }
                }
                return best;
            }

            var idxA = nearestIndex(cutA);
            var idxB = nearestIndex(cutB);
            if (idxA === -1 || idxB === -1 || idxA === idxB) continue;

            var newPath = [];
            newPath.push(anchors[idxB]);
            var k = (idxB + 1) % anchors.length;
            while (k !== idxA) {
                newPath.push(anchors[k]);
                k = (k + 1) % anchors.length;
                if (newPath.length > anchors.length + 2) break;
            }
            newPath.push(anchors[idxA]);
            dart.closed = false;
            dart.setEntirePath(newPath);
            var dartPts = dart.pathPoints;
            if (dartPts) {
                for (var d = 0; d < dartPts.length; d++) {
                    dartPts[d].leftDirection = [dartPts[d].anchor[0], dartPts[d].anchor[1]];
                    dartPts[d].rightDirection = [dartPts[d].anchor[0], dartPts[d].anchor[1]];
                    dartPts[d].pointType = PointType.CORNER;
                }
            }
        }
    }

    function trimShoulderBladeLineAtCB() {
        if (!backBodiceLayer) return;
        var cbPaths = findPathItems(backBodiceLayer, 'New Centre Back (CB)');
        var bladePaths = findPathItems(backBodiceLayer, 'Shoulder Blade Line');
        if (!cbPaths.length || !bladePaths.length) return;
        var cb = cbPaths[0];
        if (!cb || cb.typename !== 'PathItem' || cb.pathPoints.length < 2) return;
        var cbSegs = [];
        var cbPts = cb.pathPoints;
        for (var c = 0; c < cbPts.length - 1; c++) {
            var ca = cbPts[c].anchor;
            var cbn = cbPts[c + 1].anchor;
            if (ca && cbn) cbSegs.push([ca, cbn]);
        }
        if (cb.closed && cbPts.length > 2) {
            var ca0 = cbPts[cbPts.length - 1].anchor;
            var cb0 = cbPts[0].anchor;
            if (ca0 && cb0) cbSegs.push([ca0, cb0]);
        }
        if (!cbSegs.length) return;
        for (var i = 0; i < bladePaths.length; i++) {
            var blade = bladePaths[i];
            if (!blade || blade.typename !== 'PathItem') continue;
            var pts = blade.pathPoints;
            if (!pts || pts.length < 2) continue;
            var hit = null;
            var segStart = null;
            var segEnd = null;
            for (var p = 0; p < pts.length - 1 && !hit; p++) {
                var a = pts[p].anchor;
                var b = pts[p + 1].anchor;
                if (!a || !b) continue;
                segStart = a;
                segEnd = b;
                for (var s = 0; s < cbSegs.length && !hit; s++) {
                    hit = segmentIntersection(a, b, cbSegs[s][0], cbSegs[s][1]);
                }
            }
            if (!hit || !segStart || !segEnd) continue;
            var keepStart = segStart;
            if (segEnd[0] < keepStart[0]) keepStart = segEnd;
            blade.setEntirePath([
                [keepStart[0], keepStart[1]],
                [hit[0], hit[1]]
            ]);
            var bladePts = blade.pathPoints;
            if (bladePts && bladePts.length) {
                for (var bpi = 0; bpi < bladePts.length; bpi++) {
                    bladePts[bpi].leftDirection = [bladePts[bpi].anchor[0], bladePts[bpi].anchor[1]];
                    bladePts[bpi].rightDirection = [bladePts[bpi].anchor[0], bladePts[bpi].anchor[1]];
                    bladePts[bpi].pointType = PointType.CORNER;
                }
            }
        }
    }

    function extendShoulderBladeLineToDart() {
        if (!backBodiceLayer) return;
        var bladePaths = findPathItems(backBodiceLayer, 'Shoulder Blade Line');
        var dartPaths = findPathItems(backBodiceLayer, 'Back Shoulder Dart');
        if (!dartPaths.length && dartsSubLayer) dartPaths = findPathItems(dartsSubLayer, 'Back Shoulder Dart');
        if (!bladePaths.length || !dartPaths.length) return;
        var blade = bladePaths[0];
        if (!blade || blade.typename !== 'PathItem') return;
        var pts = blade.pathPoints;
        if (!pts || pts.length < 2) return;
        var startAnchor = pts[0].anchor;
        var endAnchor = pts[pts.length - 1].anchor;
        if (!startAnchor || !endAnchor) return;
        var lineStart = [startAnchor[0] - 10000, startAnchor[1]];
        var lineEnd = [startAnchor[0], startAnchor[1]];
        var best = null;
        for (var d = 0; d < dartPaths.length; d++) {
            var dart = dartPaths[d];
            if (!dart || dart.typename !== 'PathItem') continue;
            var dPts = dart.pathPoints;
            if (!dPts || dPts.length < 2) continue;
            var segCount = dart.closed ? dPts.length : dPts.length - 1;
            for (var s = 0; s < segCount; s++) {
                var a = dPts[s].anchor;
                var b = dPts[(s + 1) % dPts.length].anchor;
                if (!a || !b) continue;
                var hit = segmentIntersection(lineStart, lineEnd, a, b);
                if (!hit) continue;
                if (hit[0] <= startAnchor[0]) {
                    if (!best || hit[0] > best[0]) {
                        best = hit;
                    }
                }
            }
        }
        if (!best) return;
        pts[0].anchor = [best[0], best[1]];
        pts[0].leftDirection = [best[0], best[1]];
        pts[0].rightDirection = [best[0], best[1]];
        pts[0].pointType = PointType.CORNER;
        blade.setEntirePath([pts[0].anchor, endAnchor]);
        var updatedPts = blade.pathPoints;
        if (updatedPts && updatedPts.length) {
            var first = updatedPts[0];
            first.leftDirection = [first.anchor[0], first.anchor[1]];
            first.rightDirection = [first.anchor[0], first.anchor[1]];
            first.pointType = PointType.CORNER;
            var last = updatedPts[updatedPts.length - 1];
            last.leftDirection = [last.anchor[0], last.anchor[1]];
            last.rightDirection = [last.anchor[0], last.anchor[1]];
        }
    }

    function removeBackBustLine() {
        if (!backBodiceLayer) return;
        removePageItemsByName(backBodiceLayer, 'Bust Line');
    }

    function trimBackWaistLineAtSide() {
        if (!backBodiceLayer || !point7) return;
        var waistPaths = findPathItems(backBodiceLayer, 'New Back Waist Line');
        var sidePaths = findPathItems(backBodiceLayer, 'Lower Back Side Curve');
        if (!waistPaths.length || !sidePaths.length) return;
        var waistY = null;
        var firstWaist = waistPaths[0];
        if (firstWaist && firstWaist.typename === 'PathItem' && firstWaist.pathPoints.length) {
            var a = firstWaist.pathPoints[0].anchor;
            if (a) waistY = a[1];
        }
        if (waistY === null) return;
        var intersection = null;
        for (var s = 0; s < sidePaths.length && !intersection; s++) {
            var curve = sidePaths[s];
            if (!curve || curve.typename !== 'PathItem') continue;
            var pts = curve.pathPoints;
            if (!pts || pts.length < 2) continue;
            for (var i = 0; i < pts.length - 1; i++) {
                var pA = pts[i].anchor;
                var pB = pts[i + 1].anchor;
                if (!pA || !pB) continue;
                if ((waistY >= pA[1] && waistY <= pB[1]) || (waistY >= pB[1] && waistY <= pA[1])) {
                    var t = (Math.abs(pB[1] - pA[1]) < 1e-6) ? 0 : (waistY - pA[1]) / (pB[1] - pA[1]);
                    if (t < 0 || t > 1) continue;
                    var x = pA[0] + t * (pB[0] - pA[0]);
                    intersection = [x, waistY];
                    break;
                }
            }
        }
        if (!intersection) return;
        for (var w = 0; w < waistPaths.length; w++) {
            var waist = waistPaths[w];
            if (!waist || waist.typename !== 'PathItem') continue;
            var ptsWaist = waist.pathPoints;
            if (!ptsWaist || ptsWaist.length < 2) continue;
            var newPath = [];
            newPath.push([intersection[0], intersection[1]]);
            for (var p = 0; p < ptsWaist.length; p++) {
                var aPt = ptsWaist[p].anchor;
                if (!aPt) continue;
                if (aPt[0] >= intersection[0]) newPath.push([aPt[0], aPt[1]]);
            }
            if (newPath.length >= 2) {
                waist.setEntirePath(newPath);
                try {
                    waist.strokeDashes = DASH_PATTERN;
                } catch (eDashWaistBack) {}
                var waistPts = waist.pathPoints;
                if (waistPts && waistPts.length) {
                    for (var idx = 0; idx < waistPts.length; idx++) {
                        waistPts[idx].leftDirection = [waistPts[idx].anchor[0], waistPts[idx].anchor[1]];
                        waistPts[idx].rightDirection = [waistPts[idx].anchor[0], waistPts[idx].anchor[1]];
                        waistPts[idx].pointType = PointType.CORNER;
                    }
                }
            }
        }
    }

    function connectBackSideToHem() {
        if (!backBodiceLayer) return;
        var sidePaths = findPathItems(backBodiceLayer, 'Lower Back Side Curve');
        var hemPaths = findPathItems(backBodiceLayer, 'New Back Hem Line');
        // Fallback: look in source layers if copies are missing.
        if (!sidePaths.length && shapingLayer) sidePaths = findPathItems(shapingLayer, 'Lower Back Side Curve');
        if (!hemPaths.length && casualLayer) hemPaths = findPathItems(casualLayer, 'New Back Hem Line');
        if (!sidePaths.length || !hemPaths.length) return;
        var sideEnd = null;
        var hemStart = null;
        // lowest point on side curve
        var side = sidePaths[0];
        if (side && side.typename === 'PathItem') {
            var pts = side.pathPoints;
            for (var i = 0; i < pts.length; i++) {
                var a = pts[i].anchor;
                if (!a) continue;
                if (!sideEnd || a[1] < sideEnd[1]) sideEnd = [a[0], a[1]];
            }
        }
        var hem = hemPaths[0];
        if (hem && hem.typename === 'PathItem') {
            var hpts = hem.pathPoints;
            for (var h = 0; h < hpts.length; h++) {
                var ha = hpts[h].anchor;
                if (!ha) continue;
                if (!hemStart || ha[0] < hemStart[0]) hemStart = [ha[0], ha[1]];
            }
        }
        var sideModel = artToModel(sideEnd);
        var hemModel = artToModel(hemStart);
        if (!sideModel || !hemModel) return;
        var connectorA = drawBackBodiceLine('Back Side Hem Connector', sideModel, hemModel, false);
        var connectorB = drawBackBodiceLine('Lower Back Side Connector', sideModel, hemModel, false);
        if (connectorA) {
            try {
                connectorA.zOrder(ZOrderMethod.BRINGTOFRONT);
            } catch (eOrderA) {}
        }
        if (connectorB) {
            try {
                connectorB.zOrder(ZOrderMethod.BRINGTOFRONT);
            } catch (eOrderB) {}
        }
    }

    function resetBackShoulderFromFrame() {
        if (!backBodiceLayer || !basicFrameLayer) return;
        var names = ['Back Shoulder Dart', 'Back Shoulder Line (Original)', 'Back Shoulder Dart Leg'];
        for (var i = 0; i < names.length; i++) {
            removePageItemsByName(backBodiceLayer, names[i]);
        }
        // Copy geometry from the frame for line/leg and from the darts layer for the dart path.
        copyItemsByNameToLayer(basicFrameLayer, 'Back Shoulder Line (Original)', backBodiceLayer, null);
        copyItemsByNameToLayer(basicFrameLayer, 'Back Shoulder Dart Leg', backBodiceLayer, null);
        if (dartsSubLayer) copyItemsByNameToLayer(dartsSubLayer, 'Back Shoulder Dart', backBodiceLayer, null);
    }

    function dashBackHipLine() {
        if (!backBodiceLayer) return;
        var hipPaths = findPathItems(backBodiceLayer, 'New Back Hip Line');
        for (var i = 0; i < hipPaths.length; i++) {
            var hip = hipPaths[i];
            if (!hip || hip.typename !== 'PathItem') continue;
            try {
                hip.strokeDashes = DASH_PATTERN;
            } catch (eHipDash) {}
        }
    }


    function connectFrontHemToFrontSide() {
        if (!frontBodiceLayer) return;
        var hemMatches = [];
        var lowerSideMatches = [];
        collectPageItemsByName(frontBodiceLayer, 'Front Hem Line' + FRONT_BODICE_COPY_SUFFIX, hemMatches);
        collectPageItemsByName(frontBodiceLayer, 'Lower Front Side Curve' + FRONT_BODICE_COPY_SUFFIX, lowerSideMatches);
        if (!hemMatches.length || !lowerSideMatches.length) return;
        var hemEnd = null;
        for (var i = 0; i < hemMatches.length && !hemEnd; i++) {
            var hemItem = hemMatches[i];
            if (!hemItem || hemItem.typename !== 'PathItem') continue;
            var hemPts = hemItem.pathPoints;
            if (!hemPts || hemPts.length === 0) continue;
            var rightmost = null;
            for (var h = 0; h < hemPts.length; h++) {
                var anchor = hemPts[h].anchor;
                if (!anchor) continue;
                if (!rightmost || anchor[0] > rightmost[0]) rightmost = [anchor[0], anchor[1]];
            }
            if (rightmost) hemEnd = rightmost;
        }
        if (!hemEnd) return;
        var sideEnd = null;
        for (var j = 0; j < lowerSideMatches.length && !sideEnd; j++) {
            var sideItem = lowerSideMatches[j];
            if (!sideItem || sideItem.typename !== 'PathItem') continue;
            var sidePts = sideItem.pathPoints;
            if (!sidePts || sidePts.length === 0) continue;
            var lowest = null;
            for (var s = 0; s < sidePts.length; s++) {
                var anchorSide = sidePts[s].anchor;
                if (!anchorSide) continue;
                if (!lowest || anchorSide[1] < lowest[1]) lowest = [anchorSide[0], anchorSide[1]];
            }
            if (lowest) sideEnd = lowest;
        }
        if (!sideEnd) return;
        if (Math.abs(hemEnd[0] - sideEnd[0]) < 0.01 && Math.abs(hemEnd[1] - sideEnd[1]) < 0.01) return;
        var connector = frontBodiceLayer.pathItems.add();
        connector.name = 'Front Side Hem Connector';
        connector.stroked = true;
        connector.strokeWidth = 1;
        connector.strokeColor = FRAME_COLOR;
        connector.filled = false;
        connector.closed = false;
        connector.setEntirePath([hemEnd, sideEnd]);
        try {
            connector.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (eConnectorOrder) {}
    }

    function drawFrontBodiceCentreFrontSegment(startPoint, endPoint) {
        if (!frontBodiceLayer) return;
        removePageItemsByName(frontBodiceLayer, 'Centre Front (CF)' + FRONT_BODICE_COPY_SUFFIX);
        if (!startPoint || !endPoint) return;
        var start = toArt(startPoint);
        var end = toArt(endPoint);
        if (!start || !end) return;
        var segment = frontBodiceLayer.pathItems.add();
        segment.name = 'Centre Front (CF)' + FRONT_BODICE_COPY_SUFFIX;
        segment.stroked = true;
        segment.strokeWidth = 1;
        segment.strokeColor = FRAME_COLOR;
        segment.filled = false;
        segment.closed = false;
        segment.setEntirePath([start, end]);
        try {
            segment.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (eSegmentOrder) {}
    }

    function drawFrontBodiceDashedLine(name, startPoint, endPoint) {
        if (!frontBodiceLayer || !startPoint || !endPoint) return;
        // Clean up legacy name used previously.
        removePageItemsByName(frontBodiceLayer, 'Front Bodice 19-29');
        removePageItemsByName(frontBodiceLayer, name);
        var start = toArt(startPoint);
        var end = toArt(endPoint);
        if (!start || !end) return;
        var path = frontBodiceLayer.pathItems.add();
        path.name = name;
        path.stroked = true;
        path.strokeWidth = 1;
        path.strokeColor = FRAME_COLOR;
        path.filled = false;
        path.closed = false;
        path.setEntirePath([start, end]);
        try {
            path.strokeDashes = DASH_PATTERN;
        } catch (eDashAssign) {}
        try {
            path.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (eFrontOrder) {}
    }

    function trimFrontBodiceWaistAtSide() {
        if (!frontBodiceLayer || !point7) return;
        var waistMatches = [];
        var sideMatches = [];
        collectPageItemsByName(frontBodiceLayer, 'Waist Line' + FRONT_BODICE_COPY_SUFFIX, waistMatches);
        collectPageItemsByName(frontBodiceLayer, 'Lower Front Side Curve' + FRONT_BODICE_COPY_SUFFIX, sideMatches);
        if (!waistMatches.length || !sideMatches.length) return;
        var waistAnchor = toArt(point7);
        if (!waistAnchor) return;
        var waistY = waistAnchor[1];
        var intersection = null;
        for (var s = 0; s < sideMatches.length && !intersection; s++) {
            var curve = sideMatches[s];
            if (!curve || curve.typename !== 'PathItem') continue;
            var pts = curve.pathPoints;
            if (!pts || pts.length < 2) continue;
            for (var i = 0; i < pts.length - 1; i++) {
                var a = pts[i].anchor;
                var b = pts[i + 1].anchor;
                if (!a || !b) continue;
                var ya = a[1];
                var yb = b[1];
                if ((waistY >= ya && waistY <= yb) || (waistY >= yb && waistY <= ya)) {
                    var t = (Math.abs(yb - ya) < 1e-6) ? 0 : (waistY - ya) / (yb - ya);
                    if (t < 0 || t > 1) continue;
                    var x = a[0] + t * (b[0] - a[0]);
                    intersection = [x, waistY];
                    break;
                }
            }
        }
        if (!intersection) return;
        for (var w = 0; w < waistMatches.length; w++) {
            var waistItem = waistMatches[w];
            if (!waistItem || waistItem.typename !== 'PathItem') continue;
            var pts = waistItem.pathPoints;
            if (!pts || pts.length < 2) continue;
            var leftMost = pts[0].anchor ? [pts[0].anchor[0], pts[0].anchor[1]] : null;
            for (var p = 1; p < pts.length; p++) {
                var anchor = pts[p].anchor;
                if (!anchor) continue;
                if (!leftMost || anchor[0] < leftMost[0]) leftMost = [anchor[0], anchor[1]];
            }
            var newPath = leftMost ? [leftMost, [intersection[0], intersection[1]]] : [];
            if (newPath.length >= 2) {
                waistItem.setEntirePath(newPath);
                var waistPts = waistItem.pathPoints;
                if (waistPts && waistPts.length) {
                    for (var pAdj = 0; pAdj < waistPts.length; pAdj++) {
                        waistPts[pAdj].leftDirection = [waistPts[pAdj].anchor[0], waistPts[pAdj].anchor[1]];
                        waistPts[pAdj].rightDirection = [waistPts[pAdj].anchor[0], waistPts[pAdj].anchor[1]];
                        waistPts[pAdj].pointType = PointType.CORNER;
                    }
                }
            }
        }
    }

    function removeFrontBodiceShoulderDartRightLeg() {
        if (!frontBodiceLayer) return;
        var dartName = 'Front Shoulder Dart' + FRONT_BODICE_COPY_SUFFIX;
        var matches = [];
        collectPageItemsByName(frontBodiceLayer, dartName, matches);
        if (!matches.length) return;
        for (var i = 0; i < matches.length; i++) {
            var pathItem = matches[i];
            if (!pathItem || pathItem.typename !== 'PathItem') continue;
            var pts = pathItem.pathPoints;
            if (!pts || pts.length < 3) continue;
            var leftShoulder = pts[0].anchor;
            var apex = pts[1].anchor;
            var rightShoulder = pts[2].anchor;
            if (!leftShoulder || !apex || !rightShoulder) continue;
            try {
                pathItem.remove();
            } catch (eRemoveDart) {}
            var leftLeg = frontBodiceLayer.pathItems.add();
            leftLeg.name = 'Front Shoulder Dart Left' + FRONT_BODICE_COPY_SUFFIX;
            leftLeg.stroked = true;
            leftLeg.strokeWidth = 1;
            leftLeg.strokeColor = FRAME_COLOR;
            leftLeg.filled = false;
            leftLeg.closed = false;
            leftLeg.setEntirePath([
                [leftShoulder[0], leftShoulder[1]],
                [apex[0], apex[1]]
            ]);
            var topLine = frontBodiceLayer.pathItems.add();
            topLine.name = 'Front Shoulder Dart Top' + FRONT_BODICE_COPY_SUFFIX;
            topLine.stroked = true;
            topLine.strokeWidth = 1;
            topLine.strokeColor = FRAME_COLOR;
            topLine.filled = false;
            topLine.closed = false;
            topLine.setEntirePath([
                [leftShoulder[0], leftShoulder[1]],
                [rightShoulder[0], rightShoulder[1]]
            ]);
        }
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
        try {
            labelsLayer.locked = false;
        } catch (eMarkerLabelsLock) {}
        try {
            labelsLayer.visible = true;
        } catch (eMarkerLabelsVis) {}
        if (markersGroup) {
            try {
                markersGroup.locked = false;
            } catch (eMarkerGroupLock) {}
        }
        if (numbersGroup) {
            try {
                numbersGroup.locked = false;
            } catch (eNumbersGroupLock) {}
        }
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
    var backShoulderEndOriginal = null;
    var backShoulderEndRotated = null;
    var backShoulderDartRotation = 0;
    var point16 = null;
    var point16Original = null;
    var point16Rotated = null;
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
        x: point11.x - 10,
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
        backShoulderEndOriginal = {
            x: backShoulderEnd.x,
            y: backShoulderEnd.y
        };
        if (tIntersect !== null) {
            point16 = {
                x: point10.x,
                y: point1a.y + dirY * tIntersect
            };
            point16Original = {
                x: point16.x,
                y: point16.y
            };
        }
        var backShoulderLineOriginal = drawFrameLine([point1a.x, point1a.y], [backShoulderEnd.x, backShoulderEnd.y], 'Back Shoulder Line (Original)');
        if (backShoulderLineOriginal) {
            try {
                backShoulderLineOriginal.strokeDashes = [];
            } catch (ebackOrig) {}
        }
    }
    if (point16) {
        placeFrameMarker([point16.x, point16.y], '16');
    }
    var point17 = null;
    var point17Guide = null;
    var point17GuideRotated = null;
    var point17a = null;
    var point17aOriginal = null;
    var point17b = null;
    var point17c = null;
    var point18 = null;
    var point34 = null;
    var point35 = null;
    var point35Original = null;
    var point35Rotated = null;
    var shoulderBladeEnd = null;
    var point32 = null;
    var point32Top = null;
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
        shoulderBladeEnd = {
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
        point17Guide = {
            x: midpoint1710.x - 1.5,
            y: midpoint1710.y
        };
        var halfBackShoulderDart = backShoulderDartIntake / 2;
        if (!isFinite(halfBackShoulderDart) || halfBackShoulderDart <= 0) halfBackShoulderDart = defaults.BackShoulderDartIntake / 2;
        point17a = {
            x: point17.x,
            y: point17.y - halfBackShoulderDart
        };
        point17aOriginal = {
            x: point17a.x,
            y: point17a.y
        };
        point17c = {
            x: point17.x,
            y: point17.y + halfBackShoulderDart
        };
        placeFrameMarker([point17a.x, point17a.y], '17a');
        placeFrameMarker([point17c.x, point17c.y], '17c');
        var dartGuideLine = drawFrameLine([point17a.x, point17a.y], [point17c.x, point17c.y], 'Armhole Dart Cap');
        if (dartGuideLine) {
            try {
                dartGuideLine.strokeDashes = DASH_PATTERN;
            } catch (eDartGuide) {}
        }
        point34 = {
            x: (point17.x + shoulderBladeEnd.x) / 2,
            y: (point17.y + shoulderBladeEnd.y) / 2
        };
        placeFrameMarker([point34.x, point34.y], '34');
        var line17a34 = drawFrameLine([point17a.x, point17a.y], [point34.x, point34.y], 'Armhole Dart Leg (17a-34)');
        if (line17a34) {
            try {
                line17a34.strokeDashes = [];
            } catch (eSolid17a34) {}
        }
        var line17c34 = drawFrameLine([point17c.x, point17c.y], [point34.x, point34.y], 'Armhole Dart Leg (17c-34)');
        if (line17c34) {
            try {
                line17c34.strokeDashes = [];
            } catch (eSolid17c34) {}
        }

        function intersectParallelToBack(origin, direction, segStart, segEnd) {
            var segDir = {
                x: segEnd.x - segStart.x,
                y: segEnd.y - segStart.y
            };
            var denom = (direction.x * segDir.y) - (direction.y * segDir.x);
            if (Math.abs(denom) < 1e-6) return null;
            var diff = {
                x: segStart.x - origin.x,
                y: segStart.y - origin.y
            };
            var t = (diff.x * segDir.y - diff.y * segDir.x) / denom;
            var u = (diff.x * direction.y - diff.y * direction.x) / denom;
            if (u < -0.01 || u > 1.01) return null;
            return {
                x: origin.x + direction.x * t,
                y: origin.y + direction.y * t
            };
        }

        if (point2 && point8 && point1a && backShoulderEnd) {
            var dir28 = {
                x: point8.x - point2.x,
                y: point8.y - point2.y
            };
            var dir28Len = Math.sqrt((dir28.x * dir28.x) + (dir28.y * dir28.y));
            if (dir28Len < 1e-6) {
                dir28 = {
                    x: 0,
                    y: 1
                };
            } else {
                dir28.x /= dir28Len;
                dir28.y /= dir28Len;
            }
            var intersection = intersectParallelToBack(point34, dir28, point1a, backShoulderEnd);
            if (!intersection) {
                intersection = intersectParallelToBack(point34, {
                    x: -dir28.x,
                    y: -dir28.y
                }, point1a, backShoulderEnd);
            }
            if (intersection) {
                point35 = intersection;
                point35Original = {
                    x: point35.x,
                    y: point35.y
                };
                placeFrameMarker([point35.x, point35.y], '35');
        var line3435 = drawFrameLine([point34.x, point34.y], [point35.x, point35.y], 'Back Shoulder Dart Leg');
        if (line3435) {
            try {
                line3435.strokeDashes = [];
            } catch (eSolid3435) {}
        }
            }
        }

        if (point17a && point17c && point34) {
            var angleStart = Math.atan2(point17aOriginal.y - point34.y, point17aOriginal.x - point34.x);
            var angleEnd = Math.atan2(point17c.y - point34.y, point17c.x - point34.x);
            if (isFinite(angleStart) && isFinite(angleEnd)) {
                backShoulderDartRotation = angleEnd - angleStart;
                while (backShoulderDartRotation > Math.PI) backShoulderDartRotation -= Math.PI * 2;
                while (backShoulderDartRotation < -Math.PI) backShoulderDartRotation += Math.PI * 2;
            } else {
                backShoulderDartRotation = 0;
            }
        }

        function rotateAroundPivot(pt) {
            if (!pt) return pt;
            if (!point34 || !isFinite(backShoulderDartRotation) || Math.abs(backShoulderDartRotation) < 1e-6) {
                return {
                    x: pt.x,
                    y: pt.y
                };
            }
            var rotated = rotatePoint(pt, point34, backShoulderDartRotation);
            if (rotated) return rotated;
            return {
                x: pt.x,
                y: pt.y
            };
        }

        var rotated17a = rotateAroundPivot(point17aOriginal);
        if (rotated17a) {
            point17b = {
                x: rotated17a.x,
                y: rotated17a.y
            };
        } else {
            point17b = {
                x: point17aOriginal.x,
                y: point17aOriginal.y
            };
        }
        placeFrameMarker([point17b.x, point17b.y], '17b');

        point17GuideRotated = rotateAroundPivot(point17Guide);
        if (point17GuideRotated) point17Guide = point17GuideRotated;

        point35Rotated = rotateAroundPivot(point35);
        if (point35Rotated) point35 = point35Rotated;

        point16Rotated = rotateAroundPivot(point16);
        if (point16Rotated) point16 = point16Rotated;

        backShoulderEndRotated = rotateAroundPivot(backShoulderEnd);
        if (backShoulderEndRotated) backShoulderEnd = backShoulderEndRotated;

        if (dartsGroup && point17b && point34 && point35 && backShoulderEnd && point16) {
            var rotatedWedge = dartsGroup.pathItems.add();
            rotatedWedge.name = 'Back Shoulder Dart';
            rotatedWedge.stroked = true;
            rotatedWedge.strokeWidth = 1;
            rotatedWedge.strokeColor = DART_TRIANGLE_COLOR;
            rotatedWedge.filled = false;
            rotatedWedge.closed = true;
            try {
                rotatedWedge.strokeDashes = [];
            } catch (eRotDash) {}
            rotatedWedge.setEntirePath([toArt(point17b), toArt(point34), toArt(point35), toArt(point16), toArt(backShoulderEnd), toArt(point17b)]);
        }

        if (point17Guide) {
            point18 = {
                x: point13.x,
                y: point17Guide.y
            };
        } else {
            point18 = {
                x: point13.x,
                y: point17 ? point17.y : midpoint1710.y
            };
        }
        var line17a18 = null;
        if (point17Guide && point18) {
            line17a18 = drawFrameLine([point17Guide.x, point17Guide.y], [point18.x, point18.y], '17 Guide - 18');
        }
        if (line17a18) {
            try {
                line17a18.strokeDashes = DASH_PATTERN;
            } catch (eDash17a18) {}
        }
        placeFrameMarker([point18.x, point18.y], '18');
        if (point18) {
            var BrCValue = finalResults['BrC'];
            if (isNaN(BrCValue) && rowRefs && rowRefs['BrC'] && rowRefs['BrC'].measField) {
                BrCValue = parseField(rowRefs['BrC'].measField, defaults.BrC);
            }
            if (isNaN(BrCValue)) BrCValue = defaults.BrC;
            if (!isNaN(BrCValue)) {
                var point32Offset = (BrCValue / 20) + 1;
                point32 = {
                    x: point18.x + point32Offset,
                    y: point18.y
                };
                placeFrameMarker([point32.x, point32.y], '32');
                point32Top = {
                    x: point32.x,
                    y: point32.y - 10
                };
                var line32Guide = drawFrameLine([point32.x, point32.y], [point32Top.x, point32Top.y], 'Shoulder Dart Guide Line');
                if (line32Guide) {
                    try {
                        line32Guide.strokeDashes = DASH_PATTERN;
                    } catch (eLine32) {}
                }
            }
        }
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
        line1125 = drawFrameLine([point11.x, point11.y], [point25.x, point25.y], 'Provisional Back Side Line', FRAME_COLOR, casualLayer);
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
        measurementResults.HiG = hiG;
        easeResults.HiG = 0;
        finalResults.HiG = hiG;
    }
    if (hipShortage !== null) {
        measurementResults.HipShortage = hipShortage;
        easeResults.HipShortage = 0;
        finalResults.HipShortage = hipShortage;
    }
    if (latestDerived) {
        latestDerived.hiG = hiG;
        latestDerived.hipShortage = hipShortage;
        latestDerived.hipSpanBack = hipSpanBack;
        latestDerived.hipSpanFront = hipSpanFront;
        latestDerived.backShoulderDartIntake = backShoulderDartIntake;
        latestDerived.backShoulderDartRotation = backShoulderDartRotation;
        latestDerived.backShoulderEndOriginal = backShoulderEndOriginal;
        latestDerived.backShoulderEndRotated = backShoulderEndRotated;
        latestDerived.waistShaping = waistShaping;
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
    var point29Waist = null;
    var point30Waist = null;
    var hipLinePoint6a = null;
    var hemConnectorEnd = null;
    var waistConnectorEnd = null;
    var point36 = null;
    var point36Guide = null;
    var point37 = null;
    var point37Guide = null;
    var point38 = null;
    var point39 = null;
    var point40 = null;
    var point41 = null;
    var point42 = null;

    if (point30) {
        var point30Hem = extendLineToY(point11, point30, hemLineY);
        var line1130 = null;
        if (point30Hem) {
            line1130 = drawFrameLine([point11.x, point11.y], [point30Hem.x, point30Hem.y], 'New Back Side Line', FRAME_COLOR, casualLayer);
            if (line1130) {
                try {
                    line1130.strokeDashes = [];
                } catch (eSolid1130) {}
            }
        }

        point30Waist = (typeof waistLineY === 'number') ? extendLineToY(point11, point30, waistLineY) : null;

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
            var line306a = drawFrameLine([point30.x, point30.y], [hipLinePoint6a.x, hipLinePoint6a.y], 'New Back Hip Line', FRAME_COLOR, casualLayer);
            if (line306a) {
                try {
                    line306a.strokeDashes = [];
                } catch (eSolid306a) {}
            }
        }

        if (point30Hem) {
            hemConnectorEnd = perpendicularFootOnBackDiagonal(point30Hem);
            if (hemConnectorEnd) {
                var hemConnector = drawFrameLine([point30Hem.x, point30Hem.y], [hemConnectorEnd.x, hemConnectorEnd.y], 'New Back Hem Line', FRAME_COLOR, casualLayer);
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
                var waistConnector = drawFrameLine([point30Waist.x, point30Waist.y], [waistConnectorEnd.x, waistConnectorEnd.y], 'New Back Waist Line', FRAME_COLOR, casualLayer);
                if (waistConnector) {
                    try {
                        waistConnector.strokeDashes = [];
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
        drawFrameLine([point2.x, point2.y], [point8.x, point8.y], 'New Centre Back (CB)', FRAME_COLOR, casualLayer);
    }

    var point29Hem = null;
    if (point29) {
        point29Hem = extendLineToY(point12, point29, hemLineY);
        point29Waist = (typeof waistLineY === 'number') ? extendLineToY(point12, point29, waistLineY) : null;
        var newFrontSideLine = drawFrameLine([point12.x, point12.y], [point29Hem.x, point29Hem.y], 'New Front Side Line', FRAME_COLOR, casualLayer);
        if (newFrontSideLine) {
            try {
                newFrontSideLine.strokeDashes = [];
            } catch (eSolid1229) {}
        }
    }

    if (point30Waist) {
        point36 = {
            x: point30Waist.x,
            y: point30Waist.y - 1
        };
        placeFrameMarker([point36.x, point36.y], '36');
        point36Guide = {
            x: point36.x + 3,
            y: point36.y
        };
        var line36Guide = drawFrameLine([point36.x, point36.y], [point36Guide.x, point36Guide.y], '36 Guide', FRAME_COLOR, shapingTarget);
        if (line36Guide) {
            try {
                line36Guide.strokeDashes = [];
            } catch (eLine36) {}
        }
        point38 = {
            x: point36.x + waistShaping,
            y: point36.y
        };
        placeFrameMarker([point38.x, point38.y], '38');
        if (shapingTarget) {
            drawFrameLine([point38.x, point38.y], [point38.x + 1, point38.y], 'Point38 Tick (Front)', FRAME_COLOR, shapingTarget);
        }
        if (backBodiceLayer) {
            drawBackBodiceLine('Back Bodice Tick 38', point38, {
                x: point38.x + 1,
                y: point38.y
            }, false);
        }
    if (point30) {
        var waistShapingMagnitude = waistShaping;
        if (isNaN(waistShapingMagnitude)) waistShapingMagnitude = defaults.WaistShaping;
        var waistToHipFraction = 2 / 3;
        var point38CurveEnd = {
            x: point30Waist.x + (point30.x - point30Waist.x) * waistToHipFraction,
                y: point30Waist.y + (point30.y - point30Waist.y) * waistToHipFraction
            };
            var curveName38 = 'Lower Back Side Curve';
            if (linesLayer) {
                try {
                    var existingCurves38 = linesLayer.pathItems;
                    for (var idx38 = existingCurves38.length - 1; idx38 >= 0; idx38--) {
                        var candidateCurve38 = existingCurves38[idx38];
                        if (!candidateCurve38) continue;
                        var candidateName38 = '';
                        try {
                            candidateName38 = candidateCurve38.name;
                        } catch (eCurve38Name) {}
                        if (candidateName38 === curveName38) {
                            try {
                                candidateCurve38.remove();
                            } catch (eCurve38Remove) {}
                        }
                    }
                } catch (eCurve38Existing) {}
            }
            var bulgeBase38 = Math.max(0.2, Math.abs(waistShapingMagnitude) * 0.45);
            var outwardSign38 = (point38.x - point30Waist.x) >= 0 ? 1 : -1;
            drawCurveBetween(point38, point38CurveEnd, {
                name: curveName38,
                bulgeCm: -bulgeBase38 * outwardSign38,
                flattenEnd: true,
                target: shapingTarget
            });
        }
    }
    if (point38 && shapingTarget) {
        drawFrameLine([point38.x, point38.y], [point38.x + 1, point38.y], 'Point38 Tick (Back)', FRAME_COLOR, shapingTarget);
    }

    if (point29Waist) {
        point37 = {
            x: point29Waist.x,
            y: point29Waist.y - 1
        };
        placeFrameMarker([point37.x, point37.y], '37');
        point37Guide = {
            x: point37.x - 3,
            y: point37.y
        };
        var line37Guide = drawFrameLine([point37.x, point37.y], [point37Guide.x, point37Guide.y], '37 Guide', FRAME_COLOR, shapingTarget);
        if (line37Guide) {
            try {
                line37Guide.strokeDashes = [];
            } catch (eLine37) {}
        }
        point39 = {
            x: point37.x - waistShaping,
            y: point37.y
        };
        placeFrameMarker([point39.x, point39.y], '39');
        if (shapingTarget) {
            drawFrameLine([point39.x - 1, point39.y], [point39.x, point39.y], 'Point39 Tick (Front)', FRAME_COLOR, shapingTarget);
        }
        if (backBodiceLayer) {
            drawBackBodiceLine('Back Bodice Tick 39', {
                x: point39.x - 1,
                y: point39.y
            }, {
                x: point39.x,
                y: point39.y
            }, false);
        }
        if (point29) {
            var waistShapingMagnitude39 = waistShaping;
            if (isNaN(waistShapingMagnitude39)) waistShapingMagnitude39 = defaults.WaistShaping;
            var waistToHipFraction39 = 2 / 3;
            var point39CurveEnd = {
                x: point29Waist.x + (point29.x - point29Waist.x) * waistToHipFraction39,
                y: point29Waist.y + (point29.y - point29Waist.y) * waistToHipFraction39
            };
            var curveName39 = 'Lower Front Side Curve';
            if (linesLayer) {
                try {
                    var existingCurves39 = linesLayer.pathItems;
                    for (var idx39 = existingCurves39.length - 1; idx39 >= 0; idx39--) {
                        var candidateCurve39 = existingCurves39[idx39];
                        if (!candidateCurve39) continue;
                        var candidateName39 = '';
                        try {
                            candidateName39 = candidateCurve39.name;
                        } catch (eCurve39Name) {}
                        if (candidateName39 === curveName39) {
                            try {
                                candidateCurve39.remove();
                            } catch (eCurve39Remove) {}
                        }
                    }
                } catch (eCurve39Existing) {}
            }
            var bulgeBase39 = Math.max(0.2, Math.abs(waistShapingMagnitude39) * 0.45);
            var outwardSign39 = (point39.x - point29Waist.x) >= 0 ? 1 : -1;
            drawCurveBetween(point39, point39CurveEnd, {
                name: curveName39,
                bulgeCm: -bulgeBase39 * outwardSign39,
                flattenEnd: true,
                target: shapingTarget
            });
        }
    }
    if (point39 && shapingTarget) {
        drawFrameLine([point39.x - 1, point39.y], [point39.x, point39.y], 'Point39 Tick (Back)', FRAME_COLOR, shapingTarget);
    }

    if (point38 && point11) {
        var line1138 = drawFrameLine([point11.x, point11.y], [point38.x, point38.y], 'Upper Back Side Line', FRAME_COLOR, shapingTarget);
        if (line1138) {
            try {
                line1138.strokeDashes = [];
            } catch (eLine1138) {}
        }
    }

    if (point39 && point12) {
        var line1239 = drawFrameLine([point12.x, point12.y], [point39.x, point39.y], 'Upper Front Side Line', FRAME_COLOR, shapingTarget);
        if (line1239) {
            try {
                line1239.strokeDashes = [];
            } catch (eLine1239) {}
        }
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
    var point31 = frontDartTop;
    if (point31) {
        placeFrameMarker([point31.x, point31.y], '31');
        if (point20a) {
            drawFrameLine([point20a.x, point20a.y], [point31.x, point31.y], 'Front Shoulder Line');
        }
    }
    var trianglePoint31 = point31;
    var dartShoulderPoint = point24;
    if (point22 && point24 && point31 && point32 && point13) {
        var targetX = (point32.x + point13.x) / 2;
        var pivotPoint = point22;
        var relX = point24.x - pivotPoint.x;
        var relY = point24.y - pivotPoint.y;
        var radius = Math.sqrt((relX * relX) + (relY * relY));
        if (radius > 1e-6) {
            var desiredCos = (targetX - pivotPoint.x) / radius;
            if (desiredCos < -1) desiredCos = -1;
            if (desiredCos > 1) desiredCos = 1;
            var baseAngle = Math.atan2(relY, relX);
            var acosValue = Math.acos(desiredCos);
            if (!isNaN(acosValue)) {
                var candidateAngles = [acosValue, -acosValue];
                var bestRotation = null;
                var bestDiff = null;
                for (var cIndex = 0; cIndex < candidateAngles.length; cIndex++) {
                    var targetAngle = candidateAngles[cIndex];
                    var rotationCandidate = targetAngle - baseAngle;
                    var rotated24Candidate = rotatePoint(point24, pivotPoint, rotationCandidate);
                    if (!rotated24Candidate) continue;
                    var diffX = Math.abs(rotated24Candidate.x - targetX);
                    if (bestRotation === null || diffX < bestDiff - 1e-6 || (Math.abs(diffX - bestDiff) < 1e-6 && Math.abs(rotationCandidate) < Math.abs(bestRotation))) {
                        bestRotation = rotationCandidate;
                        bestDiff = diffX;
                        dartShoulderPoint = rotated24Candidate;
                    }
                }
                if (bestRotation !== null) {
                    var rotated31Candidate = rotatePoint(point31, pivotPoint, bestRotation);
                    if (rotated31Candidate) trianglePoint31 = rotated31Candidate;
                }
            }
        }
    }
    var point33 = dartShoulderPoint;
        if (dartsGroup && trianglePoint31 && point22 && dartShoulderPoint) {
            var shoulderTriangle = dartsGroup.pathItems.add();
            shoulderTriangle.name = 'Front Shoulder Dart';
            shoulderTriangle.stroked = true;
            shoulderTriangle.strokeWidth = 1;
        shoulderTriangle.strokeColor = DART_TRIANGLE_COLOR;
        shoulderTriangle.filled = false;
        shoulderTriangle.closed = true;
            try {
                shoulderTriangle.strokeDashes = [];
            } catch (eShoulderDash) {}
            shoulderTriangle.setEntirePath([toArt(trianglePoint31), toArt(point22), toArt(dartShoulderPoint), toArt(trianglePoint31)]);
            point42 = {
                x: trianglePoint31.x,
                y: trianglePoint31.y
            };
            placeFrameMarker([point42.x, point42.y], '42');
            if (point33) placeFrameMarker([point33.x, point33.y], '33');
        }
    var frontDartBottom = {
        x: point22.x,
        y: hemLineY
    };
    var centreBackLinePath = drawLineWithLabel({
        x: point2.x,
        y: topLineY
    }, {
        x: point2.x,
        y: hemLineY
    }, 'Centre Back (CB)', FRAME_COLOR);
    if (centreBackLinePath) {
        try {
            centreBackLinePath.strokeDashes = DASH_PATTERN;
        } catch (eDashCentreBack) {}
    }
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
    }, 'Back Side Line', FRAME_COLOR);
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
    if (point31 && point22) {
        var frontDartUpperPath = drawFrameLine([point31.x, point31.y], [point22.x, point22.y], 'Front Dart Line Upper', FRAME_COLOR);
        if (frontDartUpperPath) {
            try {
                frontDartUpperPath.strokeDashes = [];
            } catch (eFrontDartUpper) {}
        }
    }
    var frontDartLinePath = drawLineWithLabel(point22, frontDartBottom, 'Front Dart Line', FRAME_COLOR);
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
    var hemLinePath = drawLineWithLabel(hemLineStart, hemEnd, 'Hem Line', FRAME_COLOR);
    if (hemLinePath) {
        try {
            hemLinePath.strokeDashes = DASH_PATTERN;
        } catch (eDashHem) {}
    }

    if (point29Hem && point14) {
        var centreFrontBottom = {
            x: point14.x,
            y: hemLineY
        };
        point40 = {
            x: point29Hem.x,
            y: point29Hem.y
        };
        point41 = {
            x: centreFrontBottom.x,
            y: centreFrontBottom.y
        };
        placeFrameMarker([point40.x, point40.y], '40');
        placeFrameMarker([point41.x, point41.y], '41');
        var frontHemLine = drawFrameLine([point29Hem.x, point29Hem.y], [centreFrontBottom.x, centreFrontBottom.y], 'Front Hem Line', FRAME_COLOR, casualLayer);
        if (frontHemLine) {
            try {
                frontHemLine.strokeDashes = [];
            } catch (eSolidFrontHem) {}
        }
    }
    var baseArtboardIndex = doc.artboards.getActiveArtboardIndex();
    cropArtboardAroundItems(baseArtboardIndex, [linesLayer, casualLayer, markersGroup, numbersGroup, labelsGroup], 10);
    var bustLinePath = drawLineWithLabel(bustLineStart, bustEnd, 'Bust Line', FRAME_COLOR);
    try {
        bustLinePath.strokeDashes = DASH_PATTERN;
    } catch (eDashBust) {}
    var waistLinePath = drawLineWithLabel(waistLineStart, waistEnd, 'Waist Line', FRAME_COLOR);
    try {
        waistLinePath.strokeDashes = DASH_PATTERN;
    } catch (eDashWaist) {}
    var hipLinePath = drawLineWithLabel(hipLineStart, hipEnd, 'Hip Line', FRAME_COLOR);
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
        point31: point31,
        point32: point32,
        point32Top: point32Top,
        point33: point33,
        frontDartTop: frontDartTop,
        frontDartBottom: frontDartBottom,
        hemEnd: hemEnd,
        bustEnd: bustEnd,
        waistEnd: waistEnd,
        hipEnd: hipEnd,
        frontShoulderEnd: frontShoulderEnd,
        backShoulderEnd: backShoulderEnd,
        backShoulderEndOriginal: backShoulderEndOriginal,
        backShoulderEndRotated: backShoulderEndRotated,
        p16: point16,
        point16Original: point16Original,
        point16Rotated: point16Rotated,
        p17: point17,
        point17Guide: point17Guide,
        point17GuideRotated: point17GuideRotated,
        p17a: point17a,
        point17aOriginal: point17aOriginal,
        p17b: point17b,
        p17c: point17c,
        point34: point34,
        point35: point35,
        point35Original: point35Original,
        point35Rotated: point35Rotated,
        point36: point36,
        point36Guide: point36Guide,
        point37: point37,
        point37Guide: point37Guide,
        point38: point38,
        point39: point39,
        point40: point40,
        point41: point41,
        point42: point42,
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

        BackContour: BackContour,

        backContourAuto: backContourAuto,

        backShoulderDartIntake: backShoulderDartIntake,

        backShoulderDartRotation: backShoulderDartRotation,

        waistShaping: waistShaping,

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
        if (!point12 || !point13 || !point18 || !point33) return;

        var casualLayer = ensureLayer('Casual Bodice');
        try {
            casualLayer.locked = false;
        } catch (eCasLockFront) {}
        try {
            casualLayer.visible = true;
        } catch (eCasVisFront) {}
        var casualStroke = makeRGB(0, 0, 0);

        removeGroupByName(casualLayer, 'Front Armhole');
        removePageItemsByName(casualLayer, 'Front Armhole Curve');

        var startAnchor = point12;
        var endAnchor = point33;
        var vec12to13 = {
            x: point13.x - point12.x,
            y: point13.y - point12.y
        };
        var len12to13 = Math.sqrt((vec12to13.x * vec12to13.x) + (vec12to13.y * vec12to13.y));
        if (len12to13 < 1e-6) return;
        var unit12to13 = {
            x: vec12to13.x / len12to13,
            y: vec12to13.y / len12to13
        };
        var handleFrom12 = {
            x: point13.x + unit12to13.x * 1.71,
            y: point13.y + unit12to13.y * 1.71
        };

        var path = casualLayer.pathItems.add();
        path.name = 'Front Armhole Curve';
        path.stroked = true;
        path.strokeWidth = 1;
        path.strokeColor = casualStroke;
        path.filled = false;
        path.closed = false;

        var artAnchors = [toArt(startAnchor), toArt(endAnchor)];
        path.setEntirePath(artAnchors);

        var pts = path.pathPoints;
        if (pts.length === artAnchors.length) {
            var startPt = pts[0];
            startPt.pointType = PointType.SMOOTH;
            startPt.leftDirection = startPt.anchor;
            startPt.rightDirection = toArt(handleFrom12);

            var endPt = pts[1];
            endPt.pointType = PointType.SMOOTH;
            // End handle must sit 0.3 cm left and 0.3 cm up from point 18
            endPt.leftDirection = toArt({
                x: point18.x - 0.3,
                y: point18.y + 0.3
            });
            endPt.rightDirection = endPt.anchor;
        }
    }

    function drawCasualBackArmhole() {
        if (!backShoulderEnd || !point17 || !point17Guide || !point11 || !point4) return;

        var casualLayer = ensureLayer('Casual Bodice');
        try {
            casualLayer.locked = false;
        } catch (eCasLockBack) {}
        try {
            casualLayer.visible = true;
        } catch (eCasVisBack) {}
        var casualStroke = makeRGB(0, 0, 0);

        removeGroupByName(casualLayer, 'Back Casual');
        removePageItemsByName(casualLayer, 'Back Armhole Curve (17-11)');

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
        var midAnchor = point17b || point17;
        var guidePoint = point17Guide;
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

        var waistPath = casualLayer.pathItems.add();
        waistPath.name = 'Back Armhole Curve (17-11)';
        waistPath.stroked = true;
        waistPath.strokeWidth = 1;
        waistPath.strokeColor = casualStroke;
        waistPath.filled = false;
        waistPath.closed = false;
        waistPath.setEntirePath([toArt(midAnchor), toArt(endAnchor)]);

        var pts2 = waistPath.pathPoints;
        if (pts2.length === 2) {
            var waistStart = pts2[0];
            waistStart.pointType = PointType.SMOOTH;
            waistStart.leftDirection = toArt(midIncomingHandle);
            waistStart.rightDirection = toArt(midOutgoingHandle);

            var waistEnd = pts2[1];
            waistEnd.pointType = PointType.SMOOTH;
            waistEnd.leftDirection = toArt(endLeftHandle);
            waistEnd.rightDirection = toArt(endAnchor);
        }
    }

    drawCasualFrontArmhole();
    drawCasualBackArmhole();
    copyFrontComponentsToPatternsLayer();
    copyBackComponentsToPatternsLayer();
    connectFrontHemToFrontSide();
    drawFrontBodiceCentreFrontSegment(point23, point41);
    drawFrontBodiceDashedLine('Front Hip Line', point19a, point29);
    drawFrontBodiceDashedLine('Front 9-11', point9, point11);
    drawFrontBodiceDashedLine('Front 14-12', point14, point12);
    removeFrontBodiceShoulderDartRightLeg();
    trimFrontBodiceWaistAtSide();
    // Back bodice adjustments
    resetBackShoulderFromFrame();
    trimBackShoulderDartLegs();
    trimBackDartLegAtShoulderLine();
    cutBackShoulderLineAtDartLeg();
    trimShoulderBladeLineAtCB();
    extendShoulderBladeLineToDart();
    removeBackBustLine();
    trimBackWaistLineAtSide();
    connectBackSideToHem();
    dashBackHipLine();

    try {
        app.executeMenuCommand('fitin');
    } catch (eFit) {}
})();
