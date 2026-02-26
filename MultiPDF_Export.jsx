// ============================================
// Multi PDF Exporter - InDesign ScriptUI
// ============================================

#target indesign

(function () {

    if (!app.activeDocument) {
        alert("Lutfen once bir InDesign belgesi acin.");
        return;
    }

    var doc = app.activeDocument;

    // Belge adini al (uzantisiz)
    var docName = "Belge";
    if (doc.saved) {
        var fullName = doc.name; // or: Katalog.indd
        var dotIndex = fullName.lastIndexOf(".");
        docName = (dotIndex > -1) ? fullName.substring(0, dotIndex) : fullName;
    }

    // --- Ayar dosyasi (script ile ayni klasorde settings.json) ---
    var settingsFile = new File($.fileName.replace(/[^\/\\]+$/, "") + "MultiPDF_Settings.json");

    function loadSettings() {
        var defaults = {
            ranges: "1-10\n11-20\n21-30",
            prefix: "",
            suffix: "",
            preset: "",
            folder: "",
            openFolder: true,
            exportFull: true
        };
        if (!settingsFile.exists) return defaults;
        try {
            settingsFile.open("r");
            var raw = settingsFile.read();
            settingsFile.close();
            // Basit JSON parse - key:value
            var s = defaults;
            var pairs = [["ranges","ranges"], ["prefix","prefix"], ["suffix","suffix"],
                         ["preset","preset"], ["folder","folder"],
                         ["openFolder","openFolder"], ["exportFull","exportFull"]];
            for (var pi2 = 0; pi2 < pairs.length; pi2++) {
                var key = pairs[pi2][0];
                var rx = '"' + key + '"\s*:\s*';
                var idx = raw.indexOf('"' + key + '"');
                if (idx === -1) continue;
                var colon = raw.indexOf(':', idx);
                if (colon === -1) continue;
                var rest = raw.substring(colon + 1);
                // Bos ve boslukları atla
                var vi = 0;
                while (vi < rest.length && (rest.charAt(vi) === ' ' || rest.charAt(vi) === '\n' || rest.charAt(vi) === '\r')) vi++;
                var firstChar = rest.charAt(vi);
                if (firstChar === '"') {
                    // String deger
                    var end = rest.indexOf('"', vi + 1);
                    while (end !== -1 && rest.charAt(end - 1) === '\\') end = rest.indexOf('"', end + 1);
                    if (end !== -1) {
                        var val = rest.substring(vi + 1, end);
                        // Escape coz
                        val = val.split('\\n').join('\n');
                        val = val.split('\\\"').join('"');
                        s[key] = val;
                    }
                } else if (firstChar === 't' || firstChar === 'f') {
                    s[key] = (firstChar === 't');
                }
            }
            return s;
        } catch(ex) { return defaults; }
    }

    function saveSettings(ranges, prefix, suffix, preset, folder, openFolder, exportFull) {
        try {
            var escapedRanges = ranges.split('\n').join('\\n').split('"').join('\\"');
            var escapedPrefix = prefix.split('"').join('\\"');
            var escapedSuffix = suffix.split('"').join('\\"');
            var escapedPreset = preset.split('"').join('\\"');
            var escapedFolder = folder.split('\\').join('\\\\').split('"').join('\\"');
            var json = '{\n' +
                '  "ranges": "' + escapedRanges + '",\n' +
                '  "prefix": "' + escapedPrefix + '",\n' +
                '  "suffix": "' + escapedSuffix + '",\n' +
                '  "preset": "' + escapedPreset + '",\n' +
                '  "folder": "' + escapedFolder + '",\n' +
                '  "openFolder": ' + (openFolder ? 'true' : 'false') + ',\n' +
                '  "exportFull": ' + (exportFull ? 'true' : 'false') + '\n' +
                '}';
            settingsFile.open("w");
            settingsFile.write(json);
            settingsFile.close();
        } catch(ex) { /* sessizce gec */ }
    }

    var cfg = loadSettings();

    // --- Preset listesini al ---
    var presetNames = [];
    for (var p = 0; p < app.pdfExportPresets.length; p++) {
        presetNames.push(app.pdfExportPresets[p].name);
    }
    if (presetNames.length === 0) presetNames = ["[Press Quality]"];

    // --- Kayit klasoru ---
    var saveFolder = (cfg.folder && cfg.folder !== "") ? new Folder(cfg.folder) : (doc.saved ? doc.filePath : Folder.desktop);
    if (!saveFolder.exists) saveFolder = doc.saved ? doc.filePath : Folder.desktop;

    // ============================================================
    // PENCERE
    // ============================================================
    var win = new Window("dialog", "Multi PDF Exporter", undefined, { resizeable: false });
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.spacing = 12;
    win.margins = 18;

    // Baslik
    var titleText = win.add("statictext", undefined, "InDesign - Coklu PDF Export");
    titleText.graphics.font = ScriptUI.newFont("dialog", ScriptUI.FontStyle.BOLD, 13);

    // --- PDF Preset ---
    var presetGroup = win.add("group");
    presetGroup.orientation = "row";
    presetGroup.alignChildren = ["left", "center"];
    var presetLabel = presetGroup.add("statictext", undefined, "PDF Preset:");
    presetLabel.preferredSize.width = 90;
    var presetDropdown = presetGroup.add("dropdownlist", undefined, presetNames);
    presetDropdown.preferredSize.width = 260;
    presetDropdown.selection = 0;
    for (var pi = 0; pi < presetNames.length; pi++) {
        if (cfg.preset !== "" && presetNames[pi] === cfg.preset) {
            presetDropdown.selection = pi; break;
        } else if (cfg.preset === "" && presetNames[pi].indexOf("Press Quality") !== -1) {
            presetDropdown.selection = pi; break;
        }
    }

    // --- Kayit Klasoru ---
    var folderGroup = win.add("group");
    folderGroup.orientation = "row";
    folderGroup.alignChildren = ["left", "center"];
    var folderStaticLabel = folderGroup.add("statictext", undefined, "Kayit Klasoru:");
    folderStaticLabel.preferredSize.width = 90;
    var folderLabel = folderGroup.add("edittext", undefined, (saveFolder.fsName || saveFolder.toString()));
    folderLabel.preferredSize.width = 220;
    folderLabel.enabled = false;
    var browseBtn = folderGroup.add("button", undefined, "Sec...");
    browseBtn.preferredSize.width = 50;
    browseBtn.onClick = function () {
        var chosen = Folder.selectDialog("Kayit klasorunu secin:", saveFolder);
        if (chosen) {
            saveFolder = chosen;
            folderLabel.text = chosen.fsName;
        }
    };

    // --- Ayrac ---
    win.add("panel", undefined, "").preferredSize.height = 1;

    // --- Aciklama ---
    var infoText = win.add("statictext", undefined,
        "Her satira bir aralik girin (baslangic-bitis).\nDosya adi: " + docName + "_1-10.pdf seklinde olusacak.",
        { multiline: true });
    infoText.preferredSize.width = 380;

    // --- Metin kutusu ---
    var textPanel = win.add("panel", undefined, "Sayfa Araliklari");
    textPanel.orientation = "column";
    textPanel.alignChildren = ["fill", "top"];
    textPanel.margins = [10, 15, 10, 10];

    var rangeBox = textPanel.add("edittext", undefined, cfg.ranges, { multiline: true, scrolling: true });
    rangeBox.preferredSize = { width: 360, height: 140 };

    // Onizleme etiketi
    var previewLabel = textPanel.add("statictext", undefined, "", { multiline: true });
    previewLabel.preferredSize = { width: 360, height: 36 };

    // Satiri parse et
    function parseLine(line) {
        var trimmed = "";
        for (var ci = 0; ci < line.length; ci++) {
            var ch = line.charAt(ci);
            if (ch !== " " && ch !== "\r" && ch !== "\t") trimmed += ch;
        }
        if (trimmed.length === 0) return null;
        // "full" anahtar kelimesi: tum belgeyi al
        if (trimmed.toLowerCase() === "full") return { start: -1, end: -1, isFull: true };
        var dashIdx = trimmed.indexOf("-");
        if (dashIdx < 1) return null;
        var s = parseInt(trimmed.substring(0, dashIdx), 10);
        var e = parseInt(trimmed.substring(dashIdx + 1), 10);
        if (isNaN(s) || isNaN(e) || s < 1 || e < s) return null;
        return { start: s, end: e };
    }

    // Canli onizleme
    function updatePreview() {
        var lines = rangeBox.text.split("\n");
        var valid = 0;
        var firstRange = "";
        for (var li = 0; li < lines.length; li++) {
            var parsed = parseLine(lines[li]);
            if (parsed) {
                valid++;
                if (firstRange === "") firstRange = parsed.isFull ? "full" : parsed.start + "-" + parsed.end;
            }
        }
        if (valid === 0) {
            previewLabel.text = "Gecerli aralik bulunamadi.";
        } else {
            previewLabel.text = valid + " adet PDF olusturulacak. Or: " + docName + "_" + firstRange + ".pdf";
        }
    }

    rangeBox.onChanging = updatePreview;
    updatePreview();

    // --- Onay kutulari ---
    var checkGroup = win.add("group");
    checkGroup.orientation = "column";
    checkGroup.alignChildren = ["left", "top"];
    checkGroup.spacing = 6;

    var chkOpenFolder = checkGroup.add("checkbox", undefined, "Export sonunda kayit klasorunu ac");
    chkOpenFolder.value = cfg.openFolder;

    var chkExportFull = checkGroup.add("checkbox", undefined, "Aralikların yanı sıra tam belgeyi de export et (_full.pdf)");
    chkExportFull.value = cfg.exportFull;

    // --- Onek / Sonek ---
    var affixPanel = win.add("panel", undefined, "On Ek / Son Ek (opsiyonel)");
    affixPanel.orientation = "row";
    affixPanel.alignChildren = ["left", "center"];
    affixPanel.margins = [10, 15, 10, 10];
    affixPanel.spacing = 10;

    affixPanel.add("statictext", undefined, "On Ek:");
    var prefixField = affixPanel.add("edittext", undefined, cfg.prefix);
    prefixField.preferredSize.width = 120;
    prefixField.helpTip = "Dosya adinin basina eklenir. Or: BASKI_ -> BASKI_Katalog_1-10.pdf";

    affixPanel.add("statictext", undefined, "Son Ek:");
    var suffixField = affixPanel.add("edittext", undefined, cfg.suffix);
    suffixField.preferredSize.width = 120;
    suffixField.helpTip = "Dosya adinin sonuna eklenir. Or: _v2 -> Katalog_1-10_v2.pdf";

    // --- Ayrac ---
    win.add("panel", undefined, "").preferredSize.height = 1;

    // --- Butonlar ---
    var btnGroup = win.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignment = "right";
    btnGroup.spacing = 8;

    var cancelBtn = btnGroup.add("button", undefined, "Iptal", { name: "cancel" });
    cancelBtn.preferredSize.width = 80;

    var exportBtn = btnGroup.add("button", undefined, "PDF'leri Olustur ->", { name: "ok" });
    exportBtn.preferredSize.width = 160;

    // ============================================================
    // EXPORT
    // ============================================================
    exportBtn.onClick = function () {

        var lines = rangeBox.text.split("\n");
        var validRows = [];

        for (var li = 0; li < lines.length; li++) {
            var line = lines[li];
            var parsed = parseLine(line);

            // Bos satiri atla
            var trimCheck = "";
            for (var ci2 = 0; ci2 < line.length; ci2++) {
                var ch2 = line.charAt(ci2);
                if (ch2 !== " " && ch2 !== "\r" && ch2 !== "\t" && ch2 !== "\n") trimCheck += ch2;
            }
            if (trimCheck.length === 0) continue;

            if (!parsed) {
                alert((li + 1) + ". satir gecersiz: \"" + line + "\"\nDogru format: 1-10");
                return;
            }

            var baseName = parsed.isFull
                ? docName + "_full"
                : docName + "_" + parsed.start + "-" + parsed.end;
            var fileName = prefixField.text + baseName + suffixField.text;
            validRows.push({ start: parsed.start, end: parsed.end, name: fileName, isFull: parsed.isFull });
        }

        // chkExportFull isaretliyse full'u da ekle
        if (chkExportFull.value) {
            validRows.push({ start: -1, end: -1, name: prefixField.text + docName + "_full" + suffixField.text, isFull: true });
        }

        if (validRows.length === 0) {
            alert("Hic gecerli aralik girilmedi.");
            return;
        }

        var preset;
        try {
            preset = app.pdfExportPresets.item(presetDropdown.selection.text);
        } catch (ex) {
            alert("PDF Preset bulunamadi: " + presetDropdown.selection.text);
            return;
        }

        // Ayarlari kaydet
        saveSettings(
            rangeBox.text,
            prefixField.text,
            suffixField.text,
            presetDropdown.selection ? presetDropdown.selection.text : "",
            saveFolder.fsName || "",
            chkOpenFolder.value,
            chkExportFull.value
        );

        win.close();

        var total = validRows.length;
        var errors = [];
        var successCount = 0;

        // --- Progress penceresi ---
        var prog = new Window("palette", "PDF Export Yapiliyor...");
        prog.orientation = "column";
        prog.alignChildren = ["fill", "center"];
        prog.spacing = 10;
        prog.margins = 20;
        prog.preferredSize.width = 420;

        var progTitle = prog.add("statictext", undefined, "Hazirlanıyor...");
        progTitle.graphics.font = ScriptUI.newFont("dialog", ScriptUI.FontStyle.BOLD, 11);
        progTitle.preferredSize.width = 380;

        var progSub = prog.add("statictext", undefined, "");
        progSub.preferredSize.width = 380;

        // Progress bar
        var progBar = prog.add("progressbar", undefined, 0, total);
        progBar.preferredSize = { width: 380, height: 16 };

        var progCount = prog.add("statictext", undefined, "0 / " + total);
        progCount.alignment = "center";

        prog.center();
        prog.show();
        app.doScript(function() {}, ScriptLanguage.JAVASCRIPT); // UI flush

        for (var j = 0; j < total; j++) {
            var r = validRows[j];

            // UI guncelle
            var label = r.isFull ? "Tam belge (full)" : "Sayfa " + r.start + " - " + r.end;
            progTitle.text = "Kaydediliyor: " + r.name + ".pdf";
            progSub.text = label + "  (" + (j + 1) + ". dosya)";
            progBar.value = j;
            progCount.text = j + " / " + total;
            prog.update();

            try {
                var outputFile = new File(saveFolder.fsName + "/" + r.name + ".pdf");
                if (r.isFull) {
                    app.pdfExportPreferences.pageRange = "";
                    doc.exportFile(ExportFormat.PDF_TYPE, outputFile, false, preset);
                } else {
                    app.pdfExportPreferences.pageRange = r.start + "-" + r.end;
                    doc.exportFile(ExportFormat.PDF_TYPE, outputFile, false, preset);
                }
                successCount++;
            } catch (ex) {
                errors.push("- " + r.name + ".pdf: " + ex.message);
            }

            progBar.value = j + 1;
            progCount.text = (j + 1) + " / " + total;
            prog.update();
        }

        // Tamamlandi
        progTitle.text = "Tamamlandi!";
        progSub.text = successCount + " adet PDF basariyla olusturuldu.";
        progBar.value = total;
        progCount.text = total + " / " + total;
        prog.update();

        // Kisa bekleme - kullanici gorabilsin
        var waitUntil = (new Date()).getTime() + 900;
        while ((new Date()).getTime() < waitUntil) { prog.update(); }

        prog.close();

        if (errors.length === 0) {
            alert(successCount + " adet PDF basariyla olusturuldu!\n\nKonum: " + saveFolder.fsName);
        } else {
            alert(successCount + " PDF olusturuldu, " + errors.length + " hata:\n\n" + errors.join("\n"));
        }

        // Klasoru ac
        if (chkOpenFolder.value) {
            saveFolder.execute();
        }
    };

    win.center();
    win.show();

})();
