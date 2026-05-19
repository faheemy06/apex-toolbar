// ========================================================
// 1. OFFICE ENGINE INITIALIZATION
// ========================================================
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initializeDragEngine();
        initializeFinanceButtons();
    }
});

// ========================================================
// 2. THE DRAG & DROP ENGINE (The 6 Dots)
// ========================================================
function initializeDragEngine() {
    const dragHandle = document.getElementById("drag-handle");
    const toolbar = document.getElementById("toolbar");

    let isDragging = false;
    let startX, startY, initialLeft, initialTop;

    dragHandle.addEventListener("mousedown", (e) => {
        isDragging = true;
        startX = e.clientX;
        startY = e.clientY;
        
        const rect = toolbar.getBoundingClientRect();
        initialLeft = rect.left;
        initialTop = rect.top;
        
        document.addEventListener("mousemove", dragMove);
        document.addEventListener("mouseup", dragStop);
        e.preventDefault(); 
    });

    function dragMove(e) {
        if (!isDragging) return;
        const dx = e.clientX - startX;
        const dy = e.clientY - startY;
        
        toolbar.style.left = (initialLeft + dx) + "px";
        toolbar.style.top = (initialTop + dy) + "px";
    }

    function dragStop() {
        isDragging = false;
        document.removeEventListener("mousemove", dragMove);
        document.removeEventListener("mouseup", dragStop);
    }
}

// ========================================================
// 3. CORE FINANCE LOGIC & FULL-SHEET AUDIT X-RAY
// ========================================================
function initializeFinanceButtons() {
    
    // --- Automated Full-Sheet Audit X-Ray ---
    document.getElementById("btn-audit").onclick = async () => {
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const usedRange = sheet.getUsedRange();
                
                // Scan for distinct cell architectures
                const formulas = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
                const constants = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.constants);
                const formulaErrors = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas, Excel.SpecialCellValueType.errors);
                
                usedRange.load(["cellCount", "values"]);
                await context.sync();

                let formulaCount = 0;
                let constantCount = 0;
                let errorCount = 0;

                // 1. Color-code dynamic calculation models
                if (!formulas.isNullObject) {
                    formulas.load("cellCount");
                    formulas.format.fill.color = "#cce5ff"; // Corporate Blue
                    await context.sync();
                    formulaCount = formulas.cellCount;
                }
                
                // 2. Color-code hardcoded inputs
                if (!constants.isNullObject) {
                    constants.load("cellCount");
                    constants.format.fill.color = "#ffe0b2"; // Amber
                    await context.sync();
                    constantCount = constants.cellCount;
                }

                // 3. Color-code true formula calculation breaks
                if (!formulaErrors.isNullObject) {
                    formulaErrors.load("cellCount");
                    formulaErrors.format.fill.color = "#ffcdd2"; // Alert Red
                    await context.sync();
                    errorCount = formulaErrors.cellCount;
                }

                // 4. Catch text-based broken string flags (e.g., typed or pasted #REF!)
                const rows = usedRange.values;
                for (let r = 0; r < rows.length; r++) {
                    for (let c = 0; c < rows[r].length; c++) {
                        const cellValue = String(rows[r][c]);
                        if (cellValue.includes("#REF!") || cellValue.includes("#VALUE!") || cellValue.includes("#DIV/0!")) {
                            const errorCell = usedRange.getCell(r, c);
                            errorCell.format.fill.color = "#ffcdd2"; // Overwrite to Alert Red
                            errorCount++;
                            if (constantCount > 0) constantCount--; // Correct our metrics count
                        }
                    }
                }
                await context.sync();

                // Display summary profile
                alert(
                    "📊 EL-TOOLBAR AUDIT SUMMARY\n" +
                    "-----------------------------------------\n" +
                    "✅ Total Active Footprint: " + usedRange.cellCount + " cells scanned.\n\n" +
                    "🔗 Dynamic Formulas Found: " + formulaCount + " (Highlighted Blue)\n" +
                    "✏️ Hardcoded Values Found: " + constantCount + " (Highlighted Orange)\n" +
                    "⚠️ Broken Errors Caught: " + errorCount + " (Highlighted Red)\n" +
                    "-----------------------------------------\n" +
                    "Scan Complete. Model structural mapping applied."
                );
            });
        } catch (e) {
            console.log("Audit Error: " + e.message);
            alert("No active data matrix detected to scan on this sheet.");
        }
    };

    // --- Professional Formatting ---
    document.getElementById("btn-format").onclick = () => runExcelAction(r => {
        r.numberFormat = [["£#,##0.00"]];
        r.format.font.bold = true;
    });

    // --- Smart AutoSum (Horizontal & Vertical) ---
    document.getElementById("btn-sum").onclick = () => runExcelAction(async (range) => {
        range.load(["address", "rowCount", "columnCount"]);
        await range.context.sync();
        
        let target = range.rowCount >= range.columnCount ? 
                     range.getLastRow().getOffsetRange(1, 0) : 
                     range.getLastColumn().getOffsetRange(0, 1);
        
        target.formulas = [["=SUM(" + range.address + ")"]];
        target.format.font.bold = true;
        target.format.fill.color = "#e2efda"; 
    });

    // --- Add 20% VAT (With Smart AutoFit Column Fix) ---
    document.getElementById("btn-vat").onclick = () => runExcelAction(async (r) => {
        r.load("address"); await r.context.sync();
        const target = r.getLastColumn().getOffsetRange(0, 1);
        target.formulas = [["=" + r.address + "*1.2"]];
        target.numberFormat = [["£#,##0.00"]];
        target.format.fill.color = "#fff2cc"; 
        
        // Kills the '###' bug by instantly fitting the data width
        target.getEntireColumn().format.autofitColumns();
    });

    // --- Variance Δ% (Smart Direction) ---
    document.getElementById("btn-delta").onclick = () => runExcelAction(async (range) => {
        range.load(["address", "rowCount", "columnCount"]);
        await range.context.sync();
        const cells = range.address.split(":");
        const f = cells[0], s = cells.length > 1 ? cells[1] : cells[0];
        
        let target = range.rowCount >= range.columnCount ? 
                     range.getLastRow().getOffsetRange(1, 0) : 
                     range.getLastColumn().getOffsetRange(0, 1);
                     
        target.formulas = [["=(" + s + "-" + f + ")/" + f]];
        target.numberFormat = [["0.00%"]];
        target.format.fill.color = "#e0f2f1"; 
    });

    // --- Navigation & Clean ---
    document.getElementById("btn-freeze").onclick = () => runExcelAction(async (r) => { r.worksheet.freezePanes.freezeRows(1); });
    document.getElementById("btn-unfreeze").onclick = () => runExcelAction(async (r) => { r.worksheet.freezePanes.unfreeze(); });
    document.getElementById("btn-clean").onclick = () => runExcelAction(r => { 
        r.clear(Excel.ClearApplyTo.formats); 
        r.format.fill.clear(); 
    });
}

// ========================================================
// 4. CORE RUNTIME PIPELINE (For Selection Actions)
// ========================================================
async function runExcelAction(callback) {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            await callback(range);
            await context.sync();
        });
    } catch (e) { console.log("Logic Error: " + e.message); }
}
