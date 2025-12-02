const ExcelJS = require("exceljs");

/**
 * Get cell value including merged cell fallback
 */
function getCellValueWithMerge(sheet, row, col) {
    const cell = sheet.getCell(row, col);

    if (cell.value !== null && cell.value !== undefined && cell.value !== "") {
        return cell.value;
    }

    // Scan merged regions
    // const merge = sheet._merges.find(m =>
    //     row >= m.top && row <= m.bottom &&
    //     col >= m.left && col <= m.right
    // );

    // if (merge) {
    //     return sheet.getCell(merge.top, merge.left).value;
    // }

    return "";
}

/**
 * Parse Excel file to nested structured JSON
 */
async function parseActionPlan(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const sheet = workbook.worksheets[0];

    let result = [];

    // Data starts at row 3 (based on your template)
    for (let r = 3; r <= sheet.rowCount; r++) {
        const mucTieu = getCellValueWithMerge(sheet, r, 2);
        const kpi = getCellValueWithMerge(sheet, r, 3);
        const action = getCellValueWithMerge(sheet, r, 4);

        const tenCv = getCellValueWithMerge(sheet, r, 5);
        // const nguoi = getCellValueWithMerge(sheet, r, 6);
        // const maBp = getCellValueWithMerge(sheet, r, 7);
        // const tgBd = getCellValueWithMerge(sheet, r, 8);
        // const tgKt = getCellValueWithMerge(sheet, r, 9);
        // const dvt = getCellValueWithMerge(sheet, r, 10);
        // const ns = getCellValueWithMerge(sheet, r, 11);

        if (!mucTieu) continue;

        // === Level 1: Mục tiêu ===
        let index = result.findIndex(x => x.action_plan === action);
        if (index < 0) {
            const mt = { action_plan: action, kpi, tasks: [{
                ten_cong_viec: tenCv,
                target: mucTieu,
            }] };
            result.push(mt);
            continue;
        }

        // // === Level 2: KPI ===
        // let kpiItem = mt.kpis.find(x => x.kpi === kpi);
        // if (!kpiItem) {
        //     kpiItem = { kpi, actions: [] };
        //     mt.kpis.push(kpiItem);
        // }

        // // === Level 3: Action ===
        // let act = kpiItem.actions.find(x => x.action_plan === action);
        // if (!act) {
        //     act = { action_plan: action, kpi: kpi, tasks: [] };
        //     kpiItem.actions.push(act);
        // }

        // === Level 4: Task ===
        result[index].tasks.push({
            ten_cong_viec: tenCv,
            target: mucTieu,
            // nguoi_thuc_hien: nguoi,
            // ma_bo_phan: maBp,
            // thoi_gian_bat_dau: tgBd,
            // thoi_gian_ket_thuc: tgKt,
            // don_vi_tinh: dvt,
            // ngan_sach: ns
        });
    }

    return result;
}

module.exports = {
    parseActionPlan
};
