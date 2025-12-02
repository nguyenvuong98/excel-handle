const ExcelJS = require("exceljs");

/**
 * Get cell value including merged cell fallback
 */
let mergeCount = 0;
function getCellValue(sheet, row, col) {
    const cell = sheet.getCell(row, col);
    if (cell.value !== null && cell.value !== undefined && cell.value !== "") {
        return cell.value;
    }

    return "";
}

function getMergeCount(sheet, row, col) {
    return sheet.getCell(row, col)._mergeCount;
}
/**
 * Parse Excel file to nested structured JSON
 */
async function parseActionPlan(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const sheet = workbook.worksheets[0];

    let result = [];
    let countMerge = 0;
    // Data starts at row 3 (based on your template)
    for (let r = 3; r <= sheet.rowCount; r++) {
        const mergeCol = getMergeCount(sheet, r, 4);
        const mucTieu = getCellValue(sheet, r, 2);
        const kpi = getCellValue(sheet, r, 3);
        const action = getCellValue(sheet, r, 4);

        const tenCv = getCellValue(sheet, r, 5);
        if (!action) continue;
        if (countMerge == 0) {
            countMerge = mergeCol;
            const mt = { action_plan: action, kpi, tasks: [{
                ten_cong_viec: tenCv,
                target: mucTieu,
            }] };
            result.push(mt);
            continue;
        } else {
            result[result.length - 1].tasks.push({
                ten_cong_viec: tenCv,
                target: mucTieu,
                // nguoi_thuc_hien: nguoi,
                // ma_bo_phan: maBp,
                // thoi_gian_bat_dau: tgBd,
                // thoi_gian_ket_thuc: tgKt,
                // don_vi_tinh: dvt,
                // ngan_sach: ns
            });
            countMerge = countMerge == 0 ? countMerge : countMerge - 1;
        }
        // === Level 1: Mục tiêu ===
        // let index = result.findIndex(x => x.action_plan === action);
        // if (index < 0) {
        //     const mt = { action_plan: action, kpi, tasks: [{
        //         ten_cong_viec: tenCv,
        //         target: mucTieu,
        //     }] };
        //     result.push(mt);
        //     continue;
        // }

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
        // result[index].tasks.push({
        //     ten_cong_viec: tenCv,
        //     target: mucTieu,
            // nguoi_thuc_hien: nguoi,
            // ma_bo_phan: maBp,
            // thoi_gian_bat_dau: tgBd,
            // thoi_gian_ket_thuc: tgKt,
            // don_vi_tinh: dvt,
            // ngan_sach: ns
        //});
    }

    return result;
}

module.exports = {
    parseActionPlan
};
