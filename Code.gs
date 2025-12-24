function updateAllocationSheet() {
  // SOURCE FILE
  const sourceFileId = "1nZsJ8ZUOvW0XnEcL8yUizHvvuh43WCdXU10qntQdk7Q";  // <-- your source file ID
  const sourceSheetName = "Previous Month"; // <-- sheet to copy from

  // DESTINATION FILE (this file)
  const destSheetName = "Allocation"; // <-- sheet to overwrite

  // Open source
  const source = SpreadsheetApp.openById(sourceFileId);
  const sourceSheet = source.getSheetByName(sourceSheetName);

  // Open destination
  const dest = SpreadsheetApp.getActive();
  let destSheet = dest.getSheetByName(destSheetName);

  if (!destSheet) {
    destSheet = dest.insertSheet(destSheetName);
  }

  // Get data
  const range = sourceSheet.getDataRange();
  const values = range.getValues();

  // Overwrite destination
  destSheet.clearContents();
  destSheet.getRange(1, 1, values.length, values[0].length).setValues(values);

  Logger.log("Allocation sheet updated successfully.");
}



/**
 * Dashboard backend
 * - getDropdownValues(): reads Dropdown sheet, parses months, chooses most recent
 * - getHeadcountByGeo(filters): sum Headcount grouped by Geo
 * - getHeadcountBySite(filters): sum Headcount grouped by Site
 * - getHeadcountByProgram(filters): sum Headcount grouped by Program
 *
 * HC sheet layout (columns):
 * A Project Code
 * B Class
 * C Geo
 * D Program
 * E Site
 * F Headcount
 * G Month (string helper, e.g., "October 2025")
 * I Cluster Lead
 * J Director
 * K SOM
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle("Headcount and Ratio Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDropdownValues() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Dropdown");
  if (!sh) return {
    months: [], programs: [], geos: [], sites: [],
    clusterLeads: [], directors: [], soms: [],
    defaultMonth: ''
  };

  const lastRow = Math.max(2, sh.getLastRow());

  // Read raw columns
  const monthsRaw = sh.getRange("A2:A" + lastRow).getValues().flat().map(v => String(v||'').trim()).filter(Boolean);
  const programsRaw = sh.getRange("C2:C" + lastRow).getValues().flat().map(v => String(v||'').trim()).filter(Boolean);
  const geosRaw = sh.getRange("D2:D" + lastRow).getValues().flat().map(v => String(v||'').trim()).filter(Boolean);
  const sitesRaw = sh.getRange("E2:E" + lastRow).getValues().flat().map(v => String(v||'').trim()).filter(Boolean);

  const clusterLeadsRaw = sh.getRange("F2:F" + lastRow).getValues().flat().map(v => String(v||'').trim()).filter(Boolean);
  const directorsRaw = sh.getRange("G2:G" + lastRow).getValues().flat().map(v => String(v||'').trim()).filter(Boolean);
  const somsRaw = sh.getRange("H2:H" + lastRow).getValues().flat().map(v => String(v||'').trim()).filter(Boolean);

  // Keep unique values (preserve order)
  const unique = arr => Array.from(new Set(arr));

  // Determine defaultMonth by parsing month strings into Dates and finding max
  const parsed = monthsRaw
    .map(m => ({ raw: m, date: tryParseMonth(m) }))
    .filter(x => x.date instanceof Date && !isNaN(x.date));

  // Sort descending by real date
  parsed.sort((a,b) => b.date - a.date);
  const defaultMonth = parsed.length ? parsed[0].raw : (monthsRaw[monthsRaw.length-1] || '');

  return {
    months: unique(monthsRaw),
    programs: unique(programsRaw),
    geos: unique(geosRaw),
    sites: unique(sitesRaw),
    clusterLeads: unique(clusterLeadsRaw),
    directors: unique(directorsRaw),
    soms: unique(somsRaw),
    defaultMonth: defaultMonth
  };
}

// Try convert "October 2025" or "Oct 2025" to a date (first of month)
function tryParseMonth(str) {
  if (!str) return null;
  // try adding a day and letting JS parse
  const t = new Date(str + " 1");
  if (t instanceof Date && !isNaN(t)) return t;
  // fallback: try replacing common short names
  try {
    const alt = str.replace(/Sept(?=\s|$)/i, "Sep");
    const t2 = new Date(alt + " 1");
    return (t2 instanceof Date && !isNaN(t2)) ? t2 : null;
  } catch (e) {
    return null;
  }
}


/**
 * Backend updates: include Billable count (HC sheet column V -> zero-based index 21)
 * - Reads columns A..V (22 columns) so we can access Billable
 * - aggregateHeadcountBy now accumulates `billable` per group
 * - getHeadcountByGeo / getHeadcountBySite / getHeadcountByProgram return billable per row & totals.totalBillable
 */

/**
 * Generic aggregator used by Geo / Site / Program headcount endpoints.
 * groupIdx = column index to group by (0-based; e.g. geo = 2, program = 3, site = 4)
 */
function aggregateHeadcountBy(groupIdx, filters) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("HC");
  if (!sh) return { rows: [], totals: {} };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { rows: [], totals: {} };

  // Read columns A..V (22 columns) so we can access Billable (V = index 21)
  const data = sh.getRange(2, 1, lastRow - 1, 22).getValues();

  const monthFilter = filters?.month?.trim() || '';
  const programFilter = filters?.program?.trim() || '';
  const siteFilter = filters?.site?.trim() || '';
  const geoFilter = filters?.geo?.trim() || '';
  const clusterLeadFilter = filters?.clusterLead?.trim() || '';
  const directorFilter = filters?.director?.trim() || '';
  const somFilter = filters?.som?.trim() || '';

  const agg = {};
  let totalHeadcount = 0;
  let totalBillable = 0;

  data.forEach(row => {
    const geo = (row[2] || '').toString().trim();
    const program = (row[3] || '').toString().trim();
    const site = (row[4] || '').toString().trim();
    const headcountRaw = row[5];
    const month = (row[6] || '').toString().trim();
    const clusterLead = (row[8] || '').toString().trim();
    const director = (row[9] || '').toString().trim();
    const som = (row[10] || '').toString().trim();
    const billableRaw = row[21]; // column V

    // Apply filters
    if (monthFilter && month !== monthFilter) return;
    if (programFilter && program !== programFilter) return;
    if (siteFilter && site !== siteFilter) return;
    if (geoFilter && geo !== geoFilter) return;
    if (clusterLeadFilter && clusterLead !== clusterLeadFilter) return;
    if (directorFilter && director !== directorFilter) return;
    if (somFilter && som !== somFilter) return;

    const hc = (typeof headcountRaw === 'number')
      ? headcountRaw
      : parseFloat(String(headcountRaw).replace(/,/g, '')) || 0;

    // Robust billable parsing:
    // - If numeric: use numeric (some sheets may already have 0/1 counts)
    // - If string: treat "yes/y/true/1" as 1, otherwise 0
    let billableCount = 0;
    if (billableRaw !== null && billableRaw !== undefined && String(billableRaw).toString().trim() !== '') {
      if (typeof billableRaw === 'number') {
        // accept numeric values (sum them)
        billableCount = Number(billableRaw) || 0;
      } else {
        const b = String(billableRaw).trim().toLowerCase();
        if (b === 'yes' || b === 'y' || b === 'true' || b === '1') billableCount = 1;
        else {
          // try parse number fallback
          const parsed = parseFloat(b.replace(/,/g,''));
          if (!isNaN(parsed)) billableCount = parsed;
        }
      }
    }

    const key = (row[groupIdx] || '').toString().trim() || 'Unspecified';

    if (!agg[key]) {
      agg[key] = { headcount: 0, billable: 0 };
    }

    agg[key].headcount += hc;
    agg[key].billable += billableCount;

    totalHeadcount += hc;
    totalBillable += billableCount;
  });

  const rows = Object.keys(agg)
    .map(k => ({ key: k, headcount: agg[k].headcount, billable: agg[k].billable }))
    .sort((a, b) => b.headcount - a.headcount);

  // NEW: compute scorecard values
  const activeGroups = rows.filter(r => r.headcount > 0).length;
  const topGroup = rows.length ? rows[0].key : '—';

  return {
    rows: rows,
    totals: {
      totalHeadcount,
      totalBillable,
      topGroup,
      activeGroups
    }
  };
}


// --- PUBLIC: Headcount endpoints with billable included ---

function getHeadcountByGeo(filters) {
  const res = aggregateHeadcountBy(2, filters); // group by geo (col index 2)
  return {
    rows: res.rows.map(r => ({ geo: r.key, headcount: r.headcount, billable: r.billable })),
    totals: {
      totalHeadcount: res.totals.totalHeadcount,
      totalBillable: res.totals.totalBillable,
      topGeo: res.totals.topGroup,
      activeGeos: res.totals.activeGroups
    }
  };
}

function getHeadcountBySite(filters) {
  const res = aggregateHeadcountBy(4, filters); // group by site (col index 4)
  return {
    rows: res.rows.map(r => ({ site: r.key, headcount: r.headcount, billable: r.billable })),
    totals: {
      totalHeadcount: res.totals.totalHeadcount,
      totalBillable: res.totals.totalBillable,
      topSite: res.totals.topGroup,
      activeSites: res.totals.activeGroups
    }
  };
}

function getHeadcountByProgram(filters) {
  const res = aggregateHeadcountBy(3, filters); // group by program (col index 3)
  return {
    rows: res.rows.map(r => ({ program: r.key, headcount: r.headcount, billable: r.billable })),
    totals: {
      totalHeadcount: res.totals.totalHeadcount,
      totalBillable: res.totals.totalBillable,
      topProgram: res.totals.topGroup,
      activePrograms: res.totals.activeGroups
    }
  };
}


// --- Allocation loaders (read wider range so we don't break) ---
// These functions previously read fewer columns; expand to 22 columns for safety.

function getProgramAllocation(filters) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("HC");
  if (!sh) return { rows: [] };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { rows: [] };

  // Read columns A..V (22 columns) so future needed fields are available
  const data = sh.getRange(2, 1, lastRow - 1, 22).getValues();

  const monthFilter = (filters && filters.month) ? String(filters.month).trim() : '';
  const geoFilter = (filters && filters.geo) ? String(filters.geo).trim() : '';
  const siteFilter = (filters && filters.site) ? String(filters.site).trim() : '';
  const programFilter = (filters && filters.program) ? String(filters.program).trim() : '';
  const clusterLeadFilter = (filters && filters.clusterLead) ? String(filters.clusterLead).trim() : '';
  const directorFilter = (filters && filters.director) ? String(filters.director).trim() : '';
  const somFilter = (filters && filters.som) ? String(filters.som).trim() : '';

  const agg = {};

  data.forEach(row => {
    const geo = String(row[2] || '').trim();
    const program = String(row[3] || '').trim();
    const site = String(row[4] || '').trim();
    const month = String(row[6] || '').trim();
    const allocation = Number(row[11]) || 0; // column L (index 11)

    if (monthFilter && month !== monthFilter) return;
    if (geoFilter && geo !== geoFilter) return;
    if (siteFilter && site !== siteFilter) return;
    if (programFilter && program !== programFilter) return;
    if (clusterLeadFilter && String(row[8]||'').trim() !== clusterLeadFilter) return;
    if (directorFilter && String(row[9]||'').trim() !== directorFilter) return;
    if (somFilter && String(row[10]||'').trim() !== somFilter) return;

    agg[program] = (agg[program] || 0) + allocation;
  });

  const rows = Object.keys(agg).map(k => ({ program: k, allocation: agg[k] }))
    .sort((a,b)=>b.allocation - a.allocation);

  return { rows };
}

function getGeoAllocation(filters) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("HC");
  if (!sh) return { rows: [] };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { rows: [] };

  const data = sh.getRange(2, 1, lastRow - 1, 22).getValues();

  const monthFilter = (filters && filters.month) ? String(filters.month).trim() : '';
  const geoFilter = (filters && filters.geo) ? String(filters.geo).trim() : '';
  const siteFilter = (filters && filters.site) ? String(filters.site).trim() : '';
  const programFilter = (filters && filters.program) ? String(filters.program).trim() : '';
  const clusterLeadFilter = (filters && filters.clusterLead) ? String(filters.clusterLead).trim() : '';
  const directorFilter = (filters && filters.director) ? String(filters.director).trim() : '';
  const somFilter = (filters && filters.som) ? String(filters.som).trim() : '';

  const agg = {};

  data.forEach(row => {
    const geo = String(row[2] || '').trim();
    const allocation = Number(row[11]) || 0;

    if (monthFilter && String(row[6]||'').trim() !== monthFilter) return;
    if (geoFilter && geo !== geoFilter) return;
    if (siteFilter && String(row[4]||'').trim() !== siteFilter) return;
    if (programFilter && String(row[3]||'').trim() !== programFilter) return;
    if (clusterLeadFilter && String(row[8]||'').trim() !== clusterLeadFilter) return;
    if (directorFilter && String(row[9]||'').trim() !== directorFilter) return;
    if (somFilter && String(row[10]||'').trim() !== somFilter) return;

    agg[geo] = (agg[geo] || 0) + allocation;
  });

  return {
    rows: Object.keys(agg).map(k => ({ geo: k, allocation: agg[k] }))
  };
}

function getSiteAllocation(filters) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("HC");
  if (!sh) return { rows: [] };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { rows: [] };

  const data = sh.getRange(2, 1, lastRow - 1, 22).getValues();

  const monthFilter = (filters && filters.month) ? String(filters.month).trim() : '';
  const geoFilter = (filters && filters.geo) ? String(filters.geo).trim() : '';
  const siteFilter = (filters && filters.site) ? String(filters.site).trim() : '';
  const programFilter = (filters && filters.program) ? String(filters.program).trim() : '';
  const clusterLeadFilter = (filters && filters.clusterLead) ? String(filters.clusterLead).trim() : '';
  const directorFilter = (filters && filters.director) ? String(filters.director).trim() : '';
  const somFilter = (filters && filters.som) ? String(filters.som).trim() : '';

  const agg = {};

  data.forEach(row => {
    const site = String(row[4] || '').trim();
    const allocation = Number(row[11]) || 0;

    if (monthFilter && String(row[6]||'').trim() !== monthFilter) return;
    if (geoFilter && String(row[2]||'').trim() !== geoFilter) return;
    if (siteFilter && site !== siteFilter) return;
    if (programFilter && String(row[3]||'').trim() !== programFilter) return;
    if (clusterLeadFilter && String(row[8]||'').trim() !== clusterLeadFilter) return;
    if (directorFilter && String(row[9]||'').trim() !== directorFilter) return;
    if (somFilter && String(row[10]||'').trim() !== somFilter) return;

    agg[site] = (agg[site] || 0) + allocation;
  });

  return {
    rows: Object.keys(agg).map(k => ({ site: k, allocation: agg[k] }))
  };
}


function getGeoRatioDetails(filters, type) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('HC');
    if (!sh) return [];

    const values = sh.getDataRange().getValues();
    if (values.length <= 1) return [];

    // Column indexes (0-based)
   const COL = {
  geo: 2,        // C
  program: 3,    // D
  site: 4,       // E
  headcount: 5,  // F
  month: 6,      // G
  allocation: 11,// L
  TL: 12,        // M
  RTA: 13,       // N
  QA1: 14,       // O (QA CX)
  Trainer: 15,   // P
  OM: 16,        // Q
  QA2: 17,       // R (QA RCO)

  // ⭐ NEW BILLABLE COLUMNS ⭐
  billTL: 22,      // W
  billRTA: 23,     // X
  billQA1: 24,     // Y
  billQA2: 25,     // Z
  billTrainer: 26, // AA
  billOM: 27       // AB
};


    const typeToColMap = {
      TL: COL.TL,
      RTA: COL.RTA,
      QA1: COL.QA1,
      QA2: COL.QA2,
      Trainer: COL.Trainer,
      OM: COL.OM
    };

const billableColMap = {
  TL: COL.billTL,
  RTA: COL.billRTA,
  QA1: COL.billQA1,
  QA2: COL.billQA2,
  Trainer: COL.billTrainer,
  OM: COL.billOM
};


const billableIdx = billableColMap[type];
if (billableIdx === undefined) throw new Error("Invalid billable type");

    const actualIdx = typeToColMap[type];
    if (actualIdx === undefined) throw new Error("Invalid ratio type");

    const divisorMap = {
      TL: 15,
      RTA: 75,
      QA1: 50,
      QA2: 32,
      Trainer: 75,
      OM: 100
    };

    const divisor = divisorMap[type];

    const agg = {};

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const geo = row[COL.geo];
      if (!geo) continue;

      // ---------- FILTERS ----------
      if (filters.month && row[COL.month] !== filters.month) continue;
      if (filters.geo && filters.geo !== "All" && filters.geo !== row[COL.geo]) continue;
      if (filters.site && filters.site !== "All" && filters.site !== row[COL.site]) continue;
      if (filters.program && filters.program !== "All" && filters.program !== row[COL.program]) continue;
      if (filters.clusterLead && filters.clusterLead !== "All" && filters.clusterLead !== row[8]) continue;
      if (filters.director && filters.director !== "All" && filters.director !== row[9]) continue;
      if (filters.som && filters.som !== "All" && filters.som !== row[10]) continue;
      // ---------------------------------

      const headcount = parseFloatOrZero(row[COL.headcount]);
      const allocation = parseFloatOrZero(row[COL.allocation]);
      const actual = parseFloatOrZero(row[actualIdx]);
      const month = row[COL.month];

      if (!agg[geo]) {
  agg[geo] = {
    headcount: 0,
    allocation: 0,
    actual: 0,
    billable: 0,
    months: new Set()
  };
}



      agg[geo].headcount += headcount;
      agg[geo].allocation += allocation;
      agg[geo].actual += actual;
      agg[geo].billable += parseFloatOrZero(row[billableIdx]);

      if (month) agg[geo].months.add(String(month));
    }

    const results = [];

    for (const g in agg) {
      const item = agg[g];

      const actualRounded = Math.round(item.actual);
      const limit = Number((item.headcount / divisor).toFixed(2));
      const safeOver = Number((limit - actualRounded).toFixed(2));

      const monthVal = filters.month || (item.months.size === 1 ? [...item.months][0] : "All");

      results.push({
  geo: g,
  headcount: item.headcount,
  month: monthVal,
  allocation: item.allocation,
  actual: actualRounded,
  limit: limit,
  safeOver: safeOver,
  billable: item.billable
});



    }

    return results.sort((a, b) => b.headcount - a.headcount);

  } catch (e) {
    throw new Error("getGeoRatioDetails error: " + e.message);
  }
}


function getSiteRatioDetails(filters, type) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('HC');
    if (!sh) return [];

    const values = sh.getDataRange().getValues();
    if (values.length <= 1) return [];

    const COL = {
      geo: 2,
      program: 3,
      site: 4,
      headcount: 5,
      month: 6,
      allocation: 11,
      TL: 12,
      RTA: 13,
      QA1: 14,
      Trainer: 15,
      OM: 16,
      QA2: 17,

      // Billable columns
      billTL: 22,      // W
      billRTA: 23,     // X
      billQA1: 24,     // Y
      billQA2: 25,     // Z
      billTrainer: 26, // AA
      billOM: 27       // AB
    };

    const typeToColMap = {
      TL: COL.TL,
      RTA: COL.RTA,
      QA1: COL.QA1,
      QA2: COL.QA2,
      Trainer: COL.Trainer,
      OM: COL.OM
    };

    const billableColMap = {
      TL: COL.billTL,
      RTA: COL.billRTA,
      QA1: COL.billQA1,
      QA2: COL.billQA2,
      Trainer: COL.billTrainer,
      OM: COL.billOM
    };

    const actualIdx = typeToColMap[type];
    const billableIdx = billableColMap[type];

    if (actualIdx === undefined) throw new Error("Invalid ratio type");
    if (billableIdx === undefined) throw new Error("Invalid billable type");

    const divisorMap = {
      TL: 15,
      RTA: 75,
      QA1: 50,
      QA2: 32,
      Trainer: 75,
      OM: 100
    };

    const divisor = divisorMap[type];

    const agg = {};

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const site = row[COL.site];
      if (!site) continue;

      // ---------- FILTERS ----------
      if (filters.month && row[COL.month] !== filters.month) continue;
      if (filters.geo && filters.geo !== "All" && filters.geo !== row[COL.geo]) continue;
      if (filters.site && filters.site !== "All" && filters.site !== row[COL.site]) continue;
      if (filters.program && filters.program !== "All" && filters.program !== row[COL.program]) continue;
      if (filters.clusterLead && filters.clusterLead !== "All" && filters.clusterLead !== row[8]) continue;
      if (filters.director && filters.director !== "All" && filters.director !== row[9]) continue;
      if (filters.som && filters.som !== "All" && filters.som !== row[10]) continue;
      // ---------------------------------

      const headcount = parseFloatOrZero(row[COL.headcount]);
      const actual = parseFloatOrZero(row[actualIdx]);
      const billable = parseFloatOrZero(row[billableIdx]);
      const month = row[COL.month];

      if (!agg[site]) {
        agg[site] = { headcount: 0, actual: 0, billable: 0, months: new Set() };
      }

      agg[site].headcount += headcount;
      agg[site].actual += actual;
      agg[site].billable += billable;
      if (month) agg[site].months.add(String(month));
    }

    const results = [];

    for (const s in agg) {
      const item = agg[s];

      const actualRounded = Math.round(item.actual);
      const limit = Number((item.headcount / divisor).toFixed(2));
      const safeOver = Number((limit - actualRounded).toFixed(2));
      const monthVal = filters.month || (item.months.size === 1 ? [...item.months][0] : "All");

      results.push({
        site: s,
        headcount: item.headcount,
        month: monthVal,
        actual: actualRounded,
        limit: limit,
        safeOver: safeOver,
        billable: item.billable
      });
    }

    return results.sort((a, b) => b.headcount - a.headcount);

  } catch (e) {
    throw new Error("getSiteRatioDetails error: " + e.message);
  }
}


function getProgramRatioDetails(filters, type) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('HC');
    if (!sh) return [];

    const values = sh.getDataRange().getValues();
    if (values.length <= 1) return [];

    const COL = {
      geo: 2,
      program: 3,
      site: 4,
      headcount: 5,
      month: 6,
      allocation: 11,
      TL: 12,
      RTA: 13,
      QA1: 14,
      Trainer: 15,
      OM: 16,
      QA2: 17,

      // Billable columns
      billTL: 22,      // W
      billRTA: 23,     // X
      billQA1: 24,     // Y
      billQA2: 25,     // Z
      billTrainer: 26, // AA
      billOM: 27       // AB
    };

    const typeToColMap = {
      TL: COL.TL,
      RTA: COL.RTA,
      QA1: COL.QA1,
      QA2: COL.QA2,
      Trainer: COL.Trainer,
      OM: COL.OM
    };

    const billableColMap = {
      TL: COL.billTL,
      RTA: COL.billRTA,
      QA1: COL.billQA1,
      QA2: COL.billQA2,
      Trainer: COL.billTrainer,
      OM: COL.billOM
    };

    const actualIdx = typeToColMap[type];
    const billableIdx = billableColMap[type];

    if (actualIdx === undefined) throw new Error("Invalid ratio type");
    if (billableIdx === undefined) throw new Error("Invalid billable type");

    const divisorMap = {
      TL: 15,
      RTA: 75,
      QA1: 50,
      QA2: 32,
      Trainer: 75,
      OM: 100
    };

    const divisor = divisorMap[type];

    const agg = {};

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const program = row[COL.program];
      if (!program) continue;

      // ---------- FILTERS ----------
      if (filters.month && row[COL.month] !== filters.month) continue;
      if (filters.geo && filters.geo !== "All" && filters.geo !== row[COL.geo]) continue;
      if (filters.site && filters.site !== "All" && filters.site !== row[COL.site]) continue;
      if (filters.program && filters.program !== "All" && filters.program !== row[COL.program]) continue;
      if (filters.clusterLead && filters.clusterLead !== "All" && filters.clusterLead !== row[8]) continue;
      if (filters.director && filters.director !== "All" && filters.director !== row[9]) continue;
      if (filters.som && filters.som !== "All" && filters.som !== row[10]) continue;
      // ---------------------------------

      const headcount = parseFloatOrZero(row[COL.headcount]);
      const actual = parseFloatOrZero(row[actualIdx]);
      const billable = parseFloatOrZero(row[billableIdx]);
      const month = row[COL.month];

      if (!agg[program]) {
        agg[program] = { headcount: 0, actual: 0, billable: 0, months: new Set() };
      }

      agg[program].headcount += headcount;
      agg[program].actual += actual;
      agg[program].billable += billable;
      if (month) agg[program].months.add(String(month));
    }

    const results = [];

    for (const p in agg) {
      const item = agg[p];

      const actualRounded = Math.round(item.actual);
      const limit = Number((item.headcount / divisor).toFixed(2));
      const safeOver = Number((limit - actualRounded).toFixed(2));
      const monthVal = filters.month || (item.months.size === 1 ? [...item.months][0] : "All");

      results.push({
        program: p,
        headcount: item.headcount,
        month: monthVal,
        actual: actualRounded,
        limit: limit,
        safeOver: safeOver,
        billable: item.billable
      });
    }

    return results.sort((a, b) => b.headcount - a.headcount);

  } catch (e) {
    throw new Error("getProgramRatioDetails error: " + e.message);
  }
}






function parseFloatOrZero(v) {
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

/***********************
 * Others modal helpers
 ***********************/

// helper: parse month string "October 2025" -> Date and return string for previous month in same format
function previousMonthString(monthStr) {
  if (!monthStr) return '';
  try {
    // some month strings may be "Oct 2025" or "October 2025"
    var d = new Date(monthStr + " 1");
    if (isNaN(d)) return '';
    // subtract one month
    d.setMonth(d.getMonth() - 1);
    // format like "October 2025"
    var opts = { year: 'numeric', month: 'long' };
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM yyyy');
  } catch (e) {
    return '';
  }
}

// safe number parse
function parseFloatOrZero(v) {
  var n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

/**
 * Generic function to build Others-type results.
 * level: "geo" | "site" | "program"
 * type: "OthersOps" | "OthersComp" | "OthersLearn"
 *
 * Column mappings (0-based):
 *  C = 2 (geo), D = 3 (program), E = 4 (site), F = 5 (headcount), G = 6 (month)
 *  Others Ops actual = column S (0-based 18)
 *  Others Comp actual = column T (19)
 *  Others Learn actual = column U (20)
 */
function getOthersGenericDetails(filters, type, level) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('HC');
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];

  var COL = {
    geo: 2,
    program: 3,
    site: 4,
    headcount: 5,
    month: 6,
    ops: 18,    // S
    comp: 20,   // U
    learn: 19   // T
  };

  var actualCol;
  if (type === 'OthersOps') actualCol = COL.ops;
  else if (type === 'OthersComp') actualCol = COL.comp;
  else if (type === 'OthersLearn') actualCol = COL.learn;
  else throw new Error('Invalid Others type: ' + type);

  var keyCol = (level === 'geo') ? COL.geo :
               (level === 'site') ? COL.site : COL.program;

  var monthFilter = filters && filters.month ? String(filters.month).trim() : '';
  var prevMonth = monthFilter ? previousMonthString(monthFilter) : '';

  var agg = {};

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var key = row[keyCol];
    if (!key) continue;
    key = String(key).trim();

    var rowMonth = String(row[COL.month] || '').trim();

    // Apply geo/site/program + director/clusterLead/som filters
    if (filters) {
      if (filters.geo && filters.geo !== '' && filters.geo !== 'All' && filters.geo !== row[COL.geo]) continue;
      if (filters.site && filters.site !== '' && filters.site !== 'All' && filters.site !== row[COL.site]) continue;
      if (filters.program && filters.program !== '' && filters.program !== 'All' && filters.program !== row[COL.program]) continue;
      if (filters.clusterLead && filters.clusterLead !== '' && filters.clusterLead !== 'All' && filters.clusterLead !== row[8]) continue;
      if (filters.director && filters.director !== '' && filters.director !== 'All' && filters.director !== row[9]) continue;
      if (filters.som && filters.som !== '' && filters.som !== 'All' && filters.som !== row[10]) continue;
    }

    if (!agg[key]) {
      agg[key] = { headcountThis:0, headcountPrev:0, actualThis:0, actualPrev:0, months:new Set() };
    }

    var hc = parseFloatOrZero(row[COL.headcount]);
    var act = parseFloatOrZero(row[actualCol]);

    if (monthFilter) {
      if (rowMonth === monthFilter) {
        agg[key].headcountThis += hc;
        agg[key].actualThis += act;
      } else if (rowMonth === prevMonth) {
        agg[key].headcountPrev += hc;
        agg[key].actualPrev += act;
      }
    } else {
      agg[key].headcountThis += hc;
      agg[key].actualThis += act;
      if (rowMonth === prevMonth) {
        agg[key].headcountPrev += hc;
        agg[key].actualPrev += act;
      }
    }

    if (rowMonth) agg[key].months.add(rowMonth);
  }

  var results = [];
  for (var k in agg) {
    var it = agg[k];

    var monthVal = monthFilter || (it.months.size === 1 ? [...it.months][0] : 'All');

    var hcThis = it.headcountThis || 0;

    var pctThis = (hcThis > 0) ? (it.actualThis / hcThis) : 0;
    var pctPrev = (it.headcountPrev > 0) ? (it.actualPrev / it.headcountPrev) : 0;

    results.push({
      key: k,
      headcount: hcThis,
      month: monthVal,

      // 🔥 FIXED — rounded values
      actual: Math.round(it.actualThis || 0),
      pct: pctThis,

      prevMonth: prevMonth,
      prevActual: Math.round(it.actualPrev || 0), // 🔥 FIXED
      prevPct: pctPrev,

      overUnder: pctThis - pctPrev
    });
  }

  results.sort(function(a,b){ return b.headcount - a.headcount; });
  return results;
}



// Public wrappers (Geo / Site / Program) for Apps Script client calls
function getOthersOperationsDetails(filters, level) { return getOthersGenericDetails(filters, 'OthersOps', level); }
function getOthersLearningDetails(filters, level) { return getOthersGenericDetails(filters, 'OthersLearn', level); }


/**
 * Normalize month value from Allocation sheet.
 * Handles: "September 2025", " Sep 2025 ", "11/1/2025", date objects, timestamps.
 */
function normalizeMonth(value) {
  if (!value) return "";

  // Case 1: If it is a Date object (Apps Script returns Date objects for date-formatted cells)
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "MMMM yyyy");
  }

  // Case 2: If it's a string
  if (typeof value === "string") {
    const cleaned = value.trim();

    // Already in "Month Year" format (e.g., "November 2025")
    if (/^[A-Za-z]+\s+\d{4}$/.test(cleaned)) return cleaned;

    // Handle date strings like "11/1/2025" or "2025-11-01"
    const d = new Date(cleaned);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "MMMM yyyy");
    }

    return cleaned;
  }

  return "";
}


function getProgramAllocationDetails(program, monthStr) {
  if (!program || !monthStr) return { rows: [], totalCurrent: 0, totalPrev: 0 };

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Allocation");
  if (!sh) return { rows: [], totalCurrent: 0, totalPrev: 0 };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { rows: [], totalCurrent: 0, totalPrev: 0 };

  const data = sh.getRange(2, 1, lastRow - 1, 29).getValues();

  const programKey = program.toString().trim().toLowerCase();
  const monthKey = normalizeMonth(monthStr).toLowerCase();

  // Determine previous month
  const prevMonthKey = getPreviousMonthString(monthStr).toLowerCase();

  const results = [];

  let totalCurrent = 0;
  let totalPrev = 0;  // <-- true program-level total

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rProgram = String(row[0] || "").trim().toLowerCase();
    // Normalize month to handle both Date objects and strings
    const rMonth = normalizeMonth(row[27]).toLowerCase();

    if (rProgram !== programKey) continue;

    const alloc = parseFloatOrZero(row[12]);

    // ============ SELECTED MONTH ROWS =============
    if (rMonth === monthKey) {
      results.push({
        program: row[0],
        site: row[1],
        tower: row[2],
        eid: row[3],
        name: row[4],
        position: row[5],
        status: row[9],
        allocation: alloc,
        month: row[27],
         gprGroup: row[28] || "",   // <-- NEW FIELD
        prevAlloc: 0 // will be filled below
      });

      totalCurrent += alloc;
    }

    // ============ PREVIOUS MONTH TOTAL =============
    if (rMonth === normalizeMonth(prevMonthKey).toLowerCase()) {
      totalPrev += alloc;
    }
  }

  // Match prevAlloc by EID
  results.forEach(r => {
    const match = data.find(row =>
      String(row[3]).trim() === String(r.eid).trim() &&
      normalizeMonth(row[27]).toLowerCase() === normalizeMonth(prevMonthKey).toLowerCase()
    );
    r.prevAlloc = match ? parseFloatOrZero(match[12]) : 0;
  });

  return {
    rows: results,
    totalCurrent,
    totalPrev
  };
}

function getPreviousMonthString(monthStr) {
  const d = new Date(monthStr + " 1");
  d.setMonth(d.getMonth() - 1);
  return d.toLocaleString("en-US", { month: "long", year: "numeric" });
}



// Helper (kept simple)
function parseFloatOrZero(v) {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
}

/**
 * Return allocations for a given EID + Month from the Allocation sheet.
 * Also returns totalAllocation (sum of column M) for the matched rows.
 *
 * Columns we read:
 * A (0) Program
 * B (1) Location
 * C (2) Tower
 * D (3) EID
 * E (4) Employee Name
 * F (5) Position
 * ...
 * M (12) Allocation
 * ...
 * AB (27) Month
 * AC (28) GPR Group
 */
function getEmployeeMonthlyDetails(eid, monthStr) {
  if (!eid || !monthStr) return { rows: [], totalAllocation: 0 };

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Allocation");
  if (!sh) return { rows: [], totalAllocation: 0 };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { rows: [], totalAllocation: 0 };

  // Read columns A..AC => 29 columns (indices 0..28)
  const NUM_COLS = 29;
  const data = sh.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();

  const eidKey = String(eid || "").trim();
  const monthKey = String(monthStr || "").trim();

  const results = [];
  let totalAlloc = 0;

  data.forEach((row, idx) => {
    const rEID = String(row[3] || "").trim();
    const rMonth = String(row[27] || "").trim();
    if (rEID === "" || rMonth === "") return;
    if (rEID === eidKey && rMonth === monthKey) {
      const alloc = parseFloatOrZero(row[12]); // column M index 12
      totalAlloc += alloc;

      results.push({
        tower: row[2] || "",
        location: row[1] || "",
        eid: row[3] || "",
        name: row[4] || "",
        position: row[5] || "",
        program: row[0] || "",
        allocation: alloc,
        month: row[27] || "",
        gprGroup: row[28] || ""
      });
    }
  });

  return { rows: results, totalAllocation: totalAlloc };
}


function getPreviousMonthLabel(monthStr) {
  try {
    const d = new Date(monthStr + " 1");
    if (isNaN(d)) return "";

    d.setMonth(d.getMonth() - 1);

    const options = { year: "numeric", month: "long" };
    return d.toLocaleDateString("en-US", options);
  } catch (e) {
    return "";
  }
}

function getPrevAllocForEID(data, eid, prevMonthKey) {
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rEID = String(row[3] || "").trim();
    const rMonth = String(row[27] || "").trim().toLowerCase();

    if (rEID === eid && rMonth === prevMonthKey.toLowerCase()) {
      return parseFloatOrZero(row[12]); // allocation column M
    }
  }
  return 0;
}

function getProgramPrevMonthDetails(program, monthStr) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Allocation");
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, 28).getValues();

  // Determine previous month string
  const prevMonth = getPrevMonthString(monthStr);

  const programKey = program.toLowerCase();
  const prevMonthKey = prevMonth.toLowerCase();

  const results = [];

  data.forEach(row => {
    if (
      String(row[0]).trim().toLowerCase() === programKey &&
      String(row[27]).trim().toLowerCase() === prevMonthKey
    ) {
      results.push({
        program: row[0],
        site: row[1],
        tower: row[2],
        eid: row[3],
        name: row[4],
        position: row[5],
        status: row[9],
        allocation: parseFloatOrZero(row[12]),
        month: row[27]
      });
    }
  });

  return results;
}

// Convert "October 2025" → "September 2025"
function getPrevMonthString(monthStr) {
  const d = new Date(monthStr + " 1");
  d.setMonth(d.getMonth() - 1);
  const opts = { month: "long", year: "numeric" };
  return d.toLocaleDateString("en-US", opts);
}


/**
 * Send allocation table by email as CSV attachment
 */
function sendAllocationEmail(emailList, program, monthStr, tableRows) {
  if (!emailList || !emailList.length) return;

  // ==== Build CSV (clean text only, no icons) ====
  let csv = "";

  tableRows.forEach(row => {

    // row = array of 10 fields:
    // [Program, Site, Tower, EID, Name, Position, Status, AllocationPct, Variance, Month]

    let cleanRow = row.map((v, index) => {

      // CLEAN VARIANCE COLUMN (index 8)
      if (index === 8) {
        // Remove arrow icons and just keep number
        const cleaned = String(v)
          .replace(/▲|▼|●/g, "")   // remove icons
          .trim();

        return `"${cleaned.replace(/"/g, '""')}"`;
      }

      // Normal column
      return `"${String(v).replace(/"/g, '""')}"`;
    });

    csv += cleanRow.join(",") + "\n";
  });

  const blob = Utilities.newBlob(
    csv,
    "text/csv",
    `${program}_Allocation_${monthStr}.csv`
  );

  // ==== Email Subject (with timestamp) ====
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "MM/dd/yyyy hh:mm a"
  );

  const subject = `${program} Allocation Details – ${monthStr} [${timestamp}]`;

  // ==== Email Body ====
  const body =
    `Hello,\n\n` +
    `Please find attached the allocation details for ${program} for ${monthStr}.\n\n` +
    `If you have any questions, please reach out to fpa.cfm@ubiquity.com.\n\n` +
    `Thank you,\n` +
    `FP&A Team`;

  // ==== Send to multiple recipients ====
  emailList.forEach(email => {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      attachments: [blob]
    });
  });
}


function getEmployeeDetailsFiltered(month, category) {

  const cache = CacheService.getDocumentCache();
  const cacheKey = `empDetails_${month}_${category}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    try { return JSON.parse(cached); }
    catch (err) {}
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Allocation");
  if (!sh) return [];

  const values = sh.getDataRange().getValues();

  const COL = {
    site: 1,       // B
    eid: 3,        // D
    name: 4,       // E
    position: 5,   // F
    tagging: 26,   // AA
    gpr: 28,       // AC
    month: 27,     // AB
    allocation: 12 // M
  };

  const resultsMap = {}; // Unique by EID

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const rMonth = String(row[COL.month] || "").trim();
    const rTag = String(row[COL.tagging] || "").trim();
    const rGpr = String(row[COL.gpr] || "").trim();
    const alloc = parseFloat(row[COL.allocation]) || 0;
    const eid = row[COL.eid];

    // FILTERS
    if (rMonth !== month) continue;
    if (rTag !== category) continue;
    if (rGpr !== "IL") continue;
    if (!eid) continue;
    if (!alloc || alloc === 0) continue; // EXCLUDE ZERO ALLOCATION

    // UNIQUE BY EID
    resultsMap[eid] = {
      eid: eid,
      name: row[COL.name],
      position: row[COL.position],
      tagging: rTag,
      gpr: rGpr,
      month: rMonth
    };
  }

  const results = Object.values(resultsMap);

  // Save to cache for 60 seconds
  cache.put(cacheKey, JSON.stringify(results), 60);

  return results;
}
