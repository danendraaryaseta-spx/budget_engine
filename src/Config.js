var LH = LH || {};
LH.Config = {
  timezone: 'Asia/Jakarta',
  chunkCells: 16000,
  executionMs: 220000,
  maxDatasetCells: 2500000,
  maxWorkbookCells: 10000000,
  cachePrefix: '__src_',
  stagePrefix: '__stage_',
  // These are import-connection canaries, not budget inputs. Retire them explicitly.
  registrySheet: 'Base Assumptions',
  registryColumn: 1,
  registryFirstRow: 33,
  registryLastRow: 49,
  // Sources already refreshed by the previous script, rather than by formulas.
  feeds: [
    { key: 'schedule', sourceCell: "'Base Assumptions'!B33", tab: 'Compile Normal Schedule', range: 'D3:AQ', dest: 'Schedule v5', row: 3, col: 1, header: 2 },
    { key: 'dc_schedule', sourceCell: "'Base Assumptions'!B34", latestDated: true, range: 'A1:AW', dest: 'Schedule DC to DC', row: 1, col: 1, header: 8 },
    { key: 'fm_ado', sourceCell: "'RAW_ADO FM Projection'!A1", tab: 'FM Average ADO', range: 'A3:B', dest: 'RAW_ADO FM Projection', row: 3, col: 1, header: 2 },
    { key: 'lm_ado', sourceSetting: 'LM_SOURCE_ID', tab: 'Final Proj (Budget)', range: 'A19:ZZ', dest: 'RAW_ADO LM Projection', row: 5, col: 2,
      sourceHeader: 18, columnsByHeader: ['station_name','Region','Alloc','Source Hub'].concat(
        ['26','27'].reduce(function (a, y) { return a.concat(['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'].map(function (m) { return m + '-' + y; })); }, [])) }
  ],
  // Seed once from existing hardcoded inputs, then edit these tables only in the hub.
  // Mixed formula/input areas are deliberately not bulk-copied.
  manualInputs: [
    { key: 'ado_pairs', sheet: 'ADO Projection', a1: 'C6:D2428' },
    { key: 'ado_std', sheet: 'ADO Projection', a1: 'AB6:AV2428' },
    { key: 'ado_eco', sheet: 'ADO Projection', a1: 'AX6:BR2428' },
    { key: 'weight_std', sheet: 'ADO Projection', a1: 'EH6:FB2428' },
    { key: 'weight_eco', sheet: 'ADO Projection', a1: 'FD6:FX2428' },
    { key: 'bulky_origin', sheet: 'ADO Bulky Proj_Origin', a1: 'B4:W57' },
    { key: 'bulky_dest', sheet: 'ADO Bulky Proj_Destination', a1: 'B4:W59' },
    { key: 'vehicle_capacity', sheet: 'Capacity Unit', a1: 'A1:C34' },
    { key: 'distance', sheet: 'Raw_Distance', a1: 'A1:D275' }
  ]
};
