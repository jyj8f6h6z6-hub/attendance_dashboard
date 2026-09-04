(() => {
  'use strict';

  const STATUS = {
    WEEKLY: '週週聚會',
    REGULAR: '常聚會',
    OCCASIONAL: '偶聚會',
    SPORADIC: '零星聚會',
    INACTIVE: '近期未聚會'
  };

  const STATUS_ORDER = [
    STATUS.WEEKLY,
    STATUS.REGULAR,
    STATUS.OCCASIONAL,
    STATUS.SPORADIC,
    STATUS.INACTIVE
  ];

  const STATUS_CLASS = {
    [STATUS.WEEKLY]: 'weekly',
    [STATUS.REGULAR]: 'regular',
    [STATUS.OCCASIONAL]: 'occasional',
    [STATUS.SPORADIC]: 'sporadic',
    [STATUS.INACTIVE]: 'inactive'
  };

  const MS_WEEK =
    7 * 24 * 60 * 60 * 1000;


  const state = {
    fileName: '',
    rows: [],
    dateColumns: [],
    people: [],
    windowWeeks: 12,
    indexes: {},
    recentCols: [],

    baptismFileName: '',
    baptismByName: new Map(),

    selectedGroups: new Set(),
    selectedStatuses: new Set(),

    /*
     * 羣組成長趨勢：按鈕可複選。
     * 第一次載入時預設顯示所有可用羣組。
     */
    selectedTrendGroups: new Set(),
    trendGroupsInitialized: false,

    /*
     * 人員明細區域複選篩選
     * 只影響人員明細，不影響四張分析圖。
     */
    selectedPeopleDistricts: new Set(),
    selectedPeopleSmallDistricts: new Set(),

    /*
     * 圖表點擊連動
     */
    chartType: '',
    chartGroup: '',
    chartDistrict: ''
  };


  const $ =
    id =>
      document.getElementById(id);


  const els = {
    fileInput: $('fileInput'),
    fileName: $('fileName'),
    baptismFileInput: $('baptismFileInput'),
    baptismFileName: $('baptismFileName'),
    error: $('errorMessage'),
    dashboard: $('dashboard'),
    clearBtn: $('clearBtn'),

    infoFile: $('infoFile'),
    infoPeriod: $('infoPeriod'),
    infoWeeks: $('infoWeeks'),
    infoPeople: $('infoPeople'),

    ruleText: $('ruleText'),

    summaryCards: $('summaryCards'),

    trendDistrictFilter:
      $('trendDistrictFilter'),

    trendSmallDistrictFilter:
      $('trendSmallDistrictFilter'),

    trendDistrictButtons:
      $('trendDistrictButtons'),

    trendSmallDistrictButtons:
      $('trendSmallDistrictButtons'),

    trendSmallDistrictGroup:
      $('trendSmallDistrictGroup'),

    trendScopeLabel:
      $('trendScopeLabel'),

    trendLatestCount:
      $('trendLatestCount'),

    trendAverageCount:
      $('trendAverageCount'),

    trendChangeText:
      $('trendChangeText'),

    trendChart:
      $('trendChart'),

    groupTrendScopeLabel:
      $('groupTrendScopeLabel'),

    groupTrendButtons:
      $('groupTrendButtons'),

    groupTrendSelectedText:
      $('groupTrendSelectedText'),

    groupTrendChart:
      $('groupTrendChart'),

    overallBar: $('overallBar'),
    overallLegend: $('overallLegend'),

    districtBody:
      document.querySelector(
        '#districtTable tbody'
      ),

    searchInput: $('searchInput'),
    analysisDistrictButtons: $('analysisDistrictButtons'),
    analysisSmallDistrictButtons: $('analysisSmallDistrictButtons'),
    analysisSmallDistrictRow: $('analysisSmallDistrictRow'),
    analysisGroupButtons: $('analysisGroupButtons'),
    analysisStatusButtons: $('analysisStatusButtons'),
    analysisBelieverButtons: $('analysisBelieverButtons'),
    districtFilter: $('districtFilter'),
    smallDistrictFilter:
      $('smallDistrictFilter'),

    totalAttendanceFilter:
      $('totalAttendanceFilter'),

    absenceWeeksFilter:
      $('absenceWeeksFilter'),

    populationOriginalCount:
      $('populationOriginalCount'),

    populationActiveCount:
      $('populationActiveCount'),

    populationExcludedCount:
      $('populationExcludedCount'),

    newBelieverFilter:
      $('newBelieverFilter'),

    groupMultiSelect:
      $('groupMultiSelect'),

    groupFilterSummary:
      $('groupFilterSummary'),

    groupFilterOptions:
      $('groupFilterOptions'),

    clearGroupFilter:
      $('clearGroupFilter'),

    closeGroupFilter:
      $('closeGroupFilter'),

    statusMultiSelect:
      $('statusMultiSelect'),

    statusFilterSummary:
      $('statusFilterSummary'),

    statusFilterOptions:
      $('statusFilterOptions'),

    clearStatusFilter:
      $('clearStatusFilter'),

    closeStatusFilter:
      $('closeStatusFilter'),

    resultCount:
      $('resultCount'),

    peopleBody:
      document.querySelector(
        '#peopleTable tbody'
      ),

    peopleDistrictMultiSelect:
      $('peopleDistrictMultiSelect'),

    peopleDistrictSummary:
      $('peopleDistrictSummary'),

    peopleDistrictOptions:
      $('peopleDistrictOptions'),

    clearPeopleDistrictFilter:
      $('clearPeopleDistrictFilter'),

    closePeopleDistrictFilter:
      $('closePeopleDistrictFilter'),

    peopleSmallDistrictMultiSelect:
      $('peopleSmallDistrictMultiSelect'),

    peopleSmallDistrictSummary:
      $('peopleSmallDistrictSummary'),

    peopleSmallDistrictOptions:
      $('peopleSmallDistrictOptions'),

    clearPeopleSmallDistrictFilter:
      $('clearPeopleSmallDistrictFilter'),

    closePeopleSmallDistrictFilter:
      $('closePeopleSmallDistrictFilter'),

    clearPeopleAreaFilter:
      $('clearPeopleAreaFilter'),

    chartFilterIndicator:
      $('chartFilterIndicator'),

    chartFilterText:
      $('chartFilterText'),

    clearChartFilter:
      $('clearChartFilter'),

    peoplePanel:
      $('peoplePanel')
  };


  function excelDateToDate(value) {

    if (
      value instanceof Date &&
      !Number.isNaN(
        value.getTime()
      )
    ) {
      return startOfDay(value);
    }


    if (
      typeof value === 'number'
    ) {

      const parts =
        XLSX.SSF.parse_date_code(
          value
        );

      if (!parts) {
        return null;
      }

      return new Date(
        parts.y,
        parts.m - 1,
        parts.d
      );
    }


    if (
      typeof value === 'string'
    ) {

      const s =
        value.trim();

      if (!s) {
        return null;
      }


      let m =
        s.match(
          /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/
        );

      if (m) {
        return new Date(
          Number(m[1]),
          Number(m[2]) - 1,
          Number(m[3])
        );
      }


      m =
        s.match(
          /^(\d{4})[-\/.](\d{1,2})$/
        );

      if (m) {
        return new Date(
          Number(m[1]),
          Number(m[2]) - 1,
          1
        );
      }
    }


    return null;
  }


  function startOfDay(d) {

    return new Date(
      d.getFullYear(),
      d.getMonth(),
      d.getDate()
    );
  }


  /*
   * 新版點名系統的每週日期欄位以「週一」代表該週，
   * 但本系統分析的是主日聚會，因此遇到週一時自動 +6 天，
   * 換算成同一週的主日。
   *
   * 舊版報表本來就是週日日期，因此保持不變。
   */
  function normalizeAttendanceDateToSunday(d) {

    if (!d) {
      return null;
    }

    const normalized =
      startOfDay(d);

    // JavaScript：0 = 週日，1 = 週一
    if (normalized.getDay() === 1) {
      normalized.setDate(
        normalized.getDate() + 6
      );
    }

    return normalized;
  }


  function baptismValueToDate(value) {

    if (
      typeof value === 'string'
    ) {

      const s =
        value.trim();


      const m =
        s.match(
          /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/
        );


      if (m) {

        const year = Number(m[1]);
        const month = Number(m[2]);
        const day = Number(m[3]);


        if (
          month < 1 ||
          month > 12 ||
          day < 1 ||
          day > 31
        ) {
          return null;
        }


        const d =
          new Date(
            year,
            month - 1,
            day
          );


        if (
          d.getFullYear() !== year ||
          d.getMonth() !== month - 1 ||
          d.getDate() !== day
        ) {
          return null;
        }


        return d;
      }
    }


    return excelDateToDate(
      value
    );
  }


  function isAttendance(value) {

    if (
      typeof value === 'number'
    ) {
      return value > 0;
    }


    if (
      typeof value === 'boolean'
    ) {
      return value;
    }


    if (
      typeof value === 'string'
    ) {

      const s =
        value
          .trim()
          .toLowerCase();

      return [
        '1',
        'v',
        '✓',
        '✔',
        'true',
        '有',
        '出席'
      ].includes(s);
    }


    return false;
  }


  function fmtDate(d) {

    return d
      ? `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`
      : '—';
  }


  function analyzeWorkbook(
    arrayBuffer,
    fileName
  ) {

    const workbook =
      XLSX.read(
        arrayBuffer,
        {
          type: 'array',
          cellDates: false,
          raw: true
        }
      );


    if (
      !workbook.SheetNames.length
    ) {
      throw new Error(
        '找不到工作表。'
      );
    }


    /*
     * 新版點名系統固定讀取「出席」工作表；
     * 舊版 export.xlsx 沒有此工作表時，維持讀第一張工作表。
     */
    const attendanceSheetName =
      workbook.SheetNames.includes('出席')
        ? '出席'
        : workbook.SheetNames[0];


    const sheet =
      workbook.Sheets[
        attendanceSheetName
      ];


    const rows =
      XLSX.utils.sheet_to_json(
        sheet,
        {
          header: 1,
          raw: true,
          defval: ''
        }
      );


    if (
      rows.length < 2
    ) {
      throw new Error(
        '出席工作表沒有足夠資料。'
      );
    }


    const header =
      rows[0].map(
        v =>
          String(
            v ?? ''
          ).trim()
      );


    const findHeaderIndex =
      names =>
        names
          .map(
            name =>
              header.indexOf(name)
          )
          .find(
            i => i >= 0
          ) ?? -1;


    const indexes = {
      district:
        findHeaderIndex(['大區']),
      smallDistrict:
        findHeaderIndex(['小區']),
      name:
        findHeaderIndex(['姓名']),
      group:
        findHeaderIndex(['羣組', '群組'])
    };


    const missing = [];

    if (indexes.district < 0) {
      missing.push('大區');
    }

    if (indexes.smallDistrict < 0) {
      missing.push('小區');
    }

    if (indexes.name < 0) {
      missing.push('姓名');
    }

    if (indexes.group < 0) {
      missing.push('羣組／群組');
    }


    if (
      missing.length
    ) {
      throw new Error(
        `缺少必要欄位：${missing.join('、')}`
      );
    }


    const baptismIndex =
      findHeaderIndex([
        '受浸日期',
        '受浸日',
        '受浸'
      ]);


    const today =
      new Date();


    today.setHours(
      23,
      59,
      59,
      999
    );


    const dateColumns = [];


    rows[0].forEach(
      (value, col) => {

        const rawDate =
          excelDateToDate(
            value
          );

        const d =
          normalizeAttendanceDateToSunday(
            rawDate
          );

        if (
          d &&
          d <= today
        ) {
          dateColumns.push({
            col,
            date: d
          });
        }
      }
    );


    dateColumns.sort(
      (a, b) =>
        a.date - b.date
    );


    if (
      !dateColumns.length
    ) {
      throw new Error(
        '找不到可辨識的聚會日期欄位。'
      );
    }


    state.fileName =
      fileName;

    state.rows =
      rows;

    state.dateColumns =
      dateColumns;


    state.indexes = {
      ...indexes,
      baptism:
        baptismIndex
    };


    /*
     * 匯入新點名報表時，
     * 清除上一份圖表點擊與人員明細區域篩選狀態。
     * 已匯入的受浸報表保留，可直接重新配對。
     */
    state.selectedPeopleDistricts =
      new Set();

    state.selectedPeopleSmallDistricts =
      new Set();

    state.chartType = '';
    state.chartGroup = '';
    state.chartDistrict = '';

    state.selectedTrendGroups =
      new Set();

    state.trendGroupsInitialized =
      false;

    updateChartFilterUI();


    recalculate();


    if (
      baptismIndex < 0 &&
      !state.baptismByName.size
    ) {

      els.error.textContent =
        '提醒：點名報表沒有受浸日期，請再匯入受浸報表以顯示受浸日並判斷初信。';
    }
  }


  function analyzeBaptismWorkbook(
    arrayBuffer,
    fileName
  ) {

    const workbook =
      XLSX.read(
        arrayBuffer,
        {
          type: 'array',
          cellDates: false,
          raw: true
        }
      );


    const baptismSheetName =
      workbook.SheetNames.find(
        name =>
          String(name).trim() ===
          '受浸名單'
      );


    if (!baptismSheetName) {
      throw new Error(
        '受浸報表找不到第 5 個工作表「受浸名單」。'
      );
    }


    const rows =
      XLSX.utils.sheet_to_json(
        workbook.Sheets[
          baptismSheetName
        ],
        {
          header: 1,
          raw: true,
          defval: ''
        }
      );


    if (
      rows.length < 2
    ) {
      throw new Error(
        '「受浸名單」工作表沒有足夠資料。'
      );
    }


    const header =
      rows[0].map(
        v =>
          String(
            v ?? ''
          ).trim()
      );


    const nameIndex =
      header.indexOf('姓名');


    const baptismIndex =
      ['受浸日', '受浸日期', '受浸']
        .map(
          name =>
            header.indexOf(name)
        )
        .find(
          i => i >= 0
        ) ?? -1;


    if (
      nameIndex < 0 ||
      baptismIndex < 0
    ) {
      throw new Error(
        '「受浸名單」缺少「姓名」或「受浸日」欄位。'
      );
    }


    const baptismByName =
      new Map();


    rows
      .slice(1)
      .forEach(
        row => {

          const name =
            String(
              row[nameIndex] ?? ''
            ).trim();


          const baptismDate =
            baptismValueToDate(
              row[baptismIndex]
            );


          if (
            !name ||
            !baptismDate
          ) {
            return;
          }


          /*
           * 同姓名若在受浸名單出現多筆，採較新的受浸日。
           * 一般情況不會影響單一姓名資料。
           */
          const existing =
            baptismByName.get(name);


          if (
            !existing ||
            baptismDate > existing
          ) {
            baptismByName.set(
              name,
              baptismDate
            );
          }
        }
      );


    state.baptismFileName =
      fileName;

    state.baptismByName =
      baptismByName;


    els.baptismFileName.textContent =
      `${fileName}｜${baptismByName.size} 人有受浸日期`;


    if (
      state.rows.length
    ) {
      recalculate();
    }
  }


  function classify({
    attended,
    totalWeeks,
    weekly12Attendance,
    weekly12Weeks,
    weeksSinceLastAttendance
  }) {

    if (
      !totalWeeks ||
      weeksSinceLastAttendance === null ||
      weeksSinceLastAttendance > 8
    ) {
      return STATUS.INACTIVE;
    }


    if (
      weekly12Weeks >= 12 &&
      weekly12Attendance ===
        weekly12Weeks
    ) {
      return STATUS.WEEKLY;
    }


    const regularNeeded =
      Math.ceil(
        totalWeeks / 3
      );


    if (
      attended >=
        regularNeeded &&
      weeksSinceLastAttendance <= 3
    ) {
      return STATUS.REGULAR;
    }


    const occasionalNeeded =
      Math.ceil(
        totalWeeks / 4
      );


    if (
      attended >=
        occasionalNeeded &&
      weeksSinceLastAttendance <= 4
    ) {
      return STATUS.OCCASIONAL;
    }


    return STATUS.SPORADIC;
  }


  function recalculate() {

    if (
      !state.rows.length
    ) {
      return;
    }


    // 聚會情況固定以最近 12 週判定，不再提供切換。
    state.windowWeeks = 12;


    const recentCols =
      state.dateColumns.slice(
        -state.windowWeeks
      );


    state.recentCols =
      recentCols;


    const referenceDate =
      recentCols.length
        ? recentCols[
            recentCols.length - 1
          ].date
        : null;


    const { indexes } =
      state;


    const weekly12Cols =
      state.dateColumns.slice(-12);


    state.people =
      state.rows
        .slice(1)
        .map(
          (row, i) => {

            const name =
              String(
                row[
                  indexes.name
                ] ?? ''
              ).trim();


            if (!name) {
              return null;
            }


            const recentAttendance =
              recentCols.reduce(
                (sum, x) =>
                  sum +
                  (
                    isAttendance(
                      row[x.col]
                    )
                      ? 1
                      : 0
                  ),
                0
              );


            const weekly12Attendance =
              weekly12Cols.reduce(
                (sum, x) =>
                  sum +
                  (
                    isAttendance(
                      row[x.col]
                    )
                      ? 1
                      : 0
                  ),
                0
              );


            const attendedDates =
              state.dateColumns
                .filter(
                  x =>
                    isAttendance(
                      row[x.col]
                    )
                )
                .map(
                  x => x.date
                );


            const lastAttendanceDate =
              attendedDates.length
                ? attendedDates[
                    attendedDates.length - 1
                  ]
                : null;


            const weeksSinceLastAttendance =
              lastAttendanceDate &&
              referenceDate
                ? Math.floor(
                    (
                      referenceDate -
                      lastAttendanceDate
                    ) /
                    MS_WEEK
                  )
                : null;


            const totalAttendance =
              attendedDates.length;


            const baptismDate =
              indexes.baptism >= 0
                ? excelDateToDate(
                    row[
                      indexes.baptism
                    ]
                  )
                : (
                    state.baptismByName.get(
                      name
                    ) || null
                  );


            const weeksSinceBaptism =
              baptismDate &&
              referenceDate
                ? Math.floor(
                    (
                      referenceDate -
                      baptismDate
                    ) /
                    MS_WEEK
                  )
                : null;


            let newBelieverStatus =
              'unknown';


            if (
              weeksSinceBaptism !== null
            ) {

              newBelieverStatus =
                (
                  weeksSinceBaptism >= 0 &&
                  weeksSinceBaptism <= 104
                )
                  ? 'yes'
                  : 'no';
            }


            return {

              rowNumber:
                i + 2,

              district:
                String(
                  row[
                    indexes.district
                  ] ?? ''
                ).trim(),

              smallDistrict:
                String(
                  row[
                    indexes.smallDistrict
                  ] ?? ''
                ).trim(),

              name,

              group:
                String(
                  row[
                    indexes.group
                  ] ?? ''
                ).trim(),

              baptismDate,

              newBelieverStatus,

              isNewBeliever:
                newBelieverStatus ===
                'yes',

              recentAttendance,

              recentWeeks:
                recentCols.length,

              recentRate:
                recentCols.length
                  ? recentAttendance /
                    recentCols.length
                  : 0,

              totalAttendance,

              lastAttendanceDate,

              weeksSinceLastAttendance,

              weekly12Attendance,

              weekly12Weeks:
                weekly12Cols.length,

              status:
                classify({
                  attended:
                    recentAttendance,

                  totalWeeks:
                    recentCols.length,

                  weekly12Attendance,

                  weekly12Weeks:
                    weekly12Cols.length,

                  weeksSinceLastAttendance
                })
            };
          }
        )
        .filter(Boolean);


    renderAll();
  }


  function countStatuses(
    people
  ) {

    const counts =
      Object.fromEntries(
        STATUS_ORDER.map(
          s => [s, 0]
        )
      );


    people.forEach(
      p => {
        counts[p.status]++;
      }
    );


    return counts;
  }


  function renderAll() {

    renderInfo();

    renderRuleText();

    renderPopulationScope();

    renderSummary();

    buildTrendFilters();

    renderTrend();

    buildGroupTrendButtons();

    renderGroupTrend();

    renderOverall();

    renderDistricts();

    buildFilters();

    renderPeople();

    updateChartFilterUI();


    els.dashboard.classList.remove(
      'hidden'
    );


    els.clearBtn.disabled =
      false;
  }


  function renderInfo() {

    const allDates =
      state.dateColumns;

    const recent =
      state.recentCols;


    const latestDate =
      allDates[
        allDates.length - 1
      ].date;


    els.infoFile.textContent =
      state.fileName;


    els.infoPeriod.textContent =
      `${fmtDate(
        allDates[0].date
      )} – ${fmtDate(
        latestDate
      )}`;


    els.infoWeeks.textContent =
      `${allDates.length} 週｜目前分析：截至 ${fmtDate(
        latestDate
      )} 最近 ${recent.length} 週`;


    const population =
      getPopulationBase();

    els.infoPeople.textContent =
      population.exclusionActive
        ? `${population.people.length} 人（原 ${population.originalCount} 人）`
        : `${population.people.length} 人`;
  }


  function renderRuleText() {

    const actual =
      state.recentCols.length;


    const latestDate =
      state.dateColumns.length
        ? state.dateColumns[
            state.dateColumns.length - 1
          ].date
        : null;


    const earliestRecent =
      actual
        ? state.recentCols[0].date
        : null;


    const periodText =
      actual <
      state.windowWeeks
        ? `檔案目前只有 ${actual} 個有效週次`
        : `${fmtDate(
            earliestRecent
          )} – ${fmtDate(
            latestDate
          )}`;


    els.ruleText.textContent =
      `以匯入資料中最新主日 ${fmtDate(
        latestDate
      )} 為「目前」，累計最近 ${state.windowWeeks} 週（${periodText}）。初信＝受浸日期在目前截止日前 104 週內。`;
  }


  function renderSummary() {

    const people =
      getPopulationBase().people;

    const counts =
      countStatuses(
        people
      );


    const total =
      people.length || 1;


    const newBelievers =
      people.filter(
        p =>
          p.newBelieverStatus ===
          'yes'
      ).length;


    const unknownBaptism =
      people.filter(
        p =>
          p.newBelieverStatus ===
          'unknown'
      ).length;


    const cards = [

      [
        '總人數',
        people.length,
        '分析母體'
      ],

      [
        '初信',
        newBelievers,
        `${(
          newBelievers /
          total *
          100
        ).toFixed(1)}%`
      ],

      [
        '日期不明',
        unknownBaptism,
        `${(
          unknownBaptism /
          total *
          100
        ).toFixed(1)}%`
      ],

      ...STATUS_ORDER.map(
        s => [
          s,
          counts[s],
          `${(
            counts[s] /
            total *
            100
          ).toFixed(1)}%`
        ]
      )
    ];


    els.summaryCards.innerHTML =
      cards
        .map(
          ([
            label,
            value,
            note
          ]) => `
            <div class="card">

              <span class="label">
                ${escapeHtml(label)}
              </span>

              <strong class="value">
                ${value}
              </strong>

              <span class="note">
                ${escapeHtml(note)}
              </span>

            </div>
          `
        )
        .join('');
  }


  /*
   * ========================================
   * 每週聚會人數趨勢
   * ========================================
   *
   * 趨勢圖使用「分析母體」為基準，不受下方一般分析篩選影響。
   * 大區空白 = 全會所；選大區但小區空白 = 該大區；
   * 大區 + 小區都有選 = 該小區。
   */

  function renderAreaButtons({
    container,
    items,
    selectedValue,
    allLabel,
    dataAttr
  }) {

    if (!container) {
      return;
    }

    const allSelected = !selectedValue;

    container.innerHTML = [
      `
        <button
          type="button"
          class="ios-filter-btn${allSelected ? ' is-selected' : ''}"
          ${dataAttr}=""
        >
          <span class="ios-filter-dot" aria-hidden="true"></span>
          ${escapeHtml(allLabel)}
        </button>
      `,
      ...items.map(value => `
        <button
          type="button"
          class="ios-filter-btn${selectedValue === value ? ' is-selected' : ''}"
          ${dataAttr}="${escapeAttr(value)}"
        >
          <span class="ios-filter-dot" aria-hidden="true"></span>
          ${escapeHtml(value)}
        </button>
      `)
    ].join('');
  }


  function buildTrendFilters() {

    if (
      !els.trendDistrictFilter ||
      !els.trendSmallDistrictFilter ||
      !els.trendDistrictButtons
    ) {
      return;
    }

    const basePeople =
      getPopulationBase().people;

    const districts =
      unique(
        basePeople
          .map(p => p.district)
          .filter(Boolean)
      ).sort(
        (a, b) =>
          a.localeCompare(
            b,
            'zh-Hant',
            { numeric: true }
          )
      );

    const previousDistrict =
      els.trendDistrictFilter.value;

    if (
      previousDistrict &&
      !districts.includes(previousDistrict)
    ) {
      els.trendDistrictFilter.value = '';
      els.trendSmallDistrictFilter.value = '';
    }

    renderAreaButtons({
      container: els.trendDistrictButtons,
      items: districts,
      selectedValue: els.trendDistrictFilter.value,
      allLabel: '全會所',
      dataAttr: 'data-trend-district'
    });

    updateTrendSmallDistrictOptions();
  }


  function updateTrendSmallDistrictOptions() {

    if (
      !els.trendDistrictFilter ||
      !els.trendSmallDistrictFilter ||
      !els.trendSmallDistrictButtons
    ) {
      return;
    }

    const district =
      els.trendDistrictFilter.value;

    if (!district) {
      els.trendSmallDistrictFilter.value = '';

      if (els.trendSmallDistrictGroup) {
        els.trendSmallDistrictGroup.classList.add('hidden');
      }

      els.trendSmallDistrictButtons.innerHTML = '';
      return;
    }

    const smallDistricts =
      unique(
        getPopulationBase().people
          .filter(
            p =>
              p.district === district
          )
          .map(
            p => p.smallDistrict
          )
          .filter(Boolean)
      ).sort(
        (a, b) =>
          a.localeCompare(
            b,
            'zh-Hant',
            { numeric: true }
          )
      );

    if (
      els.trendSmallDistrictFilter.value &&
      !smallDistricts.includes(
        els.trendSmallDistrictFilter.value
      )
    ) {
      els.trendSmallDistrictFilter.value = '';
    }

    if (els.trendSmallDistrictGroup) {
      els.trendSmallDistrictGroup.classList.remove('hidden');
    }

    renderAreaButtons({
      container: els.trendSmallDistrictButtons,
      items: smallDistricts,
      selectedValue: els.trendSmallDistrictFilter.value,
      allLabel: '全部小區',
      dataAttr: 'data-trend-small-district'
    });
  }


  function getTrendPeople() {

    const district =
      els.trendDistrictFilter?.value || '';

    const smallDistrict =
      els.trendSmallDistrictFilter?.value || '';

    return getPopulationBase().people.filter(
      person =>
        (
          !district ||
          person.district === district
        ) &&
        (
          !smallDistrict ||
          person.smallDistrict === smallDistrict
        )
    );
  }


  function getTrendScopeLabel() {

    const district =
      els.trendDistrictFilter?.value || '';

    const smallDistrict =
      els.trendSmallDistrictFilter?.value || '';

    if (!district) {
      return '全會所';
    }

    if (!smallDistrict) {
      return district;
    }

    return `${district}｜${smallDistrict}`;
  }


  function renderTrend() {

    if (!els.trendChart) {
      return;
    }

    const people =
      getTrendPeople();

    const points =
      state.dateColumns.map(
        dateColumn => {

          const count =
            people.reduce(
              (sum, person) => {

                const row =
                  state.rows[
                    person.rowNumber - 1
                  ];

                return sum +
                  (
                    row &&
                    isAttendance(
                      row[dateColumn.col]
                    )
                      ? 1
                      : 0
                  );
              },
              0
            );

          return {
            date: dateColumn.date,
            count
          };
        }
      );

    const scopeLabel =
      getTrendScopeLabel();

    const latest =
      points.length
        ? points[points.length - 1].count
        : 0;

    const previous =
      points.length >= 2
        ? points[points.length - 2].count
        : null;

    const average =
      points.length
        ? points.reduce(
            (sum, point) =>
              sum + point.count,
            0
          ) / points.length
        : 0;

    if (els.trendScopeLabel) {
      els.trendScopeLabel.textContent =
        `${scopeLabel}｜${people.length} 人母體`;
    }

    if (els.trendLatestCount) {
      els.trendLatestCount.textContent =
        latest;
    }

    if (els.trendAverageCount) {
      els.trendAverageCount.textContent =
        average.toFixed(1);
    }

    if (els.trendChangeText) {
      if (previous === null) {
        els.trendChangeText.textContent =
          '較前一週 —';
      } else {
        const diff = latest - previous;
        const sign =
          diff > 0
            ? '+'
            : '';

        els.trendChangeText.textContent =
          `較前一週 ${sign}${diff} 人`;

        els.trendChangeText.classList.toggle(
          'is-up',
          diff > 0
        );

        els.trendChangeText.classList.toggle(
          'is-down',
          diff < 0
        );
      }
    }

    if (!points.length) {
      els.trendChart.innerHTML =
        '<div class="trend-empty">沒有可繪製的每週資料</div>';
      return;
    }

    const chartWidth =
      Math.max(
        760,
        points.length * 38
      );

    const chartHeight = 350;
    const left = 58;
    const right = 24;
    const top = 22;
    const bottom = 54;
    const plotWidth =
      chartWidth - left - right;
    const plotHeight =
      chartHeight - top - bottom;

    const rawMax =
      Math.max(
        ...points.map(p => p.count),
        1
      );

    const step =
      rawMax <= 20
        ? 5
        : rawMax <= 60
          ? 10
          : rawMax <= 150
            ? 25
            : rawMax <= 300
              ? 50
              : 100;

    const yMax =
      Math.ceil(rawMax / step) * step;

    const xAt =
      index =>
        points.length === 1
          ? left + plotWidth / 2
          : left +
            index /
            (points.length - 1) *
            plotWidth;

    const yAt =
      value =>
        top +
        plotHeight -
        value / yMax * plotHeight;

    const linePoints =
      points
        .map(
          (point, index) =>
            `${xAt(index).toFixed(1)},${yAt(point.count).toFixed(1)}`
        )
        .join(' ');

    const yTicks = 4;
    const grid =
      Array.from(
        { length: yTicks + 1 },
        (_, i) => {
          const value =
            Math.round(
              yMax *
              (yTicks - i) /
              yTicks
            );

          const y =
            top +
            i / yTicks *
            plotHeight;

          return `
            <line
              class="trend-grid-line"
              x1="${left}"
              y1="${y}"
              x2="${chartWidth - right}"
              y2="${y}"
            />
            <text
              class="trend-y-label"
              x="${left - 12}"
              y="${y + 4}"
              text-anchor="end"
            >${value}</text>
          `;
        }
      ).join('');

    const labelEvery =
      Math.max(
        1,
        Math.ceil(points.length / 10)
      );

    const xLabels =
      points
        .map(
          (point, index) => {
            const isLast =
              index === points.length - 1;

            if (
              index % labelEvery !== 0 &&
              !isLast
            ) {
              return '';
            }

            const x = xAt(index);
            const label =
              `${point.date.getMonth() + 1}/${point.date.getDate()}`;

            return `
              <text
                class="trend-x-label"
                x="${x}"
                y="${chartHeight - 20}"
                text-anchor="middle"
              >${label}</text>
            `;
          }
        )
        .join('');

    const circles =
      points
        .map(
          (point, index) => {
            const x = xAt(index);
            const y = yAt(point.count);
            const dateText =
              fmtDate(point.date);

            return `
              <g class="trend-point-group">
                <circle
                  class="trend-point-hit"
                  cx="${x}"
                  cy="${y}"
                  r="11"
                  tabindex="0"
                >
                  <title>${escapeHtml(dateText)}：${point.count} 人</title>
                </circle>
                <circle
                  class="trend-point"
                  cx="${x}"
                  cy="${y}"
                  r="4"
                  aria-hidden="true"
                />
              </g>
            `;
          }
        )
        .join('');

    els.trendChart.innerHTML = `
      <svg
        class="trend-svg"
        width="${chartWidth}"
        height="${chartHeight}"
        viewBox="0 0 ${chartWidth} ${chartHeight}"
        aria-hidden="true"
      >
        ${grid}

        <polyline
          class="trend-line"
          points="${linePoints}"
        />

        ${circles}
        ${xLabels}
      </svg>
    `;
  }


  /*
   * ========================================
   * 羣組成長趨勢
   * ========================================
   *
   * 1. 使用與上方每週人數相同的「全會所 / 大區 / 小區」範圍。
   * 2. 國小、學齡前不納入。
   * 3. 國中、高中、中學統一合併為「青少年」。
   * 4. 每條線採 4 週移動平均，不顯示每週折點。
   */

  const GROUP_TREND_COLORS = [
    '#315b66',
    '#6f8793',
    '#9a755f',
    '#6f8061',
    '#756b8a',
    '#a07f45',
    '#8a5f69',
    '#567b78'
  ];


  function cleanGroupText(value) {
    return String(value ?? '')
      .trim()
      .replace(/\s+/g, '');
  }


  function normalizeTrendGroup(group) {
    const value =
      cleanGroupText(group);

    if (!value) {
      return null;
    }

    if (
      value === '國小' ||
      value === '學齡前'
    ) {
      return null;
    }

    if (
      [
        '國中',
        '高中',
        '中學'
      ].includes(value)
    ) {
      return '青少年';
    }

    return value;
  }


  function getAvailableTrendGroups() {
    const found =
      unique(
        getTrendPeople()
          .map(
            person =>
              normalizeTrendGroup(
                person.group
              )
          )
          .filter(Boolean)
      );

    const preferred = [
      '年長',
      '中壯',
      '青壯',
      '青職',
      '大學',
      '大專',
      '青少年'
    ];

    return [
      ...preferred.filter(
        group => found.includes(group)
      ),
      ...found
        .filter(
          group => !preferred.includes(group)
        )
        .sort(
          (a, b) =>
            a.localeCompare(
              b,
              'zh-Hant',
              { numeric: true }
            )
        )
    ];
  }


  function getTrendGroupColor(
    group,
    availableGroups
  ) {
    const index =
      Math.max(
        0,
        availableGroups.indexOf(group)
      );

    return GROUP_TREND_COLORS[
      index % GROUP_TREND_COLORS.length
    ];
  }


  function buildGroupTrendButtons() {
    if (!els.groupTrendButtons) {
      return;
    }

    const groups =
      getAvailableTrendGroups();

    state.selectedTrendGroups =
      new Set(
        [...state.selectedTrendGroups]
          .filter(
            group => groups.includes(group)
          )
      );

    if (
      !state.trendGroupsInitialized
    ) {
      state.selectedTrendGroups =
        new Set(groups);

      state.trendGroupsInitialized =
        true;
    }

    els.groupTrendButtons.innerHTML =
      groups.length
        ? groups
            .map(
              group => {
                const selected =
                  state.selectedTrendGroups.has(
                    group
                  );

                const color =
                  getTrendGroupColor(
                    group,
                    groups
                  );

                return `
                  <button
                    type="button"
                    class="group-trend-btn${selected ? ' is-selected' : ''}"
                    data-trend-group="${escapeAttr(group)}"
                    aria-pressed="${selected ? 'true' : 'false'}"
                    style="--group-trend-color:${escapeAttr(color)}"
                  >
                    <span class="group-trend-btn-dot" aria-hidden="true"></span>
                    <span>${escapeHtml(group)}</span>
                  </button>
                `;
              }
            )
            .join('')
        : '<div class="group-trend-no-groups">目前沒有可分析的羣組</div>';

    updateGroupTrendMeta();
  }


  function updateGroupTrendMeta() {
    if (els.groupTrendScopeLabel) {
      els.groupTrendScopeLabel.textContent =
        getTrendScopeLabel();
    }

    if (!els.groupTrendSelectedText) {
      return;
    }

    const selected =
      getAvailableTrendGroups()
        .filter(
          group =>
            state.selectedTrendGroups.has(
              group
            )
        );

    els.groupTrendSelectedText.textContent =
      selected.length
        ? `目前顯示：${selected.join('＋')}`
        : '目前未選擇羣組';
  }


  function movingAverage4(values) {
    return values.map(
      (_, index) => {
        if (index < 3) {
          return null;
        }

        const window =
          values.slice(
            index - 3,
            index + 1
          );

        return window.reduce(
          (sum, value) => sum + value,
          0
        ) / 4;
      }
    );
  }


  function smoothSvgPath(points) {
    if (!points.length) {
      return '';
    }

    if (points.length === 1) {
      return `M ${points[0].x} ${points[0].y}`;
    }

    let path =
      `M ${points[0].x} ${points[0].y}`;

    for (
      let i = 1;
      i < points.length;
      i++
    ) {
      const previous =
        points[i - 1];

      const current =
        points[i];

      const middleX =
        (previous.x + current.x) / 2;

      path +=
        ` C ${middleX} ${previous.y}, ${middleX} ${current.y}, ${current.x} ${current.y}`;
    }

    return path;
  }


  function renderGroupTrend() {
    if (!els.groupTrendChart) {
      return;
    }

    updateGroupTrendMeta();

    const availableGroups =
      getAvailableTrendGroups();

    const selectedGroups =
      availableGroups.filter(
        group =>
          state.selectedTrendGroups.has(
            group
          )
      );

    if (!selectedGroups.length) {
      els.groupTrendChart.innerHTML =
        '<div class="trend-empty">請點選至少一個羣組查看成長趨勢</div>';
      return;
    }

    const people =
      getTrendPeople();

    const series =
      selectedGroups.map(
        group => {
          const groupPeople =
            people.filter(
              person =>
                normalizeTrendGroup(
                  person.group
                ) === group
            );

          const weeklyCounts =
            state.dateColumns.map(
              dateColumn =>
                groupPeople.reduce(
                  (sum, person) => {
                    const row =
                      state.rows[
                        person.rowNumber - 1
                      ];

                    return sum +
                      (
                        row &&
                        isAttendance(
                          row[dateColumn.col]
                        )
                          ? 1
                          : 0
                      );
                  },
                  0
                )
            );

          return {
            group,
            color:
              getTrendGroupColor(
                group,
                availableGroups
              ),
            values:
              movingAverage4(
                weeklyCounts
              )
          };
        }
      );

    if (
      state.dateColumns.length < 4
    ) {
      els.groupTrendChart.innerHTML =
        '<div class="trend-empty">至少需要 4 週資料才能計算 4 週移動平均</div>';
      return;
    }

    const chartWidth =
      Math.max(
        760,
        state.dateColumns.length * 28
      );

    const chartHeight = 350;
    const left = 58;
    const right = 24;
    const top = 24;
    const bottom = 54;
    const plotWidth =
      chartWidth - left - right;
    const plotHeight =
      chartHeight - top - bottom;

    const allValues =
      series.flatMap(
        item =>
          item.values.filter(
            value => value !== null
          )
      );

    const rawMax =
      Math.max(
        ...allValues,
        1
      );

    const step =
      rawMax <= 20
        ? 5
        : rawMax <= 60
          ? 10
          : rawMax <= 150
            ? 25
            : rawMax <= 300
              ? 50
              : 100;

    const yMax =
      Math.ceil(rawMax / step) * step;

    const xAt =
      index =>
        state.dateColumns.length === 1
          ? left + plotWidth / 2
          : left +
            index /
            (state.dateColumns.length - 1) *
            plotWidth;

    const yAt =
      value =>
        top +
        plotHeight -
        value / yMax * plotHeight;

    const yTicks = 4;

    const grid =
      Array.from(
        { length: yTicks + 1 },
        (_, i) => {
          const value =
            Math.round(
              yMax *
              (yTicks - i) /
              yTicks
            );

          const y =
            top +
            i / yTicks *
            plotHeight;

          return `
            <line
              class="trend-grid-line"
              x1="${left}"
              y1="${y}"
              x2="${chartWidth - right}"
              y2="${y}"
            />
            <text
              class="trend-y-label"
              x="${left - 12}"
              y="${y + 4}"
              text-anchor="end"
            >${value}</text>
          `;
        }
      ).join('');

    const labelEvery =
      Math.max(
        1,
        Math.ceil(
          state.dateColumns.length / 10
        )
      );

    const xLabels =
      state.dateColumns
        .map(
          (dateColumn, index) => {
            const isLast =
              index ===
              state.dateColumns.length - 1;

            if (
              index % labelEvery !== 0 &&
              !isLast
            ) {
              return '';
            }

            const label =
              `${dateColumn.date.getMonth() + 1}/${dateColumn.date.getDate()}`;

            return `
              <text
                class="trend-x-label"
                x="${xAt(index)}"
                y="${chartHeight - 20}"
                text-anchor="middle"
              >${label}</text>
            `;
          }
        )
        .join('');

    const paths =
      series
        .map(
          item => {
            const chartPoints =
              item.values
                .map(
                  (value, index) =>
                    value === null
                      ? null
                      : {
                          x: Number(
                            xAt(index).toFixed(1)
                          ),
                          y: Number(
                            yAt(value).toFixed(1)
                          ),
                          value,
                          index
                        }
                )
                .filter(Boolean);

            const path =
              smoothSvgPath(
                chartPoints
              );

            const hitPoints =
              chartPoints
                .map(
                  point => {
                    const dateText =
                      fmtDate(
                        state.dateColumns[
                          point.index
                        ].date
                      );

                    return `
                      <circle
                        class="group-trend-hit"
                        cx="${point.x}"
                        cy="${point.y}"
                        r="10"
                        tabindex="0"
                      >
                        <title>${escapeHtml(item.group)}｜${escapeHtml(dateText)}｜4週平均 ${point.value.toFixed(1)} 人</title>
                      </circle>
                    `;
                  }
                )
                .join('');

            return `
              <g class="group-trend-series">
                <path
                  class="group-trend-line"
                  d="${path}"
                  style="stroke:${escapeAttr(item.color)}"
                />
                ${hitPoints}
              </g>
            `;
          }
        )
        .join('');

    els.groupTrendChart.innerHTML = `
      <svg
        class="trend-svg"
        width="${chartWidth}"
        height="${chartHeight}"
        viewBox="0 0 ${chartWidth} ${chartHeight}"
        aria-hidden="true"
      >
        ${grid}
        ${paths}
        ${xLabels}
      </svg>
    `;
  }


  function renderOverall() {

    const people =
      getPopulationBase().people;

    const counts =
      countStatuses(
        people
      );


    const total =
      people.length || 1;


    els.overallBar.innerHTML =
      STATUS_ORDER
        .map(
          s => {

            const pct =
              counts[s] /
              total *
              100;


            return `
              <div
                class="seg-${STATUS_CLASS[s]}"
                style="width:${pct}%"
                title="${escapeAttr(
                  `${s} ${counts[s]} 人 (${pct.toFixed(1)}%)`
                )}"
              ></div>
            `;
          }
        )
        .join('');


    els.overallLegend.innerHTML =
      STATUS_ORDER
        .map(
          s => {

            const pct =
              counts[s] /
              total *
              100;


            return `
              <div class="legend-item">

                <span>
                  <i class="dot ${STATUS_CLASS[s]}"></i>
                  ${escapeHtml(s)}
                </span>

                <strong>
                  ${counts[s]}
                  <small>
                    ${pct.toFixed(1)}%
                  </small>
                </strong>

              </div>
            `;
          }
        )
        .join('');
  }


  function renderDistricts() {

    const groups =
      groupBy(
        getPopulationBase().people,
        p =>
          p.district ||
          '未分類'
      );


    els.districtBody.innerHTML =
      [...groups.entries()]
        .sort(
          (a, b) =>
            a[0].localeCompare(
              b[0],
              'zh-Hant'
            )
        )
        .map(
          ([
            district,
            people
          ]) => {

            const c =
              countStatuses(
                people
              );


            const stableRate =
              people.length
                ? (
                    c[STATUS.WEEKLY] +
                    c[STATUS.REGULAR]
                  ) /
                  people.length *
                  100
                : 0;


            const nb =
              people.filter(
                p =>
                  p.newBelieverStatus ===
                  'yes'
              ).length;


            const unknown =
              people.filter(
                p =>
                  p.newBelieverStatus ===
                  'unknown'
              ).length;


            return `
              <tr>

                <td>
                  <strong>
                    ${escapeHtml(district)}
                  </strong>
                </td>

                <td>${people.length}</td>
                <td>${nb}</td>
                <td>${unknown}</td>
                <td>${c[STATUS.WEEKLY]}</td>
                <td>${c[STATUS.REGULAR]}</td>
                <td>${c[STATUS.OCCASIONAL]}</td>
                <td>${c[STATUS.SPORADIC]}</td>
                <td>${c[STATUS.INACTIVE]}</td>
                <td>${stableRate.toFixed(1)}%</td>

              </tr>
            `;
          }
        )
        .join('');
  }


  function analysisFilterGroups() {
    const found = unique(
      getPopulationBase().people
        .map(p => {
          const g = cleanGroupText(p.group);
          return ['國中', '高中', '中學'].includes(g) ? '青少年' : g;
        })
        .filter(Boolean)
    );
    const preferred = ['年長','中壯','青壯','青職','大學','大專','青少年','國小','學齡前'];
    return [...preferred.filter(g => found.includes(g)), ...found.filter(g => !preferred.includes(g)).sort((a,b)=>a.localeCompare(b,'zh-Hant'))];
  }

  function makeAnalysisButton(value, label, selected, attr, detail='') {
    return `<button type="button" class="ios-filter-btn ${selected ? 'is-selected' : ''}" ${attr}="${escapeAttr(value)}" aria-pressed="${selected}">
      <span class="analysis-group-main"><i class="ios-filter-dot"></i><span>${escapeHtml(label)}</span></span>
      ${detail ? `<span class="analysis-group-detail">${escapeHtml(detail)}</span>` : ''}
    </button>`;
  }

  function renderAnalysisFilterButtons() {
    if (!els.analysisDistrictButtons) return;
    const basePeople = getPopulationBase().people;
    const districts = unique(basePeople.map(p=>p.district).filter(Boolean)).sort((a,b)=>a.localeCompare(b,'zh-Hant'));
    const d = els.districtFilter.value;
    els.analysisDistrictButtons.innerHTML = makeAnalysisButton('', '全部', !d, 'data-analysis-district') + districts.map(x=>makeAnalysisButton(x,x,d===x,'data-analysis-district')).join('');

    const smalls = unique(basePeople.filter(p=>!d || p.district===d).map(p=>p.smallDistrict).filter(Boolean)).sort((a,b)=>a.localeCompare(b,'zh-Hant'));
    const sm = els.smallDistrictFilter.value;
    els.analysisSmallDistrictRow.classList.toggle('hidden', !d);
    els.analysisSmallDistrictButtons.innerHTML = makeAnalysisButton('', '全部', !sm, 'data-analysis-small') + smalls.map(x=>makeAnalysisButton(x,x,sm===x,'data-analysis-small')).join('');

    const areaPeople = basePeople.filter(p=>(!d||p.district===d)&&(!sm||p.smallDistrict===sm));
    els.analysisGroupButtons.innerHTML = analysisFilterGroups().map(g=>{
      let detail='';
      if (g==='青少年') {
        const counts={國中:0,高中:0,中學:0};
        areaPeople.forEach(p=>{const x=cleanGroupText(p.group); if (x in counts) counts[x]++;});
        const total=counts.國中+counts.高中+counts.中學;
        detail=`共 ${total}｜國中 ${counts.國中}・高中 ${counts.高中}・中學 ${counts.中學}`;
      }
      return makeAnalysisButton(g,g,state.selectedGroups.has(g),'data-analysis-group',detail);
    }).join('');

    els.analysisStatusButtons.innerHTML = STATUS_ORDER.map(x=>makeAnalysisButton(x,x,state.selectedStatuses.has(x),'data-analysis-status')).join('');
    const nb=els.newBelieverFilter.value;
    els.analysisBelieverButtons.innerHTML = [['','全部'],['yes','初信'],['no','非初信'],['unknown','日期不明']].map(([v,l])=>makeAnalysisButton(v,l,nb===v,'data-analysis-believer')).join('');
  }

  function buildFilters() {

    preserveOptions(
      els.districtFilter,
      unique(
        getPopulationBase().people.map(
          p => p.district
        )
      )
    );


    updateSmallDistrictOptions();

    buildGroupOptions();

    buildStatusOptions();
    renderAnalysisFilterButtons();
  }


  function availableGroups() {

    return unique(
      getPopulationBase().people
        .map(p => {
          const g = cleanGroupText(p.group);
          return ['國中', '高中', '中學'].includes(g) ? '青少年' : g;
        })
        .filter(Boolean)
    )
      .sort(
        (a, b) =>
          a.localeCompare(
            b,
            'zh-Hant'
          )
      );
  }


  function buildGroupOptions() {

    const groups =
      availableGroups();


    state.selectedGroups =
      new Set(
        [
          ...state.selectedGroups
        ].filter(
          g =>
            groups.includes(g)
        )
      );


    els.groupFilterOptions.innerHTML =
      groups
        .map(
          g => `
            <label class="group-option">

              <input
                type="checkbox"
                value="${escapeAttr(g)}"
                ${
                  state.selectedGroups.has(g)
                    ? 'checked'
                    : ''
                }
              >

              <span>
                ${escapeHtml(g)}
              </span>

            </label>
          `
        )
        .join('');


    updateGroupSummary();
  }


  function updateGroupSummary() {

    const selected =
      [
        ...state.selectedGroups
      ];


    if (
      !selected.length
    ) {

      els.groupFilterSummary.textContent =
        '全部';

    } else if (
      selected.length <= 2
    ) {

      els.groupFilterSummary.textContent =
        selected.join('＋');

    } else {

      els.groupFilterSummary.textContent =
        `已選 ${selected.length} 組`;
    }
  }


  function setSelectedGroups(
    groups
  ) {

    const available =
      availableGroups();


    const aliases = {

      college: [
        '大專',
        '大學'
      ],

      youth: [
        '青職'
      ],

      junior: [
        '國中'
      ],

      senior: [
        '高中'
      ],

      child: [
        '國小'
      ]
    };


    const resolve =
      key =>
        aliases[key]
          .filter(
            x =>
              available.includes(x)
          );


    let wanted = [];


    if (
      groups === 'all'
    ) {
      wanted = [];
    }


    if (
      groups ===
      'college-youth'
    ) {

      wanted = [
        ...resolve('college'),
        ...resolve('youth')
      ];
    }


    if (
      groups === 'teen'
    ) {

      wanted = ['青少年'];
    }


    if (
      groups === 'student'
    ) {

      wanted = [
        ...resolve('college'),
        ...resolve('junior'),
        ...resolve('senior')
      ];
    }


    if (
      groups === 'child'
    ) {

      wanted = [
        ...resolve('child')
      ];
    }


    state.selectedGroups =
      new Set(wanted);


    buildGroupOptions();
    renderAnalysisFilterButtons();

    renderPeople();
  }


  function buildStatusOptions() {

    state.selectedStatuses =
      new Set(
        [
          ...state.selectedStatuses
        ].filter(
          s =>
            STATUS_ORDER.includes(s)
        )
      );


    els.statusFilterOptions.innerHTML =
      STATUS_ORDER
        .map(
          status => `
            <label class="group-option">

              <input
                type="checkbox"
                value="${escapeAttr(status)}"
                ${
                  state.selectedStatuses.has(status)
                    ? 'checked'
                    : ''
                }
              >

              <span>
                <i class="dot ${STATUS_CLASS[status]}"></i>
                ${escapeHtml(status)}
              </span>

            </label>
          `
        )
        .join('');


    updateStatusSummary();
  }


  function updateStatusSummary() {

    const selected =
      STATUS_ORDER.filter(
        s =>
          state.selectedStatuses.has(s)
      );


    if (
      !selected.length
    ) {

      els.statusFilterSummary.textContent =
        '全部';

    } else if (
      selected.length === 1
    ) {

      els.statusFilterSummary.textContent =
        selected[0];

    } else {

      els.statusFilterSummary.textContent =
        `已選 ${selected.length} 項`;
    }
  }


  function setSelectedStatuses(
    statuses
  ) {

    state.selectedStatuses =
      new Set(statuses);


    buildStatusOptions();
    renderAnalysisFilterButtons();

    renderPeople();
  }


  function updateSmallDistrictOptions() {

    const selectedDistrict =
      els.districtFilter.value;


    const peopleInDistrict =
      selectedDistrict
        ? getPopulationBase().people.filter(
            p =>
              p.district ===
              selectedDistrict
          )
        : getPopulationBase().people;


    preserveOptions(
      els.smallDistrictFilter,
      unique(
        peopleInDistrict.map(
          p =>
            p.smallDistrict
        )
      )
    );
  }


  function preserveOptions(
    select,
    values
  ) {

    const previous =
      select.value;


    select.innerHTML =
      '<option value="">全部</option>' +

      values
        .filter(Boolean)
        .sort(
          (a, b) =>
            a.localeCompare(
              b,
              'zh-Hant'
            )
        )
        .map(
          v => `
            <option value="${escapeAttr(v)}">
              ${escapeHtml(v)}
            </option>
          `
        )
        .join('');


    if (
      [
        ...select.options
      ].some(
        o =>
          o.value ===
          previous
      )
    ) {

      select.value =
        previous;
    }
  }


  /* ========================================
     排除功能
  ======================================== */

  function getAvailableHistoryWeeks() {

    if (
      !state.dateColumns.length
    ) {
      return 0;
    }


    if (
      state.dateColumns.length === 1
    ) {
      return 1;
    }


    const firstDate =
      state.dateColumns[0].date;


    const lastDate =
      state.dateColumns[
        state.dateColumns.length - 1
      ].date;


    return (
      Math.floor(
        (
          lastDate -
          firstDate
        ) /
        MS_WEEK
      ) + 1
    );
  }


  function shouldExcludePerson(
    person,
    maxTotalAttendance,
    absenceWeeks,
    availableHistoryWeeks
  ) {

    const hasTotalCondition =
      maxTotalAttendance !== null;


    const hasAbsenceCondition =
      absenceWeeks !== null;


    if (
      !hasTotalCondition &&
      !hasAbsenceCondition
    ) {
      return false;
    }


    const matchesLowAttendance =
      hasTotalCondition
        ? person.totalAttendance <=
          maxTotalAttendance
        : false;


    let matchesLongAbsence =
      false;


    if (
      hasAbsenceCondition
    ) {

      if (
        person.weeksSinceLastAttendance !==
        null
      ) {

        matchesLongAbsence =
          person.weeksSinceLastAttendance >=
          absenceWeeks;

      } else {

        matchesLongAbsence =
          availableHistoryWeeks >=
          absenceWeeks;
      }
    }


    if (
      hasTotalCondition &&
      hasAbsenceCondition
    ) {

      return (
        matchesLowAttendance &&
        matchesLongAbsence
      );
    }


    if (
      hasTotalCondition
    ) {

      return matchesLowAttendance;
    }


    return matchesLongAbsence;
  }


  /*
   * ========================================
   * 分析母體
   * ========================================
   *
   * 這兩個排除條件不是一般「分析篩選」，
   * 而是先決定整份報表要分析哪些人。
   * 上方摘要、比例、大區比較與四張群組圖，
   * 都以這裡排除後的人員為母體。
   */

  function getPopulationBase() {

    const maxTotalAttendance =
      els.totalAttendanceFilter.value === ''
        ? null
        : Number(els.totalAttendanceFilter.value);

    const absenceWeeks =
      els.absenceWeeksFilter.value === ''
        ? null
        : Number(els.absenceWeeksFilter.value);

    const availableHistoryWeeks =
      getAvailableHistoryWeeks();

    const people =
      state.people.filter(
        p =>
          !shouldExcludePerson(
            p,
            maxTotalAttendance,
            absenceWeeks,
            availableHistoryWeeks
          )
      );

    return {
      people,
      originalCount: state.people.length,
      excludedCount: state.people.length - people.length,
      exclusionActive:
        maxTotalAttendance !== null ||
        absenceWeeks !== null
    };
  }


  function renderPopulationScope() {

    const base =
      getPopulationBase();

    if (els.populationOriginalCount) {
      els.populationOriginalCount.textContent =
        base.originalCount;
    }

    if (els.populationActiveCount) {
      els.populationActiveCount.textContent =
        base.people.length;
    }

    if (els.populationExcludedCount) {
      els.populationExcludedCount.textContent =
        base.excludedCount;
    }
  }


  /*
   * ========================================
   * 圖表群組判斷
   * ========================================
   */

  function matchesChartGroup(
    person,
    chartGroup
  ) {

    if (!chartGroup) {
      return true;
    }

    if (chartGroup === '青少年') {
      return [
        '國中',
        '高中',
        '中學'
      ].includes(person.group);
    }

    return person.group === chartGroup;
  }


  function getGeneralFilteredPeople() {

    const q =
      els.searchInput.value
        .trim()
        .toLowerCase();

    const district =
      els.districtFilter.value;

    const small =
      els.smallDistrictFilter.value;

    const nb =
      els.newBelieverFilter.value;

    const selectedGroups =
      state.selectedGroups;

    const selectedStatuses =
      state.selectedStatuses;

    const population =
      getPopulationBase();

    const visible =
      population.people.filter(
        p => {

          const matchesSearch =
            !q ||
            p.name.toLowerCase().includes(q);

          const matchesDistrict =
            !district ||
            p.district === district;

          const matchesSmallDistrict =
            !small ||
            p.smallDistrict === small;

          const normalizedFilterGroup =
            ['國中', '高中', '中學'].includes(cleanGroupText(p.group))
              ? '青少年'
              : cleanGroupText(p.group);

          const matchesGroup =
            !selectedGroups.size ||
            selectedGroups.has(normalizedFilterGroup);

          const matchesStatus =
            !selectedStatuses.size ||
            selectedStatuses.has(p.status);

          const matchesNewBeliever =
            !nb ||
            p.newBelieverStatus === nb;

          return (
            matchesSearch &&
            matchesDistrict &&
            matchesSmallDistrict &&
            matchesGroup &&
            matchesStatus &&
            matchesNewBeliever
          );
        }
      );

    return {
      people: visible,
      populationCount: population.people.length,
      originalCount: population.originalCount,
      excludedCount: population.excludedCount,
      exclusionActive: population.exclusionActive
    };
  }


  function matchesChartTypeStatus(
    person,
    chartType
  ) {

    if (!chartType) {
      return true;
    }

    if (chartType === 'stable') {
      return (
        person.status === STATUS.WEEKLY ||
        person.status === STATUS.REGULAR
      );
    }

    if (chartType === 'care') {
      return (
        person.status === STATUS.OCCASIONAL ||
        person.status === STATUS.SPORADIC
      );
    }

    return true;
  }


  function peopleSmallDistrictKey(
    person
  ) {

    return `${person.district}::${person.smallDistrict}`;
  }


  function sortZh(values) {

    return [...values].sort(
      (a, b) =>
        a.localeCompare(
          b,
          'zh-Hant',
          {
            numeric: true
          }
        )
    );
  }


  function updatePeopleAreaFilterUI(
    candidates
  ) {

    if (
      !els.peopleDistrictOptions ||
      !els.peopleDistrictSummary ||
      !els.peopleSmallDistrictOptions ||
      !els.peopleSmallDistrictSummary
    ) {
      return;
    }


    const availableDistricts =
      new Set(
        candidates
          .map(p => p.district)
          .filter(Boolean)
      );


    state.selectedPeopleDistricts =
      new Set(
        [...state.selectedPeopleDistricts]
          .filter(
            district =>
              availableDistricts.has(district)
          )
      );


    const districtList =
      sortZh(availableDistricts);


    els.peopleDistrictOptions.innerHTML =
      districtList.length
        ? [
            `
              <button
                type="button"
                class="ios-filter-btn people-area-btn${!state.selectedPeopleDistricts.size ? ' is-selected' : ''}"
                data-people-district=""
              >
                <span class="ios-filter-dot" aria-hidden="true"></span>
                全部
              </button>
            `,
            ...districtList.map(
              district => {

                const count =
                  candidates.filter(
                    p => p.district === district
                  ).length;

                return `
                  <button
                    type="button"
                    class="ios-filter-btn people-area-btn${state.selectedPeopleDistricts.has(district) ? ' is-selected' : ''}"
                    data-people-district="${escapeAttr(district)}"
                  >
                    <span class="ios-filter-dot" aria-hidden="true"></span>
                    <span>${escapeHtml(district)}</span>
                    <small>${count}</small>
                  </button>
                `;
              }
            )
          ].join('')
        : '<div class="people-filter-empty">目前沒有可選大區</div>';


    const selectedDistricts =
      sortZh(state.selectedPeopleDistricts);


    els.peopleDistrictSummary.textContent =
      !selectedDistricts.length
        ? '大區：全部'
        : selectedDistricts.length === 1
          ? `大區：${selectedDistricts[0]}`
          : `大區：已選 ${selectedDistricts.length} 項`;


    const smallCandidates =
      selectedDistricts.length
        ? candidates.filter(
            p =>
              state.selectedPeopleDistricts.has(
                p.district
              )
          )
        : [];


    const availableSmallMap =
      new Map();


    smallCandidates.forEach(
      p => {

        if (!p.smallDistrict) {
          return;
        }

        const key =
          peopleSmallDistrictKey(p);

        if (!availableSmallMap.has(key)) {
          availableSmallMap.set(
            key,
            {
              key,
              district: p.district,
              smallDistrict: p.smallDistrict,
              count: 0
            }
          );
        }

        availableSmallMap.get(key).count++;
      }
    );


    state.selectedPeopleSmallDistricts =
      new Set(
        [...state.selectedPeopleSmallDistricts]
          .filter(
            key =>
              availableSmallMap.has(key)
          )
      );


    const smallList =
      [...availableSmallMap.values()]
        .sort(
          (a, b) => {

            const districtCompare =
              a.district.localeCompare(
                b.district,
                'zh-Hant',
                { numeric: true }
              );

            if (districtCompare) {
              return districtCompare;
            }

            return a.smallDistrict.localeCompare(
              b.smallDistrict,
              'zh-Hant',
              { numeric: true }
            );
          }
        );


    if (
      els.peopleSmallDistrictMultiSelect
    ) {

      els.peopleSmallDistrictMultiSelect
        .classList
        .toggle(
          'hidden',
          !selectedDistricts.length
        );

      if (!selectedDistricts.length) {
        els.peopleSmallDistrictMultiSelect.open = false;
      }
    }


    els.peopleSmallDistrictOptions.innerHTML =
      smallList.length
        ? [
            `
              <button
                type="button"
                class="ios-filter-btn people-area-btn${!state.selectedPeopleSmallDistricts.size ? ' is-selected' : ''}"
                data-people-small-district=""
              >
                <span class="ios-filter-dot" aria-hidden="true"></span>
                全部小區
              </button>
            `,
            ...smallList.map(
              item => `
                <button
                  type="button"
                  class="ios-filter-btn people-area-btn${state.selectedPeopleSmallDistricts.has(item.key) ? ' is-selected' : ''}"
                  data-people-small-district="${escapeAttr(item.key)}"
                  title="${escapeAttr(`${item.district}｜${item.smallDistrict} ${item.count} 人`)}"
                >
                  <span class="ios-filter-dot" aria-hidden="true"></span>
                  <span>${escapeHtml(item.smallDistrict)}</span>
                  <small>${item.count}</small>
                </button>
              `
            )
          ].join('')
        : '<div class="people-filter-empty">所選大區目前沒有小區資料</div>';


    const selectedSmallCount =
      state.selectedPeopleSmallDistricts.size;


    els.peopleSmallDistrictSummary.textContent =
      !selectedSmallCount
        ? '小區：全部'
        : selectedSmallCount === 1
          ? `小區：${(
              availableSmallMap.get(
                [...state.selectedPeopleSmallDistricts][0]
              )?.smallDistrict || ''
            )}`
          : `小區：已選 ${selectedSmallCount} 項`;


    const localActive =
      Boolean(
        state.selectedPeopleDistricts.size ||
        state.selectedPeopleSmallDistricts.size
      );


    if (els.clearPeopleAreaFilter) {
      els.clearPeopleAreaFilter
        .classList
        .toggle(
          'hidden',
          !localActive
        );
    }
  }


  function clearPeopleAreaFilters() {

    state.selectedPeopleDistricts =
      new Set();

    state.selectedPeopleSmallDistricts =
      new Set();

    renderPeople();
  }


  function getPeopleAreaFilteredPeople(
    candidates
  ) {

    return candidates.filter(
      person => {

        const matchesDistrict =
          !state.selectedPeopleDistricts.size ||
          state.selectedPeopleDistricts.has(
            person.district
          );

        const matchesSmallDistrict =
          !state.selectedPeopleSmallDistricts.size ||
          state.selectedPeopleSmallDistricts.has(
            peopleSmallDistrictKey(person)
          );

        return (
          matchesDistrict &&
          matchesSmallDistrict
        );
      }
    );
  }


  function renderPeople() {

    const base =
      getGeneralFilteredPeople();

    const chartType =
      state.chartType;

    const chartGroup =
      state.chartGroup;

    const chartDistrict =
      state.chartDistrict;

    const chartVisible =
      base.people.filter(
        p =>
          matchesChartGroup(p, chartGroup) &&
          (
            !chartDistrict ||
            p.district === chartDistrict
          ) &&
          matchesChartTypeStatus(p, chartType)
      );


    /*
     * 名單區的大區 / 小區複選，
     * 只以目前已列出的 chartVisible 為選項來源。
     */
    updatePeopleAreaFilterUI(
      chartVisible
    );


    const visible =
      getPeopleAreaFilteredPeople(
        chartVisible
      );


    const peopleAreaActive =
      Boolean(
        state.selectedPeopleDistricts.size ||
        state.selectedPeopleSmallDistricts.size
      );

    const chartActive =
      Boolean(
        chartType ||
        chartGroup ||
        chartDistrict
      );

    if (chartActive || peopleAreaActive) {

      const sourceCount =
        chartVisible.length;

      const labels = [];

      if (chartActive) {
        labels.push('圖表篩選');
      }

      if (peopleAreaActive) {
        labels.push('大區／小區篩選');
      }

      els.resultCount.textContent =
        `目前${labels.join('＋')}：${visible.length} / ${sourceCount} 人`;

    } else {
      els.resultCount.textContent =
        `顯示 ${visible.length} / ${base.populationCount} 人`;
    }

    els.peopleBody.innerHTML =
      visible
        .map(
          p => `
            <tr>
              <td>${escapeHtml(p.district)}</td>
              <td>${escapeHtml(p.smallDistrict)}</td>
              <td><strong>${escapeHtml(p.name)}</strong></td>
              <td>${escapeHtml(p.group)}</td>
              <td>${p.baptismDate ? fmtDate(p.baptismDate) : '—'}</td>
              <td>
                ${
                  p.newBelieverStatus === 'yes'
                    ? '<span class="new-believer">初信</span>'
                    : p.newBelieverStatus === 'no'
                      ? '非初信'
                      : '日期不明'
                }
              </td>
              <td>
                <span class="status ${STATUS_CLASS[p.status]}">
                  ${escapeHtml(p.status)}
                </span>
              </td>
              <td>${p.lastAttendanceDate ? fmtDate(p.lastAttendanceDate) : '—'}</td>
              <td>${p.recentAttendance} / ${p.recentWeeks}</td>
              <td>${(p.recentRate * 100).toFixed(1)}%</td>
              <td>${p.totalAttendance}</td>
            </tr>
          `
        )
        .join('');

    document.dispatchEvent(
      new CustomEvent('analysisBaseChanged')
    );
  }


  globalThis.AttendanceDashboardAPI = {
    getAnalysisBasePeople() {
      return getGeneralFilteredPeople()
        .people
        .map(
          p => ({
            district: p.district,
            group: p.group,
            status: p.status
          })
        );
    }
  };


  /*
   * ========================================
   * 圖表篩選 UI
   * ========================================
   */

  function updateChartFilterUI() {

    if (
      !els.chartFilterIndicator ||
      !els.chartFilterText
    ) {
      return;
    }


    if (
      !state.chartType ||
      !state.chartGroup
    ) {

      els.chartFilterIndicator
        .classList
        .add('hidden');


      els.chartFilterText.textContent =
        '';


      document
        .querySelectorAll(
          '.age-filter-trigger'
        )
        .forEach(
          element =>
            element.classList.remove(
              'is-selected'
            )
        );


      return;
    }


    const typeLabel =
      state.chartType === 'care'
        ? '需加強牧養'
        : '穩定聚會';

    const description =
      state.chartDistrict
        ? `${state.chartGroup} × ${state.chartDistrict}`
        : state.chartGroup;

    els.chartFilterText.textContent =
      `圖表篩選：${typeLabel}｜${description}`;


    els.chartFilterIndicator
      .classList
      .remove('hidden');


    /*
     * 更新圖表選取外觀
     */
    document
      .querySelectorAll(
        '.age-filter-trigger'
      )
      .forEach(
        element => {

          const group =
            element.dataset.ageGroup ||
            '';


          const district =
            element.dataset.ageDistrict ||
            '';


          const type =
            element.dataset.chartType ||
            'stable';

          const selected =
            type === state.chartType &&
            group === state.chartGroup &&
            district === state.chartDistrict;


          element.classList.toggle(
            'is-selected',
            selected
          );
        }
      );
  }


  function clearChartFilter() {

    state.chartType = '';

    state.chartGroup = '';

    state.chartDistrict = '';


    renderPeople();

    updateChartFilterUI();
  }


  function applyChartFilter(
    type,
    group,
    district
  ) {

    /*
     * 點同一個項目第二次：
     * 取消。
     */
    if (
      state.chartType === type &&
      state.chartGroup === group &&
      state.chartDistrict === district
    ) {

      clearChartFilter();

      return;
    }


    state.chartType =
      type;


    state.chartGroup =
      group;


    state.chartDistrict =
      district;


    renderPeople();

    updateChartFilterUI();


    /*
     * 自動捲到人員明細。
     */
    if (
      els.peoplePanel
    ) {

      els.peoplePanel.scrollIntoView({
        behavior: 'smooth',
        block: 'start'
      });
    }
  }


  function closeMultiSelects(
    except = null
  ) {

    [
      els.groupMultiSelect,
      els.statusMultiSelect,
      els.peopleDistrictMultiSelect,
      els.peopleSmallDistrictMultiSelect
    ].forEach(
      details => {

        if (
          details &&
          details !== except
        ) {

          details.open =
            false;
        }
      }
    );
  }


  function groupBy(
    items,
    fn
  ) {

    const m =
      new Map();


    items.forEach(
      x => {

        const k =
          fn(x);


        if (
          !m.has(k)
        ) {

          m.set(
            k,
            []
          );
        }


        m.get(k).push(x);
      }
    );


    return m;
  }


  function unique(arr) {

    return [
      ...new Set(arr)
    ];
  }


  function escapeHtml(v) {

    return String(
      v ?? ''
    ).replace(
      /[&<>'"]/g,
      c => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        "'": '&#39;',
        '"': '&quot;'
      }[c])
    );
  }


  function escapeAttr(v) {

    return escapeHtml(v);
  }


  async function onAttendanceFile(file) {

    els.error.textContent =
      '';


    els.fileName.textContent =
      file.name;


    try {

      if (
        !globalThis.XLSX
      ) {

        throw new Error(
          'Excel 解析元件尚未載入。請確認網路連線，或依 README 將 SheetJS 改為本機 vendor。'
        );
      }


      const buffer =
        await file.arrayBuffer();


      analyzeWorkbook(
        buffer,
        file.name
      );


    } catch (err) {

      console.error(err);


      els.error.textContent =
        err?.message ||
        '無法讀取這份 Excel。';


      els.dashboard.classList.add(
        'hidden'
      );
    }
  }


  async function onBaptismFile(file) {

    els.error.textContent =
      '';


    els.baptismFileName.textContent =
      file.name;


    try {

      if (
        !globalThis.XLSX
      ) {

        throw new Error(
          'Excel 解析元件尚未載入。請確認網路連線，或依 README 將 SheetJS 改為本機 vendor。'
        );
      }


      const buffer =
        await file.arrayBuffer();


      analyzeBaptismWorkbook(
        buffer,
        file.name
      );


    } catch (err) {

      console.error(err);


      els.error.textContent =
        err?.message ||
        '無法讀取這份受浸報表。';
    }
  }


  function clearData() {

    state.fileName = '';

    state.rows = [];

    state.dateColumns = [];

    state.people = [];

    state.recentCols = [];

    state.baptismFileName = '';

    state.baptismByName =
      new Map();

    state.windowWeeks = 12;


    state.selectedGroups =
      new Set();


    state.selectedStatuses =
      new Set();


    state.selectedPeopleDistricts =
      new Set();

    state.selectedPeopleSmallDistricts =
      new Set();


    state.chartType = '';

    state.chartGroup = '';

    state.chartDistrict = '';


    state.selectedTrendGroups =
      new Set();

    state.trendGroupsInitialized =
      false;


    updateChartFilterUI();


    els.fileInput.value =
      '';


    els.baptismFileInput.value =
      '';


    els.fileName.textContent =
      '尚未選擇點名報表';


    els.baptismFileName.textContent =
      '尚未選擇受浸報表';


    els.error.textContent =
      '';


    els.dashboard.classList.add(
      'hidden'
    );


    els.clearBtn.disabled =
      true;


    els.searchInput.value =
      '';


    els.districtFilter.value =
      '';


    els.smallDistrictFilter.value =
      '';


    if (els.trendDistrictFilter) {
      els.trendDistrictFilter.value = '';
    }

    if (els.trendSmallDistrictFilter) {
      els.trendSmallDistrictFilter.value = '';
      els.trendSmallDistrictFilter.disabled = true;
    }


    els.totalAttendanceFilter.value =
      '';


    els.absenceWeeksFilter.value =
      '';


    els.newBelieverFilter.value =
      '';


    closeMultiSelects();
  }


  /*
   * ========================================
   * Excel
   * ========================================
   */

  els.fileInput.addEventListener(
    'change',
    e => {

      const file =
        e.target.files?.[0];


      if (file) {

        onAttendanceFile(file);
      }
    }
  );


  els.baptismFileInput.addEventListener(
    'change',
    e => {

      const file =
        e.target.files?.[0];


      if (file) {

        onBaptismFile(file);
      }
    }
  );


  /*
   * ========================================
   * 每週趨勢篩選
   * ========================================
   */

  if (els.trendDistrictButtons) {

    els.trendDistrictButtons.addEventListener(
      'click',
      e => {
        const button = e.target.closest('[data-trend-district]');
        if (!button) {
          return;
        }

        const nextDistrict =
          button.dataset.trendDistrict || '';

        if (els.trendDistrictFilter.value !== nextDistrict) {
          els.trendDistrictFilter.value = nextDistrict;
          els.trendSmallDistrictFilter.value = '';
        }

        buildTrendFilters();
        renderTrend();
        buildGroupTrendButtons();
        renderGroupTrend();
      }
    );
  }


  if (els.trendSmallDistrictButtons) {

    els.trendSmallDistrictButtons.addEventListener(
      'click',
      e => {
        const button = e.target.closest('[data-trend-small-district]');
        if (!button) {
          return;
        }

        els.trendSmallDistrictFilter.value =
          button.dataset.trendSmallDistrict || '';

        updateTrendSmallDistrictOptions();
        renderTrend();
        buildGroupTrendButtons();
        renderGroupTrend();
      }
    );
  }


  /*
   * ========================================
   * 羣組成長趨勢：iPhone App 式複選按鈕
   * ========================================
   */

  document.addEventListener('click', e => {
    const districtBtn = e.target.closest('[data-analysis-district]');
    if (districtBtn) {
      els.districtFilter.value = districtBtn.dataset.analysisDistrict || '';
      updateSmallDistrictOptions();
      renderAnalysisFilterButtons();
      renderPeople();
      return;
    }
    const smallBtn = e.target.closest('[data-analysis-small]');
    if (smallBtn) {
      els.smallDistrictFilter.value = smallBtn.dataset.analysisSmall || '';
      renderAnalysisFilterButtons(); renderPeople(); return;
    }
    const groupBtn = e.target.closest('[data-analysis-group]');
    if (groupBtn) {
      const g=groupBtn.dataset.analysisGroup;
      if (state.selectedGroups.has(g)) state.selectedGroups.delete(g); else state.selectedGroups.add(g);
      buildGroupOptions(); renderAnalysisFilterButtons(); renderPeople(); return;
    }
    const statusBtn = e.target.closest('[data-analysis-status]');
    if (statusBtn) {
      const st=statusBtn.dataset.analysisStatus;
      if (state.selectedStatuses.has(st)) state.selectedStatuses.delete(st); else state.selectedStatuses.add(st);
      buildStatusOptions(); renderAnalysisFilterButtons(); renderPeople(); return;
    }
    const believerBtn = e.target.closest('[data-analysis-believer]');
    if (believerBtn) {
      els.newBelieverFilter.value=believerBtn.dataset.analysisBeliever || '';
      renderAnalysisFilterButtons(); renderPeople(); return;
    }
  });

  if (els.groupTrendButtons) {
    els.groupTrendButtons.addEventListener(
      'click',
      e => {
        const button =
          e.target.closest(
            '[data-trend-group]'
          );

        if (!button) {
          return;
        }

        const group =
          button.dataset.trendGroup || '';

        if (!group) {
          return;
        }

        if (
          state.selectedTrendGroups.has(
            group
          )
        ) {
          state.selectedTrendGroups.delete(
            group
          );
        } else {
          state.selectedTrendGroups.add(
            group
          );
        }

        buildGroupTrendButtons();
        renderGroupTrend();
      }
    );
  }


  /*
   * ========================================
   * 大區
   * ========================================
   */

  els.districtFilter.addEventListener(
    'change',
    () => {

      closeMultiSelects();

      updateSmallDistrictOptions();
      renderAnalysisFilterButtons();

      renderPeople();
    }
  );


  /*
   * ========================================
   * 分析母體設定
   * ========================================
   */

  [
    els.totalAttendanceFilter,
    els.absenceWeeksFilter
  ].forEach(
    el => {

      const refreshPopulation =
        () => {

          closeMultiSelects();

          state.chartType = '';
          state.chartGroup = '';
          state.chartDistrict = '';

          state.selectedPeopleDistricts =
            new Set();

          state.selectedPeopleSmallDistricts =
            new Set();

          renderAll();
        };

      el.addEventListener(
        'change',
        refreshPopulation
      );
    }
  );


  /*
   * ========================================
   * 一般篩選
   * ========================================
   */

  [
    els.searchInput,
    els.smallDistrictFilter,
    els.newBelieverFilter
  ].forEach(
    el => {

      el.addEventListener(
        'input',
        () => {

          closeMultiSelects();

          renderAnalysisFilterButtons();
          renderPeople();
        }
      );


      el.addEventListener(
        'change',
        () => {

          closeMultiSelects();

          renderAnalysisFilterButtons();
          renderPeople();
        }
      );


      el.addEventListener(
        'focus',
        () =>
          closeMultiSelects()
      );
    }
  );


  /*
   * ========================================
   * 羣組
   * ========================================
   */

  els.groupFilterOptions.addEventListener(
    'change',
    e => {

      if (
        !e.target.matches(
          'input[type="checkbox"]'
        )
      ) {
        return;
      }


      if (
        e.target.checked
      ) {

        state.selectedGroups.add(
          e.target.value
        );

      } else {

        state.selectedGroups.delete(
          e.target.value
        );
      }


      updateGroupSummary();
      renderAnalysisFilterButtons();

      renderPeople();
    }
  );


  document
    .querySelectorAll(
      '[data-group-preset]'
    )
    .forEach(
      btn => {

        btn.addEventListener(
          'click',
          () =>
            setSelectedGroups(
              btn.dataset.groupPreset
            )
        );
      }
    );


  els.clearGroupFilter.addEventListener(
    'click',
    () =>
      setSelectedGroups(
        'all'
      )
  );


  els.closeGroupFilter.addEventListener(
    'click',
    () => {

      els.groupMultiSelect.open =
        false;
    }
  );


  /*
   * ========================================
   * 聚會情況
   * ========================================
   */

  els.statusFilterOptions.addEventListener(
    'change',
    e => {

      if (
        !e.target.matches(
          'input[type="checkbox"]'
        )
      ) {
        return;
      }


      if (
        e.target.checked
      ) {

        state.selectedStatuses.add(
          e.target.value
        );

      } else {

        state.selectedStatuses.delete(
          e.target.value
        );
      }


      updateStatusSummary();
      renderAnalysisFilterButtons();

      renderPeople();
    }
  );


  els.clearStatusFilter.addEventListener(
    'click',
    () =>
      setSelectedStatuses([])
  );


  els.closeStatusFilter.addEventListener(
    'click',
    () => {

      els.statusMultiSelect.open =
        false;
    }
  );


  /*
   * ========================================
   * Multi select
   * ========================================
   */

  els.groupMultiSelect.addEventListener(
    'toggle',
    () => {

      if (
        els.groupMultiSelect.open
      ) {

        closeMultiSelects(
          els.groupMultiSelect
        );
      }
    }
  );


  els.statusMultiSelect.addEventListener(
    'toggle',
    () => {

      if (
        els.statusMultiSelect.open
      ) {

        closeMultiSelects(
          els.statusMultiSelect
        );
      }
    }
  );


  document
    .querySelectorAll(
      'select'
    )
    .forEach(
      select => {

        select.addEventListener(
          'pointerdown',
          () =>
            closeMultiSelects()
        );


        select.addEventListener(
          'focus',
          () =>
            closeMultiSelects()
        );
      }
    );


  document.addEventListener(
    'pointerdown',
    e => {

      const insideGroup =
        els.groupMultiSelect.contains(
          e.target
        );


      const insideStatus =
        els.statusMultiSelect.contains(
          e.target
        );


      const insidePeopleDistrict =
        els.peopleDistrictMultiSelect &&
        els.peopleDistrictMultiSelect.contains(
          e.target
        );


      const insidePeopleSmallDistrict =
        els.peopleSmallDistrictMultiSelect &&
        els.peopleSmallDistrictMultiSelect.contains(
          e.target
        );


      if (
        !insideGroup &&
        !insideStatus &&
        !insidePeopleDistrict &&
        !insidePeopleSmallDistrict
      ) {

        closeMultiSelects();
      }
    }
  );


  /*
   * ========================================
   * 人員明細：大區 / 小區複選
   * ========================================
   */

  if (els.peopleDistrictOptions) {

    els.peopleDistrictOptions.addEventListener(
      'click',
      e => {

        const button =
          e.target.closest(
            '[data-people-district]'
          );

        if (!button) {
          return;
        }

        const district =
          button.dataset.peopleDistrict || '';

        if (!district) {
          state.selectedPeopleDistricts = new Set();
          state.selectedPeopleSmallDistricts = new Set();
          renderPeople();
          return;
        }

        if (state.selectedPeopleDistricts.has(district)) {
          state.selectedPeopleDistricts.delete(district);
        } else {
          state.selectedPeopleDistricts.add(district);
        }

        renderPeople();
      }
    );
  }


  if (els.peopleSmallDistrictOptions) {

    els.peopleSmallDistrictOptions.addEventListener(
      'click',
      e => {

        const button =
          e.target.closest(
            '[data-people-small-district]'
          );

        if (!button) {
          return;
        }

        const key =
          button.dataset.peopleSmallDistrict || '';

        if (!key) {
          state.selectedPeopleSmallDistricts = new Set();
          renderPeople();
          return;
        }

        if (state.selectedPeopleSmallDistricts.has(key)) {
          state.selectedPeopleSmallDistricts.delete(key);
        } else {
          state.selectedPeopleSmallDistricts.add(key);
        }

        renderPeople();
      }
    );
  }


  if (els.clearPeopleDistrictFilter) {

    els.clearPeopleDistrictFilter.addEventListener(
      'click',
      () => {

        state.selectedPeopleDistricts =
          new Set();

        state.selectedPeopleSmallDistricts =
          new Set();

        renderPeople();
      }
    );
  }


  if (els.clearPeopleSmallDistrictFilter) {

    els.clearPeopleSmallDistrictFilter.addEventListener(
      'click',
      () => {

        state.selectedPeopleSmallDistricts =
          new Set();

        renderPeople();
      }
    );
  }


  if (els.clearPeopleAreaFilter) {

    els.clearPeopleAreaFilter.addEventListener(
      'click',
      clearPeopleAreaFilters
    );
  }


  if (els.closePeopleDistrictFilter) {

    els.closePeopleDistrictFilter.addEventListener(
      'click',
      () => {
        els.peopleDistrictMultiSelect.open = false;
      }
    );
  }


  if (els.closePeopleSmallDistrictFilter) {

    els.closePeopleSmallDistrictFilter.addEventListener(
      'click',
      () => {
        els.peopleSmallDistrictMultiSelect.open = false;
      }
    );
  }


  if (els.peopleDistrictMultiSelect) {

    els.peopleDistrictMultiSelect.addEventListener(
      'toggle',
      () => {

        if (els.peopleDistrictMultiSelect.open) {
          closeMultiSelects(
            els.peopleDistrictMultiSelect
          );
        }
      }
    );
  }


  if (els.peopleSmallDistrictMultiSelect) {

    els.peopleSmallDistrictMultiSelect.addEventListener(
      'toggle',
      () => {

        if (els.peopleSmallDistrictMultiSelect.open) {
          closeMultiSelects(
            els.peopleSmallDistrictMultiSelect
          );
        }
      }
    );
  }


  /*
   * ========================================
   * 圖表 → 人員明細
   * ========================================
   *
   * age-analysis.js 產生的按鈕
   * 都會帶有 .age-filter-trigger。
   */

  document.addEventListener(
    'click',
    e => {

      const trigger =
        e.target.closest(
          '.age-filter-trigger'
        );


      if (
        !trigger
      ) {
        return;
      }


      const group =
        trigger.dataset.ageGroup ||
        '';


      const district =
        trigger.dataset.ageDistrict ||
        '';


      const type =
        trigger.dataset.chartType ||
        'stable';


      if (
        !group
      ) {
        return;
      }


      applyChartFilter(
        type,
        group,
        district
      );
    }
  );


  /*
   * 四張分析圖重新產生 DOM 後，
   * 恢復目前圖表篩選的選取外觀。
   */

  document.addEventListener(
    'ageAnalysisRendered',
    updateChartFilterUI
  );


  /*
   * 人員明細右上角：
   * 清除圖表篩選
   */

  if (
    els.clearChartFilter
  ) {

    els.clearChartFilter.addEventListener(
      'click',
      clearChartFilter
    );
  }


  /*
   * ========================================
   * 清除全部資料
   * ========================================
   */

  els.clearBtn.addEventListener(
    'click',
    clearData
  );

})();