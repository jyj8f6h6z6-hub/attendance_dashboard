(() => {
  'use strict';

  const STABLE_STATUSES = new Set([
    '週週聚會',
    '常聚會'
  ]);

  const CARE_STATUSES = new Set([
    '偶聚會',
    '零星聚會'
  ]);

  const INACTIVE_STATUSES = new Set([
    '近期未聚會'
  ]);

  const MAIN_GROUP_ORDER = [
    '年長',
    '中壯',
    '青壯',
    '青職',
    '大學',
    '青少年'
  ];

  const STUDENT_GROUP_ORDER = [
    '大學',
    '高中',
    '國中'
  ];

  const KNOWN_GROUPS = new Set([
    '年長', '中壯', '青壯', '青職', '大學',
    '高中', '國中', '中學', '國小', '學齡前'
  ]);

  const $ = id => document.getElementById(id);

  function escapeHtml(value) {
    return String(value ?? '').replace(
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

  function escapeAttr(value) {
    return escapeHtml(value);
  }

  function cleanText(value) {
    return String(value ?? '')
      .trim()
      .replace(/\s+/g, '');
  }

  function getMainGroup(group) {
    const value = cleanText(group);

    if ([
      '年長', '中壯', '青壯', '青職', '大學'
    ].includes(value)) {
      return value;
    }

    if ([
      '高中', '國中', '中學'
    ].includes(value)) {
      return '青少年';
    }

    return null;
  }

  function getAnalysisBasePeople() {
    const api = globalThis.AttendanceDashboardAPI;

    if (!api || typeof api.getAnalysisBasePeople !== 'function') {
      return [];
    }

    return api.getAnalysisBasePeople()
      .map(person => ({
        district: String(person.district ?? '').trim(),
        group: cleanText(person.group),
        status: String(person.status ?? '').trim()
      }));
  }

  function getDistrictOrder(people) {
    const preferred = [
      '一大區', '二大區', '三大區'
    ];

    const found = [
      ...new Set(
        people
          .map(p => p.district)
          .filter(Boolean)
      )
    ];

    return [
      ...preferred.filter(d => found.includes(d)),
      ...found
        .filter(d => !preferred.includes(d))
        .sort((a, b) => a.localeCompare(b, 'zh-Hant'))
    ];
  }

  function createGroupStats(groups, districts) {
    const stats = {};

    groups.forEach(group => {
      stats[group] = {
        total: 0,
        districts: {}
      };

      districts.forEach(district => {
        stats[group].districts[district] = 0;
      });
    });

    return stats;
  }

  function districtClass(index) {
    return `district-color-${(index % 6) + 1}`;
  }

  function buildBars(
    order,
    stats,
    total,
    districts,
    chartType
  ) {
    if (!total) {
      return `
        <div class="age-empty">
          目前沒有可分析資料
        </div>
      `;
    }

    const maxCount = Math.max(
      ...order.map(group => stats[group]?.total || 0),
      1
    );

    return order.map(group => {
      const item = stats[group] || {
        total: 0,
        districts: {}
      };

      const count = item.total;
      const overallPct = total ? count / total * 100 : 0;
      const barWidth = count ? count / maxCount * 100 : 0;

      const segments = districts.map((district, index) => {
        const districtCount = item.districts[district] || 0;
        if (!districtCount) {
          return '';
        }

        const districtPct = count
          ? districtCount / count * 100
          : 0;

        const label = districtPct >= 10
          ? `${districtCount}`
          : '';

        return `
          <button
            type="button"
            class="age-district-segment age-filter-trigger ${districtClass(index)}"
            data-chart-type="${escapeAttr(chartType)}"
            data-age-group="${escapeAttr(group)}"
            data-age-district="${escapeAttr(district)}"
            style="width:${districtPct}%;"
            title="${escapeAttr(`查看 ${group}・${district}：${districtCount} 人`)}"
            aria-label="${escapeAttr(`查看 ${group} ${district} ${districtCount} 人`)}"
          >
            ${label ? `<span>${districtCount}</span>` : ''}
          </button>
        `;
      }).join('');

      return `
        <div class="age-row">
          <button
            type="button"
            class="age-label age-filter-trigger"
            data-chart-type="${escapeAttr(chartType)}"
            data-age-group="${escapeAttr(group)}"
            title="${escapeAttr(`查看全部${group}人員`)}"
          >
            ${escapeHtml(group)}
          </button>

          <div class="age-track">
            <div
              class="age-stack"
              style="width:${barWidth}%;"
            >
              ${segments}
            </div>
          </div>

          <div class="age-number">
            <strong>${count}</strong>
            <span>人</span>
            <small>${overallPct.toFixed(1)}%</small>
          </div>
        </div>
      `;
    }).join('');
  }

  function buildDistrictLegend(districts, people) {
    if (!districts.length) {
      return '';
    }

    const total = people.length;

    return `
      <div class="district-legend">
        ${districts.map((district, index) => {
          const count = people.filter(
            person => person.district === district
          ).length;

          const pct = total
            ? count / total * 100
            : 0;

          return `
            <div class="district-legend-item">
              <i class="district-dot ${districtClass(index)}"></i>
              <span class="district-name">${escapeHtml(district)}</span>
              <strong>${count}</strong>
              <span class="district-unit">人</span>
              <small>${pct.toFixed(1)}%</small>
            </div>
          `;
        }).join('')}
      </div>
    `;
  }

  function makeStats(people, districts) {
    const mainStats = createGroupStats(
      MAIN_GROUP_ORDER,
      districts
    );

    const studentStats = createGroupStats(
      STUDENT_GROUP_ORDER,
      districts
    );

    const unknown = new Map();

    people.forEach(person => {
      const originalGroup = cleanText(person.group);
      const district = person.district;
      const mainGroup = getMainGroup(originalGroup);

      if (mainGroup) {
        mainStats[mainGroup].total++;
        if (mainStats[mainGroup].districts[district] === undefined) {
          mainStats[mainGroup].districts[district] = 0;
        }
        mainStats[mainGroup].districts[district]++;
      }

      if (STUDENT_GROUP_ORDER.includes(originalGroup)) {
        studentStats[originalGroup].total++;
        if (studentStats[originalGroup].districts[district] === undefined) {
          studentStats[originalGroup].districts[district] = 0;
        }
        studentStats[originalGroup].districts[district]++;
      }

      if (originalGroup && !KNOWN_GROUPS.has(originalGroup)) {
        unknown.set(
          originalGroup,
          (unknown.get(originalGroup) || 0) + 1
        );
      }
    });

    return {
      mainStats,
      studentStats,
      unknown
    };
  }

  function renderSet({
    people,
    chartType,
    ageTotalId,
    ageDistributionId,
    studentTotalId,
    studentDistributionId,
    studentShareId,
    warningId,
    studentLabel,
    shareLabel,
    warningSubject
  }) {
    const districts = getDistrictOrder(people);
    const {
      mainStats,
      studentStats,
      unknown
    } = makeStats(people, districts);

    const mainTotal = MAIN_GROUP_ORDER.reduce(
      (sum, group) => sum + mainStats[group].total,
      0
    );

    const ageTotal = $(ageTotalId);
    if (ageTotal) {
      ageTotal.textContent = `${mainTotal} 人`;
    }

    const ageDistribution = $(ageDistributionId);
    if (ageDistribution) {
      const mainPeople = people.filter(
        person => getMainGroup(person.group)
      );

      ageDistribution.innerHTML =
        buildBars(
          MAIN_GROUP_ORDER,
          mainStats,
          mainTotal,
          districts,
          chartType
        ) +
        buildDistrictLegend(
          districts,
          mainPeople
        );
    }

    const studentTotal = STUDENT_GROUP_ORDER.reduce(
      (sum, group) => sum + studentStats[group].total,
      0
    );

    const studentTotalEl = $(studentTotalId);
    if (studentTotalEl) {
      studentTotalEl.textContent = `${studentTotal} 人`;
    }

    const studentDistribution = $(studentDistributionId);
    if (studentDistribution) {
      const studentPeople = people.filter(
        person => STUDENT_GROUP_ORDER.includes(person.group)
      );

      studentDistribution.innerHTML =
        buildBars(
          STUDENT_GROUP_ORDER,
          studentStats,
          studentTotal,
          districts,
          chartType
        ) +
        buildDistrictLegend(
          districts,
          studentPeople
        );
    }

    const studentPct = mainTotal
      ? studentTotal / mainTotal * 100
      : 0;

    const studentShare = $(studentShareId);
    if (studentShare) {
      studentShare.innerHTML = `
        <span>
          ${escapeHtml(studentLabel)}
          <strong>${studentTotal}</strong>
          人
        </span>
        <span>
          ${escapeHtml(shareLabel)}
          <strong>${studentPct.toFixed(1)}%</strong>
        </span>
      `;
    }

    const warning = $(warningId);
    if (warning) {
      const unknownCount = [...unknown.values()]
        .reduce((sum, count) => sum + count, 0);

      if (unknownCount) {
        const details = [...unknown.entries()]
          .map(([group, count]) =>
            `${escapeHtml(group)} ${count}人`
          )
          .join('、');

        warning.innerHTML = `
          發現
          <strong>${unknownCount}</strong>
          位${escapeHtml(warningSubject)}的群組名稱不在目前設定中
          （${details}）
        `;
        warning.classList.remove('hidden');
      } else {
        warning.textContent = '';
        warning.classList.add('hidden');
      }
    }
  }

  function renderAnalysis() {
    const basePeople = getAnalysisBasePeople();

    const stablePeople = basePeople.filter(
      person => STABLE_STATUSES.has(person.status)
    );

    const carePeople = basePeople.filter(
      person => CARE_STATUSES.has(person.status)
    );

    const inactivePeople = basePeople.filter(
      person => INACTIVE_STATUSES.has(person.status)
    );

    renderSet({
      people: stablePeople,
      chartType: 'stable',
      ageTotalId: 'ageTotal',
      ageDistributionId: 'ageDistribution',
      studentTotalId: 'studentAgeTotal',
      studentDistributionId: 'studentAgeDistribution',
      studentShareId: 'studentShare',
      warningId: 'ageUnknown',
      studentLabel: '穩定聚會學生',
      shareLabel: '占穩定聚會者',
      warningSubject: '穩定聚會者'
    });

    renderSet({
      people: carePeople,
      chartType: 'care',
      ageTotalId: 'careAgeTotal',
      ageDistributionId: 'careAgeDistribution',
      studentTotalId: 'careStudentAgeTotal',
      studentDistributionId: 'careStudentAgeDistribution',
      studentShareId: 'careStudentShare',
      warningId: 'careAgeUnknown',
      studentLabel: '需加強牧養學生',
      shareLabel: '占需加強牧養者',
      warningSubject: '需加強牧養者'
    });

    renderSet({
      people: inactivePeople,
      chartType: 'inactive',
      ageTotalId: 'inactiveAgeTotal',
      ageDistributionId: 'inactiveAgeDistribution',
      studentTotalId: 'inactiveStudentAgeTotal',
      studentDistributionId: 'inactiveStudentAgeDistribution',
      studentShareId: 'inactiveStudentShare',
      warningId: 'inactiveAgeUnknown',
      studentLabel: '待關心聖徒學生',
      shareLabel: '占待關心聖徒',
      warningSubject: '待關心聖徒'
    });

    document.dispatchEvent(
      new CustomEvent('ageAnalysisRendered')
    );
  }

  function clearAnalysis() {
    [
      'ageTotal',
      'studentAgeTotal',
      'careAgeTotal',
      'careStudentAgeTotal',
      'inactiveAgeTotal',
      'inactiveStudentAgeTotal'
    ].forEach(id => {
      const element = $(id);
      if (element) {
        element.textContent = '—';
      }
    });

    [
      'ageDistribution',
      'studentAgeDistribution',
      'studentShare',
      'careAgeDistribution',
      'careStudentAgeDistribution',
      'careStudentShare',
      'inactiveAgeDistribution',
      'inactiveStudentAgeDistribution',
      'inactiveStudentShare'
    ].forEach(id => {
      const element = $(id);
      if (element) {
        element.innerHTML = '';
      }
    });

    [
      'ageUnknown',
      'careAgeUnknown',
      'inactiveAgeUnknown'
    ].forEach(id => {
      const warning = $(id);
      if (warning) {
        warning.textContent = '';
        warning.classList.add('hidden');
      }
    });
  }

  function install() {
    document.addEventListener(
      'analysisBaseChanged',
      renderAnalysis
    );

    const clearBtn = $('clearBtn');
    if (clearBtn) {
      clearBtn.addEventListener(
        'click',
        () => setTimeout(clearAnalysis, 0)
      );
    }

    renderAnalysis();
  }

  if (document.readyState === 'loading') {
    document.addEventListener(
      'DOMContentLoaded',
      install
    );
  } else {
    install();
  }
})();

/* =========================================================
   v1.1.23｜羣組季度比較＋人員對照
   - 比較各羣組「平均每週聚會人數」：本季、上季、去年同季
   - 本季尚未結束時，去年同季只取相同的前幾週，條件才公平
   - 平均少於 10 人時只顯示數字，不標示升降
   - 點羣組名稱：在表格下方展開「消失／新增／持續」名單（小區＋姓名）
   - 附本季人員流動：新出現、恢復聚會、轉為近期未聚會
========================================================= */
(() => {
  'use strict';

  const GROUPS = ['年長', '中壯', '青壯', '青職', '大學', '青少年'];
  const STUDENT_GROUPS = new Set(['大學', '青少年']);
  const PREFERRED_DISTRICTS = ['一大區', '二大區', '三大區'];
  const SMALL_THRESHOLD = 10;
  const FLAT_THRESHOLD = 0.05;
  const INACTIVE_WEEKS = 8;
  const MS_WEEK = 7 * 24 * 60 * 60 * 1000;
  const MODE_KEY = 'attendance-dashboard-group-quarter-mode-v1';

  let mode = 'all';
  try {
    if (localStorage.getItem(MODE_KEY) === 'districts') mode = 'districts';
  } catch (_) {}

  /* 目前展開的名單：{ scope, group }；compareTo：lastYear／previous */
  let openDetail = null;
  let compareTo = 'lastYear';

  let cache = { rows: null, key: '', persons: [] };

  const $ = id => document.getElementById(id);

  function esc(value) {
    return String(value ?? '').replace(/[&<>'"]/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
    }[c]));
  }

  function normalizeGroup(value) {
    const g = String(value ?? '').trim().replace(/\s+/g, '');
    if (['國中', '高中', '中學'].includes(g)) return '青少年';
    if (g === '大專') return '大學';
    return GROUPS.includes(g) ? g : '';
  }

  /* 季度：1–4 月＝Q1、5–8 月＝Q2、9–12 月＝Q3 */
  function quarterOf(date) {
    const m = date.getMonth() + 1;
    return { year: date.getFullYear(), quarter: m <= 4 ? 1 : m <= 8 ? 2 : 3 };
  }

  function previousQuarter(p) {
    return p.quarter > 1
      ? { year: p.year, quarter: p.quarter - 1 }
      : { year: p.year - 1, quarter: 3 };
  }

  function quarterKey(p) {
    return `${p.year}Q${p.quarter}`;
  }

  function firstSunday(p) {
    const d = new Date(p.year, (p.quarter - 1) * 4, 1);
    d.setDate(d.getDate() + (7 - d.getDay()) % 7);
    return d;
  }

  function quarterEnd(p) {
    return new Date(p.year, (p.quarter - 1) * 4 + 4, 0);
  }

  function weekOrdinal(date, p) {
    return Math.round((date - firstSunday(p)) / MS_WEEK);
  }

  function totalSundays(p) {
    return Math.floor((quarterEnd(p) - firstSunday(p)) / MS_WEEK) + 1;
  }

  function columnIndexes(dateColumns, p) {
    const key = quarterKey(p);
    const list = [];
    dateColumns.forEach((c, i) => {
      if (quarterKey(quarterOf(c.date)) === key) list.push(i);
    });
    return list;
  }

  function getPeriods(dateColumns) {
    const latestDate = dateColumns[dateColumns.length - 1].date;
    const current = quarterOf(latestDate);
    const previous = previousQuarter(current);
    const lastYear = { year: current.year - 1, quarter: current.quarter };
    const total = totalSundays(current);
    const elapsed = Math.min(total, weekOrdinal(latestDate, current) + 1);
    const inProgress = elapsed < total;
    const lastYearAll = columnIndexes(dateColumns, lastYear);

    return {
      current,
      previous,
      lastYear,
      total,
      elapsed,
      inProgress,
      latestDate,
      currentIdx: columnIndexes(dateColumns, current),
      previousIdx: columnIndexes(dateColumns, previous),
      lastYearIdx: inProgress
        ? lastYearAll.filter(i => weekOrdinal(dateColumns[i].date, lastYear) < elapsed)
        : lastYearAll
    };
  }

  function getPersons(data) {
    const key = `${data.dateColumns.length}|${data.people.map(p => p.rowNumber).join(',')}`;
    if (cache.rows === data.rows && cache.key === key) return cache.persons;

    const persons = [];
    data.people.forEach(p => {
      const group = normalizeGroup(p.group);
      if (!group) return;
      const row = data.rows[p.rowNumber - 1] || [];
      const att = new Uint8Array(data.dateColumns.length);
      data.dateColumns.forEach((c, i) => {
        if (data.isAttendance(row[c.col])) att[i] = 1;
      });
      persons.push({
        name: String(p.name ?? '').trim(),
        smallDistrict: String(p.smallDistrict ?? '').trim(),
        district: String(p.district ?? '').trim(),
        group,
        att
      });
    });

    cache = { rows: data.rows, key, persons };
    return persons;
  }

  function average(persons, idxs) {
    if (!idxs.length) return null;
    let sum = 0;
    persons.forEach(p => idxs.forEach(i => { sum += p.att[i]; }));
    return sum / idxs.length;
  }

  function getFlows(persons, dateColumns, periods) {
    const idxs = periods.currentIdx;
    if (!idxs.length) return null;

    const start = idxs[0];
    const quarterStart = dateColumns[start].date;
    let newcomers = 0;
    let returned = 0;
    let lapsed = 0;

    persons.forEach(p => {
      const firstThis = idxs.find(i => p.att[i]);

      if (firstThis !== undefined) {
        let before = -1;
        for (let i = start - 1; i >= 0; i--) {
          if (p.att[i]) { before = i; break; }
        }

        if (before < 0) {
          newcomers++;
        } else {
          const missed = Math.round(
            (dateColumns[firstThis].date - dateColumns[before].date) / MS_WEEK
          ) - 1;
          if (missed > INACTIVE_WEEKS) returned++;
        }
      }

      let last = -1;
      for (let i = p.att.length - 1; i >= 0; i--) {
        if (p.att[i]) { last = i; break; }
      }

      if (last >= 0) {
        const lapseDate = new Date(
          dateColumns[last].date.getTime() + (INACTIVE_WEEKS + 1) * MS_WEEK
        );
        if (lapseDate >= quarterStart && lapseDate <= periods.latestDate) lapsed++;
      }
    });

    return { newcomers, returned, lapsed };
  }

  function fmt(value) {
    return value === null ? '—' : value.toFixed(1);
  }

  function diffCell(cur, base) {
    if (cur === null || base === null) {
      return '<td class="gq-diff is-na">—</td>';
    }

    const d = cur - base;
    const abs = Math.abs(d).toFixed(1);

    if (Math.max(cur, base) < SMALL_THRESHOLD) {
      return `<td class="gq-diff is-small" title="人數少，增減僅供參考">${d > 0 ? '+' : d < 0 ? '−' : ''}${abs}</td>`;
    }

    if (Math.abs(d) < FLAT_THRESHOLD) {
      return '<td class="gq-diff is-flat">― 0.0</td>';
    }

    return d > 0
      ? `<td class="gq-diff is-up">▲ ${abs}</td>`
      : `<td class="gq-diff is-down">▼ ${abs}</td>`;
  }

  function sortPeople(list) {
    return list.sort((a, b) =>
      a.smallDistrict.localeCompare(b.smallDistrict, 'zh-Hant', { numeric: true }) ||
      a.name.localeCompare(b.name, 'zh-Hant')
    );
  }

  function buildDetail(groupPeople, periods, scopeLabel, group) {
    const cmpIdx = compareTo === 'previous' ? periods.previousIdx : periods.lastYearIdx;
    const cmpLabel = compareTo === 'previous'
      ? `上季 ${quarterKey(periods.previous)}（${cmpIdx.length} 週）`
      : `去年同季 ${quarterKey(periods.lastYear)}（${periods.inProgress ? '前 ' : ''}${cmpIdx.length} 週）`;

    const curAvg = average(groupPeople, periods.currentIdx);
    const cmpAvg = average(groupPeople, cmpIdx);

    const lost = [];
    const added = [];
    const kept = [];

    groupPeople.forEach(p => {
      const now = periods.currentIdx.some(i => p.att[i]);
      const then = cmpIdx.some(i => p.att[i]);
      if (now && !then) added.push(p);
      else if (!now && then) lost.push(p);
      else if (now && then) kept.push(p);
    });

    const column = (title, cls, list) => `
      <div class="gq-detail-col ${cls}">
        <h5>${title} <strong>${list.length}</strong> 人</h5>
        ${list.length
          ? `<ul>${sortPeople(list).map(p => `<li><span>${esc(p.smallDistrict || '—')}</span>${esc(p.name)}</li>`).join('')}</ul>`
          : '<p class="gq-detail-empty">無</p>'}
      </div>
    `;

    const compareButton = (value, label) => `
      <button type="button" class="group-trend-mode-btn${compareTo === value ? ' is-selected' : ''}" data-gq-compare="${value}" aria-pressed="${compareTo === value}">${label}</button>
    `;

    return `
      <div class="gq-detail">
        <div class="gq-detail-head">
          <div>
            <strong>${esc(scopeLabel)}｜${esc(group)}</strong>
            <span>本季 ${quarterKey(periods.current)}（${periods.currentIdx.length} 週）對比 ${cmpLabel}；平均 ${fmt(cmpAvg)} → ${fmt(curAvg)} 人</span>
          </div>
          <div class="gq-detail-actions">
            <div class="group-trend-mode">
              ${compareButton('lastYear', '對去年同季')}
              ${compareButton('previous', '對上季')}
            </div>
            <button type="button" class="gq-detail-close" data-gq-close>收起 ✕</button>
          </div>
        </div>
        ${cmpIdx.length
          ? `<div class="gq-detail-grid">
              ${column('消失', 'is-lost', lost)}
              ${column('新增', 'is-added', added)}
              ${column('持續', 'is-kept', kept)}
            </div>`
          : '<p class="gq-detail-empty">比較期間沒有資料。</p>'}
        <p class="gq-detail-note">消失＝比較期間有出席、本季尚未出席；新增＝本季有出席、比較期間未出席；持續＝兩段期間都有出席。已從名冊移除或被統計名單設定排除的人不會列出。</p>
      </div>
    `;
  }

  function buildTable(persons, periods, dateColumns, title, scopeKey) {
    const items = GROUPS
      .map(g => ({
        label: g,
        people: persons.filter(p => p.group === g),
        student: STUDENT_GROUPS.has(g)
      }))
      .filter(item => item.people.length);

    if (!items.length) return '';

    items.push({ label: '合計', people: persons, total: true });

    const body = items.map(item => {
      const cur = average(item.people, periods.currentIdx);
      const prev = average(item.people, periods.previousIdx);
      const ly = average(item.people, periods.lastYearIdx);
      const mark = item.student
        ? '<span class="gq-mark" title="依目前名冊判定，僅供參考">※</span>'
        : '';
      const isOpen = Boolean(
        openDetail && openDetail.scope === scopeKey && openDetail.group === item.label
      );

      const label = item.total
        ? esc(item.label)
        : `<button type="button" class="gq-row-btn${isOpen ? ' is-open' : ''}" data-gq-open="${esc(scopeKey)}" data-gq-group="${esc(item.label)}" aria-expanded="${isOpen}" title="查看名單">${esc(item.label)}${mark}<span class="gq-caret" aria-hidden="true">${isOpen ? '▾' : '▸'}</span></button>`;

      return `
        <tr${item.total ? ' class="gq-total"' : ''}>
          <th scope="row">${label}</th>
          <td>${fmt(cur)}</td>
          <td>${fmt(prev)}</td>
          ${diffCell(cur, prev)}
          <td>${fmt(ly)}</td>
          ${diffCell(cur, ly)}
        </tr>
      `;
    }).join('');

    const openItem = openDetail && openDetail.scope === scopeKey
      ? items.find(item => !item.total && item.label === openDetail.group)
      : null;

    const detailHtml = openItem
      ? buildDetail(openItem.people, periods, title || '全會所', openItem.label)
      : '';

    const lyNote = periods.inProgress ? `・前 ${periods.elapsed} 週` : '';
    const f = getFlows(persons, dateColumns, periods);

    const flowHtml = f
      ? `
        <div class="gq-flow">
          <span>本季人員流動</span>
          <span>新出現 <strong>${f.newcomers}</strong> 人</span>
          <span>恢復聚會 <strong>${f.returned}</strong> 人</span>
          <span>轉為近期未聚會 <strong>${f.lapsed}</strong> 人</span>
        </div>
      `
      : '';

    return `
      <div class="gq-block">
        ${title ? `<h4 class="gq-block-title">${esc(title)}</h4>` : ''}
        <div class="table-wrap">
          <table class="gq-table">
            <thead>
              <tr>
                <th>羣組</th>
                <th>本季平均<small>${quarterKey(periods.current)}</small></th>
                <th>上季<small>${quarterKey(periods.previous)}</small></th>
                <th>較上季</th>
                <th>去年同季<small>${quarterKey(periods.lastYear)}${lyNote}</small></th>
                <th>較去年同季</th>
              </tr>
            </thead>
            <tbody>${body}</tbody>
          </table>
        </div>
        ${detailHtml}
        ${flowHtml}
      </div>
    `;
  }

  function updateModeButtons() {
    document.querySelectorAll('[data-gq-mode]').forEach(button => {
      const selected = button.dataset.gqMode === mode;
      button.classList.toggle('is-selected', selected);
      button.setAttribute('aria-pressed', selected ? 'true' : 'false');
    });
  }

  function render() {
    const container = $('groupQuarterTables');
    if (!container) return;

    updateModeButtons();

    const api = globalThis.AttendanceDashboardAPI;
    const data = api && typeof api.getGroupQuarterData === 'function'
      ? api.getGroupQuarterData()
      : null;
    const meta = $('groupQuarterMeta');

    if (!data || !data.dateColumns.length) {
      container.innerHTML = '<div class="trend-empty compact">沒有可分析資料</div>';
      if (meta) meta.innerHTML = '';
      return;
    }

    const persons = getPersons(data);
    const periods = getPeriods(data.dateColumns);

    if (meta) {
      meta.innerHTML = `
        <span>本季 ${quarterKey(periods.current)}${periods.inProgress ? `：進行中（已 ${periods.elapsed}／${periods.total} 週）` : '：已完整'}｜點羣組名稱可查看名單</span>
        <span>${periods.inProgress ? `去年同季取相同的前 ${periods.elapsed} 週比較` : '去年同季取整季比較'}</span>
      `;
    }

    if (mode === 'districts') {
      const found = [...new Set(persons.map(p => p.district).filter(Boolean))];
      const districts = [
        ...PREFERRED_DISTRICTS.filter(d => found.includes(d)),
        ...found
          .filter(d => !PREFERRED_DISTRICTS.includes(d))
          .sort((a, b) => a.localeCompare(b, 'zh-Hant', { numeric: true }))
      ];

      container.innerHTML = districts
        .map(d => buildTable(persons.filter(p => p.district === d), periods, data.dateColumns, d, d))
        .join('') || '<div class="trend-empty compact">沒有大區資料</div>';
    } else {
      container.innerHTML =
        buildTable(persons, periods, data.dateColumns, '', 'all') ||
        '<div class="trend-empty compact">沒有可分析的羣組</div>';
    }
  }

  document.addEventListener('click', e => {
    const modeButton = e.target.closest('[data-gq-mode]');
    if (modeButton) {
      mode = modeButton.dataset.gqMode === 'districts' ? 'districts' : 'all';
      openDetail = null;
      try { localStorage.setItem(MODE_KEY, mode); } catch (_) {}
      render();
      return;
    }

    const openButton = e.target.closest('[data-gq-open]');
    if (openButton) {
      const scope = openButton.dataset.gqOpen || '';
      const group = openButton.dataset.gqGroup || '';
      openDetail = openDetail && openDetail.scope === scope && openDetail.group === group
        ? null
        : { scope, group };
      render();
      return;
    }

    const compareButton = e.target.closest('[data-gq-compare]');
    if (compareButton) {
      compareTo = compareButton.dataset.gqCompare === 'previous' ? 'previous' : 'lastYear';
      render();
      return;
    }

    if (e.target.closest('[data-gq-close]')) {
      openDetail = null;
      render();
    }
  });

  document.addEventListener('analysisBaseChanged', render);

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', render);
  } else {
    render();
  }
})();