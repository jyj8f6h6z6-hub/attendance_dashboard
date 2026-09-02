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

    document.dispatchEvent(
      new CustomEvent('ageAnalysisRendered')
    );
  }

  function clearAnalysis() {
    [
      'ageTotal',
      'studentAgeTotal',
      'careAgeTotal',
      'careStudentAgeTotal'
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
      'careStudentShare'
    ].forEach(id => {
      const element = $(id);
      if (element) {
        element.innerHTML = '';
      }
    });

    [
      'ageUnknown',
      'careAgeUnknown'
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
