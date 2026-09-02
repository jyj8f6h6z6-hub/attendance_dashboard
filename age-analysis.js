(() => {
  'use strict';


  const STABLE_STATUSES =
    new Set([
      '週週聚會',
      '常聚會'
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


  const KNOWN_GROUPS =
    new Set([
      '年長',
      '中壯',
      '青壯',
      '青職',
      '大學',
      '高中',
      '國中',
      '中學',
      '國小',
      '學齡前'
    ]);


  const $ =
    id =>
      document.getElementById(id);


  function escapeHtml(value) {

    return String(
      value ?? ''
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


  function escapeAttr(value) {

    return escapeHtml(value);
  }


  function cleanText(value) {

    return String(
      value ?? ''
    )
      .trim()
      .replace(
        /\s+/g,
        ''
      );
  }


  function getMainGroup(group) {

    const value =
      cleanText(group);


    if (
      value === '年長'
    ) {
      return '年長';
    }


    if (
      value === '中壯'
    ) {
      return '中壯';
    }


    if (
      value === '青壯'
    ) {
      return '青壯';
    }


    if (
      value === '青職'
    ) {
      return '青職';
    }


    if (
      value === '大學'
    ) {
      return '大學';
    }


    if (
      value === '高中' ||
      value === '國中' ||
      value === '中學'
    ) {
      return '青少年';
    }


    return null;
  }


  /*
   * 直接從目前「人員明細」
   * 取得畫面上正在顯示的人。
   *
   * 因此原本上方分析篩選
   * 仍然可以影響兩張圖。
   */

  function getVisibleStablePeople() {

    const tbody =
      document.querySelector(
        '#peopleTable tbody'
      );


    if (!tbody) {
      return [];
    }


    return [
      ...tbody.querySelectorAll(
        'tr'
      )
    ]
      .map(
        row => {

          const cells =
            row.querySelectorAll(
              'td'
            );


          if (
            cells.length < 7
          ) {
            return null;
          }


          const district =
            String(
              cells[0]
                .textContent ??
              ''
            ).trim();


          const group =
            cleanText(
              cells[3]
                .textContent
            );


          const status =
            String(
              cells[6]
                .textContent ??
              ''
            ).trim();


          if (
            !STABLE_STATUSES.has(
              status
            )
          ) {
            return null;
          }


          return {
            district,
            group,
            status
          };
        }
      )
      .filter(Boolean);
  }


  function getDistrictOrder(
    people
  ) {

    const preferred = [
      '一大區',
      '二大區',
      '三大區'
    ];


    const found =
      [
        ...new Set(
          people
            .map(
              p =>
                p.district
            )
            .filter(Boolean)
        )
      ];


    const result = [];


    preferred.forEach(
      district => {

        if (
          found.includes(
            district
          )
        ) {

          result.push(
            district
          );
        }
      }
    );


    found
      .filter(
        district =>
          !result.includes(
            district
          )
      )
      .sort(
        (a, b) =>
          a.localeCompare(
            b,
            'zh-Hant'
          )
      )
      .forEach(
        district =>
          result.push(
            district
          )
      );


    return result;
  }


  function createGroupStats(
    groups,
    districts
  ) {

    const stats = {};


    groups.forEach(
      group => {

        stats[group] = {
          total: 0,
          districts: {}
        };


        districts.forEach(
          district => {

            stats[group]
              .districts[
                district
              ] = 0;
          }
        );
      }
    );


    return stats;
  }


  function districtClass(
    index
  ) {

    const number =
      (
        index % 6
      ) + 1;


    return (
      `district-color-${number}`
    );
  }


  /*
   * ========================================
   * 長條圖
   * ========================================
   *
   * 群組名稱：
   *   data-age-group
   *
   * 彩色區塊：
   *   data-age-group
   *   data-age-district
   *
   * script.js 會負責接收點擊。
   */

  function buildBars(
    order,
    stats,
    total,
    districts
  ) {

    if (!total) {

      return `
        <div class="age-empty">
          目前沒有可分析資料
        </div>
      `;
    }


    const maxCount =
      Math.max(
        ...order.map(
          group =>
            stats[group]?.total ||
            0
        ),
        1
      );


    return order
      .map(
        group => {

          const item =
            stats[group] || {
              total: 0,
              districts: {}
            };


          const count =
            item.total;


          const overallPct =
            total
              ? count /
                total *
                100
              : 0;


          const barWidth =
            count
              ? count /
                maxCount *
                100
              : 0;


          const segments =
            districts
              .map(
                (
                  district,
                  index
                ) => {

                  const districtCount =
                    item
                      .districts[
                        district
                      ] || 0;


                  if (
                    !districtCount
                  ) {
                    return '';
                  }


                  const districtPct =
                    count
                      ? districtCount /
                        count *
                        100
                      : 0;


                  const label =
                    districtPct >= 10
                      ? `${districtCount}`
                      : '';


                  return `
                    <button
                      type="button"

                      class="
                        age-district-segment
                        age-filter-trigger
                        ${districtClass(index)}
                      "

                      data-age-group="${escapeAttr(group)}"

                      data-age-district="${escapeAttr(district)}"

                      style="
                        width:${districtPct}%;
                      "

                      title="${escapeAttr(
                        `查看 ${group}・${district}：${districtCount} 人`
                      )}"

                      aria-label="${escapeAttr(
                        `查看 ${group} ${district} ${districtCount} 人`
                      )}"
                    >

                      ${
                        label
                          ? `
                            <span>
                              ${districtCount}
                            </span>
                          `
                          : ''
                      }

                    </button>
                  `;
                }
              )
              .join('');


          return `
            <div class="age-row">


              <button
                type="button"

                class="
                  age-label
                  age-filter-trigger
                "

                data-age-group="${escapeAttr(group)}"

                title="${escapeAttr(
                  `查看全部${group}人員`
                )}"
              >
                ${escapeHtml(group)}
              </button>


              <div class="age-track">

                <div
                  class="age-stack"

                  style="
                    width:${barWidth}%;
                  "
                >
                  ${segments}
                </div>

              </div>


              <div class="age-number">

                <strong>
                  ${count}
                </strong>

                <span>
                  人
                </span>

                <small>
                  ${overallPct.toFixed(1)}%
                </small>

              </div>


            </div>
          `;
        }
      )
      .join('');
  }


  function buildDistrictLegend(
    districts,
    people
  ) {

    if (
      !districts.length
    ) {
      return '';
    }


    const total =
      people.length;


    return `
      <div class="district-legend">

        ${
          districts
            .map(
              (
                district,
                index
              ) => {

                const count =
                  people.filter(
                    person =>
                      person.district ===
                      district
                  ).length;


                const pct =
                  total
                    ? count /
                      total *
                      100
                    : 0;


                return `
                  <div class="district-legend-item">

                    <i
                      class="
                        district-dot
                        ${districtClass(index)}
                      "
                    ></i>

                    <span class="district-name">
                      ${escapeHtml(district)}
                    </span>

                    <strong>
                      ${count}
                    </strong>

                    <span class="district-unit">
                      人
                    </span>

                    <small>
                      ${pct.toFixed(1)}%
                    </small>

                  </div>
                `;
              }
            )
            .join('')
        }

      </div>
    `;
  }


  function renderAnalysis() {

    const people =
      getVisibleStablePeople();


    const districts =
      getDistrictOrder(
        people
      );


    const mainStats =
      createGroupStats(
        MAIN_GROUP_ORDER,
        districts
      );


    const studentStats =
      createGroupStats(
        STUDENT_GROUP_ORDER,
        districts
      );


    const unknown =
      new Map();


    people.forEach(
      person => {

        const originalGroup =
          cleanText(
            person.group
          );


        const district =
          person.district;


        const mainGroup =
          getMainGroup(
            originalGroup
          );


        if (
          mainGroup
        ) {

          mainStats[
            mainGroup
          ].total++;


          if (
            mainStats[
              mainGroup
            ].districts[
              district
            ] === undefined
          ) {

            mainStats[
              mainGroup
            ].districts[
              district
            ] = 0;
          }


          mainStats[
            mainGroup
          ].districts[
            district
          ]++;
        }


        if (
          STUDENT_GROUP_ORDER.includes(
            originalGroup
          )
        ) {

          studentStats[
            originalGroup
          ].total++;


          if (
            studentStats[
              originalGroup
            ].districts[
              district
            ] === undefined
          ) {

            studentStats[
              originalGroup
            ].districts[
              district
            ] = 0;
          }


          studentStats[
            originalGroup
          ].districts[
            district
          ]++;
        }


        if (
          originalGroup &&
          !KNOWN_GROUPS.has(
            originalGroup
          )
        ) {

          unknown.set(
            originalGroup,
            (
              unknown.get(
                originalGroup
              ) || 0
            ) + 1
          );
        }
      }
    );


    const mainTotal =
      MAIN_GROUP_ORDER.reduce(
        (sum, group) =>
          sum +
          mainStats[
            group
          ].total,
        0
      );


    const ageTotal =
      $('ageTotal');


    if (
      ageTotal
    ) {

      ageTotal.textContent =
        `${mainTotal} 人`;
    }


    const ageDistribution =
      $('ageDistribution');


    if (
      ageDistribution
    ) {

      const mainPeople =
        people.filter(
          person =>
            getMainGroup(
              person.group
            )
        );


      ageDistribution.innerHTML =
        buildBars(
          MAIN_GROUP_ORDER,
          mainStats,
          mainTotal,
          districts
        ) +

        buildDistrictLegend(
          districts,
          mainPeople
        );
    }


    /*
     * ========================================
     * 學生
     * ========================================
     */

    const studentTotal =
      STUDENT_GROUP_ORDER.reduce(
        (sum, group) =>
          sum +
          studentStats[
            group
          ].total,
        0
      );


    const studentAgeTotal =
      $('studentAgeTotal');


    if (
      studentAgeTotal
    ) {

      studentAgeTotal.textContent =
        `${studentTotal} 人`;
    }


    const studentAgeDistribution =
      $('studentAgeDistribution');


    if (
      studentAgeDistribution
    ) {

      const studentPeople =
        people.filter(
          person =>
            STUDENT_GROUP_ORDER.includes(
              person.group
            )
        );


      studentAgeDistribution.innerHTML =
        buildBars(
          STUDENT_GROUP_ORDER,
          studentStats,
          studentTotal,
          districts
        ) +

        buildDistrictLegend(
          districts,
          studentPeople
        );
    }


    const studentPct =
      mainTotal
        ? studentTotal /
          mainTotal *
          100
        : 0;


    const studentShare =
      $('studentShare');


    if (
      studentShare
    ) {

      studentShare.innerHTML = `
        <span>

          穩定聚會學生

          <strong>
            ${studentTotal}
          </strong>

          人

        </span>


        <span>

          占穩定聚會者

          <strong>
            ${studentPct.toFixed(1)}%
          </strong>

        </span>
      `;
    }


    /*
     * ========================================
     * 未知群組
     * ========================================
     */

    const warning =
      $('ageUnknown');


    if (
      !warning
    ) {
      return;
    }


    const unknownCount =
      [
        ...unknown.values()
      ].reduce(
        (
          sum,
          count
        ) =>
          sum + count,
        0
      );


    if (
      unknownCount
    ) {

      const details =
        [
          ...unknown.entries()
        ]
          .map(
            ([
              group,
              count
            ]) =>
              `${escapeHtml(group)} ${count}人`
          )
          .join('、');


      warning.innerHTML = `
        發現

        <strong>
          ${unknownCount}
        </strong>

        位穩定聚會者的群組名稱不在目前設定中
        （${details}）
      `;


      warning.classList.remove(
        'hidden'
      );


    } else {

      warning.textContent =
        '';


      warning.classList.add(
        'hidden'
      );
    }


    /*
     * 圖表每次重新產生 DOM 後，
     * 通知 script.js 重新標記目前選取狀態。
     */
    document.dispatchEvent(
      new CustomEvent(
        'ageAnalysisRendered'
      )
    );
  }


  function clearAnalysis() {

    [
      'ageTotal',
      'studentAgeTotal'
    ]
      .forEach(
        id => {

          const element =
            $(id);


          if (
            element
          ) {

            element.textContent =
              '—';
          }
        }
      );


    [
      'ageDistribution',
      'studentAgeDistribution',
      'studentShare'
    ]
      .forEach(
        id => {

          const element =
            $(id);


          if (
            element
          ) {

            element.innerHTML =
              '';
          }
        }
      );


    const warning =
      $('ageUnknown');


    if (
      warning
    ) {

      warning.textContent =
        '';


      warning.classList.add(
        'hidden'
      );
    }
  }


  function install() {

    const tbody =
      document.querySelector(
        '#peopleTable tbody'
      );


    if (
      !tbody
    ) {
      return;
    }


    const observer =
      new MutationObserver(
        () => {

          renderAnalysis();
        }
      );


    observer.observe(
      tbody,
      {
        childList: true,
        subtree: true
      }
    );


    const filterIds = [
      'searchInput',
      'districtFilter',
      'smallDistrictFilter',
      'totalAttendanceFilter',
      'absenceWeeksFilter',
      'newBelieverFilter'
    ];


    filterIds.forEach(
      id => {

        const element =
          $(id);


        if (
          !element
        ) {
          return;
        }


        element.addEventListener(
          'input',
          () => {

            requestAnimationFrame(
              renderAnalysis
            );
          }
        );


        element.addEventListener(
          'change',
          () => {

            requestAnimationFrame(
              renderAnalysis
            );
          }
        );
      }
    );


    const groupOptions =
      $('groupFilterOptions');


    if (
      groupOptions
    ) {

      groupOptions.addEventListener(
        'change',
        () => {

          requestAnimationFrame(
            renderAnalysis
          );
        }
      );
    }


    const statusOptions =
      $('statusFilterOptions');


    if (
      statusOptions
    ) {

      statusOptions.addEventListener(
        'change',
        () => {

          requestAnimationFrame(
            renderAnalysis
          );
        }
      );
    }


    document
      .querySelectorAll(
        '[data-group-preset]'
      )
      .forEach(
        button => {

          button.addEventListener(
            'click',
            () => {

              requestAnimationFrame(
                renderAnalysis
              );
            }
          );
        }
      );


    const clearBtn =
      $('clearBtn');


    if (
      clearBtn
    ) {

      clearBtn.addEventListener(
        'click',
        () => {

          setTimeout(
            clearAnalysis,
            0
          );
        }
      );
    }


    renderAnalysis();
  }


  if (
    document.readyState ===
    'loading'
  ) {

    document.addEventListener(
      'DOMContentLoaded',
      install
    );

  } else {

    install();
  }

})();