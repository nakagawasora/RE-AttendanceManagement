const test = require('node:test');
const assert = require('node:assert/strict');
const AttendanceUtil = require('../src/AttendanceUtil.js');

test('重複検索日は開始24時間前から終了日までに限定する', () => {
  assert.deepEqual(
    AttendanceUtil.getDuplicateSearchDatePrefixes(
      '2026-07-30T10:00',
      '2026-07-30T18:00'
    ),
    ['2026/07/29', '2026/07/30']
  );
  assert.deepEqual(
    AttendanceUtil.getDuplicateSearchDatePrefixes(
      '2026-07-30T23:00',
      '2026-07-31T02:00'
    ),
    ['2026/07/29', '2026/07/30', '2026/07/31']
  );
});

test('隣接する勤怠は重複せず、包含・部分一致は重複する', () => {
  assert.equal(
    AttendanceUtil.intervalsOverlap(
      '2026-07-30T10:00',
      '2026-07-30T12:00',
      '2026-07-30T08:00',
      '2026-07-30T10:00'
    ),
    false
  );
  assert.equal(
    AttendanceUtil.intervalsOverlap(
      '2026-07-30T10:00',
      '2026-07-30T12:00',
      '2026-07-30T11:00',
      '2026-07-30T13:00'
    ),
    true
  );
  assert.equal(
    AttendanceUtil.intervalsOverlap(
      '2026-07-30T10:00',
      '2026-07-30T12:00',
      '2026-07-30T09:00',
      '2026-07-30T13:00'
    ),
    true
  );
});

test('TextFinder候補行だけを読み込んで重複判定する', () => {
  const rows = {
    2: ['2026/07/01 09:00:00', 'PJ1', '', '2026/07/01 09:00:00', '2026/07/01 10:00:00'],
    3: ['2026/07/30 09:00:00', 'PJ1', '', '2026/07/30 09:00:00', '2026/07/30 11:00:00'],
    4: ['2026/07/30 12:00:00', 'PJ1', '', '2026/07/30 12:00:00', '2026/07/30 13:00:00']
  };
  const rowReads = [];
  let finderPattern = '';
  const sheet = {
    getLastRow() {
      return 4;
    },
    getRange(row, column, rowCount, columnCount) {
      if (row === 2 && column === 4 && rowCount === 3 && columnCount === 1) {
        return {
          createTextFinder(pattern) {
            finderPattern = pattern;
            const matchingRows = Object.keys(rows)
              .map(Number)
              .filter(rowNumber => new RegExp(pattern).test(rows[rowNumber][3]));
            return {
              useRegularExpression() {
                return this;
              },
              matchCase() {
                return this;
              },
              findAll() {
                return matchingRows.map(rowNumber => ({
                  getRow() {
                    return rowNumber;
                  }
                }));
              }
            };
          }
        };
      }
      assert.equal(rowCount, 1);
      assert.equal(column, 1);
      assert.equal(columnCount, 5);
      rowReads.push(row);
      return {
        getValues() {
          return [rows[row]];
        }
      };
    }
  };

  const messages = AttendanceUtil.findDuplicateAttendanceMessages(
    sheet,
    '2026-07-30T10:00',
    '2026-07-30T12:00'
  );
  assert.match(finderPattern, /2026\/07\/30/);
  assert.deepEqual(rowReads, [3, 4]);
  assert.equal(messages.length, 1);
  assert.match(messages[0], /2026\/07\/30 09:00:00/);
});

test('月跨ぎ勤怠を対象月の範囲だけ集計する', () => {
  const rows = [
    ['PJ1', '', '2026/06/30 23:00:00', '2026/07/01 02:00:00'],
    ['PJ1', '', '2026/07/15 10:00:00', '2026/07/15 12:30:00'],
    ['PJ2', '', '2026/08/01 09:00:00', '2026/08/01 10:00:00']
  ];
  assert.deepEqual(
    AttendanceUtil.calculateMonthlyProjectWorkTimesFromRows(rows, 7, 2026),
    {
      PJ1: 4.5 * 60 * 60 * 1000
    }
  );
});

test('フォーム入力の必須項目と24時間制限を検証する', () => {
  assert.deepEqual(
    AttendanceUtil.getFormValidationMessages({}),
    [
      'PJを選択してください',
      '作業内容を入力してください',
      '開始時刻、終了時刻を正しく入力してください'
    ]
  );
  assert.deepEqual(
    AttendanceUtil.getFormValidationMessages({
      project: 'PJ1',
      content: '作業',
      startTime: '2026-07-30T10:00',
      endTime: '2026-07-31T10:00'
    }),
    ['勤務時間が不正です']
  );
  assert.deepEqual(
    AttendanceUtil.getFormValidationMessages({
      project: 'PJ1',
      content: '作業',
      startTime: 'invalid',
      endTime: 'invalid'
    }),
    ['勤務時間が不正です']
  );
});

test('実績明細を時給増分・役職手当・税額込みで一度に計算する', () => {
  const summary = AttendanceUtil.buildMonthlyWorkTimeSummaryFromRows(
    1,
    [
      ['PJ1', '', '2026/07/01 09:00:00', '2026/07/01 11:00:00'],
      ['PJ2', '', '2026/07/02 09:00:00', '2026/07/02 10:00:00']
    ],
    [
      [10, 'PJ1', 1000, true, ''],
      [20, 'PJ2', 2000, true, '']
    ],
    [
      [1, 1, 10, 100],
      [2, 2, 20, 999]
    ],
    [
      [1, '役職手当', 1, 500, ''],
      [2, '削除済み', 1, 500, '2026/01/01']
    ],
    7,
    2026,
    10
  );

  assert.deepEqual(summary.workTimeByProject, {
    PJ1: '2時間0分',
    PJ2: '1時間0分'
  });
  assert.deepEqual(
    summary.lineItems.map(item => ({
      type: item.type,
      name: item.name,
      unitPrice: item.unitPrice,
      amount: item.amount,
      amountWithTax: item.amountWithTax
    })),
    [
      { type: 'project', name: 'PJ1', unitPrice: 1100, amount: 2200, amountWithTax: 2420 },
      { type: 'project', name: 'PJ2', unitPrice: 2000, amount: 2000, amountWithTax: 2200 },
      { type: 'allowance', name: '役職手当', unitPrice: 500, amount: 1500, amountWithTax: 1650 }
    ]
  );
  assert.equal(summary.totalAmount, 6270);
});
