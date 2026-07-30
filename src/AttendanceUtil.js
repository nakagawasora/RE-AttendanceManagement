var AttendanceUtil = (function() {
  function formatDatePrefix(date) {
    return `${date.getFullYear()}/${String(date.getMonth() + 1).padStart(2, '0')}/${String(date.getDate()).padStart(2, '0')}`;
  }

  function getDuplicateSearchDatePrefixes(startTime, endTime) {
    const start = new Date(startTime);
    const end = new Date(endTime);
    const firstDay = new Date(start.getTime() - 86400000);
    firstDay.setHours(0, 0, 0, 0);
    const lastDay = new Date(end);
    lastDay.setHours(0, 0, 0, 0);
    const prefixes = [];
    for (let day = new Date(firstDay); day <= lastDay; day.setDate(day.getDate() + 1)) {
      prefixes.push(formatDatePrefix(day));
    }
    return prefixes;
  }

  function escapeRegularExpression(text) {
    return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function buildDuplicateSearchPattern(startTime, endTime) {
    const prefixes = getDuplicateSearchDatePrefixes(startTime, endTime);
    return `^(${prefixes.map(escapeRegularExpression).join('|')}) `;
  }

  function intervalsOverlap(startTime, endTime, previousStartTime, previousEndTime) {
    const requestedStart = new Date(startTime);
    const requestedEnd = new Date(endTime);
    const previousStart = new Date(previousStartTime);
    const previousEnd = new Date(previousEndTime);
    return !(requestedStart >= previousEnd || requestedEnd <= previousStart);
  }

  function getFormValidationMessages(parameter) {
    const messages = [];
    if (!parameter.project) {
      messages.push('PJを選択してください');
    }
    if (!parameter.content) {
      messages.push('作業内容を入力してください');
    }
    if (!parameter.startTime || !parameter.endTime) {
      messages.push('開始時刻、終了時刻を正しく入力してください');
    } else {
      const duration = new Date(parameter.endTime) - new Date(parameter.startTime);
      if (!isFinite(duration) || duration <= 0 || duration >= 86400000) {
        messages.push('勤務時間が不正です');
      }
    }
    return messages;
  }

  function findDuplicateAttendanceMessages(attendanceSheet, startTime, endTime) {
    if (attendanceSheet.getLastRow() < 2) {
      return [];
    }
    const matches = attendanceSheet
      .getRange(2, 4, attendanceSheet.getLastRow() - 1, 1)
      .createTextFinder(buildDuplicateSearchPattern(startTime, endTime))
      .useRegularExpression(true)
      .matchCase(true)
      .findAll();
    const messages = [];
    const checkedRows = {};
    matches.forEach(match => {
      const rowNumber = match.getRow();
      if (checkedRows[rowNumber]) {
        return;
      }
      checkedRows[rowNumber] = true;
      const row = attendanceSheet.getRange(rowNumber, 1, 1, 5).getValues()[0];
      if (intervalsOverlap(startTime, endTime, row[3], row[4])) {
        messages.push(`入力された勤怠はタイムスタンプ:${row[0]}に提出された勤怠と重複する時間があります。開始時刻と終了時刻を再度確認してください。`);
      }
    });
    return messages;
  }

  function calculateMonthlyProjectWorkTimesFromRows(attendanceRows, month, year) {
    const monthStart = new Date(year, month - 1, 1);
    const monthEnd = new Date(year, month, 1);
    const projectWorkTimes = {};
    (attendanceRows || []).forEach(row => {
      const projectName = row[0];
      const startTime = new Date(row[2]);
      const endTime = new Date(row[3]);
      if (!projectName || isNaN(startTime.getTime()) || isNaN(endTime.getTime())) {
        return;
      }
      const overlapStart = Math.max(startTime.getTime(), monthStart.getTime());
      const overlapEnd = Math.min(endTime.getTime(), monthEnd.getTime());
      if (overlapEnd <= overlapStart) {
        return;
      }
      projectWorkTimes[projectName] =
        Number(projectWorkTimes[projectName] || 0) + (overlapEnd - overlapStart);
    });
    return projectWorkTimes;
  }

  function msecToHourAndMinutes(msec) {
    return `${Math.trunc(msec / 3600000)}時間${Math.trunc((msec % 3600000) / 60000)}分`;
  }

  function buildMonthlyWorkTimeSummaryFromRows(
    userId,
    attendanceRows,
    projectRows,
    upgradeRows,
    allowanceRows,
    month,
    year,
    taxPercentage
  ) {
    const projectWorkTimes = calculateMonthlyProjectWorkTimesFromRows(attendanceRows, month, year);
    const workTimeByProject = {};
    Object.keys(projectWorkTimes).forEach(projectName => {
      const workTime = Number(projectWorkTimes[projectName] || 0);
      if (workTime > 0) {
        workTimeByProject[projectName] = msecToHourAndMinutes(workTime);
      }
    });

    const upgradeByProjectId = {};
    (upgradeRows || []).forEach(row => {
      if (String(row[1]) === String(userId)) {
        upgradeByProjectId[String(row[2])] = Number(row[3] || 0);
      }
    });
    const projectItems = (projectRows || []).map(row => {
      const projectId = row[0];
      const projectName = row[1];
      const workTimeMsec = Number(projectWorkTimes[projectName] || 0);
      const unitPrice = Number(row[2] || 0) +
        Number(upgradeByProjectId[String(projectId)] || 0);
      return {
        type: 'project',
        id: projectId,
        name: projectName,
        unitPrice: unitPrice,
        workTimeMsec: workTimeMsec,
        workTimeText: msecToHourAndMinutes(workTimeMsec),
        amount: Math.round(unitPrice * workTimeMsec / 3600000)
      };
    }).filter(item => item.workTimeMsec > 0);

    const totalWorkTime = Object.keys(projectWorkTimes)
      .reduce((total, key) => total + Number(projectWorkTimes[key] || 0), 0);
    const roleAllowanceItems = (allowanceRows || [])
      .filter(row => String(row[2]) === String(userId))
      .filter(row => !row[4])
      .map(row => {
        const price = Number(row[3] || 0);
        return {
          type: 'allowance',
          id: row[0],
          name: row[1],
          unitPrice: price,
          price: price,
          workTimeMsec: totalWorkTime,
          workTimeText: msecToHourAndMinutes(totalWorkTime),
          amount: Math.round(price * totalWorkTime / 3600000)
        };
      })
      .filter(item => item.name && item.unitPrice !== 0 && item.workTimeMsec > 0);

    let totalAmount = 0;
    const taxRate = 1 + (Number(taxPercentage || 0) / 100);
    const lineItems = projectItems.concat(roleAllowanceItems).map(item => {
      const amountWithTax = Math.round(item.amount * taxRate);
      totalAmount += amountWithTax;
      return {
        type: item.type,
        id: item.id,
        name: item.name,
        unitPrice: item.unitPrice,
        workTime: item.workTimeText,
        amount: item.amount,
        amountWithTax: amountWithTax
      };
    });
    return {
      workTimeByProject: workTimeByProject,
      lineItems: lineItems,
      roleAllowances: roleAllowanceItems.map(item => ({
        name: item.name,
        price: item.price,
        workTime: item.workTimeText,
        amount: item.amount,
        amountWithTax: Math.round(item.amount * taxRate)
      })),
      totalAmount: totalAmount
    };
  }

  return {
    buildDuplicateSearchPattern: buildDuplicateSearchPattern,
    buildMonthlyWorkTimeSummaryFromRows: buildMonthlyWorkTimeSummaryFromRows,
    calculateMonthlyProjectWorkTimesFromRows: calculateMonthlyProjectWorkTimesFromRows,
    findDuplicateAttendanceMessages: findDuplicateAttendanceMessages,
    getFormValidationMessages: getFormValidationMessages,
    getDuplicateSearchDatePrefixes: getDuplicateSearchDatePrefixes,
    intervalsOverlap: intervalsOverlap
  };
})();

if (typeof module !== 'undefined' && module.exports) {
  module.exports = AttendanceUtil;
}
