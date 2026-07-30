const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const AttendanceUtil = require('../src/AttendanceUtil.js');

function loadManagementContext() {
  const source = fs.readFileSync(
    path.join(__dirname, '..', 'src', 'ManagementUtil.js'),
    'utf8'
  );
  const context = {
    AttendanceUtil,
    BillUtil: { databaseSheetId: 'database' },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty() {
            return '';
          }
        };
      }
    },
    console
  };
  vm.createContext(context);
  vm.runInContext(source, context);
  return context;
}

test('勤怠RPCは認証・入力・PJエラーをプレーンメッセージで返す', () => {
  const context = loadManagementContext();
  context.authenticateAttendanceRequest = () => ({
    messages: ['アカウント情報が不正です'],
    user: null
  });
  let result = context.performAttendanceFormAction({
    action: 'submit'
  });
  assert.equal(result.status, 'error');
  assert.deepEqual(Array.from(result.messages), [
    'アカウント情報が不正です',
    'PJを選択してください',
    '作業内容を入力してください',
    '開始時刻、終了時刻を正しく入力してください'
  ]);

  context.authenticateAttendanceRequest = () => ({
    messages: [],
    user: { values: [1] }
  });
  context.getProjectAccessValidationMessage = () => '選択されたPJは現在利用できません';
  result = context.performAttendanceFormAction({
    action: 'submit',
    project: 'PJ1',
    content: '作業',
    startTime: '2026-07-30T10:00',
    endTime: '2026-07-30T11:00'
  });
  assert.equal(result.status, 'error');
  assert.deepEqual(Array.from(result.messages), ['選択されたPJは現在利用できません']);
});

test('保存と呼び出しのRPCはパスワードを返さない', () => {
  const context = loadManagementContext();
  let savedValue = '';
  const auth = {
    messages: [],
    user: {
      rowNumber: 2,
      values: [1, 'attendance', '', '', '', 'user@example.com', 'hash', '{"parameter":{"project":"PJ1","content":"保存済み","startTime":"2026-07-30T10:00","endTime":"2026-07-30T11:00"},"formatVersion":"secure_v1"}']
    },
    userSheet: {
      getRange() {
        return {
          setValue(value) {
            savedValue = value;
          }
        };
      }
    }
  };
  context.authenticateAttendanceRequest = () => auth;

  const saveResult = context.performAttendanceFormAction({
    action: 'save',
    email: 'user@example.com',
    password: 'secret',
    project: 'PJ1',
    content: '作業',
    startTime: '2026-07-30T10:00',
    endTime: '2026-07-30T11:00'
  });
  assert.equal(saveResult.status, 'success');
  assert.ok(savedValue);

  const loadResult = context.performAttendanceFormAction({
    action: 'load',
    email: 'user@example.com',
    password: 'secret'
  });
  assert.equal(loadResult.status, 'success');
  assert.equal(loadResult.fields.project, 'PJ1');
  assert.equal(Object.hasOwn(loadResult.fields, 'password'), false);
  assert.equal(JSON.stringify(loadResult).includes('secret'), false);
});

test('勤怠提出はScript Lockを使わず重複確認後に追記する', () => {
  const context = loadManagementContext();
  const events = [];
  const sheet = {
    getLastRow() {
      return 1;
    },
    getRange() {
      return {
        setValues() {
          events.push('append');
        }
      };
    }
  };
  context.AttendanceUtil = {
    ...AttendanceUtil,
    findDuplicateAttendanceMessages() {
      events.push('duplicate-check');
      return [];
    }
  };
  context.LockService = {
    getScriptLock() {
      throw new Error('勤怠提出ではScript Lockを使用しない');
    }
  };
  context.SpreadsheetApp = {
    openById() {
      return {
        getSheets() {
          return [sheet];
        }
      };
    }
  };
  context.formatTimeToGoogleForm = value => String(value);

  const values = Array(10).fill('');
  values[1] = 'attendance-sheet';
  const result = context.submitAttendanceRequest(
    {
      project: 'PJ1',
      content: '作業',
      startTime: '2026-07-30T10:00',
      endTime: '2026-07-30T11:00'
    },
    { user: { values } }
  );
  assert.equal(result.status, 'success');
  assert.deepEqual(events, ['duplicate-check', 'append']);
});
