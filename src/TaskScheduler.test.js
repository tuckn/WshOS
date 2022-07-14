/* globals Wsh: false */
/* globals __dirname: false */

/* globals describe: false */
/* globals test: false */
/* globals expect: false */

// Shorthand
var util = Wsh.Util;
var path = Wsh.Path;
var os = Wsh.OS;

var WSCRIPT = os.exefiles.wscript;
var SCHTASKS = os.exefiles.schtasks;

var _cb = function (fn/* , args */) {
  var args = Array.from(arguments).slice(1);
  return function () { fn.apply(null, args); };
};

var _getProcessIDs = function (processName, options) {
  var sWbemObjSets = os.WMI.getProcesses(processName, options);

  return sWbemObjSets.map(function (sWbemObjSet) {
    return sWbemObjSet.ProcessId;
  });
};

describe('TaskScheduler', function () {
  var network = Wsh.Network;
  if (!network) network = WScript.CreateObject('WScript.Network');

  var mockWsfGUI = path.join(__dirname, 'assets', 'MockGUI.wsf');

  var noneStrVals = [true, false, undefined, null, 0, 1, NaN, Infinity, [], {}];

  test('check_SchTask_commands', function () {
    var taskName = 'TestScheduler_' + util.createDateString();
    var args = ['//job:autoQuit0', mockWsfGUI];
    var retVal;

    // Creating
    retVal = os.Task.create(taskName, WSCRIPT, args, { isDryRun: true });
    expect(retVal).toContain(SCHTASKS + ' /Create /F /TN "' + taskName + '"' + ' /SC ONCE /ST 23:59 /IT /RL LIMITED /TR "');

    // Creating to run with highest
    retVal = os.Task.create(taskName, WSCRIPT, args, {
      runsWithHighest: true,
      isDryRun: true
    });
    expect(retVal).toContain(SCHTASKS + ' /Create /F /TN "' + taskName + '"' + ' /SC ONCE /ST 23:59 /IT /RL HIGHEST /TR "');

    // Checking
    retVal = os.Task.exists(taskName, { isDryRun: true });
    expect(retVal).toContain(SCHTASKS + ' /Query /XML /TN "' + taskName + '"');

    // Getting the task XML
    retVal = os.Task.xmlString(taskName, { isDryRun: true });
    expect(retVal).toContain(SCHTASKS + ' /Query /XML /TN "' + taskName + '"');

    // Running
    retVal = os.Task.run(taskName, { isDryRun: true });
    expect(retVal).toContain(SCHTASKS + ' /Run /I /TN "' + taskName + '"');

    // Removing
    retVal = os.Task.del(taskName, { isDryRun: true });
    expect(retVal).toContain(SCHTASKS + ' /Delete /F /TN "' + taskName + '"');
  });

  test('check_fromCreatingToRemoving', function () {
    var taskName = 'TestScheduler_' + util.createDateString();
    var args = ['//job:autoQuit1', mockWsfGUI];
    var retVal;

    // Checking throwing error
    noneStrVals.forEach(function (val) {
      expect(_cb(os.Task.create, val)).toThrowError();
      expect(_cb(os.Task.exists, val)).toThrowError();
      expect(_cb(os.Task.xmlString, val)).toThrowError();
      expect(_cb(os.Task.run, val)).toThrowError();
      expect(_cb(os.Task.del, val)).toThrowError();
    });

    // Creating dryRun
    os.Task.create(taskName, WSCRIPT, args, { isDryRun: true });
    expect(os.Task.exists(taskName)).toBe(false); // Not created
    // Creating
    retVal = os.Task.create(taskName, WSCRIPT, args, {});
    expect(retVal).toBe(undefined);
    expect(os.Task.exists(taskName)).toBe(true); // Created

    // Getting the task XML
    retVal = os.Task.xmlString(taskName);
    expect(retVal).toContain('<?xml version="1.0" encoding="UTF-16"?>');
    expect(retVal).toContain('<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">');
    expect(retVal).toContain('<URI>\\' + taskName + '</URI>');
    expect(retVal).not.toContain('<RunLevel>HighestAvailable</RunLevel>');

    // Running
    retVal = os.Task.run(taskName);
    expect(retVal).toBe(undefined);

    /**
     * run後、すぐに_getProcessIDsするとプロセスを取得できないときがある
     * var pIDs = _getProcessIDs(WSCRIPT, mockWsfGUI);
     * expect(pIDs.length > 0).toBe(true);
     */

    // Removing dryRun
    os.Task.del(taskName, { isDryRun: true });
    expect(os.Task.exists(taskName)).toBe(true); // Not removed
    // Removing
    retVal = os.Task.del(taskName);
    expect(retVal).toBe(undefined);
    expect(os.Task.exists(taskName)).toBe(false); // Removed
  });

  test('check_fromCreatingToRemoving_Admin', function () {
    var taskName = 'TestScheduler_' + util.createDateString();
    var args = ['//job:autoQuit1', mockWsfGUI];
    var retVal;

    // Checking throwing error
    noneStrVals.forEach(function (val) {
      expect(_cb(os.Task.create, val)).toThrowError();
      expect(_cb(os.Task.exists, val)).toThrowError();
      expect(_cb(os.Task.xmlString, val)).toThrowError();
      expect(_cb(os.Task.run, val)).toThrowError();
      expect(_cb(os.Task.del, val)).toThrowError();
    });

    // Creating dryRun
    os.Task.create(taskName, WSCRIPT, args, {
      runsWithHighest: true,
      isDryRun: true
    });
    expect(os.Task.exists(taskName)).toBe(false); // Not created
    // Creating
    retVal = os.Task.create(taskName, WSCRIPT, args, {
      runsWithHighest: true
    });
    expect(retVal).toBe(undefined);
    expect(os.Task.exists(taskName)).toBe(true); // Created

    // Getting the task XML
    retVal = os.Task.xmlString(taskName);
    expect(retVal).toContain('<?xml version="1.0" encoding="UTF-16"?>');
    expect(retVal).toContain('<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">');
    expect(retVal).toContain('<URI>\\' + taskName + '</URI>');
    expect(retVal).toContain('<RunLevel>HighestAvailable</RunLevel>');

    // Running
    retVal = os.Task.run(taskName);
    expect(retVal).toBe(undefined);

    /**
     * run後、すぐに_getProcessIDsするとプロセスを取得できないときがある
     * var pIDs = _getProcessIDs(WSCRIPT, mockWsfGUI);
     * expect(pIDs.length > 0).toBe(true);
     */

    // Removing dryRun
    os.Task.del(taskName, { isDryRun: true });
    expect(os.Task.exists(taskName)).toBe(true); // Not removed
    // Removing
    retVal = os.Task.del(taskName);
    expect(retVal).toBe(undefined);
    expect(os.Task.exists(taskName)).toBe(false); // Removed
  });

  test('runTemporary', function () {
    var uniqueStr = os.makeTmpPath();
    var args = ['//job:autoQuit0', mockWsfGUI, uniqueStr];

    var pIDs;
    pIDs = _getProcessIDs(WSCRIPT, { matchWords: [uniqueStr] });
    expect(pIDs).toHaveLength(0);

    os.Task.runTemporary(WSCRIPT, args, { winStyle: 'activeMin' });

    var msecTimeOut = 30000;

    do {
      pIDs = _getProcessIDs(WSCRIPT, { matchWords: [uniqueStr] });
      WScript.Sleep(300);
      msecTimeOut -= 300;
    } while (pIDs.length === 0 && msecTimeOut > 0);

    expect(msecTimeOut).toBeGreaterThan(0);

    noneStrVals.forEach(function (val) {
      expect(_cb(os.Task.ensureToDelete, val)).toThrowError();
      expect(_cb(os.Task.ensureToCreate, val)).toThrowError();
      expect(_cb(os.Task.runTemporary, val)).toThrowError();
    });
  });

  test('runTemporary_fromAdminToMidiumWIL', function () {
    expect('@TODO').toBe('tested');
  });
});
