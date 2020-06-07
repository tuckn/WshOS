/* globals Wsh: false */
/* globals __dirname: false */

/* globals describe: false */
/* globals test: false */
/* globals expect: false */

// Shorthand
var sh = Wsh.Shell;
var CD = Wsh.Constants;
var util = Wsh.Util;
var fso = Wsh.FileSystemObject;
var path = Wsh.Path;
var os = Wsh.OS;

var isNumber = util.isNumber;
var isString = util.isString;
var isSolidArray = util.isSolidArray;
var isSolidString = util.isSolidString;
var isSameMeaning = util.isSameMeaning;

var SYSDIR = String(fso.GetSpecialFolder(CD.folderSpecs.system));
var WSCRIPT = path.join(SYSDIR, 'wscript.exe');

var _cb = function (fn/* , args */) {
  var args = Array.from(arguments).slice(1);
  return function () { fn.apply(null, args); };
};

describe('WMI', function () {
  var network = Wsh.Network;
  if (!network) network = WScript.CreateObject('WScript.Network');

  var mockWsfGUI = path.join(__dirname, 'assets', 'MockGUI.wsf');

  test('execQuery', function () {
    var query, sWbemObjSets;

    var oExec = sh.Exec(WSCRIPT + ' "' + mockWsfGUI + '"');
    var mockPID = oExec.ProcessID;

    query = 'SELECT * FROM Win32_Process WHERE Caption = "' + path.basename(WSCRIPT) + '"';
    sWbemObjSets = os.WMI.execQuery(query);

    expect(sWbemObjSets.length > 0).toBe(true);

    var mockWmi = sWbemObjSets.find(function (sWbemObjSet) {
      return sWbemObjSet.ProcessId === mockPID;
    });

    expect(mockPID).toBe(mockWmi.ProcessId);
    expect(isSameMeaning(WSCRIPT, mockWmi.ExecutablePath)).toBe(true);

    // Can use methods
    var iRetVal = mockWmi.Terminate();
    expect(isNumber(iRetVal)).toBe(true);

    sWbemObjSets = os.WMI.execQuery(query);
    expect(sWbemObjSets.length === 0).toBe(true);

    // Array property
    query = 'SELECT * FROM CIM_BIOSElement';
    sWbemObjSets = os.WMI.execQuery(query);

    expect(sWbemObjSets.length > 0).toBe(true);
    var wmi0 = sWbemObjSets[0];
    // console.dir(sWbemObjSets[0]); // Empty {}
    expect(isString(wmi0.Name)).toBe(true);
    // Not Js Array
    expect(isSolidArray(wmi0.BiosCharacteristics)).toBe(false);
    // Require converting to read the array property in SWbemObjectSet
    expect(isSolidArray(new VBArray(wmi0.BiosCharacteristics).toArray())).toBe(true);
  });

  test('toJsObjects', function () {
    var query = 'SELECT * FROM CIM_BIOSElement';
    var sWbemObjSets = os.WMI.execQuery(query);
    var wmiObjs = os.WMI.toJsObjects(sWbemObjSets);

    expect(wmiObjs.length > 0).toBe(true);
    // console.dir(wmiObjs[0]);
    expect(isString(wmiObjs[0].Name)).toBe(true);
    expect(isString(wmiObjs[0].Manufacturer)).toBe(true);
    expect(isSolidArray(wmiObjs[0].BiosCharacteristics)).toBe(true);
    expect(isSolidArray(wmiObjs[0].ListOfLanguages)).toBe(true);
  });

  test('getProcesses', function () {
    var sWbemObjSets, filterdSWbemObjSets;
    var exeName = path.basename(WSCRIPT);

    var oExecA = sh.Exec(WSCRIPT + ' "' + mockWsfGUI + '"');
    var mockPIDA = oExecA.ProcessID;

    var oExecB = sh.Exec(WSCRIPT + ' "' + mockWsfGUI + '" --type B');
    var mockPIDB = oExecB.ProcessID;

    var oExecC = sh.Exec(WSCRIPT + ' "' + mockWsfGUI + '" --type C');
    var mockPIDC = oExecC.ProcessID;

    // Ex.1 Specifing a process name
    sWbemObjSets = os.WMI.getProcesses(exeName);
    expect(sWbemObjSets.length >= 3).toBe(true);

    filterdSWbemObjSets = sWbemObjSets.filter(function (sWbemObjSet) {
      return (mockPIDA === sWbemObjSet.ProcessId
        || mockPIDB === sWbemObjSet.ProcessId
        || mockPIDC === sWbemObjSet.ProcessId);
    });
    expect(filterdSWbemObjSets.length).toBe(3);

    // Ex.2 Specifing a process ID
    sWbemObjSets = os.WMI.getProcesses(mockPIDA);
    expect(sWbemObjSets.length === 1).toBe(true);

    // Ex.3 Specifing a full path
    sWbemObjSets = os.WMI.getProcesses(WSCRIPT);
    expect(sWbemObjSets.length >= 3).toBe(true);

    // Ex.4 Specifing options
    sWbemObjSets = os.WMI.getProcesses(exeName, {
      matchWords: ['--type B']
    });
    expect(sWbemObjSets.length === 1).toBe(true);

    sWbemObjSets = os.WMI.getProcesses(exeName, {
      matchWords: ['--type Z']
    });
    expect(sWbemObjSets.length === 0).toBe(true);

    sWbemObjSets = os.WMI.getProcesses(exeName, {
      matchWords: [mockWsfGUI],
      excludingWords: ['--type B', '--type C']
    });
    expect(sWbemObjSets.length === 1).toBe(true);

    // Close the mocks
    sWbemObjSets = os.WMI.getProcesses(exeName, {
      matchWords: [mockWsfGUI]
    });
    expect(sWbemObjSets.length === 3).toBe(true);

    sWbemObjSets.forEach(function (sWbemObjSet) { sWbemObjSet.Terminate(); });

    var errVals = [true, false, undefined, null, NaN, Infinity, [], {}];
    errVals.forEach(function (val) {
      expect(_cb(os.WMI.getProcesses, val)).toThrowError();
    });
  });

  test('getProcess', function () {
    var sWbemObjSet;
    var exeName = path.basename(WSCRIPT);

    var oExecA = sh.Exec(WSCRIPT + ' "' + mockWsfGUI + '"');
    var oExecB = sh.Exec(WSCRIPT + ' "' + mockWsfGUI + '" --type B');
    var oExecC = sh.Exec(WSCRIPT + ' "' + mockWsfGUI + '" --type C');

    sWbemObjSet = os.WMI.getProcess(exeName, {
      matchWords: ['--type B']
    });
    expect(sWbemObjSet.ProcessId).toBe(oExecB.ProcessID);
    expect(sWbemObjSet.Terminate()).toBe(0);
    oExecA.Terminate();
    oExecC.Terminate();

    var errVals = [true, false, undefined, null, NaN, Infinity, [], {}];
    errVals.forEach(function (val) {
      expect(_cb(os.WMI.getProcess, val)).toThrowError();
    });
  });

  test('getWithSWbemPath', function () {
    var sWbemPath = 'Win32_LogicalDisk.DeviceID="C:"';
    var sWbemObjSet = os.WMI.getWithSWbemPath(sWbemPath);

    expect(sWbemObjSet.DeviceID).toBe('C:');
    expect(sWbemObjSet.Caption).toBe('C:');
    expect(isSolidString(sWbemObjSet.FileSystem)).toBe(true);
    expect(isSolidString(sWbemObjSet.FreeSpace)).toBe(true);

    var errVals = [true, false, undefined, null, NaN, Infinity, [], {}, 0, 1, ''];
    errVals.forEach(function (val) {
      expect(_cb(os.WMI.getWithSWbemPath, val)).toThrowError();
    });

    sWbemPath = 'Win32_Process.Caption="chrome.exe"';
    // @TODO Fix
    sWbemObjSet = os.WMI.getWithSWbemPath(sWbemPath);
  });

  test('getWindowsUserAccounts', function () {
    expect('@TODO').toBe('tested');
    // var sWbemObjSets = os.WMI.getWindowsUserAccounts();

    /** @todo */
    // sWbemObjSets.forEach(function (sWbemObjSet) {
    //   console.log('AccountType: ' + sWbemObjSet.AccountType);
    //   console.log('Caption: ' + sWbemObjSet.Caption);
    //   console.log('Name: ' + sWbemObjSet.Name);
    //   console.log('Domain: ' + sWbemObjSet.Domain);
    //   console.log('Status: ' + sWbemObjSet.Status);
    // });
  });

  test('getThisProcess', function () {
    expect('@TODO').toBe('tested');
    // var thisProcess = os.WMI.getThisProcess();

    /** @todo */
    // console.log('ProcessID: ' + thisProcess.ProcessId);
    // console.log('Caption: ' + thisProcess.Caption);
    // console.log('Name: ' + thisProcess.Name);
    // console.log('CommandLine: ' + thisProcess.CommandLine);
    // console.log('ParentProcessId: ' + thisProcess.ParentProcessId);
  });
});
