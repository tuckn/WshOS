/* globals Wsh: false */
/* globals __dirname: false */
/* globals __filename: false */

/* globals describe: false */
/* globals test: false */
/* globals expect: false */

// Shorthand
var util = Wsh.Util;
var fso = Wsh.FileSystemObject;
var path = Wsh.Path;
var os = Wsh.OS;

var isPureNumber = util.isPureNumber;

var _cb = function (fn/* , args */) {
  var args = Array.from(arguments).slice(1);
  return function () { fn.apply(null, args); };
};

describe('Handler', function () {
  var network = Wsh.Network;
  if (!network) network = WScript.CreateObject('WScript.Network');

  var wshFileName = path.basename(__filename);
  var mockWsfGUI = path.join(__dirname, 'assets', 'MockGUI.wsf');

  var noneStrVals = [true, false, undefined, null, 0, 1, NaN, Infinity, [], {}];

  // Drive

  test('DriveController', function () {
    var driveInfo = os.getDrivesInfo();

    driveInfo.forEach(function (drive) {
      expect(os.isExistingDrive(drive.letter)).toBe(true);
      // @TODO More tests
    });

    var hdriveInfo = os.getHardDriveLetters();
    hdriveInfo.forEach(function (drive) {
      expect(os.isExistingDrive(drive.letter)).toBe(true);
      // @TODO More tests
    });

    var newLetter = os.getTheLastFreeDriveLetter();
    expect(os.isExistingDrive(newLetter)).toBe(false); // Non existing

    os.assignDriveLetter(__dirname, newLetter); // Assign the drive letter
    expect(os.isExistingDrive(newLetter)).toBe(true); // Existing
    expect(fso.FileExists(newLetter + ':\\' + wshFileName)).toBe(true);

    os.removeAssignedDriveLetter(newLetter); // Remove the drive
    expect(os.isExistingDrive(newLetter)).toBe(false); // Non existing

    // Check TypeError
    noneStrVals.forEach(function (val) {
      expect(_cb(os.assignDriveLetter, val, val)).toThrowError();
      expect(_cb(os.removeAssignedDriveLetter, val)).toThrowError();
    });
  });

  // Process

  test('getProcessIDs', function () {
    var exeName = path.basename(os.exefiles.wscript);

    var oExecA = os.exec(os.exefiles.wscript, [mockWsfGUI]);
    var mockPIDA = oExecA.ProcessID;

    var oExecB = os.exec(os.exefiles.wscript, [mockWsfGUI, '--type', 'B']);
    var mockPIDB = oExecB.ProcessID;

    var oExecC = os.exec(os.exefiles.wscript, [mockWsfGUI, '--type', 'C']);
    var mockPIDC = oExecC.ProcessID;

    var pIDs = os.getProcessIDs(exeName);
    expect(pIDs.length >= 3).toBe(true);
    expect(util.includes(pIDs, mockPIDA)).toBe(true);
    expect(util.includes(pIDs, mockPIDB)).toBe(true);
    expect(util.includes(pIDs, mockPIDC)).toBe(true);

    oExecA.Terminate();
    oExecB.Terminate();
    oExecC.Terminate();

    var errVals = [true, false, undefined, null, NaN, Infinity, [], {}];
    errVals.forEach(function (val) {
      expect(_cb(os.getProcessIDs, val)).toThrowError();
    });
  });

  test('activateProcess', function () {
    expect('@TODO').toBe('tested');
  });

  test('terminateProcesses', function () {
    var exeName = path.basename(os.exefiles.wscript);

    var oExecA = os.exec(os.exefiles.wscript, [mockWsfGUI]);
    var mockPIDA = oExecA.ProcessID;

    var oExecB = os.exec(os.exefiles.wscript, [mockWsfGUI, '--type', 'B']);
    var mockPIDB = oExecB.ProcessID;

    var oExecC = os.exec(os.exefiles.wscript, [mockWsfGUI, '--type', 'C']);
    var mockPIDC = oExecC.ProcessID;

    expect(os.terminateProcesses(mockPIDA)).toBe(undefined);
    expect(os.terminateProcesses(mockPIDB)).toBe(undefined);
    expect(os.terminateProcesses(mockPIDC)).toBe(undefined);

    var pIDs = os.getProcessIDs(exeName);
    expect(util.includes(pIDs, mockPIDA)).toBe(false);
    expect(util.includes(pIDs, mockPIDB)).toBe(false);
    expect(util.includes(pIDs, mockPIDC)).toBe(false);

    var errVals = [true, false, undefined, null, NaN, Infinity, [], {}, ''];
    errVals.forEach(function (val) {
      expect(_cb(os.terminateProcesses, val)).toThrowError();
    });
  });

  test('getThisProcessID', function () {
    var pID = os.getThisProcessID();
    expect(isPureNumber(pID)).toBe(true);
  });

  test('getThisParentProcessID', function () {
    var pID = os.getThisParentProcessID();
    expect(isPureNumber(pID)).toBe(true);
  });

  // User

  test('UserController', function () {
    expect('@TODO').toBe('tested');
  });
});
