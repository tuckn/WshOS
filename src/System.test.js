/* globals Wsh: false */

/* globals describe: false */
/* globals test: false */
/* globals expect: false */

// Shorthand
var CD = Wsh.Constants;
var util = Wsh.Util;
var fso = Wsh.FileSystemObject;
var path = Wsh.Path;
var shApp = Wsh.ShellApplication;
var os = Wsh.OS;

var isString = util.isString;
var isSameMeaning = util.isSameMeaning;

var _cb = function (fn/* , args */) {
  var args = Array.from(arguments).slice(1);
  return function () { fn.apply(null, args); };
};

describe('System', function () {
  var network = Wsh.Network;
  if (!network) network = WScript.CreateObject('WScript.Network');

  var COMPNAME = String(network.ComputerName); // %COMPUTERNAME%
  var SYSROOT = String(fso.GetSpecialFolder(CD.folderSpecs.windows)); // %SystemRoot%
  var SYSDIR = String(fso.GetSpecialFolder(CD.folderSpecs.system)); // %SystemRoot%\Synstem32
  var USERTEMP = String(fso.GetSpecialFolder(CD.folderSpecs.temporary)); // %TEMP% (User)
  var USERPROFILE = String(shApp.Namespace(40).Self.Path); // %USERPROFILE%
  var USERNAME = String(network.UserName); // %USERNAME%

  test('Values', function () {
    expect(os.EOL).toBe('\r\n');
    expect(os.platform()).toBe('win32');

    expect(os.envVars.USERPROFILE).toBe(USERPROFILE);
    expect(os.envVars.Path).toBeDefined();
    expect(os.envVars.TEMP).toBe(USERTEMP);
    expect(os.envVars.TMP).toBe(USERTEMP);
    expect(os.envVars.USERNAME).toBe(USERNAME);

    expect(os.arch()).toBe('amd64');
    expect(os.is64arch()).toBe(true); // @TODO Test for 32 bit
    expect(os.tmpdir()).toBe(USERTEMP);
    expect(os.homedir()).toBe(USERPROFILE);
    expect(os.hostname()).toBe(COMPNAME);
    expect(os.type()).toBe('Windows_NT'); // %OS%

    var userInfo = os.userInfo();
    expect(userInfo.username).toBe(USERNAME);
    expect(userInfo.homedir).toBe(USERPROFILE);
  });

  test('makeTmpPath', function () {
    var tmpPath1 = os.makeTmpPath();
    expect(fso.FileExists(tmpPath1)).toBe(false);
    expect(fso.FolderExists(tmpPath1)).toBe(false);
    expect(path.dirname(tmpPath1)).toBe(os.tmpdir());

    var tmpPath2 = os.makeTmpPath();
    expect(fso.FileExists(tmpPath2)).toBe(false);
    expect(fso.FolderExists(tmpPath2)).toBe(false);
    expect(path.dirname(tmpPath2)).toBe(os.tmpdir());
    expect(tmpPath1 === tmpPath2).toBe(false);

    var tmpPath3 = os.makeTmpPath('cache-', '.tmp');
    expect(path.basename(tmpPath3).indexOf('cache-')).toBe(0);
    expect(fso.GetExtensionName(tmpPath3)).toBe('tmp');
  });

  test('writeTempText', function () {
    var tmpPath1 = os.writeTempText('Foo Bar');
    expect(fso.FileExists(tmpPath1)).toBe(true);

    var tmpTxtObj = fso.OpenTextFile(tmpPath1);
    var data = tmpTxtObj.ReadAll();
    expect(data).toBe('Foo Bar');

    // Cleans
    tmpTxtObj.Close();
    fso.DeleteFile(tmpPath1, CD.fso.force.yes);
  });

  test('@TODO', function () {
    expect(os.arch()).toBe('arm'); // @TODO Test some arch
    expect(os.is64arch()).toBe(false); // @TODO Test for 32 bit
  });

  // System Information

  test('sysInfo', function () {
    var sysInfo = os.sysInfo();
    expect(isSameMeaning(sysInfo.CSName, COMPNAME)).toBe(true);
    expect(isSameMeaning(sysInfo.WindowsDirectory, SYSROOT)).toBe(true);
    expect(isSameMeaning(sysInfo.SystemDirectory, SYSDIR)).toBe(true);
    // ...
    // console.dir(sysInfo);
  });

  test('SystemStatus', function () {
    /** @todo How do I test these? Should use a Virtual machine? */
    expect(os.freemem()).toBeGreaterThan(0);
    expect(os.totalmem()).toBeGreaterThan(0);
    expect(isString(os.cmdCodeset())).toBe(true);
    expect(isString(os.release())).toBe(true);
    expect(os.hasUAC()).toBe(true);
    expect(os.isUacDisable()).toBe(false);
  });

  // Event Log

  test('writeLogEvent', function () {
    expect(os.writeLogEvent.success('Success Log')).toBe(undefined);
    expect(os.writeLogEvent.error('Error Log')).toBe(undefined);
    expect(os.writeLogEvent.warn('Warning Log')).toBe(undefined);
    expect(os.writeLogEvent.info('Information Log')).toBe(undefined);
  });

  // Clipboard

  test('setStrToClipboardAO', function () {
    // os.setStrToClipboardAO('str');
    expect('@TODO').toBe('tested');
  });

  test('getStrFromClipboard', function () {
    // os.getStrFromClipboard('str');
    expect('@TODO').toBe('tested');
  });
});
