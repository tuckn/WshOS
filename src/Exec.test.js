/* globals Wsh: false */
/* globals __dirname: false */

/* globals describe: false */
/* globals test: false */
/* globals expect: false */

// Shorthand
var CD = Wsh.Constants;
var util = Wsh.Util;
var fso = Wsh.FileSystemObject;
var path = Wsh.Path;
var os = Wsh.OS;

var objAdd = Object.assign;
var isEmpty = util.isEmpty;
var isNumber = util.isNumber;
var isPlainObject = util.isPlainObject;
var CMD = os.exefiles.cmd;
var CSCRIPT = os.exefiles.cscript;
var WSCRIPT = os.exefiles.wscript;
var NET = os.exefiles.net;
var PING = os.exefiles.ping;
var NOTEPAD = os.exefiles.notepad;

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

var _terminateProcesses = function (processName, options) {
  var sWbemObjSets = os.WMI.getProcesses(processName, options);
  if (sWbemObjSets.length === 0) return;

  sWbemObjSets.forEach(function (sWbemObjSet) { sWbemObjSet.Terminate(); });
};

describe('Exec', function () {
  var network = Wsh.Network;
  if (!network) network = WScript.CreateObject('WScript.Network');

  var mockWsfCLI = path.join(__dirname, 'assets', 'MockCLI.wsf');
  var mockWsfGUI = path.join(__dirname, 'assets', 'MockGUI.wsf');

  var noneStrVals = [true, false, undefined, null, 0, 1, NaN, Infinity, [], {}];

  test('surroundCmdArg', function () {
    var surroundCmdArg = os.surroundCmdArg;

    noneStrVals.forEach(function (val) {
      expect(_cb(surroundCmdArg, val)).toThrowError();
    });

    var hasSpacePath = 'C:\\Program Files';
    expect(surroundCmdArg(hasSpacePath)).toBe('"' + hasSpacePath + '"');

    var noSpacePath = 'C:\\Windows\\System32\\notepad.exe';
    expect(surroundCmdArg(noSpacePath)).toBe(noSpacePath);

    var wQuotationPath = '"C:\\Program Files (x86)\\Windows NT"';
    expect(surroundCmdArg(wQuotationPath)).toBe(wQuotationPath);

    var jpPath = 'D:\\2011-01-01家族で初詣';
    expect(surroundCmdArg(jpPath)).toBe('"' + jpPath + '"');

    var smbPath = '\\\\MyNAS\\Multimedia';
    expect(surroundCmdArg(smbPath)).toBe(smbPath);

    var hasSpaceSmbPath = '"\\\\192.168.12.34\\Public\\最新資料_12.xlsx"';
    expect(surroundCmdArg(hasSpaceSmbPath)).toBe(hasSpaceSmbPath);

    // Non path string
    expect(surroundCmdArg('abcd1234')).toBe('abcd1234');
    expect(surroundCmdArg('abcd 1234')).toBe('"abcd 1234"');
    expect(surroundCmdArg('あいうえお')).toBe('"あいうえお"');

    // A command control character
    expect(surroundCmdArg('|')).toBe('|');
    expect(surroundCmdArg('>')).toBe('>');
    expect(surroundCmdArg('1> "C:\\logs.txt"')).toBe('"1> "C:\\logs.txt""');
    // Inner quoted
    expect(surroundCmdArg('-p"My p@ss wo^d"')).toBe('-p"My p@ss wo^d"');
    expect(surroundCmdArg('1>"C:\\logs.txt"')).toBe('1>"C:\\logs.txt"');
  });

  test('escapeForCmd', function () {
    var escapeForCmd = os.escapeForCmd;

    var errVals = [true, false, undefined, null, [1], { a: 'A' }];
    errVals.forEach(function (val) {
      expect(_cb(escapeForCmd, val)).toThrowError();
    });

    expect(escapeForCmd('abcd1234')).toBe('abcd1234');
    expect(escapeForCmd('abcd 1234')).toBe('abcd 1234');
    expect(escapeForCmd('あいうえお')).toBe('あいうえお');
    expect(escapeForCmd('tag=R&B')).toBe('tag=R^&B');
    expect(escapeForCmd(300)).toBe('300');

    // Stdout -> Not escape
    expect(escapeForCmd('>')).toBe('>');
    expect(escapeForCmd('1>')).toBe('1>');
    expect(escapeForCmd('2>&1')).toBe('2>&1');

    expect(escapeForCmd('<')).toBe('<');
    expect(escapeForCmd('|')).toBe('|');
    expect(escapeForCmd('foo|bar')).toBe('foo^|bar');

    expect(escapeForCmd('/RegExp="^(A|The) $"')).toBe('/RegExp=\\"^^(A^|The) $\\"');
    expect(escapeForCmd('<%^[yyyy|yy]-MM-DD%>')).toBe('^<%^^[yyyy^|yy]-MM-DD%^>');

    // Path
    expect(escapeForCmd('C:\\Program Files')).toBe('C:\\Program Files');

    var noSpacePath = 'C:\\Windows\\System32\\notepad.exe';
    expect(escapeForCmd(noSpacePath)).toBe(noSpacePath);

    var wQuotationPath = '"C:\\Program Files (x86)\\Windows NT"';
    expect(escapeForCmd(wQuotationPath)).toBe('\\"C:\\Program Files (x86)\\Windows NT\\"');
  });

  test('joinCmdArgs', function () {
    var joinCmdArgs = os.joinCmdArgs;

    var noneArray = [true, false, undefined, null, 0, 1, NaN, Infinity, '', {}];
    noneArray.forEach(function (val) {
      expect(joinCmdArgs(val)).toBe('');
    });

    // String
    var argsStr = '//nologo "C:\\My Code\\test.wsf" -t Test';
    expect(joinCmdArgs(argsStr)).toBe(argsStr);

    // Empty Array
    expect(joinCmdArgs([])).toBe('');

    // Solid Array
    var args1 = [
      'C:\\Program Files (x86)\\Hoge\\foo.exe',
      '1>&2',
      'C:\\Users\\Tuckn\\err.log',
      '|',
      'tag=R&B',
      '/RegExp="^(A|The) $"',
      '<%^[yyyy|yy]-MM-DD%>'
    ];

    expect(joinCmdArgs(args1)).toBe(''
      + '"C:\\Program Files (x86)\\Hoge\\foo.exe" '
      + '1>&2 '
      + 'C:\\Users\\Tuckn\\err.log '
      + '| '
      + 'tag=R^&B '
      + '"/RegExp=\\"^^(A^|The) $\\"" '
      + '^<%^^[yyyy^|yy]-MM-DD%^>'
    );

    expect(joinCmdArgs(args1, { escapes: false })).toBe(''
      + '"C:\\Program Files (x86)\\Hoge\\foo.exe" '
      + '1>&2 '
      + 'C:\\Users\\Tuckn\\err.log '
      + '| '
      + 'tag=R&B '
      + '/RegExp="^(A|The) $" '
      + '<%^[yyyy|yy]-MM-DD%>'
    );

    var args2 = [
      'D:\\俺のフォルダ\\Hoge\\foo.exe',
      '>>',
      'E:\\Tuckn\\bar.exe',
      '2>>&1',
      '"piyo piyo.cmd" > out.txt'
    ];

    expect(joinCmdArgs(args2)).toBe(''
      + '"D:\\俺のフォルダ\\Hoge\\foo.exe" '
      + '>> '
      + 'E:\\Tuckn\\bar.exe '
      + '2>>&1 '
      + '"\\"piyo piyo.cmd\\" ^> out.txt"' // Oops! > to ^>
    );

    expect(joinCmdArgs(args2, { escapes: false })).toBe(''
      + '"D:\\俺のフォルダ\\Hoge\\foo.exe" '
      + '>> '
      + 'E:\\Tuckn\\bar.exe '
      + '2>>&1 '
      + '""piyo piyo.cmd" > out.txt"' // Oops! surrounded whole
    );
  });

  test('convToCmdlineStr', function () {
    var convToCmdlineStr = os.convToCmdlineStr;
    var command, args;

    var noneArray = [true, false, undefined, null, 0, 1, NaN, Infinity, '', {}];
    noneArray.forEach(function (val) {
      expect(_cb(os.convToCmdlineStr, val)).toThrowError();
    });

    command = convToCmdlineStr(NET);
    expect(command).toBe(NET);

    command = convToCmdlineStr(NET + ' use');
    expect(command).toBe('"' + NET + ' use"'); // Oops! Invalid command!

    command = convToCmdlineStr(NET, ['use']);
    expect(command).toBe(NET + ' use'); // OK! Executable command

    command = convToCmdlineStr(NET, ['use', '/delete']); // Array is OK
    expect(command).toBe(NET + ' use /delete');

    command = convToCmdlineStr(NET, 'use /delete'); // String is OK
    expect(command).toBe(NET + ' use /delete');

    command = convToCmdlineStr(NET, { args: ['use', '/delete'] }); // Object is NG
    expect(command).toBe(NET);

    // Option (in CommandPrompt)
    command = convToCmdlineStr(NET, ['use'], { shell: true });
    expect(command).toBe(CMD + ' /S /C"' + NET + ' use"');

    command = convToCmdlineStr(NET, 'use', { shell: true });
    expect(command).toBe(CMD + ' /S /C"' + NET + ' use"');

    command = convToCmdlineStr(NET, 'use', { shell: true, closes: false });
    expect(command).toBe(CMD + ' /S /K"' + NET + ' use"');

    // Surrounded with double quotations
    command = convToCmdlineStr('D:\\My Apps\\test.exe');
    expect(command).toBe('"D:\\My Apps\\test.exe"'); // surrounded

    command = convToCmdlineStr('"D:\\My Apps\\test.exe"');
    expect(command).toBe('"D:\\My Apps\\test.exe"'); // No surrounded

    command = convToCmdlineStr('D:\\My Apps\\test.exe', ['-n', 'Tuckn']);
    expect(command).toBe('"D:\\My Apps\\test.exe" -n Tuckn'); // Surrounded

    command = convToCmdlineStr('D:\\Apps\\test.exe', ['-n', 'Tuckn']);
    expect(command).toBe('D:\\Apps\\test.exe -n Tuckn'); // No surrounded

    // Escape
    // Array arguments is parsed
    command = convToCmdlineStr('D:\\My Apps\\test.exe', [
      '/RegExp="^(A|The) $"',
      '-f',
      'C:\\My Data\\img.doc'
    ]);
    expect(command).toBe('"D:\\My Apps\\test.exe"'
      + ' "/RegExp=\\"^^(A^|The) $\\"" -f "C:\\My Data\\img.doc"'); // executable!

    // String arguments is not parsed
    command = convToCmdlineStr(
      'D:\\My Apps\\test.exe',
      '/RegExp="^(A|The) $" -f C:\\My Data\\img.doc'
    );
    expect(command).toBe('"D:\\My Apps\\test.exe"'
      + ' /RegExp="^(A|The) $" -f C:\\My Data\\img.doc'); // Not executable!

    // Some patterns
    args = [
      '1>&2',
      'C:\\My Logs\\Tuckn\\err.log',
      '|',
      'tag=R&B',
      '/RegExp="^(A|The) $"',
      '<%^[yyyy|yy]-MM-DD%>'
    ];

    command = convToCmdlineStr(NET, args);
    expect(command).toBe(NET
      + ' 1>&2'
      + ' "C:\\My Logs\\Tuckn\\err.log"' // Surrounded
      + ' |'
      + ' tag=R^&B' // Escaped
      + ' "/RegExp=\\"^^(A^|The) $\\""'
      + ' ^<%^^[yyyy^|yy]-MM-DD%^>'
    );

    command = convToCmdlineStr(NET, args, { shell: true });
    expect(command).toBe(CMD + ' /S /C"' + NET
      + ' 1>&2'
      + ' "C:\\My Logs\\Tuckn\\err.log"' // Surrounded
      + ' |'
      + ' tag=R^&B' // Escaped
      + ' "/RegExp=\\"^^(A^|The) $\\""'
      + ' ^<%^^[yyyy^|yy]-MM-DD%^>'
      + '"'
    );

    command = convToCmdlineStr(NET, args, { escapes: false });
    expect(command).toBe(NET
      + ' 1>&2'
      + ' "C:\\My Logs\\Tuckn\\err.log"' // No surrounded
      + ' |'
      + ' tag=R&B' // No escaped
      + ' /RegExp="^(A|The) $"'
      + ' <%^[yyyy|yy]-MM-DD%>'
    );
  });

  test('dryRun', function () {
    var op = { isDryRun: true };
    var tmpPath = os.makeTmpPath();
    var retVal;
    var args;

    noneStrVals.forEach(function (val) {
      expect(_cb(os.shExec, val)).toThrowError();
    });

    retVal = os.shExec('mkdir', [tmpPath], objAdd(op, { shell: true }));
    expect(retVal).toBe('dry-run [os.shExec]: ' + CMD + ' /S /C"mkdir ' + tmpPath + '"'
    );

    args = ['//nologo', '//job:withErr', mockWsfCLI];

    retVal = os.shExec(CSCRIPT, args, op);
    expect(retVal).toContain(
      CMD + ' /S /C"' + CSCRIPT + ' ' + os.joinCmdArgs(args) + '"'
    );

    retVal = os.shExec(PING, '127.0.0.1', op);
    expect(retVal).toBe(
      'dry-run [os.shExec]: ' + CMD + ' /S /C"' + PING + ' 127.0.0.1"'
    );

    retVal = os.shExecSync('mkdir', [tmpPath], objAdd(op, { shell: true }));
    expect(retVal).toContain(CMD + ' /S /C"mkdir ' + tmpPath + '"');

    retVal = os.shExecSync(CSCRIPT, args, op);
    expect(retVal).toContain(
      CMD + ' /S /C"' + CSCRIPT + ' ' + os.joinCmdArgs(args) + '"'
    );

    retVal = os.shRun('mkdir', [tmpPath], objAdd(op, { shell: true }));
    expect(retVal).toContain(CMD + ' /S /C"mkdir ' + tmpPath + '"');

    retVal = os.shRunSync('mkdir', [tmpPath], objAdd(op, { shell: true }));
    expect(retVal).toContain(CMD + ' /S /C"mkdir ' + tmpPath + '"');

    args = ['//job:autoQuit1', mockWsfGUI];

    retVal = os.shRun(WSCRIPT, ['//job:autoQuit1', mockWsfGUI], op);
    expect(retVal).toContain(
      CMD + ' /S /C"' + WSCRIPT + ' ' + os.joinCmdArgs(args) + '"'
    );

    retVal = os.shRunSync(WSCRIPT, args, op);
    expect(retVal).toContain(
      CMD + ' /S /C"' + WSCRIPT + ' ' + os.joinCmdArgs(args) + '"'
    );

    args = ['/D', tmpPath + '-symlink', tmpPath];

    retVal = os.runAsAdmin('mklink', args, objAdd(op, { shell: true }));
    expect(retVal).toContain(CMD + ' /S /C"mklink ' + os.joinCmdArgs(args) + '"');
  });

  test('exec, CMD-Command, shell-true', function () {
    var tmpPath = os.makeTmpPath();
    expect(fso.FolderExists(tmpPath)).toBe(false);

    // Non shell option -> Fail, because mkdir is CMD command
    expect(_cb(os.shExec, 'mkdir', [tmpPath])).toThrowError();
    expect(fso.FolderExists(tmpPath)).toBe(false);
    // in CMD
    var retObj = os.shExec('mkdir', [tmpPath], { shell: true });
    expect(retObj.ExitCode).toBe(CD.runs.ok); // Always 0

    while (retObj.Status == 0) WScript.Sleep(300); // Waiting the finished

    expect(retObj.Status).toBe(1); // 0:Processing, 1:Finished
    expect(fso.FolderExists(tmpPath)).toBe(true);

    // Cleans
    fso.DeleteFolder(tmpPath, CD.fso.force.yes);
    expect(fso.FolderExists(tmpPath)).toBe(false);
  });

  test('exec, executingFile', function () {
    var oExec, stdOut, stdErr;

    // No Error Result
    oExec = os.shExec(CSCRIPT, ['//nologo', '//job:nonErr', mockWsfCLI]);

    expect(oExec.ExitCode).toBe(CD.runs.ok); // Always: 0
    expect(isNumber(oExec.ProcessID)).toBe(true); // Random

    while (oExec.Status == 0) WScript.Sleep(300); // Waiting the finished

    expect(oExec.ExitCode).toBe(CD.runs.ok);

    stdOut = oExec.StdOut.ReadAll();
    expect(stdOut.indexOf('StdOut Message') !== -1).toBe(true);
    expect(oExec.StdOut.ReadAll() === '').toBe(true); // empty after ReadAll

    stdErr = oExec.StdErr.ReadAll();
    expect(stdErr === '').toBe(true); // Empty
    expect(oExec.StdErr.ReadAll() === '').toBe(true); // empty after ReadAll

    // Error Result
    oExec = os.shExec(CSCRIPT, ['//nologo', '//job:withErr', mockWsfCLI]);

    expect(oExec.ExitCode).toBe(CD.runs.ok); // Always: 0
    expect(isNumber(oExec.ProcessID)).toBe(true); // Random

    while (oExec.Status == 0) WScript.Sleep(300); // Waiting the finished

    expect(oExec.ExitCode).toBe(CD.runs.err); // 1

    stdOut = oExec.StdOut.ReadAll();
    expect(stdOut.indexOf('StdOut Message') !== -1).toBe(true);
    expect(oExec.StdOut.ReadAll() === '').toBe(true); // empty after ReadAll

    stdErr = oExec.StdErr.ReadAll();
    expect(stdErr.indexOf('StdErr Message') !== -1).toBe(true);
    expect(oExec.StdErr.ReadAll() === '').toBe(true); // empty after ReadAll
  });

  test('exec, ping', function () {
    var oExec = os.shExec(PING, '127.0.0.1');

    expect(oExec.Status).toBe(0); // 0: Processing

    while (oExec.Status == 0) WScript.Sleep(300); // Waiting the finished

    expect(oExec.Status).toBe(1); // 1: Finished

    var stdout = oExec.StdOut.ReadAll();
    expect(/127\.0\.0\.1.+TTL/i.test(stdout)).toBe(true);
  });

  test('execSync, CMD-Command, shell-true', function () {
    var tmpPath = os.makeTmpPath();
    expect(fso.FolderExists(tmpPath)).toBe(false);

    // Non shell option -> Fail, because mkdir is CMD command
    expect(_cb(os.shExecSync, 'mkdir', [tmpPath])).toThrowError();
    expect(fso.FolderExists(tmpPath)).toBe(false);
    // in CMD
    var retObj = os.shExecSync('mkdir', [tmpPath], { shell: true });
    expect(fso.FolderExists(tmpPath)).toBe(true);
    expect(isPlainObject(retObj)).toBe(true); // A plain object
    expect(retObj.exitCode).toBe(CD.runs.ok); // OK: 0
    expect(retObj.error).toBe(false);
    expect(isEmpty(retObj.stdout)).toBe(true);
    expect(isEmpty(retObj.stderr)).toBe(true);

    // Cleans
    fso.DeleteFolder(tmpPath, CD.fso.force.yes);
    expect(fso.FolderExists(tmpPath)).toBe(false);

    noneStrVals.forEach(function (val) {
      expect(_cb(os.shExecSync, val)).toThrowError();
    });
  });

  test('execSync, executingFile', function () {
    var retObj;

    retObj = os.shExecSync(CSCRIPT, ['//nologo', '//job:withErr', mockWsfCLI]);
    expect(retObj.exitCode).toBe(CD.runs.err); // Error: 1
    expect(retObj.error).toBe(true);
    expect(retObj.stdout).toBe('StdOut Message');
    expect(retObj.stderr).toBe('StdErr Message');

    retObj = os.shExecSync(CSCRIPT, ['//nologo', '//job:nonErr', mockWsfCLI]);
    expect(retObj.exitCode).toBe(CD.runs.ok); // OK: 0
    expect(retObj.error).toBe(false);
    expect(retObj.stdout).toBe('StdOut Message');
    expect(isEmpty(retObj.stderr)).toBe(true);
  });

  test('run, CMD-Command, shell-true', function () {
    var tmpPath = os.makeTmpPath();
    var rtnVal;

    expect(fso.FolderExists(tmpPath)).toBe(false);

    // Non shell option -> Fail, because mkdir is CMD command
    expect(_cb(os.shRun, 'mkdir', [tmpPath])).toThrowError();
    expect(fso.FolderExists(tmpPath)).toBe(false);
    // in CMD
    rtnVal = os.shRun('mkdir', [tmpPath], { shell: true });
    expect(rtnVal).toBe(CD.runs.ok); // Always 0

    while (!fso.FolderExists(tmpPath)) WScript.Sleep(300);

    expect(fso.FolderExists(tmpPath)).toBe(true);

    fso.DeleteFolder(tmpPath, CD.fso.force.yes); // Cleans
    expect(fso.FolderExists(tmpPath)).toBe(false);
  });

  test('run, executingFile', function () {
    var rtnVal;

    rtnVal = os.shRun(WSCRIPT, ['//job:autoQuit0', mockWsfGUI]);
    expect(rtnVal).toBe(CD.runs.ok); // Always 0

    rtnVal = os.shRun(WSCRIPT, ['//job:autoQuit1', mockWsfGUI]);
    expect(rtnVal).toBe(CD.runs.ok); // Always 0
  });

  test('run, executingFile, winStyle: hidden', function () {
    var rtnVal, pIDs;

    // Confirm notepad.exe non-existing
    pIDs = _getProcessIDs(NOTEPAD);
    expect(pIDs).toHaveLength(0);

    rtnVal = os.shRun(NOTEPAD, [], { winStyle: 'hidden' });
    expect(rtnVal).toBe(CD.runs.ok);

    WScript.Sleep(3000); // Confirm hidden notepad.exe :-)

    // Terminate notepad.exe
    pIDs = _getProcessIDs(NOTEPAD);
    pIDs.forEach(function (pID) { _terminateProcesses(pID); });
  });

  test('run, executingFile, winStyle: activeDef', function () {
    var rtnVal, pIDs;

    // Confirm no notepad.exe existing
    pIDs = _getProcessIDs(NOTEPAD);
    expect(pIDs).toHaveLength(0);

    rtnVal = os.shRun(NOTEPAD, [], { winStyle: 'activeDef' });
    expect(rtnVal).toBe(CD.runs.ok);

    WScript.Sleep(3000); // Confirm the state of notepad.exe

    // Terminate notepad.exe
    pIDs = _getProcessIDs(NOTEPAD);
    pIDs.forEach(function (pID) { _terminateProcesses(pID); });
  });

  test('run, executingFile, winStyle: activeMin', function () {
    var rtnVal, pIDs;

    // Confirm no notepad.exe existing
    pIDs = _getProcessIDs(NOTEPAD);
    expect(pIDs).toHaveLength(0);

    rtnVal = os.shRun(NOTEPAD, [], { winStyle: 'activeMin' });
    expect(rtnVal).toBe(CD.runs.ok);

    WScript.Sleep(3000); // Confirm the state of notepad.exe

    // Terminate notepad.exe
    pIDs = _getProcessIDs(NOTEPAD);
    pIDs.forEach(function (pID) { _terminateProcesses(pID); });
  });

  test('run, executingFile, winStyle: activeMax', function () {
    var rtnVal, pIDs;

    // Confirm no notepad.exe existing
    pIDs = _getProcessIDs(NOTEPAD);
    expect(pIDs).toHaveLength(0);

    rtnVal = os.shRun(NOTEPAD, [], { winStyle: 'activeMax' });
    expect(rtnVal).toBe(CD.runs.ok);

    WScript.Sleep(3000); // Confirm the state of notepad.exe

    // Terminate notepad.exe
    pIDs = _getProcessIDs(NOTEPAD);
    pIDs.forEach(function (pID) { _terminateProcesses(pID); });
  });

  test('run, executingFile, winStyle: nonActive', function () {
    var rtnVal, pIDs;

    // Confirm no notepad.exe existing
    pIDs = _getProcessIDs(NOTEPAD);
    expect(pIDs).toHaveLength(0);

    rtnVal = os.shRun(NOTEPAD, [], { winStyle: 'nonActive' });
    expect(rtnVal).toBe(CD.runs.ok);

    WScript.Sleep(3000); // Confirm the state of notepad.exe

    // Terminate notepad.exe
    pIDs = _getProcessIDs(NOTEPAD);
    pIDs.forEach(function (pID) { _terminateProcesses(pID); });
  });

  test('runSync, CMD-Command, shell-true', function () {
    var tmpPath = os.makeTmpPath();
    expect(fso.FolderExists(tmpPath)).toBe(false);

    // Non shell option -> Fail, because mkdir is CMD command
    expect(_cb(os.shRunSync, 'mkdir', [tmpPath])).toThrowError();
    expect(fso.FolderExists(tmpPath)).toBe(false);
    // in CMD
    var rtnVal = os.shRunSync('mkdir', [tmpPath], { shell: true });
    expect(isNumber(rtnVal)).toBe(true); // A number
    expect(rtnVal).toBe(CD.runs.ok); // OK:0, Error:1
    expect(fso.FolderExists(tmpPath)).toBe(true);

    fso.DeleteFolder(tmpPath, CD.fso.force.yes); // Cleans
    expect(fso.FolderExists(tmpPath)).toBe(false);

    noneStrVals.forEach(function (val) {
      expect(_cb(os.shRunSync, val)).toThrowError();
    });
  });

  test('runSync, executingFile', function () {
    var rtnVal;

    rtnVal = os.shRunSync(WSCRIPT, ['//job:autoQuit0', mockWsfGUI]);
    expect(rtnVal).toBe(CD.runs.ok);

    rtnVal = os.shRunSync(WSCRIPT, ['//job:autoQuit1', mockWsfGUI]);
    expect(rtnVal).toBe(CD.runs.err);
  });

  test('isAdmin', function () {
    // expect(os.isAdmin()).toBe('');
    expect('@TODO').toBe('tested');
  });

  test('runAsAdmin, CMD Command, shell-true)', function () {
    // Create a test directory
    var tmpPath = os.makeTmpPath();
    expect(fso.FolderExists(tmpPath)).toBe(false);
    fso.CreateFolder(tmpPath);
    expect(fso.FolderExists(tmpPath)).toBe(true);

    // Create a sub directory in the test directory
    var subDir = path.join(tmpPath, 'SubDir');
    expect(fso.FolderExists(subDir)).toBe(false);
    fso.CreateFolder(subDir);
    expect(fso.FolderExists(subDir)).toBe(true);

    // Name a symlink path
    var subDirSymlink = path.join(tmpPath, 'SubDir-Symlink');

    // Non shell option -> Fail, because mklink is CMD command
    var mainCmd = 'mklink';
    var args = ['/D', subDirSymlink, subDir];
    // @FIXME Can not catch the error in admin process.
    // expect(_cb(os.runAsAdmin, mainCmd, args)).toThrowError();
    expect(fso.FolderExists(subDirSymlink)).toBe(false);
    // Using CMD option
    var rtnVal = os.runAsAdmin(mainCmd, args, { shell: true });
    expect(rtnVal).toBe(undefined);

    var isCreated = false;
    var msecTimeOut = 6000;
    do {
      isCreated = fso.FolderExists(subDirSymlink);
      WScript.Sleep(300);
      msecTimeOut -= 300;
    } while (!isCreated && msecTimeOut > 0);
    expect(isCreated).toBe(true);

    // Cleans
    fso.DeleteFolder(subDirSymlink, CD.fso.force.yes);
    fso.DeleteFolder(subDir, CD.fso.force.yes);
    fso.DeleteFolder(tmpPath, CD.fso.force.yes);
    expect(fso.FolderExists(tmpPath)).toBe(false);

    noneStrVals.forEach(function (val) {
      expect(_cb(os.runAsAdmin, val)).toThrowError();
    });
  });

  test('runAsAdmin, executingFile', function () {
    var rtnVal;
    rtnVal = os.runAsAdmin(NET, ['session']);
    expect(rtnVal).toBe(undefined);
  });

  test('for an emergency', function () {
    // for an emergency. e.g. Infinity roop
    // _terminateProcesses(path.basename(WSCRIPT));
  });
});
