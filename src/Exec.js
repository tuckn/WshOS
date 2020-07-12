/* globals Wsh: false */

(function () {
  // Shorthands
  var CD = Wsh.Constants;
  var util = Wsh.Util;
  var sh = Wsh.Shell;
  var shApp = Wsh.ShellApplication;

  var objAssign = Object.assign;
  var insp = util.inspect;
  var isArray = util.isArray;
  var isNumber = util.isNumber;
  var isPureNumber = util.isPureNumber;
  var isString = util.isString;
  var isSolidArray = util.isSolidArray;
  var isSolidString = util.isSolidString;
  var obtain = util.obtainPropVal;

  var os = Wsh.OS;

  /** @constant {string} */
  var MODULE_TITLE = 'WshOS/Exec.js';

  var throwErrNonStr = function (functionName, typeErrVal) {
    util.throwTypeError('string', MODULE_TITLE, functionName, typeErrVal);
  };

  // os.surroundPath {{{
  /**
   * Surrounds the file path with double quotes.
   *
   * @example
   * var os = Wsh.OS; // Shorthand
   *
   * os.surroundPath('C:\\Windows\\System32\\notepad.exe'); // No space
   * // Returns: 'C:\\Windows\\System32\\notepad.exe'
   *
   * os.surroundPath('C:\\Program Files'); // Has a space
   * // Returns: '"C:\\Program Files"'
   *
   * os.surroundPath('"C:\\Program Files (x86)\\Windows NT"'); // Already quoted
   * // Returns: '"C:\\Program Files (x86)\\Windows NT"'
   *
   * os.surroundPath('D:\\2011-01-01家族で初詣'); // Non-ASCII chars
   * // Returns: '"D:\\2011-01-01家族で初詣"'
   * @function surroundPath
   * @memberof Wsh.OS
   * @param {string} str - The path to surround.
   * @returns {string} - The surrounded path.
   */
  os.surroundPath = function (str) {
    if (!isString(str)) throwErrNonStr('os.surroundPath', str);

    if (str === '') return '';
    if (util.isASCII(str) && !/[\s"&<>^|]/.test(str)) return str;
    if (/^".*"$/.test(str)) return str;
    return '"' + str + '"';
  };
  var _srrPath = os.surroundPath;
  // }}}

  // os.escapeForCmd {{{
  /**
   * Escapes the string of argument in CMD.exe. Note the difference from {@link Wsh.OS.joinCmdArgs}. {@link http://thinca.hatenablog.com/entry/20100210/1265813598|Ref1} {@link http://output.jsbin.com/anitaz/11|Ref2}
   *
   * @example
   * var os = Wsh.OS; // Shorthand
   *
   * os.escapeForCmd('abcd1234'); // Returns: 'abcd1234'
   * os.escapeForCmd('abcd 1234'); // Returns: '"abcd 1234"
   * os.escapeForCmd('あいうえお'); // Returns: '"あいうえお"'
   * os.escapeForCmd('tag=R&B'); // Returns: '"tag=R^&B"'
   *
   * os.escapeForCmd('C:\\Program Files');
   * // Returns: '"C:\\Program Files"');
   *
   * os.escapeForCmd('C:\\Windows\\System32\\notepad.exe');
   * // Returns: 'C:\\Windows\\System32\\notepad.exe'
   *
   * os.escapeForCmd("C:\\Program Files (x86)\\Windows NT");
   * // Returns: '"\\"C:\\Program Files (x86)\\Windows NT\\""'
   *
   * os.escapeForCmd('>');
   * // Returns: '"^>"'
   *
   * os.escapeForCmd('/RegExp="^(A|The) $"');
   * // Returns: '"/RegExp=\\"^^(A^|The) $\\""'
   *
   * os.escapeForCmd('<%^[yyyy|yy]-MM-DD%>');
   * // Returns: '"^<%^^[yyyy^|yy]-MM-DD%^>"'
   * @function escapeForCmd
   * @memberof Wsh.OS
   * @param {string} str - The string to convert.
   * @returns {string} - The string escaped for Command-Prompt.
   */
  os.escapeForCmd = function (str) {
    if (!isNumber(str) && !isString(str)) {
      throwErrNonStr('os.escapeForCmd', str);
    }

    if (str === '') return '""';
    if (isNumber(str)) return String(str);
    if (util.isASCII(str) && !/[\s"&<>^|]/.test(str)) return str;

    var escArg = '"'
      + str.replace(/(["])/g, '\\$1').replace(/([&<>^|])/g, '^$1')
      + '"';

    return escArg;
  };
  var _escapeForCmd = os.escapeForCmd;
  // }}}

  // os.joinCmdArgs {{{
  /**
   * Escapes and joines the Array of arguments for Command-Prompt.
   *
   * @example
   * Wsh.OS.joinCmdArgs([
   *   'C:\\Program Files (x86)\\Hoge\\foo.exe',
   *   '1>&2', // Redirect
   *   'C:\\Logs\\err.log',
   *   '|', //Pipe
   *   'tag=R&B',
   *   '/RegExp="^(A|The) $"',
   *   '<%^[yyyy|yy]-MM-DD%>'
   * ]);
   * // Returns: '"C:\\Program Files (x86)\\Hoge\\foo.exe" 1>&2 C:\\Logs\\err.log | "tag=R^&B" "/RegExp=\\"^^(A^|The) $\\"" "^<%^^[yyyy^|yy]-MM-DD%^>"'
   * @function joinCmdArgs
   * @memberof Wsh.OS
   * @param {Array} args - The arguments to convert.
   * @param {object} [options] - Optional parameters.
   * @param {boolean} [options.escapes=true] - Escapes arguments.
   * @returns {string} - The string of escaped and joined.
   */
  os.joinCmdArgs = function (args, options) {
    if (isSolidString(args)) return String(args);
    if (!isSolidArray(args)) return '';

    var escapes = obtain(options, 'escapes', true);

    var argsStr = args.reduce(function (acc, arg) {
      if (!escapes) return acc + ' ' + arg;
      if (/^(1|2)?>{1,2}(&(1|2))?$/.test(arg)) return acc + ' ' + arg;
      if (arg === '<' || arg === '|') return acc + ' ' + arg;
      return acc + ' ' + _escapeForCmd(arg);
    }, '');

    return argsStr.trim();
  };
  var _joinCmdArgs = os.joinCmdArgs;
  // }}}

  // os.convToCmdCommand {{{
  /**
   * @typedef {object} typeConvToCommandOptions
   * @property {boolean} [shell=false] - Wraps with CMD.EXE
   * @property {boolean} [closes=true] - /C (continue) or /K (kill)
   * @property {boolean} [escapes=true] - Escapes the arguments.
   */

  /**
   * Converts the command and arguments for Command-Prompt.
   *
   * @example
   * var os = Wsh.OS;
   *
   * os.convToCmdCommand('net', ['use', '/delete']);
   * // Returns: 'net user /delete'
   *
   * os.convToCmdCommand('net', ['use', '/delete'], { shell: true });
   * // Returns: 'C:\Windows\System32\cmd.exe /S /C"net user /delete"'
   *
   * os.convToCmdCommand('net', 'use /delete', { shell: true, closes: false });
   * // Returns: 'C:\Windows\System32\cmd.exe /S /K"net user /delete"'
   *
   * // The 2nd argument: Array vs String
   * // Array is escaped
   * os.convToCmdCommand('D:\\My Apps\\app.exe',
   *   ['/RegExp="^(A|The) $"', '-f', 'C:\\My Data\\img.doc']);
   * // Returns: '"D:\\My Apps\\app.exe" "/RegExp=\\"^^(A^|The) $\\"" -f "C:\\My Data\\img.doc"'
   *
   * // String is not escaped
   * os.convToCmdCommand('D:\\My Apps\\app.exe',
   *   '/RegExp="^(A|The) $" -f C:\\My Data\\img.doc');
   * // Returns: '/RegExp="^(A|The) $" -f C:\\My Data\\img.doc'
   * @function convToCmdCommand
   * @memberof Wsh.OS
   * @param {string} cmdStr - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments to format.
   * @param {typeConvToCommandOptions} [options] - Optional parameters.
   * @returns {string} - The converted command.
   */
  os.convToCmdCommand = function (cmdStr, args, options) {
    var functionName = 'os.convToCmdCommand';
    if (!isSolidString(cmdStr)) throwErrNonStr(functionName, cmdStr);

    var command = _srrPath(cmdStr);

    var argsStr;
    if (isArray(args)) {
      argsStr = _joinCmdArgs(args, options);
    } else if (isString(args)) {
      argsStr = args;
    } else {
      argsStr = '';
      // throw new Error('Error [ERR_INVALID_ARG_TYPE]\n'
      //   + '  at ' + functionName + ' (' + MODULE_TITLE + ')\n'
      //   + '  args: ' + insp(args));
    }

    if (argsStr !== '') command += ' ' + argsStr;

    var shell = obtain(options, 'shell', false);
    var closes = obtain(options, 'closes', true);
    /**
     * /C or /K <cmd.exe op1 op2...>: cmdを実行（/Kは実行後終了しない
     * /S: /C or /K 後の""で括られた中身をコマンド全体とみなす。
     * "のエスケープに""とする必要がなくなる
     * "以外の特殊文字は^でエスケープする必要がある
     */
    if (shell) {
      if (closes) {
        command = _srrPath(os.exefiles.cmd) + ' /S /C"' + command + '"';
      } else {
        command = _srrPath(os.exefiles.cmd) + ' /S /K"' + command + '"';
      }
    }

    return command;
  }; // }}}

  // os.exec {{{
  /**
   * The object returnning from Wsh.OS.shExec ({@link https://msdn.microsoft.com/ja-jp/library/cc364375.aspx|WshScriptExec object}).
   *
   * @typedef {object} typeExecObject
   * @property {number} ExitCode
   * @property {number} ProcessID
   * @property {number} Status
   * @property {object} StdOut
   * @property {object} StdIn
   * @property {object} StdErr
   * @property {function} Terminate
   */

  /**
   * Asynchronously executes the command with WScript.Shell.Exec. Note that a DOS window is always displayed. {@link https://msdn.microsoft.com/ja-jp/library/cc364375.aspx|WshScriptExec object}
   *
   * @example
   * var os = Wsh.OS;
   *
   * // Ex.1 CMD-command
   * // Throws a Error. `mkdir` is the CMD command
   * os.exec('mkdir', 'D:\\Temp');
   *
   * // With the `shell` option
   * var oExec1 = os.exec('mkdir', 'D:\\Temp', { shell: true });
   * console.log(oExec1.ExitCode); // 0 (always)
   *
   * while (oExec1.Status == 0) WScript.Sleep(300); // Waiting the finished
   * console.log(oExec1.Status); // 1 (It means finished)
   *
   * // Ex.2 Exe-file
   * var oExec2 = os.exec('ping.exe', ['127.0.0.1']);
   * console.log(oExec2.ExitCode); // 0 (always)
   *
   * while (oExec2.Status == 0) WScript.Sleep(300); // Waiting the finished
   * console.log(oExec2.Status); // 1 (It means finished)
   *
   * var result = oExec2.StdOut.ReadAll();
   * console.log(result); // Outputs the result of ping 127.0.0.1
   * @function exec
   * @memberof Wsh.OS
   * @param {string} cmdStr - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments.
   * @param {typeConvToCommandOptions} [options] - Optional parameters.
   * @returns {typeExecObject} - See {@link Wsh.OS.typeExecObject}
   */
  os.exec = function (cmdStr, args, options) {
    var functionName = 'os.exec';
    if (!isString(cmdStr)) throwErrNonStr(functionName, cmdStr);

    var command = os.convToCmdCommand(cmdStr, args, options);
    try {
      /**
       * @function Exec
       * @memberof Wsh.Shell
       * @param {string} strCommand
       * @returns {typeExecObject}
       */
      return sh.Exec(command);
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + functionName + ' (' + MODULE_TITLE + ')\n'
        + '  command: ' + command);
    }
  }; // }}}

  // os.execSync {{{
  /**
   * @typedef {object} typeExecSyncReturn
   * @property {number} exitCode - There are applications that return 1 (NG) even if the processing is successful
   * @property {string} stdout
   * @property {string} stderr
   * @property {boolean} error
   */

  /**
   * Synchronize of {@link Wsh.OS.exec}
   *
   * @example
   * var os = Wsh.OS;
   *
   * // Ex1. CMD-command
   * // Throws a Error. `mkdir` is the CMD command
   * os.execSync('mkdir', 'D:\\Temp');
   *
   * // With the `shell` option
   * var retObj1 = os.execSync('mkdir', 'D:\\Temp', { shell: true });
   * console.dir(retObj1);
   * // Outputs: {
   * //   exitCode: 0,
   * //   stdout: "",
   * //   stderr: "",
   * //   error: false };
   *
   * // Ex2. Exe-file
   * var retObj2 = os.execSync('ping.exe', ['127.0.0.1']);
   * console.dir(retObj2);
   * // Outputs: {
   * //   exitCode: 0,
   * //   stdout: <The result of ping 127.0.0.1>,
   * //   stderr: "",
   * //   error: false };
   * @function execSync
   * @memberof Wsh.OS
   * @param {string} cmdStr - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments.
   * @param {object} [options] - See {@link Wsh.OS.convToCmdCommand}
   * @returns {typeExecSyncReturn} - See {@link typeExecSyncReturn}
   */
  os.execSync = function (cmdStr, args, options) {
    var functionName = 'os.execSync';
    if (!isString(cmdStr)) throwErrNonStr(functionName, cmdStr);

    var command = os.convToCmdCommand(cmdStr, args, options);
    if (!isSolidString(command)) throwErrNonStr(functionName, command);

    var oExec;
    var stdout = '';
    var stderr = '';
    var readSize = 4096;

    /** The problem of Exec Stream. {@link https://social.msdn.microsoft.com/Forums/ja-JP/da5a0366-6446-49f0-b4e2-adfbd4314aac/vba12391wsh12434exec123912345526045261781239832066201023090635469?forum=vbajp|Microsoft Forums}, {@link https://srad.jp/~IR.0-4/journal/572274/|Ref} */
    try {
      oExec = sh.Exec(command);

      // @note .Status -> 0:Processing, 1:Finished
      while (oExec.Status === 0) {
        stdout += oExec.StdOut.Read(readSize);
        stderr += oExec.StdErr.Read(readSize);
        WScript.Sleep(300);
      }
      // Suck out the remnants
      while (!oExec.StdOut.AtEndOfStream) {
        stdout += oExec.StdOut.Read(readSize);
      }
      while (!oExec.StdOut.AtEndOfStream) {
        stdout += oExec.StdErr.Read(readSize);
      }
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + functionName + ' (' + MODULE_TITLE + ')\n'
        + '  command: ' + command);
    }

    return {
      exitCode: oExec.ExitCode,
      stdout: stdout,
      stderr: stderr,
      error: isSolidString(stderr)
    };
  }; // }}}

  // _shRun {{{
  /**
   * @typedef {object} typeShRunOptions
   * @property {(number|string)} [winStyle="activeDef"] - See {@link https://docs.tuckn.net/WshUtil/Wsh.Constants.windowStyles.html|Wsh.Constants.windowStyles}.
   * @property {boolean} [waits=true] - See {@link https://docs.tuckn.net/WshUtil/Wsh.Constants.waits.html|Wsh.Constants.waits}.
   */

  /**
   * @private
   * @function _shRun
   * @memberof Wsh.OS
   * @param {string} command - The command and arguments.
   * @param {typeShRunOptions} [options] - Optional parameters.
   * @returns {number} - Always return 0.
   */
  function _shRun (command, options) {
    var functionName = '_shRun';
    if (!isSolidString(command)) throwErrNonStr(functionName, command);

    var winStyle = obtain(options, 'winStyle', 'activeDef');
    var winStyleCode = isPureNumber(winStyle)
      ? winStyle : CD.windowStyles[winStyle];
    var waits = obtain(options, 'waits', CD.waits.yes);

    try {
      /**
       * @function Run
       * @memberof Wsh.Shell
       * @param {string} strCommand
       * @param {number} intWindowStyle
       * @param {boolean} bWaitOnReturn
       * @returns {number}
       */
      return sh.Run(command, winStyleCode, waits);
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + functionName + ' (' + MODULE_TITLE + ')\n'
        + '  command: ' + command);
    }
  } // }}}

  // os.run {{{
  /**
   * Asynchronously runs the application with WScript.Shell.Run. {@link https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364421%28v%3dmsdn.10%29|WSH Run method}
   *
   * @example
   * var os = Wsh.OS;
   * var fso = Wsh.FileSystemObject;
   *
   * // Ex1. CMD-command
   * // Throws a Error. `mkdir` is the CMD command
   * os.run('mkdir', 'D:\\Temp');
   *
   * // With the options
   * os.run('mkdir', 'D:\\Temp', { shell: true, winStyle: 'hidden' });
   * // Returns: Always 0
   *
   * while (!fso.FolderExists('D:\\Temp')) { // Waiting the created
   *   WScript.Sleep(300);
   * }
   * console.dir(fso.FolderExists('D:\\Temp')); // true
   *
   * // Ex2. Exe-file
   * os.run('notepad.exe', ['D:\\Test.txt'], { winStyle: 'activeMax' });
   * // Returns: Always 0
   * @function run
   * @memberof Wsh.OS
   * @param {string} cmdStr - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments.
   * @param {object} [options] See {@link typeShRunOptions} and {@link typeConvToCommandOptions}
   * @returns {number} - Always return 0.
   */
  os.run = function (cmdStr, args, options) {
    var functionName = 'os.run';
    if (!isString(cmdStr)) throwErrNonStr(functionName, cmdStr);

    var command = os.convToCmdCommand(cmdStr, args, options);
    return _shRun(command, objAssign({}, options, { waits: false }));
  }; // }}}

  // os.runSync {{{
  /**
   * Synchronize of {@link Wsh.OS.run}
   *
   * @example
   * var os = Wsh.OS;
   * var fso = Wsh.FileSystemObject;
   *
   * // Ex1. CMD-command
   * // Throws a Error. `mkdir` is the CMD command
   * os.runSync('mkdir', 'D:\\Temp');
   *
   * // With the options
   * var retVal1 = os.runSync('mkdir', 'D:\\Temp', {
   *   shell: true, winStyle: 'hidden'
   * });
   *
   * console.log(retVal1); // Outputs the number of the result code
   * console.dir(fso.FolderExists('D:\\Temp')); // true
   *
   * // Ex2. Exe-file
   * var retVal2 = os.runSync('notepad.exe', ['D:\\Test.txt'], {
   *   winStyle: 'activeMax'
   * });
   *
   * // Waits until the notepad process is finished
   *
   * console.log(retVal2); // Outputs the number of the result code
   * @function runSync
   * @memberof Wsh.OS
   * @param {string} cmdStr - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments.
   * @param {object} [options] See {@link typeShRunOptions} and {@link typeConvToCommandOptions}
   * @returns {number} - The return code from the app.
   */
  os.runSync = function (cmdStr, args, options) {
    var functionName = 'os.runSync';
    if (!isString(cmdStr)) throwErrNonStr(functionName, cmdStr);

    var command = os.convToCmdCommand(cmdStr, args, options);
    return _shRun(command, objAssign({}, options, { waits: true }));
  }; // }}}

  var _isAdmin = undefined;

  // os.isAdmin {{{
  /**
   * Checks if this process is running as Administrator authority.
   *
   * @example
   * Wsh.OS.isAdmin(); // Returns: false
   * @function isAdmin
   * @memberof Wsh.OS
   * @returns {boolean} - true -> running as Administrator authority.
   */
  os.isAdmin = function () {
    if (_isAdmin !== undefined) return _isAdmin;

    try {
      /** `net session` worked on from Windows XP to 10 */
      var iRetVal = os.runSync(os.exefiles.net, ['session'], {
        winStyle: 'hidden'
      });
      _isAdmin = iRetVal === CD.runs.ok;
    } catch (e) {
      _isAdmin = false;
    }

    return _isAdmin;
  }; // }}}

  // os.runAsAdmin {{{
  /**
   * Asynchronously runs the application as administrator authority with WScript.Shell.ShellExecute. {@link https://docs.microsoft.com/en-us/windows/win32/shell/shell-shellexecute|WSH ShellExecute method}
   *
   * @example
   * var os = Wsh.OS;
   *
   * // Ex1. CMD-command
   * // Throws a Error. `mklink` is the CMD command
   * os.runAsAdmin('mklink', 'D:\\Temp-Symlink D:\\Temp');
   *
   * // With the `shell` option
   * os.runAsAdmin('mklink', 'D:\\Temp-Symlink D:\\Temp', {
   *   shell: true
   * });
   *
   * // Ex2. Exe-file
   * os.runAsAdmin('D:\\MyApp\\Everything.exe', ['-instance', 'MySet']);
   * @function runAsAdmin
   * @memberof Wsh.OS
   * @param {string} cmdStr - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments.
   * @param {object} [options] See {@link typeShRunOptions} and {@link typeConvToCommandOptions}
   * @returns {void}
   */
  os.runAsAdmin = function (cmdStr, args, options) {
    if (os.isAdmin()) return os.run(cmdStr, args, options);

    var functionName = 'os.runAsAdmin';
    if (!isString(cmdStr)) throwErrNonStr(functionName, cmdStr);

    var exePath = cmdStr;

    var argsStr;
    if (isArray(args)) {
      argsStr = _joinCmdArgs(args, options);
    } else if (isString(args)) {
      argsStr = args;
    } else {
      argsStr = '';
    }

    var shell = obtain(options, 'shell', false);
    if (shell) {
      exePath = os.exefiles.cmd;
      argsStr = '/S /C"' + _srrPath(cmdStr) + ' ' + argsStr + '"';
    }

    var winStyle = obtain(options, 'winStyle', 'activeDef');
    var winStyleCode = isPureNumber(winStyle) ? winStyle : CD.windowStyles[winStyle];

    try {
      /**
       * Performs the specified operation on a specified file. It's asyncronize. {@link https://docs.microsoft.com/en-us/windows/win32/shell/shell-shellexecute|Microsoft Docs}
       * @function ShellExecute
       * @memberof Wsh.ShellApplication
       * @param {string} sFile
       * @param {string} [vArguments]
       * @param {string} [vDirectory]
       * @param {string} [vOperation=open] - {@link https://docs.microsoft.com/en-us/windows/win32/shell/context|Microsoft Docs}
       * @param {number} [vShow]
       * @returns {void}
       */
      return shApp.ShellExecute(
        exePath,
        argsStr,
        'open',
        'runas',
        winStyleCode
      );
    } catch (e) {
      // @FIXME Can not catch the error in admin process
      throw new Error(insp(e) + '\n'
        + '  at _execFileAsAdmin (' + MODULE_TITLE + ')\n'
        + '  file: ' + insp(exePath) + '\n  argsStr: ' + insp(argsStr));
    }
  }; // }}}
})();

// vim:set foldmethod=marker commentstring=//%s :
