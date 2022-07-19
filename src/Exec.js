/* globals Wsh: false */

(function () {
  // Shorthands
  var CD = Wsh.Constants;
  var util = Wsh.Util;
  var sh = Wsh.Shell;
  var shApp = Wsh.ShellApplication;

  var objAdd = Object.assign;
  var insp = util.inspect;
  var isArray = util.isArray;
  var isNumber = util.isNumber;
  var isPureNumber = util.isPureNumber;
  var isString = util.isString;
  var isSolidArray = util.isSolidArray;
  var isSolidString = util.isSolidString;
  var includes = util.includes;
  var obtain = util.obtainPropVal;

  var os = Wsh.OS;

  /** @constant {string} */
  var MODULE_TITLE = 'WshOS/Exec.js';

  var throwErrNonStr = function (functionName, typeErrVal) {
    util.throwTypeError('string', MODULE_TITLE, functionName, typeErrVal);
  };

  // os.surroundCmdArg {{{
  /**
   * Surrounds a file path and argument with double quotes.
   *
   * @example
   * var os = Wsh.OS; // Shorthand
   *
   * os.surroundCmdArg('C:\\Windows\\System32\\notepad.exe'); // No space
   * // Returns: 'C:\\Windows\\System32\\notepad.exe'
   *
   * os.surroundCmdArg('C:\\Program Files'); // Has a space
   * // Returns: '"C:\\Program Files"'
   *
   * os.surroundCmdArg('"C:\\Program Files (x86)\\Windows NT"'); // Already quoted
   * // Returns: '"C:\\Program Files (x86)\\Windows NT"'
   *
   * os.surroundCmdArg('D:\\2011-01-01家族で初詣'); // Non-ASCII chars
   * // Returns: '"D:\\2011-01-01家族で初詣"'
   *
   * os.surroundCmdArg('D:\\Music\\R&B'); // Ampersand
   * // Returns: '"D:\\Music\\R&B"'
   *
   * os.surroundCmdArg('abcd1234'); // Returns: 'abcd1234'
   * os.surroundCmdArg('abcd 1234'); // Returns: '"abcd 1234"
   * os.surroundCmdArg('あいうえお'); // Returns: '"あいうえお"'
   *
   * // A command control character
   * os.surroundCmdArg('|'); // Returns: '|'
   * os.surroundCmdArg('>'); // Returns: '>'
   *
   * // Inner quoted
   * os.surroundCmdArg('-p"My p@ss wo^d"'); // Returns: '-p"My p@ss wo^d"'
   * os.surroundCmdArg('1> "C:\\logs.txt"'); // Returns: '"1> "C:\\logs.txt""'
   * os.surroundCmdArg('1>"C:\\logs.txt"'); // Returns: '1>"C:\\logs.txt"'
   * @function surroundCmdArg
   * @memberof Wsh.OS
   * @param {string} str - The path to surround.
   * @returns {string} - The surrounded path.
   */
  os.surroundCmdArg = function (str) {
    if (!isString(str)) throwErrNonStr('os.surroundCmdArg', str);

    if (str === '') return '';

    // Already quoted
    if (/^".*"$/.test(str)) return str;

    if (util.isASCII(str)) {
      if (!includes(str, ' ')) return str;

      // @Note CMD treats the ,=; as an argument delimiter
      // if (!/["&<>^|,=;]/.test(str)) return str;
    }

    // Inner quoted
    // Ex1. Stdout: 1>"D:\\My data\\logs.txt"
    // Ex2. 7-Zip password option: -p"My p@ss wo^d", @"D:\excluded files.txt"
    // @TODO This completion is low level ...
    if (/^[A-Za-z0-9_@>/=-]+(".+")?[A-Za-z0-9_@&/=-]*$/.test(str)) return str;

    return '"' + str + '"';
  };
  var _srrPath = os.surroundCmdArg;
  // }}}

  // os.escapeForCmd {{{
  /**
   * Escapes the string of argument in CMD.exe. Note {@link http://thinca.hatenablog.com/entry/20100210/1265813598|Ref1} {@link http://output.jsbin.com/anitaz/11|Ref2}
   *
   * @example
   * var os = Wsh.OS; // Shorthand
   *
   * os.escapeForCmd('tag=R&B'); // Returns: 'tag=R^&B'
   *
   * os.escapeForCmd('>');
   * // Returns: '>' // Not escaped. It's a redirect character.
   *
   * os.escapeForCmd('/RegExp="^(A|The) $"');
   * // Returns: '"/RegExp=\\"^^(A^|The) $\\""'
   *
   * os.escapeForCmd('<%^[yyyy|yy]-MM-DD%>');
   * // Returns: '^<%^^[yyyy^|yy]-MM-DD%^>'
   *
   * os.escapeForCmd(300);
   * // Returns: '300'
   *
   * os.escapeForCmd('C:\\Program Files');
   * // Returns: 'C:\\Program Files');
   *
   * os.escapeForCmd('C:\\Windows\\System32\\notepad.exe');
   * // Returns: 'C:\\Windows\\System32\\notepad.exe'
   *
   * os.escapeForCmd('"C:\\Program Files (x86)\\Windows NT"');
   * // Returns: '\\"C:\\Program Files (x86)\\Windows NT\\"'
   * @function escapeForCmd
   * @memberof Wsh.OS
   * @param {string} str - The string to convert.
   * @returns {string} - The string escaped for Command-Prompt.
   */
  os.escapeForCmd = function (str) {
    if (!isNumber(str) && !isString(str)) throwErrNonStr('os.escapeForCmd', str);

    if (isNumber(str)) return String(str);
    if (str === '') return str;

    // Stdout
    if (/^(1|2)?>{1,2}(&(1|2))?$/.test(str)) return str;

    // Pipe
    if (str === '<' || str === '|') return str;

    return str.replace(/(["])/g, '\\$1').replace(/([&<>^|])/g, '^$1');
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
   * // Returns: '"C:\\Program Files (x86)\\Hoge\\foo.exe" 1>&2 C:\\Logs\\err.log | tag=R^&B "/RegExp=\\"^^(A^|The) $\\"" ^<%^^[yyyy^|yy]-MM-DD%^>'
   *
   * // @NOTE If you include standard output, divide the array elements
   * Wsh.OS.joinCmdArgs(['ping.exe', 'localhost', '>', 'C:\\My Logs\\stdout.log]);
   * // Or, do not put space in between the dest and quoted it
   * Wsh.OS.joinCmdArgs(['ping.exe', 'localhost', '1>"C:\\My Logs\\stdout.log"]);
   * @function joinCmdArgs
   * @memberof Wsh.OS
   * @param {string[]|string} args - The arguments to convert. If args is String, return this string value.
   * @param {object} [options] - Optional parameters.
   * @param {boolean} [options.escapes=true] - Escapes arguments for CMD.
   * @returns {string} - The string of escaped and joined.
   */
  os.joinCmdArgs = function (args, options) {
    if (isSolidString(args)) return String(args);
    if (!isSolidArray(args)) return '';

    var escapes = obtain(options, 'escapes', true);

    var argsStr = args.reduce(function (acc, arg) {
      var str = escapes ? _escapeForCmd(arg) : arg;

      return acc + ' ' + os.surroundCmdArg(str);
    }, '');

    return argsStr.trim();
  };
  var _joinCmdArgs = os.joinCmdArgs;
  // }}}

  // os.convToCmdlineStr {{{
  /**
   * @typedef {object} typeConvToCommandOptions
   * @property {boolean} [shell=false] - Wraps with CMD.EXE
   * @property {boolean} [closes=true] - /C (Close?) or /K (Keep?)
   * @property {boolean} [escapes=true] - Escapes the arguments.
   */

  /**
   * Converts the command and arguments to a command line string.
   *
   * @example
   * var os = Wsh.OS;
   *
   * os.convToCmdlineStr('net', ['use', '/delete']);
   * // Returns: 'net user /delete'
   *
   * os.convToCmdlineStr('net', ['use', '/delete'], { shell: true });
   * // Returns: 'C:\Windows\System32\cmd.exe /S /C"net user /delete"'
   *
   * os.convToCmdlineStr('net', 'use /delete', { shell: true, closes: false });
   * // Returns: 'C:\Windows\System32\cmd.exe /S /K"net user /delete"'
   *
   * // The 2nd argument: Array vs String
   * // Array is escaped
   * os.convToCmdlineStr('D:\\My Apps\\app.exe',
   *   ['/RegExp="^(A|The) $"', '-f', 'C:\\My Data\\img.doc']);
   * // Returns: '"D:\\My Apps\\app.exe" "/RegExp=\\"^^(A^|The) $\\"" -f "C:\\My Data\\img.doc"'
   *
   * // String is not escaped
   * os.convToCmdlineStr('D:\\My Apps\\app.exe',
   *   '/RegExp="^(A|The) $" -f C:\\My Data\\img.doc');
   * // Returns: '/RegExp="^(A|The) $" -f C:\\My Data\\img.doc'
   * @function convToCmdlineStr
   * @memberof Wsh.OS
   * @param {string} cmdStr - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments to format.
   * @param {typeConvToCommandOptions} [options] - Optional parameters.
   * @returns {string} - The converted command.
   */
  os.convToCmdlineStr = function (cmdStr, args, options) {
    var FN = 'os.convToCmdlineStr';
    if (!isSolidString(cmdStr)) throwErrNonStr(FN, cmdStr);

    var command = _srrPath(cmdStr);

    var argsStr;
    if (isArray(args)) {
      argsStr = _joinCmdArgs(args, options);
    } else if (isString(args)) {
      argsStr = args;
    } else {
      argsStr = '';
      // throw new Error('Error [ERR_INVALID_ARG_TYPE]\n'
      //   + '  at ' + FN + ' (' + MODULE_TITLE + ')\n'
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

  // _shRun {{{
  /**
   * @typedef {typeOsExecOptions} typeShRunOptions
   * @property {(number|string)} [winStyle="activeDef"] - See {@link https://docs.tuckn.net/WshUtil/Wsh.Constants.windowStyles.html|Wsh.Constants.windowStyles}.
   * @property {boolean} [waits=true] - See {@link https://docs.tuckn.net/WshUtil/Wsh.Constants.waits.html|Wsh.Constants.waits}.
   */

  /**
   * @private
   * @function _shRun
   * @memberof Wsh.OS
   * @param {string} command - The command and arguments.
   * @param {typeShRunOptions} [options] - Optional parameters.
   * @returns {number|string} - Return a number except when options.isDryRun is true.
   */
  function _shRun (command, options) {
    var FN = '_shRun';
    if (!isSolidString(command)) throwErrNonStr(FN, command);

    var isDryRun = obtain(options, 'isDryRun', false);
    if (isDryRun) return 'dry-run [' + FN + ']: ' + command;

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
        + '  at ' + FN + ' (' + MODULE_TITLE + ')\n'
        + '  command: ' + command);
    }
  } // }}}

  // os.shRun {{{
  /**
   * Asynchronously runs the application with WScript.Shell.Run. {@link https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364421%28v%3dmsdn.10%29|WSH Run method}
   *
   * @example
   * var os = Wsh.OS;
   * var fso = Wsh.FileSystemObject;
   *
   * // Ex1. CMD-command
   * // Throws a Error. `mkdir` is the CMD command
   * os.shRun('mkdir', 'D:\\Temp1');
   *
   * // With the options
   * os.shRun('mkdir', 'D:\\Temp1', { shell: true, winStyle: 'hidden' });
   * // Returns: Always 0
   *
   * while (!fso.FolderExists('D:\\Temp1')) { // Waiting the created
   *   WScript.Sleep(300);
   * }
   * fso.FolderExists('D:\\Temp1'); // Returns: true
   *
   * // Ex2. Exe-file
   * os.shRun('notepad.exe', ['D:\\Test2.txt'], { winStyle: 'activeMax' });
   * // Returns: Always 0
   *
   * // Ex.3 Dry Run
   * var cmd = os.shRun('mkdir', 'D:\\Temp3', { shell: true, isDryRun: true });
   * console.log(cmd);
   * // Outputs:
   * // dry-run [_shRun]: C:\Windows\System32\cmd.exe /S /C"mkdir D:\Temp3"
   * fso.FolderExists('D:\\Temp3'); // Returns: false
   * @function run
   * @memberof Wsh.OS
   * @param {string} cmdStr - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments.
   * @param {typeShRunOptions} [options] - Optional parameters.
   * @returns {number|string} - Returns 0 except when options.isDryRun is true.
   */
  os.shRun = function (cmdStr, args, options) {
    // var FN = 'os.shRun';

    var command = os.convToCmdlineStr(cmdStr, args, options);

    return _shRun(command, objAdd({}, options, { waits: false }));
  }; // }}}

  // os.shRunSync {{{
  /**
   * Synchronize of {@link Wsh.OS.shRun}
   *
   * @example
   * var os = Wsh.OS;
   * var fso = Wsh.FileSystemObject;
   *
   * // Ex1. CMD-command
   * // Throws a Error. `mkdir` is the CMD command
   * os.shRunSync('mkdir', 'D:\\Temp1');
   *
   * // With the options
   * var retVal1 = os.shRunSync('mkdir', 'D:\\Temp1', {
   *   shell: true, winStyle: 'hidden'
   * });
   *
   * console.log(retVal1); // Outputs the number of the result code
   * fso.FolderExists('D:\\Temp1'); // Returns: true
   *
   * // Ex2. Exe-file
   * var retVal2 = os.shRunSync('notepad.exe', ['D:\\Test2.txt'], {
   *   winStyle: 'activeMax'
   * });
   *
   * // Waits until the notepad process is finished
   *
   * console.log(retVal2); // Outputs the number of the result code
   *
   * // Ex.3 Dry Run
   * var cmd = os.shRunSync('mkdir', 'D:\\Temp3', { shell: true, isDryRun: true });
   * console.log(cmd);
   * // Outputs:
   * // dry-run [_shRun]: C:\Windows\System32\cmd.exe /S /C"mkdir D:\Temp3"
   * fso.FolderExists('D:\\Temp3'); // Returns: false
   * @function runSync
   * @memberof Wsh.OS
   * @param {string} cmdStr - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments.
   * @param {object} [options] - Optional parameters. See [typeConvToCommandOptions]{@link https://docs.tuckn.net/WshOS/global.html#typeConvToCommandOptions}  and [typeShRunOptions]{@link https://docs.tuckn.net/WshOS/docs/global.html#typeShRunOptions}.
   * @returns {number|string} - Returns code from the app except when options.isDryRun is true.
   */
  os.shRunSync = function (cmdStr, args, options) {
    // var FN = 'os.shRunSync';

    var command = os.convToCmdlineStr(cmdStr, args, options);
    return _shRun(command, objAdd({}, options, { waits: true }));
  }; // }}}

  /**
   * @typedef {typeConvToCommandOptions} typeOsExecOptions
   * @property {boolean} [isDryRun=false] - No execute, returns the string of command.
   * @see In addition to the above, [typeConvToCommandOptions]{@link https://docs.tuckn.net/WshOS/global.html#typeConvToCommandOptions} can also be specified.
   */

  // os.shExec {{{
  /**
   * The object returning from Wsh.OS.shExec ({@link https://msdn.microsoft.com/ja-jp/library/cc364375.aspx|WshScriptExec object}). When `options.shell: true` is specified, exitCode may not be accurate because this value is returned by CMD.
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
   * os.shExec('mkdir', 'D:\\Temp');
   *
   * // With the `shell` option
   * var oExec1 = os.shExec('mkdir', 'D:\\Temp1', { shell: true });
   * console.log(oExec1.ExitCode); // 0 (always)
   *
   * while (oExec1.Status == 0) WScript.Sleep(300); // Waiting the finished
   * console.log(oExec1.Status); // 1 (It means finished)
   * fso.FolderExists('D:\\Temp1'); // Returns: true
   *
   * // Ex.2 Exe-file
   * var oExec2 = os.shExec('ping.exe', ['127.0.0.1']);
   * console.log(oExec2.ExitCode); // 0 (always)
   *
   * while (oExec2.Status == 0) WScript.Sleep(300); // Waiting the finished
   * console.log(oExec2.Status); // 1 (It means finished)
   *
   * var result = oExec2.StdOut.ReadAll();
   * console.log(result); // Outputs the result of ping 127.0.0.1
   *
   * // Ex.3 Dry Run
   * var cmd = os.shExec('mkdir', 'D:\\Temp3', { shell: true, isDryRun: true });
   * console.log(cmd);
   * // Outputs:
   * // dry-run [os.shExec]: C:\Windows\System32\cmd.exe /S /C"mkdir D:\Temp3"
   * fso.FolderExists('D:\\Temp3'); // Returns: false
   * @function exec
   * @memberof Wsh.OS
   * @param {string} cmdStr - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments.
   * @param {typeOsExecOptions} [options] - Optional parameters.
   * @returns {typeExecObject|string} - If options.isDryRun is true, returns string.
   */
  os.shExec = function (cmdStr, args, options) {
    var FN = 'os.shExec';
    if (!isSolidString(cmdStr)) throwErrNonStr(FN, cmdStr);

    var command = os.convToCmdlineStr(cmdStr, args, options);
    if (!isSolidString(command)) throwErrNonStr(FN, command);

    var isDryRun = obtain(options, 'isDryRun', false);
    if (isDryRun) return 'dry-run [' + FN + ']: ' + command;

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
        + '  at ' + FN + ' (' + MODULE_TITLE + ')\n'
        + '  command: ' + command);
    }
  }; // }}}

  // os.shExecSync {{{
  /**
   *  The object returning from Wsh.OS.shExecSync. When `options.shell: true` is specified, exitCode may not be accurate because this value is returned by CMD.
   *
   * @typedef {object} typeExecSyncReturn
   * @property {number} exitCode - There are applications that return 1 (NG) even if the processing is successful.
   * @property {string} command - The executed command line
   * @property {string} stdout
   * @property {string} stderr
   * @property {boolean} error
   */

  /**
   * Synchronize of {@link Wsh.OS.shExec}
   *
   * @example
   * var os = Wsh.OS;
   *
   * // Ex1. CMD-command
   * // Throws a Error. `mkdir` is the CMD command
   * os.shExecSync('mkdir', 'D:\\Temp');
   *
   * // With the `shell` option
   * var retObj1 = os.shExecSync('mkdir', 'D:\\Temp', { shell: true });
   * console.dir(retObj1);
   * // Outputs: {
   * //   exitCode: 0,
   * //   stdout: "",
   * //   stderr: "",
   * //   error: false };
   *
   * // Ex2. Exe-file
   * var retObj2 = os.shExecSync('ping.exe', ['127.0.0.1']);
   * console.dir(retObj2);
   * // Outputs: {
   * //   exitCode: 0,
   * //   stdout: <The result of ping 127.0.0.1>,
   * //   stderr: "",
   * //   error: false };
   *
   * // Ex.3 Dry Run
   * var cmd = os.shExecSync('ping.exe', '127.0.0.1', { isDryRun: true });
   * console.log(cmd);
   * // Outputs:
   * // dry-run [os.shExecSync]: C:\Windows\System32\cmd.exe /S /C"C:\Windows\System32\PING.EXE 127.0.0.1"
   * @function execSync
   * @memberof Wsh.OS
   * @param {string} cmdStr - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments.
   * @param {typeOsExecOptions} [options] - Optional parameters.
   * @returns {typeExecSyncReturn|string} - If options.isDryRun is true, returns string.
   */
  os.shExecSync = function (cmdStr, args, options) {
    var FN = 'os.shExecSync';
    if (!isSolidString(cmdStr)) throwErrNonStr(FN, cmdStr);

    var command = os.convToCmdlineStr(cmdStr, args, options);
    if (!isSolidString(command)) throwErrNonStr(FN, command);

    var isDryRun = obtain(options, 'isDryRun', false);
    if (isDryRun) return 'dry-run [' + FN + ']: ' + command;

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
        + '  at ' + FN + ' (' + MODULE_TITLE + ')\n'
        + '  command: ' + command);
    }

    return {
      command: command,
      exitCode: oExec.ExitCode,
      stdout: stdout,
      stderr: stderr,
      error: isSolidString(stderr)
    };
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
      var iRetVal = os.shRunSync(os.exefiles.net, ['session'], {
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
   * @param {typeShRunOptions} [options] - Optional parameters.
   * @returns {void|string} - Returns undefined except when options.isDryRu is =true.
   */
  os.runAsAdmin = function (cmdStr, args, options) {
    if (os.isAdmin()) return os.shRun(cmdStr, args, options);

    var FN = 'os.runAsAdmin';
    if (!isSolidString(cmdStr)) throwErrNonStr(FN, cmdStr);

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

    var isDryRun = obtain(options, 'isDryRun', false);
    if (isDryRun) return 'dry-run [' + FN + ']: ' + exePath + ' ' + argsStr;

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
