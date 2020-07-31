/* globals Wsh: false */

(function () {
  // Shorthands
  var CD = Wsh.Constants;
  var util = Wsh.Util;
  var fso = Wsh.FileSystemObject;
  var path = Wsh.Path;

  var objAdd = Object.assign;
  var insp = util.inspect;
  var isPureNumber = util.isPureNumber;
  var isString = util.isString;
  var obtain = util.obtainPropVal;

  var os = Wsh.OS;
  var WSCRIPT = os.exefiles.wscript;

  /** @constant {string} */
  var MODULE_TITLE = 'WshOS/TaskScheduler.js';

  var throwErrNonStr = function (functionName, typeErrVal) {
    util.throwTypeError('string', MODULE_TITLE, functionName, typeErrVal);
  };

  /**
   * The wrapper object for functions to handle Windows TaskScheduler.
   *
   * @namespace Task
   * @memberof Wsh.OS
   */
  os.Task = {};

  // os.Task.create {{{
  /**
   * Creates the scheduled task. To see API run "C:\>SchTasks.exe /Create /?" or {@link https://technet.microsoft.com/ja-jp/library/cc725744(v=ws.10).aspx|Schtasks create}.
   *
   * @example
   * Wsh.OS.Task.create('MyTask', 'wscript.exe', '//job:run my-task.wsf');
   * @function create
   * @memberof Wsh.OS.Task
   * @param {string} taskName - The Task name to create.
   * @param {string} cmdStr - The command to execute.
   * @param {(string[]|string)} [args] - The arguments for the command.
   * @param {typeShRunOptions} [options] - Optional Parameters.
   * @throws {string} - If an error occurs during command execution, or if the command exits with a value other than 0.
   * @throws {string} - If an error occurs during command execution, or if the command exits with a value other than 0.
   * @returns {void|string} - If isDryRun is true, returns string.
   */
  os.Task.create = function (taskName, cmdStr, args, options) {
    var FN = 'os.Task.create';
    if (!isString(taskName)) throwErrNonStr(FN, taskName);

    /**
     * タスク登録時、↓ /ST 00:00 などにすると、StdErrに"警告: /ST が現時刻よりも早いため、タスクは実行されない可能性があります。"と出力されダルいので23:59にする
     * エラー: '/TR' のオプションの値を 261 文字より多い文字で指定することはできません。
     */
    var mainCmd = os.exefiles.schtasks;
    var command = os.convToCmdCommand(cmdStr, args, options);
    var argsStr = '/Create /F /TN "' + taskName + '"'
      + ' /SC ONCE /ST 23:59 /IT /RL LIMITED /TR "'
      + command.replace(/"/g, '\\"') + '"';
    // ↑CMD /C"..."とは異なり、"は\"でエスケープしなければならない

    try {
      /**
SchTasks.exe /Create /F /TN myTask /SC ONCE /ST 23:59 /IT /RL LIMITED /TR "C:\myBat.bat\"
  exitCode: 0,
  stdout: "成功: スケジュール タスク "myTask" は正しく作成されました。
  stderr: ""
       */
      var retVal = os.runSync(
        mainCmd,
        argsStr,
        objAdd({}, options, {
          shell: false,
          escapes: false,
          winStyle: 'hidden'
        })
      );

      var isDryRun = obtain(options, 'isDryRun', false);
      if (isDryRun) return 'dry-run [' + FN + ']: ' + retVal;

      if (retVal === CD.runs.ok) return;

      throw new Error('Error [ExitCode is not Ok] "' + retVal + '"\n');
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + FN + ' (' + MODULE_TITLE + ')\n'
        + '  mainCmd: "' + mainCmd + '"\n  argsStr: "' + argsStr + '"');
    }
  }; // }}}

  // os.Task.exists {{{
  /**
   * Checks the scheduled task existing. To see API run "C:\>SchTasks.exe /Query /?" or {@link https://technet.microsoft.com/ja-jp/library/cc725744(v=ws.10).aspx|Schtasks}.
   *
   * @example
   * Wsh.OS.Task.exists('MyTask'); // Returns: true
   * @function exists
   * @memberof Wsh.OS.Task
   * @param {string} taskName - The Task name.
   * @param {object} [options] - Optional Parameters.
   * @param {boolean} [options.isDryRun=false] - No execute, returns the string of command.
   * @returns {boolean|string} - If the task is existing returns true. If isDryRun is true, returns string.
   */
  os.Task.exists = function (taskName, options) {
    var FN = 'os.Task.exists';
    if (!isString(taskName)) throwErrNonStr(FN, taskName);

    var mainCmd = os.exefiles.schtasks;
    var args = ['/Query', '/XML', '/TN', taskName];
    var isDryRun = obtain(options, 'isDryRun', false);

    try {
      /**
SchTasks.exe /Query /XML /TN myTask
[Success]
  exitCode: 0,
  stdout: "<?xml version="1.0" encoding="UTF-16"?> ..."
  stderr: ""
[Fail]
  exitCode: 1
  stdout: ""
  stderr: "エラー: 指定されたファイルが見つかりません。"
       */
      var retVal = os.runSync(mainCmd, args, {
        shell: false,
        winStyle: 'hidden',
        isDryRun: isDryRun
      });

      if (isDryRun) return 'dry-run [' + FN + ']: ' + retVal;

      return retVal === CD.runs.ok;
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + FN + ' (' + MODULE_TITLE + ')');
    }
  }; // }}}

  // os.Task.run {{{
  /**
   * Runs the scheduled task. To see API run `C:\>SchTasks.exe /Run /?` or {@link https://technet.microsoft.com/ja-jp/library/cc725744(v=ws.10).aspx|Schtasks}.
   *
   * @example
   * Wsh.OS.Task.run('MyTask');
   * @function run
   * @memberof Wsh.OS.Task
   * @param {string} taskName - The Task name
   * @param {object} [options] - Optional Parameters.
   * @param {boolean} [options.isDryRun=false] - No execute, returns the string of command.
   * @returns {void|string} - If isDryRun is true, returns string.
   */
  os.Task.run = function (taskName, options) {
    var FN = 'os.Task.run';
    if (!isString(taskName)) throwErrNonStr(FN, taskName);

    var mainCmd = os.exefiles.schtasks;
    var args = ['/Run', '/I', '/TN', taskName];
    var isDryRun = obtain(options, 'isDryRun', false);

    /**
     * スケジュール実行は即座に0を返すので、正常に起動したかどうかは判断できない。sh.Runのwaitの指定は意味がない。StdOutを取得するなど別の方法が必要
SchTasks.exe /Query /XML /TN myTask
[Success]
  exitCode: 0,
  stdout: "情報: スケジュール タスク "myTask" は現在実行中です。
      成功: スケジュール タスク "myTask" の実行が試行されました。"
  stderr: ""
     */
    try {
      var retVal = os.run(mainCmd, args, {
        shell: false,
        winStyle: 'hidden',
        isDryRun: isDryRun
      });

      if (isDryRun) return 'dry-run [' + FN + ']: ' + retVal;

      return;
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + FN + ' (' + MODULE_TITLE + ')');
    }
  }; // }}}

  // os.Task.del {{{
  /**
   * Deletes the scheduled task. To see API run "C:\>SchTasks.exe /Delete /?" or {@link https://technet.microsoft.com/ja-jp/library/cc725744(v=ws.10).aspx|Schtasks}. I wanted to name "delete" this method, but on JScript, can not use .delete as Object property name.
   *
   * @example
   * Wsh.OS.Task.del('MyTask');
   * @function del
   * @memberof Wsh.OS.Task
   * @param {string} taskName - The Task name.
   * @param {object} [options] - Optional Parameters.
   * @param {boolean} [options.isDryRun=false] - No execute, returns the string of command.
   * @throws {string} - If an error occurs during command execution, or if the command exits with a value other than 0.
   * @returns {void|string} - If isDryRun is true, returns string.
   */
  os.Task.del = function (taskName, options) {
    var FN = 'os.Task.del';
    if (!isString(taskName)) throwErrNonStr(FN, taskName);

    if (!os.Task.exists(taskName)) return;

    var mainCmd = os.exefiles.schtasks;
    var args = ['/Delete', '/F', '/TN', taskName];
    var isDryRun = obtain(options, 'isDryRun', false);

    try {
      /**
SchTasks.exe /Delete /F /TN myTask
[Success]
  exitCode: 0,
  stdout: "成功: スケジュール タスク "myTask" は正しく削除されました。
  stderr: ""
[Fail]
  exitCode: 1,
  stdout: "",
  stderr: "エラー: 指定されたファイルが見つかりません。"
       */
      var retVal = os.runSync(mainCmd, args, {
        shell: false,
        winStyle: 'hidden',
        isDryRun: isDryRun
      });

      if (isDryRun) return 'dry-run [' + FN + ']: ' + retVal;

      if (retVal === CD.runs.ok) return;

      throw new Error('Error [ExitCode is not Ok] "' + retVal + '"\n');
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + FN + ' (' + MODULE_TITLE + ')');
    }
  }; // }}}

  // os.Task.ensureToDelete {{{
  /**
   * Deletes the scheduled task. If the deletion fails, it will retry to delete for the specified number of seconds.
   *
   * @example
   * Wsh.OS.Task.ensureToDelete('MyTask', 5000); // Retry for 5sec
   * @function ensureToDelete
   * @memberof Wsh.OS.Task
   * @param {string} taskName - The task name.
   * @param {number} [msecTimeOut=10000] - default: 10sec.
   * @param {object} [options] - Optional Parameters.
   * @param {boolean} [options.isDryRun=false] - No execute, returns the string of command.
   * @returns {void|string} - If isDryRun is true, returns string.
   */
  os.Task.ensureToDelete = function (taskName, msecTimeOut, options) {
    var FN = 'os.Task.ensureToDelete';
    if (!isString(taskName)) throwErrNonStr(FN, taskName);

    if (!os.Task.exists(taskName)) return;

    msecTimeOut = isPureNumber(msecTimeOut) ? msecTimeOut : 10000;

    do {
      try {
        var retVal = os.Task.del(taskName, options);

        var isDryRun = obtain(options, 'isDryRun', false);
        if (isDryRun) return 'dry-run [' + FN + ']: ' + retVal;

        return;
      } catch (e) {
        WScript.Sleep(100);
        msecTimeOut -= 100;
      }
    } while (msecTimeOut > 0);

    throw new Error('Error [Delete Task] "' + taskName + '"\n'
      + '  at ' + FN + ' (' + MODULE_TITLE + ')');
  }; // }}}

  // os.Task.ensureToCreate {{{
  /**
   *
   * Creates the scheduled task. If the task is already existing, delete this. and If the creation fails, it will retry to create for the specified number of seconds.
   *
   * @example
   * var createTask = Wsh.OS.Task.ensureToCreate; // Shorthand
   *
   * createTask('MyTask', 'wscript.exe', '//job:run my-task.wsf', 5000);
   * @function ensureToCreate
   * @memberof Wsh.OS.Task
   * @param {string} taskName - The task name to create.
   * @param {string} mainCmd - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments for the command.
   * @param {typeShRunOptions} [options] - Optional Parameters.
   * @param {number} [msecTimeOut=10000] - default: 10sec.
   * @param {boolean} [options.isDryRun=false] - No execute, returns the string of command.
   * @returns {void|string} - If isDryRun is true, returns string.
   */
  os.Task.ensureToCreate = function (
    taskName,
    mainCmd,
    args,
    options,
    msecTimeOut
  ) {
    var FN = 'os.Task.ensureToCreate';
    if (!isString(taskName)) throwErrNonStr(FN, taskName);

    // If existing, delete the task
    if (os.Task.exists(taskName)) os.Task.ensureToDelete(taskName);

    os.Task.create(taskName, mainCmd, args, options);

    // Check the task to exist
    msecTimeOut = isPureNumber(msecTimeOut) ? msecTimeOut : 10000;

    do {
      try {
        var retVal = os.Task.exists(taskName, options);

        var isDryRun = obtain(options, 'isDryRun', false);
        if (isDryRun) return 'dry-run [' + FN + ']: ' + retVal;

        if (retVal) return;
      } catch (e) {
        WScript.Sleep(100);
        msecTimeOut -= 100;
      }
    } while (msecTimeOut > 0);

    throw new Error('Error: [Create the task(TimeOut)] ' + taskName + '\n'
      + '  at ' + FN + ' (' + MODULE_TITLE + ')\n'
      + '  mainCmd: ' + mainCmd + '\n  args: ' + insp(args));
  }; // }}}

  // os.Task.runTemporary {{{
  /**
   * Create a task for the command, execute it, and then delete it. The advantage is that the command is always executed with Medium-WIL authority.
   *
   * @example
   * var runOnce = Wsh.OS.Task.runTemporary; // Shorthand
   *
   * runOnce('wscript.exe', '//job:run my-task.wsf');
   * @function runTemporary
   * @memberof Wsh.OS.Task
   * @param {string} cmdStr - The executable file path or The command of Command-Prompt.
   * @param {(string[]|string)} [args] - The arguments for the command.
   * @param {typeShRunOptions} [options] - Optional Parameters.
   * @returns {void}
   */
  os.Task.runTemporary = function (cmdStr, args, options) {
    var FN = 'os.Task.runTemporary';
    if (!isString(cmdStr)) throwErrNonStr(FN, cmdStr);

    var command = os.convToCmdCommand(cmdStr, args, options);

    // Write a temporary JScript
    // @note SchTask.exeには261文字しか渡せないので、WSHでラップする
    var tmpCode = 'WScript.CreateObject(\'WScript.Shell\').Run(\''
        + command.replace(/\\{1}/g, '\\\\')
        + '\', ' + String(CD.windowStyles.hidden)
        + ', ' + String(CD.notWait) + ');';
    var tmpJsPath = os.writeTempText(tmpCode, '.js');

    var taskName = 'Task_' + path.basename(os.makeTmpPath())
      + '_' + util.createDateString();

    var isDryRun = obtain(options, 'isDryRun', false);
    var retLog = '';
    var retVal;

    retVal = os.Task.ensureToCreate(taskName, WSCRIPT, [tmpJsPath], options);
    if (isDryRun) retLog = 'dry-run [' + FN + ']: ' + retVal;

    retVal = os.Task.run(taskName, options);
    if (isDryRun) retLog += '\n' + retVal;

    retVal = '\n' + os.Task.ensureToDelete(taskName, options);
    if (isDryRun) retLog += '\n' + retVal;

    // console.log(tmpJsPath); // Debug
    fso.DeleteFile(tmpJsPath, CD.fso.force.yes);

    if (isDryRun) return retLog;
    return;
  }; // }}}

  // os.Task.createAsAdmin {{{
  /**
   * TODO 一度管理者権限でタスクに登録すれば、次回からUACの確認なしで実行可能。ただしタスクを保存しておく必要があり、環境も汚す
   *
   * @function createAsAdmin
   * @memberof Wsh.OS.Task
   */
  os.Task.createAsAdmin = function () {
    // @todo W.I.P
  }; // }}}
})();

// vim:set foldmethod=marker commentstring=//%s :
