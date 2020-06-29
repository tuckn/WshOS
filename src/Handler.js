/* globals Wsh: false */

(function () {
  // Shorthands
  var sh = Wsh.Shell;
  var CD = Wsh.Constants;
  var util = Wsh.Util;
  var fso = Wsh.FileSystemObject;

  var insp = util.inspect;
  var isPureNumber = util.isPureNumber;
  var isString = util.isString;
  var isSolidString = util.isSolidString;

  var os = Wsh.OS;

  /** @constant {string} */
  var MODULE_TITLE = 'WshOS/Handler.js';

  var throwErrNonStr = function (functionName, typeErrVal) {
    util.throwTypeError('string', MODULE_TITLE, functionName, typeErrVal);
  };

  // Drive

  // os.isExistingDrive {{{
  /**
   * Checks if the drive existing.
   *
   * @example
   * Wsh.OS.isExistingDrive('D');
   * // Returns: true
   * // Note that 'D:\' can not work.
   * @function isExistingDrive
   * @memberof Wsh.OS
   * @param {string} driveLetter - The drive letter to check.
   * @returns {boolean} - If the drive is existing, returns true.
   */
  os.isExistingDrive = function (driveLetter) {
    var functionName = 'os.isExistingDrive';
    if (!isString(driveLetter)) throwErrNonStr(functionName, driveLetter);

    return fso.DriveExists(driveLetter);
  }; // }}}

  // os.getDrivesInfo {{{
  /**
   * Gets the Drives-collection of all Drive objects available on the local machine. See {@link https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/drives-property|Drives property} and {@link https://docs.tuckn.net/WshUtil/Wsh.Constants.driveTypes.html|Wsh.Constants.driveTypes}.}
   *
   * @example
   * Wsh.OS.getDrivesInfo();
   * // Returns: [
   * //   0: {
   * //     letter: "C",
   * //     name: "Windows",
   * //     type: 2
   * //   },
   * //   1: {
   * //     letter: "D",
   * //     name: "BackUp",
   * //     type: 2
   * //   },
   * //   2: {
   * //     letter: "Q",
   * //     name: "CD-ROM",
   * //     type: 4
   * //   },
   * //   3: {
   * //     letter: "S",
   * //     name: "microSD",
   * //     type: 1
   * //   } ]
   * @function getDrivesInfo
   * @memberof Wsh.OS
   * @returns {Array} - The drive info.
   */
  os.getDrivesInfo = function () {
    var drvs = new Enumerator(fso.Drives);
    var drv;
    var rtnDrvs = [];

    while (!drvs.atEnd()) {
      drv = drvs.item();

      if (drv.IsReady) {
        rtnDrvs.push({
          letter: drv.DriveLetter,
          name: drv.VolumeName,
          type: drv.DriveType
        });
      }

      drvs.moveNext();
    }

    return rtnDrvs;
  }; // }}}

  // os.getTheLastFreeDriveLetter {{{
  /**
   * Gets the last Drive-letter to be able to assign.
   *
   * @example
   * Wsh.OS.getTheLastFreeDriveLetter(); // Returns: 'Y'
   * @function getTheLastFreeDriveLetter
   * @memberof Wsh.OS
   * @returns {string} - The last drive letter to be able to assign.
   */
  os.getTheLastFreeDriveLetter = function () {
    var driveLetter;
    var freeDriveLetter = '';

    // 後ろ(Z)から取得開始
    for (var i = 26, I = 2; i >= I; i--) {
      driveLetter = String.fromCharCode(64 + i);

      if (os.isExistingDrive(driveLetter)) {
        continue;
      } else {
        freeDriveLetter = driveLetter;
        break;
      }
    }

    return freeDriveLetter;
  }; // }}}

  // os.getHardDriveLetters {{{
  /**
   * Gets the Drives-collection of all Hard-Drive objects available on the local machine.
   *
   * @example
   * Wsh.OS.getHardDriveLetters();
   * // Returns: [
   * //   0: {
   * //     letter: "C",
   * //     name: "Windows",
   * //     type: 2
   * //   },
   * //   1: {
   * //     letter: "D",
   * //     name: "BackUp",
   * //     type: 2
   * //   } ]
   * @function getHardDriveLetters
   * @memberof Wsh.OS
   * @returns {Array} - The Hard-drives info.
   */
  os.getHardDriveLetters = function () {
    var drvsInfo = os.getDrivesInfo();

    return drvsInfo.filter(function (val) {
      return val.type === CD.driveTypes.fixed;
    });
  }; // }}}

  // os.assignDriveLetter {{{
  /**
   * Assigns the drive letter to the directory.
   *
   * @example
   * var assignLetter = Wsh.OS.assignDriveLetter; // Shorthand
   *
   * assignLetter('C:\\My Data', 'X');
   * // Return 'X'
   *
   * assignLetter('\\\\MyNAS\\MultiMedia', 'M', 'myname', 'mY-p@ss');
   * // Return 'M'
   * @function assignDriveLetter
   * @memberof Wsh.OS
   * @param {string} dirPath - The path of directory to assign.
   * @param {string} [letter] - The drive letter.
   * @param {string} [userName] - The user name.
   * @param {string} [password] - The password.
   * @returns {string} - The assigned letter.
   */
  os.assignDriveLetter = function (dirPath, letter, userName, password) {
    var functionName = 'os.assignDriveLetter';
    if (!isString(dirPath)) throwErrNonStr(functionName, dirPath);
    if (!isString(letter) && !letter) throwErrNonStr(functionName, letter);

    if (!isSolidString(letter)) letter = os.getTheLastFreeDriveLetter();

    var uncPath = dirPath;
    if (/^[a-z]:/i.test(dirPath)) {
      uncPath = dirPath.replace(/^([a-z]):/i, '\\\\' + os.hostname() + '\\$1$');
    }

    var mainCmd = os.exefiles.net;
    var args = ['use', letter + ':', uncPath];

    if (isSolidString(userName) && isSolidString(password)) {
      args = args.concat(['/user:' + userName, password]);
    }

    try {
      var iRetVal = os.runSync(mainCmd, args, { winStyle: 'hidden' });

      if (iRetVal !== CD.runs.ok) {
        throw new Error('Error [ExitCode is not Ok] "' + iRetVal + '"\n');
      }
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + functionName + ' (' + MODULE_TITLE + ')');
    }

    return letter;
  }; // }}}

  // os.removeAssignedDriveLetter {{{
  /**
   * Removes the assigning drive letter.
   *
   * @example
   * Wsh.OS.removeAssignedDriveLetter('Y');
   * @function removeAssignedDriveLetter
   * @memberof Wsh.OS
   * @param {string} letter - The assigning drive letter to remove.
   * @throws {string} - If an error occurs during command execution, or if the command exits with a value other than 0.
   * @returns {void}
   */
  os.removeAssignedDriveLetter = function (letter) {
    var functionName = 'os.removeAssignedDriveLetter';
    if (!isString(letter)) throwErrNonStr(functionName, letter);

    var mainCmd = os.exefiles.net;
    var args = ['use', letter + ':', '/delete'];

    try {
      var iRetVal = os.runSync(mainCmd, args, {
        shell: false,
        winStyle: 'hidden'
      });
      if (iRetVal === CD.runs.ok) return;

      throw new Error('Error [ExitCode is not Ok] "' + iRetVal + '"\n');
    } catch (e) {
      /*
       * net use Z:  /delete
       *  ExitCode: 2,
       *  StdOut: ""
       *  StdErr: "ネットワーク接続が見つかりませんでした。
       *    NET HELPMSG 2250 と入力すると、より詳しい説明が得られます。"
       */
      throw new Error(insp(e) + '\n'
        + '  at ' + functionName + ' (' + MODULE_TITLE + ')');
    }
  }; // }}}

  // Process

  // os.getProcessIDs {{{
  /**
   * Gets processes IDs that the specified process string.
   *
   * @example
   * var getPIDs = Wsh.OS.getProcessIDs; // Shorthand
   *
   * var pIDs = getPIDs('Chrome.exe');
   * // Returns: [33221, 22044, 43113, 42292, 17412]
   *
   * var pIDs = getPIDs('C:\\Program Files\\Git\\bin\\git.exe');
   * // Returns: [1732, 4316]
   * @function getProcessIDs
   * @memberof Wsh.OS
   * @param {(number|string)} processName - The PID or process name or full path.
   * @param {typeGetProcessesOptions} [options] - Optional parameters.
   * @returns {Array} - The Array of processes ID.
   */
  os.getProcessIDs = function (processName, options) {
    var functionName = 'os.getProcessIDs';
    if (!isPureNumber(processName) && !isString(processName)) {
      throwErrNonStr(functionName, processName);
    }

    var sWbemObjSets = os.WMI.getProcesses(processName, options);

    return sWbemObjSets.map(function (sWbemObjSet) {
      return sWbemObjSet.ProcessId;
    });
  }; // }}}

  // os.getProcessObjs {{{
  /**
   * Gets an Array of objects matching the specified process ID or string.
   *
   * @example
   * var getProcessObjs = Wsh.OS.getProcessObjs; // Shorthand
   *
   * var pObjs = getProcessObjs('chrome.exe');
   * // Returns: [
   * // {
   * //   Caption: 'chrome.exe',
   * //   CommandLine: ...
   * //   ...
   * // }, {
   * //   Caption: 'chrome.exe',
   * //   ...
   * // }, {
   * //   ...
   * // }]
   *
   * @function getProcessObjs
   * @memberof Wsh.OS
   * @param {(number|string)} processName - The PID or process name or full path.
   * @param {typeGetProcessesOptions} [options] - Optional parameters.
   * @returns {object[]} - The Array of objects of the process.
   */
  os.getProcessObjs = function (processName, options) {
    var functionName = 'os.getProcessObjs';
    if (!isPureNumber(processName) && !isString(processName)) {
      throwErrNonStr(functionName, processName);
    }

    var sWbemObjSets = os.WMI.getProcesses(processName, options);

    return os.WMI.toJsObjects(sWbemObjSets);
  }; // }}}

  // os.getProcessObj {{{
  /**
   * Gets an Object matching the specified process ID or string.
   *
   * @example
   * var getProcessObj = Wsh.OS.getProcessObj; // Shorthand
   *
   * var pObj = getProcessObj('wscript.exe');
   * // Returns: {
   * //   Caption: 'wscript.exe',
   * //   CommandLine: 'C:\\Windows\\System32\\wscript.exe //nologo C:\\MockGUI.wsf',
   * //   CreationClassName: 'Win32_Process',
   * //   CreationDate: '20200628110828.532484+540',
   * //   CSCreationClassName: 'Win32_ComputerSystem',
   * //   CSName: 'MYCOMPNAME',
   * //   Description: 'wscript.exe',
   * //   ExecutablePath: 'C:\\Windows\\System32\\wscript.exe',
   * //   ExecutionState: null,
   * //   Handle: '5678',
   * //   HandleCount: 241,
   * //   InstallDate: null,
   * //   KernelModeTime: '312500',
   * //   MaximumWorkingSetSize: 1380,
   * //   MinimumWorkingSetSize: 200,
   * //   Name: 'wscript.exe',
   * //   OSCreationClassName: 'Win32_OperatingSystem',
   * //   OSName: 'Microsoft Windows 10 Pro|C:\\WINDOWS|\\Device\\Harddisk0\\Partition4',
   * //   OtherOperationCount: '123',
   * //   OtherTransferCount: '4567',
   * //   PageFaults: 3462,
   * //   PageFileUsage: 3260,
   * //   ParentProcessId: 12628,
   * //   PeakPageFileUsage: 3504,
   * //   PeakVirtualSize: '2203497648128',
   * //   PeakWorkingSetSize: 12892,
   * //   Priority: 8,
   * //   PrivatePageCount: '3338240',
   * //   ProcessId: 7580,
   * //   QuotaNonPagedPoolUsage: 15,
   * //   QuotaPagedPoolUsage: 267,
   * //   QuotaPeakNonPagedPoolUsage: 15,
   * //   QuotaPeakPagedPoolUsage: 267,
   * //   ReadOperationCount: '45',
   * //   ReadTransferCount: '4631',
   * //   SessionId: 1,
   * //   Status: null,
   * //   TerminationDate: null,
   * //   ThreadCount: 11,
   * //   UserModeTime: '0',
   * //   VirtualSize: '2203497648128',
   * //   WindowsVersion: '10.0.12345',
   * //   WorkingSetSize: '13164544',
   * //   WriteOperationCount: '0',
   * //   WriteTransferCount: '0' }
   * @function getProcessObj
   * @memberof Wsh.OS
   * @param {(number|string)} processName - The PID or process name or full path.
   * @param {typeGetProcessesOptions} [options] - Optional parameters.
   * @returns {object} - The Object of the process.
   */
  os.getProcessObj = function (processName, options) {
    var functionName = 'os.getProcessObj';
    if (!isPureNumber(processName) && !isString(processName)) {
      throwErrNonStr(functionName, processName);
    }

    var sWbemObjSet = os.WMI.getProcess(processName, options);
    return os.WMI.toJsObject(sWbemObjSet);
  }; // }}}

  // os.activateProcess {{{
  /**
   * Activates the specified processes. {@link https://msdn.microsoft.com/ja-jp/library/cc364396.aspx|AppActivate}
   *
   * @example
   * Wsh.OS.activateProcess('Code.exe');
   * // Returns: 18321
   *
   * Wsh.OS.activateProcess('C:\\Windows\\System32\\cmd.exe');
   * // Returns: 12591
   * @function activateProcess
   * @memberof Wsh.OS
   * @param {(number|string)} processName - The PID or process name or full path.
   * @param {typeGetProcessesOptions} [options] - Optional parameters.
   * @returns {number} - The process ID to be activated.
   */
  os.activateProcess = function (processName, options) {
    var functionName = 'os.activateProcess';
    if (!isPureNumber(processName) && !isSolidString(processName)) {
      throwErrNonStr(functionName, processName);
    }

    var sWbemObjSet = os.WMI.getProcess(processName, options);
    if (!sWbemObjSet) return null;

    sh.AppActivate(sWbemObjSet.ProcessId);
    return sWbemObjSet.ProcessId;
  }; // }}}

  // os.terminateProcesses {{{
  /**
   * Terminates the specified processes.
   *
   * @example
   * Wsh.OS.terminateProcesses('cmd.exe');
   * // Terminates all cmd.exe processes.
   * @function terminateProcesses
   * @memberof Wsh.OS
   * @param {(number|string)} processName - The PID or process name or full path.
   * @param {typeGetProcessesOptions} [options] - Optional parameters.
   * @returns {void}
   */
  os.terminateProcesses = function (processName, options) {
    var functionName = 'os.terminateProcesses';
    if (!isPureNumber(processName) && !isSolidString(processName)) {
      throwErrNonStr(functionName, processName);
    }

    var sWbemObjSets = os.WMI.getProcesses(processName, options);
    if (sWbemObjSets.length === 0) return;

    sWbemObjSets.forEach(function (sWbemObjSet) { sWbemObjSet.Terminate(); });
  }; // }}}

  // os.getThisProcessID {{{
  /**
   * Gets this process ID.
   *
   * @example
   * Wsh.OS.getThisProcessID();
   * // Returns: 18321 (PID of cscript.exe or wscript.exe)
   * @function getThisProcessID
   * @memberof Wsh.OS
   * @returns {number} - This process ID.
   */
  os.getThisProcessID = function () {
    if (os._thisProcessID !== undefined) return os._thisProcessID;

    var thisProcess = os.WMI.getThisProcess();

    return thisProcess.ProcessId;
  }; // }}}

  // os.getThisParentProcessID {{{
  /**
   * Gets the parent-process ID of this process.
   *
   * @example
   * Wsh.OS.getThisParentProcessID();
   * // Returns: 5277
   * @function getThisParentProcessID
   * @memberof Wsh.OS
   * @returns {number} - The parent process ID.
   */
  os.getThisParentProcessID = function () {
    if (os._thisParentProcessID !== undefined) return os._thisParentProcessID;

    var thisProcess = os.WMI.getThisProcess();

    return thisProcess.ParentProcessId;
  }; // }}}

  // User

  // os.addUser {{{
  /**
   * Adds the new user in this local Windows.
   *
   * @example
   * Wsh.OS.addUser('MyUserName', 'mY-P@ss');
   * @function addUser
   * @memberof Wsh.OS
   * @param {string} userName - The user name to add.
   * @param {string} password - The user's password.
   * @throws {string} - If an error occurs during command execution, or if the command exits with a value other than 0.
   * @returns {void}
   */
  os.addUser = function (userName, password) {
    var functionName = 'os.addUser';
    if (!isString(userName)) throwErrNonStr(functionName, userName);

    var mainCmd = os.exefiles.net;
    var args = ['user', userName, password, '/ADD', '/Y'];

    try {
      /*
WScript.Shell.Exec("net user NewUserName UsersPassword /ADD /Y")
[成功時]
  ExitCode: 0,
  StdOut: "コマンドは正常に終了しました。"
  StdErr: ""
[既に指定のユーザーが存在する時]
  ExitCode: 2,
  StdOut: ""
  StdErr: "アカウントは既に存在します。
    NET HELPMSG 2224 と入力すると、より詳しい説明が得られます。"
[パスワードポリシーを満たしていない時]
  ExitCode: 2,
  StdOut: ""
  StdErr: "バスワードはバスワード ポリシーの要件を満たしていません。
    パスワードの最短の長さ、パスワードの複雑性、およびバスワード履歴の
    要件を確認してください。
    NET HELPMSG 2245 と入力すると、より詳しい説明が得られます。"
[管理者権限なし]
  ExitCode: 2,
  StdOut: ""
  StdErr: "システム エラー 5 が発生しました。
  アクセスが拒否されました。"
       */
      var iRetVal = os.runSync(mainCmd, args, {
        shell: false,
        winStyle: 'hidden'
      });
      if (iRetVal === CD.runs.ok) return;

      throw new Error('Error [ExitCode is not Ok] "' + iRetVal + '"\n');
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + functionName + ' (' + MODULE_TITLE + ')');
    }
  }; // }}}

  // os.attachAdminAuthorityToUser {{{
  /**
   * Adds the administrator authority to the user.
   *
   * @example
   * Wsh.OS.attachAdminAuthorityToUser('MyUserName');
   * @function attachAdminAuthorityToUser
   * @memberof Wsh.OS
   * @param {string} userName - The user name to add administrator authority.
   * @throws {string} - If an error occurs during command execution, or if the command exits with a value other than 0.
   * @returns {void}
   */
  os.attachAdminAuthorityToUser = function (userName) {
    var functionName = 'os.attachAdminAuthorityToUser';
    if (!isString(userName)) throwErrNonStr(functionName, userName);

    var mainCmd = os.exefiles.net;
    var args = ['localgroup', 'Administrators', userName, '/ADD', '/Y'];

    try {
      /*
WScript.Shell.Exec("net localgroup Administrators UserName /ADD /Y")
[成功時]
  ExitCode: 0,
  StdOut: "コマンドは正常に終了しました。"
  StdErr: ""
[既にメンバーの時]
  ExitCode: 2,
  StdOut: ""
  StdErr: "システム エラー 1378 が発生しました。
    指定されたアカウント名は既にグループのメンバーです。"
[ユーザーが存在しない時]
  ExitCode: 1,
  StdOut: ""
  StdErr: "次のグローバル ユーザーまたはグループは存在しません:
    ken。
  NET HELPMSG 3783 と入力すると、より詳しい説明が得られます。"
       */
      var iRetVal = os.runSync(mainCmd, args, {
        shell: false,
        winStyle: 'hidden'
      });
      if (iRetVal === CD.runs.ok) return;

      throw new Error('Error [ExitCode is not Ok] "' + iRetVal + '"\n');
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + functionName + ' (' + MODULE_TITLE + ')');
    }
  }; // }}}

  // os.deleteUser {{{
  /**
   * Delete the user in this local Windows.
   *
   * @example
   * Wsh.OS.deleteUser('MyUserName');
   * @function deleteUser
   * @memberof Wsh.OS
   * @param {string} userName - The user name to add.
   * @throws {string} - If an error occurs during command execution, or if the command exits with a value other than 0.
   * @returns {void}
   */
  os.deleteUser = function (userName) {
    var functionName = 'os.deleteUser';
    if (!isString(userName)) throwErrNonStr(functionName, userName);

    var mainCmd = os.exefiles.net;
    var args = ['user', userName, '/DELETE', '/Y'];

    try {
      /*
net user UserName /delete
[成功時]
  ExitCode: 0,
  StdOut: "コマンドは正常に終了しました。"
  StdErr: ""
[失敗時 対象ユーザーが存在しない]
  ExitCode: 2,
  StdOut: ""
  StdErr: "ユーザー名が見つかりません。
    NET HELPMSG 2221 と入力すると、より詳しい説明が得られます。"
       */
      var iRetVal = os.runSync(mainCmd, args, {
        shell: false,
        winStyle: 'hidden'
      });
      if (iRetVal === CD.runs.ok) return;

      throw new Error('Error [ExitCode is not Ok] "' + iRetVal + '"\n');
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + functionName + ' (' + MODULE_TITLE + ')');
    }
  }; // }}}
})();

// vim:set foldmethod=marker commentstring=//%s :
