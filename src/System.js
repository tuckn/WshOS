/* globals Wsh: false */

(function () {
  // Shorthands
  var sh = Wsh.Shell;
  var CD = Wsh.Constants;
  var util = Wsh.Util;
  var fso = Wsh.FileSystemObject;
  var path = Wsh.Path;

  var insp = util.inspect;
  var isString = util.isString;
  var isSolidString = util.isSolidString;
  var includes = util.includes;

  var os = Wsh.OS;

  /** @constant {string} */
  var MODULE_TITLE = 'WshOS/System.js';

  var throwErrNonStr = function (functionName, typeErrVal) {
    util.throwTypeError('string', MODULE_TITLE, functionName, typeErrVal);
  };

  /**
   * Windows system-specific end-of-line marker. Similar to {@link https://nodejs.org/api/os.html#os_os_eol|Node.js os.eol}.
   *
   * @example
   * Wsh.OS.EOL; // Returns: '\r\n'
   * @name EOL
   * @memberof Wsh.OS
   * @constant {string}
   */
  os.EOL = '\r\n';

  // os.platform {{{
  /**
   * Returns win32 which is the string identifying Windows. Similar to {@link https://nodejs.org/api/os.html#os_os_platform|Node.js os.platform()}.
   *
   * @example
   * Wsh.OS.platform(); // Returns: 'win32'
   * @function platform
   * @memberof Wsh.OS
   * @returns {string} - Returns 'win32'
   */
  os.platform = function () {
    return 'win32';
  }; // }}}

  os._cwd = sh.CurrentDirectory;

  // os.exefiles {{{
  // %SystemRoot%\Synstem32
  var SYSDIR = String(fso.GetSpecialFolder(CD.folderSpecs.system));

  /**
   * Paths of system binary file. (Why defining full paths?): For example, if a file called Net.cmd exists in the current directory, the net use command cannot be executed. In this case, we can execute by specifying the full path (C:\Windows\System32\net.exe use).
   *
   * @namespace exefiles
   * @memberof Wsh.OS
   */
  /** @lends Wsh.OS.exefiles */
  os.exefiles = {
    /** @constant {string} */
    cmd: path.join(SYSDIR, 'cmd.exe'),
    /** @constant {string} */
    certutil: path.join(SYSDIR, 'certutil.exe'),
    /** @constant {string} */
    cscript: path.join(SYSDIR, 'cscript.exe'),
    /** @constant {string} */
    wscript: path.join(SYSDIR, 'wscript.exe'),
    /** @constant {string} */
    net: path.join(SYSDIR, 'net.exe'),
    /** @constant {string} */
    netsh: path.join(SYSDIR, 'netsh.exe'),
    /** @constant {string} */
    notepad: path.join(SYSDIR, 'notepad.exe'),
    /** @constant {string} */
    ping: path.join(SYSDIR, 'PING.EXE'),
    /** @constant {string} */
    schtasks: path.join(SYSDIR, 'schtasks.exe'),
    /** @constant {string} */
    timeout: path.join(SYSDIR, 'timeout.exe'),
    /** @constant {string} */
    xcopy: path.join(SYSDIR, 'xcopy.exe')
  }; // }}}

  // os.getEnvVars {{{
  /**
   * Returns Windows Environment Variables. The execution result of this function is already stored in the global variable `{@link Wsh.OS.envVars}`.
   *
   * @example
   * Wsh.OS.getEnvVars();
   * // Returns: {
   * //   ALLUSERSPROFILE: "C:\ProgramData",
   * //   APPDATA: "C:\Users\UserName\AppData\Roaming",
   * //   CommonProgramFiles: "C:\Program Files\Common Files",
   * //   CommonProgramFiles(x86): "C:\Program Files (x86)\Common Files",
   * //   CommonProgramW6432: "C:\Program Files\Common Files",
   * //   COMPUTERNAME: "MYPC0123",
   * //   ComSpec: "C:\WINDOWS\system32\cmd.exe",
   * //   HOMEDRIVE: "C:",
   * //   HOMEPATH: "\Users\UserName",
   * //   ... }
   * @function getEnvVars
   * @memberof Wsh.OS
   * @returns {object} - Windows Environment Variables.
   */
  os.getEnvVars = function () {
    var environmentVars = {};

    var enmSet = new Enumerator(sh.Environment('PROCESS'));
    var itm;
    var i = 0;
    for (; ! enmSet.atEnd(); enmSet.moveNext()) {
      itm = enmSet.item();

      if (/^=/.test(itm) || !includes(itm, '=')) continue;

      i = itm.indexOf('=');
      environmentVars[itm.slice(0, i)] = itm.slice(i + 1);
    }

    return environmentVars;
  }; // }}}

  // os.envVars {{{
  /**
   * Windows Environment Variables. `{@link Wsh.OS.getEnvVars}` execution result and the referer of `{@link https://docs.tuckn.net/WshModeJs/process.html#.env|process.env}`.
   *
   * @example
   * console.dir(Wsh.OS.envVars);
   * // Returns: {
   * //   ALLUSERSPROFILE: "C:\ProgramData",
   * //   APPDATA: "C:\Users\UserName\AppData\Roaming",
   * //   CommonProgramFiles: "C:\Program Files\Common Files",
   * //   CommonProgramFiles(x86): "C:\Program Files (x86)\Common Files",
   * //   CommonProgramW6432: "C:\Program Files\Common Files",
   * //   COMPUTERNAME: "MYPC0123",
   * //   ComSpec: "C:\WINDOWS\system32\cmd.exe",
   * //   HOMEDRIVE: "C:",
   * //   HOMEPATH: "\Users\UserName",
   * //   ... }
   * @name envVars
   * @memberof Wsh.OS
   * @constant {object}
   */
  os.envVars = os.getEnvVars(); // Caching
  // }}}

  var _osArchitecture = undefined;

  // os.arch {{{
  /**
   * Returns the operating system CPU architecture. Similar to {@link https://nodejs.org/api/os.html#os_os_arch|Node.js os.arch()}.
   *
   * @example
   * Wsh.OS.arch(); // Returns: 'amd64'
   * @function arch
   * @memberof Wsh.OS
   * @returns {string} - The operating system CPU architecture.
   */
  os.arch = function () {
    if (_osArchitecture !== undefined) return _osArchitecture;

    _osArchitecture = os.envVars.PROCESSOR_ARCHITECTURE.toLowerCase();

    if (_osArchitecture === 'x86'
        && isSolidString(os.envVars.PROCESSOR_ARCHITEW6432)) {
      _osArchitecture = os.envVars.PROCESSOR_ARCHITEW6432.toLowerCase();
    }

    return _osArchitecture;
  }; // }}}

  var _osIs64Arch = undefined;

  // os.is64arch {{{
  /**
   * Identifies whether this Windows CPU architecture is 64bit or not.
   *
   * @example
   * Wsh.OS.is64arch(); // Returns: true
   * @function is64arch
   * @memberof Wsh.OS
   * @returns {boolean} - Returns true if the architecture is 64bit, else false.
   */
  os.is64arch = function () {
    if (_osIs64Arch !== undefined) return _osIs64Arch;

    var arch = os.arch();

    if (arch === 'amd64' || arch === 'ia64' || arch !== 'x86') {
      _osIs64Arch = true;
    } else {
      _osIs64Arch = false;
    }

    return _osIs64Arch;
  }; // }}}

  // os.tmpdir {{{
  /**
   * Returns the operating system's default directory for temporary files as a string. Similar to {@link https://nodejs.org/api/os.html#os_os_tmpdir|Node.js os.tmpdir()}.
   *
   * @example
   * Wsh.OS.tmpdir();
   * // Returns: 'C:\\Users\\UserName\\AppData\\Local\\Temp'
   * @function tmpdir
   * @memberof Wsh.OS
   * @returns {string} - The path of directory for temporary.
   */
  os.tmpdir = function () {
    return os.envVars.TMP;
  }; // }}}

  // os.makeTmpPath {{{
  /**
   * Creates the unique path on the %TEMP%
   *
   * @example
   * var os = Wsh.OS; // Shorthand
   *
   * os.makeTmpPath();
   * // Returns: 'C:\\Users\\UserName\\AppData\\Local\\Temp\\rad6E884.tmp'
   *
   * os.makeTmpPath();
   * // Returns: 'C:\\Users\\UserName\\AppData\\Local\\Temp\\rad20D90.tmp'
   *
   * os.makeTmpPath('cache-', '.tmp');
   * // Returns: 'C:\\Users\\UserName\\AppData\\Local\\Temp\\cache-radA8812.tmp.tmp'
   * @function makeTmpPath
   * @memberof Wsh.OS
   * @param {string} [prefix] - The prefix name of the path.
   * @param {string} [postfix] - The postfix name of the path..
   * @returns {string} - The unique path on the %TEMP%.
   */
  os.makeTmpPath = function (prefix, postfix) {
    var tmpName = fso.GetTempName();

    if (!isSolidString(prefix)) prefix = '';
    if (!isSolidString(postfix)) postfix = '';

    return path.join(os.tmpdir(), prefix + tmpName + postfix);
  }; // }}}

  // os.homedir {{{
  /**
   * Returns the string path of the current user's home directory. Similar to {@link https://nodejs.org/api/os.html#os_os_homedir|Node.js os.homedir()}.
   *
   * @example
   * Wsh.OS.homedir(); // Returns: 'C:\\Users\\UserName'
   * @function homedir
   * @memberof Wsh.OS
   * @returns {string} - The path of the current user's home directory.
   */
  os.homedir = function () {
    return path.join(os.envVars.HOMEDRIVE, os.envVars.HOMEPATH);
  }; // }}}

  // os.hostname {{{
  /**
   * Returns the computer name of the operating system as a string. Similar to {@link https://nodejs.org/api/os.html#os_os_hostname|Node.js os.hostname()}.
   *
   * @example
   * Wsh.OS.hostname(); // Returns: 'MYPC0123'
   * @function hostname
   * @memberof Wsh.OS
   * @returns {string} - The computer name.
   */
  os.hostname = function () {
    return os.envVars.COMPUTERNAME;
  }; // }}}

  // os.type {{{
  /**
   * Returns the operating system name. Similar to {@link https://nodejs.org/api/os.html#os_os_type|Node.js os.type()}
   *
   * @example
   * Wsh.OS.type(); // Returns: 'Windows_NT'
   * @function type
   * @memberof Wsh.OS
   * @returns {string} - The operating system name.
   */
  os.type = function () {
    return os.envVars.OS;
  }; // }}}

  // os.userInfo {{{
  /**
   * Returns information about the currently effective user. Similar to {@link https://nodejs.org/api/os.html#os_os_userinfo_options|Node.js os.userInfo()}.
   *
   * @example
   * Wsh.OS.userInfo();
   * // Returns: {
   * //   uid: -1,
   * //   gid: -1,
   * //   username: 'UserName',
   * //   homedir: 'C:\Users\UserName',
   * //   shell: null }
   * @function userInfo
   * @memberof Wsh.OS
   * @returns {object} - The object of effective user.
   */
  os.userInfo = function (/* options */) {
    return {
      uid: -1, // On Windows
      gid: -1, // On Windows
      username: os.envVars.USERNAME,
      homedir: os.envVars.USERPROFILE,
      shell: null // On Windows
    };
  }; // }}}

  // os.writeTempText {{{
  /**
   * Writes the string on a temporary file and returns the file-path.
   *
   * @example
   * var os = Wsh.OS; // Shorthand
   *
   * var tmpPath = os.writeTempText('Foo Bar');
   * // Returns: 'C:\\Users\\UserName\\AppData\\Local\\Temp\\rad91A81.tmp'
   * @function writeTempText
   * @memberof Wsh.OS
   * @param {string} txtData - Thes string data.
   * @param {string} [ext=''] - The extension of the tempporary text file.
   * @returns {string} - The temporary file path.
   */
  os.writeTempText = function (txtData, ext) {
    var functionName = 'os.writeTempText';
    if (!isString(txtData)) throwErrNonStr(functionName, txtData);
    if (!isSolidString(ext)) ext = '';

    var tmpTxtPath = os.makeTmpPath('osWriteTempWsh_', ext);
    var tmpTxtObj = fso.CreateTextFile(tmpTxtPath, CD.fso.creates.yes);
    tmpTxtObj.Write(txtData);
    tmpTxtObj.Close();

    if (!fso.FileExists(tmpTxtPath)) {
      throw new Error('Error: [Create TmpFile] ' + tmpTxtPath + '\n'
        + '  at ' + functionName + ' (' + MODULE_TITLE + ')');
    }

    return tmpTxtPath;
  }; // }}}

  // System Information

  // os.sysInfo {{{
  var _osSysInfo = undefined;

  /**
   * Returns the local computer OS system infomation. {@link https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-operatingsystem|Win32_OperatingSystem class}.
   *
   * @example
   * Wsh.OS.sysInfo();
   * // Returns: {
   * //   BuildNumber: '12345',
   * //   BuildType: 'Multiprocessor Free',
   * //   Caption: 'Microsoft Windows 10 Pro',
   * //   CodeSet: '987',
   * //   CountryCode: '81',
   * //   CreationClassName: 'Win32_OperatingSystem',
   * //   CSCreationClassName: 'Win32_ComputerSystem',
   * //   CSDVersion: null,
   * //   CSName: 'COMPNAME',
   * //   CurrentTimeZone: 540,
   * //   DataExecutionPrevention_32BitApplications: true,
   * //   DataExecutionPrevention_Available: true,
   * //   DataExecutionPrevention_Drivers: true,
   * //   DataExecutionPrevention_SupportPolicy: 2,
   * //   Debug: false,
   * //   Description: '',
   * //   Distributed: false,
   * //   EncryptionLevel: 256,
   * //   ForegroundApplicationBoost: 2,
   * //   FreePhysicalMemory: '15674120',
   * //   FreeSpaceInPagingFiles: '8773316',
   * //   FreeVirtualMemory: '18820380',
   * //   InstallDate: '20l71221152304.000000+540',
   * //   LargeSystemCache: null,
   * //   LastBootUpTime: '20181003042152.994157+540',
   * //   LocalDateTime: '20181003051042.787000+540',
   * //   Locale: '0411',
   * //   Manufacturer: 'Microsoft Corporation',
   * //   MaxNumberOfProcesses: -1,
   * //   MaxProcessMemorySize: '137438953344',
   * //   MUILanguages: [ 'ja-JP' ],
   * //   Name: 'Microsoft Windows 10 Pro|C:\WINDOWS|\Device\Harddisk0\Partition4',
   * //   NumberOfLicensedUsers: 0,
   * //   NumberOfProcesses: 212,
   * //   NumberOfUsers: 2,
   * //   OperatingSystemSKU: 48,
   * //   Organization: '',
   * //   OSArchitecture: '64 ビット',
   * //   OSLanguage: 1041,
   * //   OSProductSuite: 256,
   * //   OSType: 18,
   * //   OtherTypeDescription: null,
   * //   PAEEnabled: null,
   * //   PlusProductID: null,
   * //   PlusVersionNumber: null,
   * //   PortableOperatingSystem: false,
   * //   Primary: true,
   * //   ProductType: 1,
   * //   RegisteredUser: 'Username',
   * //   SerialNumber: '12345-67890-12345-ABOEM',
   * //   ServicePackMajorVersion: 0,
   * //   ServicePackMinorVersion: 0,
   * //   SizeStoredInPagingFiles: '8912896',
   * //   Status: 'OK',
   * //   SuiteMask: 272,
   * //   SystemDevice: '\Device\HarddiskVolume4',
   * //   SystemDirectory: 'C:\WINDOWS\system32',
   * //   SystemDrive: 'C:',
   * //   TotalSwapSpaceSize: null,
   * //   TotalVirtualMemorySize: '42392460',
   * //   TotalVisibleMemorySize: '33479564',
   * //   Version: '10.0.12345',
   * //   WindowsDirectory: 'C:\WINDOWS' }
   * @function sysInfo
   * @memberof Wsh.OS
   * @returns {object} - The local computer OS system infomation.
   */
  os.sysInfo = function () {
    if (_osSysInfo !== undefined) return _osSysInfo;

    var sWbemObjSets = os.WMI.execQuery('SELECT * FROM Win32_OperatingSystem');
    var wmiObj = os.WMI.toJsObject(sWbemObjSets[0]); /** @todo [0] only? */

    _osSysInfo = wmiObj;
    return _osSysInfo;
  }; // }}}

  // os.cmdCodeset {{{
  /**
   * Returns the chara-code of CommandPrompt on this Windows.
   *
   * @example
   * Wsh.OS.cmdCodeset(); // Returns: 'shift_jis'
   * @function cmdCodeset
   * @memberof Wsh.OS
   * @returns {string} - The code page value an operating system uses. See {@link https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-operatingsystem|CodeSet}.
   */
  os.cmdCodeset = function () {
    return CD.codePageIdentifiers[os.sysInfo().CodeSet];
  }; // }}}

  // os.freemem {{{
  /**
   * Returns the amount of free system memory in bytes as an integer. Similar to {@link https://nodejs.org/api/os.html#os_os_freemem|Node.js os.freemem()}.
   *
   * @example
   * Wsh.OS.freemem(); // Returns: 176822642
   * @function freemem
   * @memberof Wsh.OS
   * @returns {number} - The amount of free system memory in bytes as an integer.
   */
  os.freemem = function () {
    return Number(os.sysInfo().FreePhysicalMemory) * 1024;
  }; // }}}

  // os.release {{{
  /**
   * Returns the operating system as a string. Similar to {@link https://nodejs.org/api/os.html#os_os_release|Node.js os.release()}.
   *
   * @example
   * Wsh.OS.release(); // Returns: 10.0.18363
   * @function release
   * @memberof Wsh.OS
   * @returns {string} - The version of the operating system as a string.
   */
  os.release = function () {
    return os.sysInfo().Version;
  }; // }}}

  // os.totalmem {{{
  /**
   * Returns the total amount of system memory in bytes as an integer. Similar to {@link https://nodejs.org/api/os.html#os_os_totalmem|Node.js os.totalmem()}.
   *
   * @example
   * Wsh.OS.release(); // Returns: 34276515840
   * @function totalmem
   * @memberof Wsh.OS
   * @returns {number} - The total amount of system memory in bytes as an integer.
   */
  os.totalmem = function () {
    /** Node.js returns Byte, TotalVisibleMemorySize returns MByte. */
    return Number(os.sysInfo().TotalVisibleMemorySize) * 1024;
  }; // }}}

  var _hasUAC = undefined;

  // os.hasUAC {{{
  /**
   * Checks whether or not that this Windows have UAC
   *
   * @example
   * Wsh.OS.hasUAC(); // Returns: true
   * @function hasUAC
   * @memberof Wsh.OS
   * @returns {boolean} - If having UAC returns true.
   */
  os.hasUAC = function () {
    if (_hasUAC !== undefined) return _hasUAC;

    if (parseFloat(os.sysInfo().Version.slice(0, 3)) >= 6.0) {
      _hasUAC = true;
    } else {
      _hasUAC = true;
    }

    return _hasUAC;
  }; // }}}

  // os.isUacDisable {{{
  /**
   * Checks whether enabled or disabled UAC.
   *
   * @example
   * Wsh.OS.isUacDisable(); // Returns: true
   * @function isUacDisable
   * @memberof Wsh.OS
   * @returns {boolean} - If enabled UAC returns true.
   */
  os.isUacDisable = function () {
    var regUac = 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System';
    var cpba = sh.RegRead(regUac + '\\ConsentPromptBehaviorAdmin');
    var posd = sh.RegRead(regUac + '\\PromptOnSecureDesktop');
    var elua = sh.RegRead(regUac + '\\EnableLUA');

    return cpba + posd + elua === 0;
  }; // }}}

  // Event Log

  // os.writeLogEvent {{{
  function _writeLogEvent (type, text) {
    var functionName = '_writeLogEvent';
    if (!isString(text)) throwErrNonStr(functionName, text);

    try {
      /**
       * Logs the event in Windows Event Log.
       *
       * @function LogEvent
       * @memberof Wsh.Shell
       * @param {number} intType - The type of the event. See {@link https://docs.tuckn.net/WshUtil/Wsh.Constants.logLevels.html|Wsh.Constants.logLevels}
       * @param {string} strMessage
       * @param {string} [strTarget]
       * @returns {void}
       */
      sh.LogEvent(type, text);
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + functionName + ' (' + MODULE_TITLE + ')\n'
        + '  Failed to write the log event.\n  text: ' + text);
    }
  }

  /**
   * The wrapper object for functions to handle Windows Event Log.
   *
   * @namespace writeLogEvent
   * @memberof Wsh.OS
   */
  os.writeLogEvent = {
    /**
     * Logs the success event in Windows Event Log.
     *
     * @example
     * Wsh.OS.writeLogEvent.success('Success Log');
     * @function success
     * @memberof Wsh.OS.writeLogEvent
     * @param {string} text - The message string.
     * @returns {void}
     */
    success: function (text) {
      _writeLogEvent(CD.logLevels.success, text);
    },

    /**
     * Logs the error event in Windows Event Log.
     *
     * @example
     * Wsh.OS.writeLogEvent.error('Error Log');
     * @function error
     * @memberof Wsh.OS.writeLogEvent
     * @param {string} text - The message string.
     * @returns {void}
     */
    error: function (text) {
      _writeLogEvent(CD.logLevels.error, text);
    },

    /**
     * Logs the warn event in Windows Event Log.
     *
     * @example
     * Wsh.OS.writeLogEvent.success('Warn Log');
     * @function warn
     * @memberof Wsh.OS.writeLogEvent
     * @param {string} text - The message string.
     * @returns {void}
     */
    warn: function (text) {
      _writeLogEvent(CD.logLevels.warning, text);
    },

    /**
     * Logs the Information event in Windows Event Log.
     *
     * @example
     * Wsh.OS.writeLogEvent.info('Information Log');
     * @function info
     * @memberof Wsh.OS.writeLogEvent
     * @param {string} text - The message string.
     * @returns {void}
     */
    info: function (text) {
      _writeLogEvent(CD.logLevels.infomation, text);
    }
  }; // }}}

  // Clipboard

  // os.setStrToClipboardAO {{{
  /**
   * [Experimenta] Sets the string to the Clipboard.
   *
   * @todo Not work!!
   * @function setStrToClipboardAO
   * @memberof Wsh.OS
   * @param {string} str
   * @returns {void}
   */
  os.setStrToClipboardAO = function (str) {
    var dynwrap = new ActiveXObject('DynamicWrapper');

    dynwrap.register('ole32', 'OleInitialize', 'f=s', 'i=l', 'r=l');
    dynwrap.OleInitialize(0);
    new ActiveXObject('htmlfile').parentWindow.clipboardData.setData(
      'text',
      str
    );
  }; // }}}

  // os.getStrFromClipboard {{{
  /**
   * [Experimental] Gets the string from Clipboard.
   *
   * @function getStrFromClipboard
   * @memberof Wsh.OS
   * @returns {string}
   */
  os.getStrFromClipboard = function () {
    return new ActiveXObject('htmlfile').parentWindow.clipboardData.getData(
      'text'
    );
  }; // }}}
})();

// vim:set foldmethod=marker commentstring=//%s :
