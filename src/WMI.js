/* globals Wsh: false */

(function () {
  if (Wsh && Wsh.OS) return;

  /**
   * Shorthand of WScript.CreateObject('Shell.Application')
   *
   * @namespace ShellApplication
   * @memberof Wsh
   */
  if (!Wsh.ShellApplication) {
    Wsh.ShellApplication = WScript.CreateObject('Shell.Application');
  }

  /**
   * This module takes charge of basis handling Windows OS (Partly similar to Node.js-OS).
   *
   * @namespace OS
   * @memberof Wsh
   * @requires {@link https://github.com/tuckn/WshPath|tuckn/WshPath}
   */
  Wsh.OS = {};

  // Shorthands
  var CD = Wsh.Constants;
  var util = Wsh.Util;
  var fso = Wsh.FileSystemObject;
  var path = Wsh.Path;

  var insp = util.inspect;
  var isArray = util.isArray;
  var isPureNumber = util.isPureNumber;
  var isString = util.isString;
  var isSolidArray = util.isSolidArray;
  var isSolidString = util.isSolidString;
  var isSameMeaning = util.isSameMeaning;
  var obtain = util.obtainPropVal;
  var includes = util.includes;

  var os = Wsh.OS;

  /** @constant {string} */
  var MODULE_TITLE = 'WshOS/WMI.js';

  var throwErrNonStr = function (functionName, typeErrVal) {
    util.throwTypeError('string', MODULE_TITLE, functionName, typeErrVal);
  };

  /**
   * The wrapper object for functions to handle {@link https://docs.microsoft.com/en-us/windows/win32/wmisdk/wmi-start-page|WMI} (Windows Management Instrumentation. WBEM for Windows).
   *
   * @namespace WMI
   * @memberof Wsh.OS
   */
  os.WMI = {};

  /**
   * An SWbemObjectSet object is a collection of {@link https://docs.microsoft.com/en-us/windows/win32/wmisdk/swbemobjectset|SWbemObjectSet object}
   *
   * @name sWbemObjectSet
   */

  // os.WMI.execQuery {{{
  /**
   * @typedef {object} typeExecWmiQueryOptions
   * @property {string} [compName=null] - Ex1. server1.network.fabrikam Ex2. 111.222.333.444
   * @property {string} [domainName="."]
   * @property {string} [namespace="."]
   */

  /**
   * Executes the query to WMI (Windows Management Instrumentation. WBEM for Windows). {@link https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/ms974592(v=msdn.10)?redirectedfrom=MSDN|Microsoft Docs} {@link http://www.wmifun.net/library/|Ref.1}
   *
   * @example
   * var wmi = Wsh.OS.WMI; // Shorthand
   *
   * // Ex.1 Gets all processes
   * var query = 'SELECT * FROM Win32_Process';
   * var sWbemObjSets = wmi.execQuery(query);
   *
   * // Ex.2 Add WHERE phrase (Do not omit the extension)
   * var query = 'SELECT * FROM Win32_Process WHERE Caption = "notepad.exe"';
   * var sWbemObjSets = wmi.execQuery(query);
   *
   * sWbemObjSets.forEach(function (sWbemObjSet) {
   *   console.log('ProcessID: ' + sWbemObjSet.ProcessId);
   *   console.log('ExecutablePath: ' + sWbemObjSet.ExecutablePath);
   *   var iRetVal = sWbemObjSet.Terminate();
   *   console.log('Terminated it with returning value ' + iRetVal);
   * });
   * // Outputs:
   * //   ProcessID: 8356
   * //   ExecutablePath: C:\WINDOWS\system32\notepad.exe
   * //   Terminated it with returning value 0
   * //   ProcessID: 26712
   * //   ExecutablePath: C:\WINDOWS\system32\notepad.exe
   * //   Terminated it with returning value 0
   * //   ....
   *
   * // The below code does not work. Because the return object is not JS objects.
   * // Use Wsh.OS.WMI.toJsObjects to convert to JS Objects
   * Object.keys(sWbemObjSets[0]).forEach(function (propName) {
   *   console.log(propName + ': ' + sWbemObjSets[0][propName]);
   * });
   * // No display
   *
   * // Ex.3 Specifing options
   * var query = 'SELECT * FROM CIM_BIOSElement';
   * var sWbemObjSets = wmi.execQuery(query, {
   *   compName: 'otherComp.office.local',
   *   domainName: 'office.local'
   * });
   *
   * // Require converting to read the array property in SWbemObjectSet
   * console.log(new VBArray(sWbemObjSets[0].BiosCharacteristics).toArray());
   *
   * // Ex.4 Specifing the other namespace
   * var query = 'SELECT * FROM CIM_LogicalElement';
   * var sWbemObjSets = wmi.execQuery(query, {
   *   namespace: '\\\\.\\root\\Hardware'
   * });
   * @function execQuery
   * @memberof Wsh.OS.WMI
   * @param {string} query - The query to execute.
   * @param {typeExecWmiQueryOptions} [options] - Optional parameters.
   * @returns {sWbemObjectSet[]} - The array of enumerated SWbem-Object-Sets.
   */
  os.WMI.execQuery = function (query, options) {
    var FN = 'os.WMI.execQuery';
    if (!isSolidString(query)) throwErrNonStr(FN, query);

    var compName = obtain(options, 'compName', null);
    var domainName = obtain(options, 'domainName', '.');
    var namespace = obtain(options, 'namespace', '\\\\' + domainName + '\\root\\CIMV2');

    /** WBEM:Web-Based Enterprise Management */
    var wbem = WScript.CreateObject('WbemScripting.SWbemLocator');
    // var swbemServices = GetObject(namespace);
    /** @note The Above code equal WScript.CreateObject('WbemScripting.SWbemLocator').ConnectServer(null, namespace) */

    try {
      // Gets SWbemServices object
      /**
       * root\CIMV2 is the local computer namespace. CIM = Common Information Model. {@link https://docs.microsoft.com/en-us/windows/win32/wmisdk/swbemlocator-connectserver|Microsoft Docs}
       */
      var swbemServices = wbem.ConnectServer(compName, namespace);
      // Execute a query. Gets SWbemObjectSet object.
      var swbemSet = swbemServices.ExecQuery(query);
      // Convert SWbemObjectSet object to Enumerator object for JScript.
      var enumeSet = new Enumerator(swbemSet);
      var sWbemObjSets = [];
      if (!sWbemObjSets) return [];

      while (!enumeSet.atEnd()) {
        sWbemObjSets.push(enumeSet.item());
        enumeSet.moveNext();
      }

      return sWbemObjSets;
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + FN + ' (' + MODULE_TITLE + ')\n'
        + '  query: ' + query + '\n  options: ' + insp(options));
    }
  }; // }}}

  // os.WMI.toJsObject {{{
  /**
   * Converts Enumerated SWbemObjectSets to JScript Object. {@link https://qiita.com/tnakagawa/items/ae579f19d74dd86e40c6|Ref.2}
   *
   * @example
   * var wmi = Wsh.OS.WMI;
   *
   * var sWbemObjSets = wmi.execQuery('SELECT * FROM CIM_BIOSElement');
   * var biosElements = wmi.toJsObject(sWbemObjSets[0]);
   * console.dir(biosElements.ListOfLanguages);
   * // Outputs: [
   * //   "en|US|iso8859-1",
   * //   "fr|FR|iso8859-1",
   * //   "zh|TW|unicode",
   * //   "zh|CN|unicode",
   * //   "ja|JP|unicode",
   * //   "de|DE|iso8859-1",
   * //   "es|ES|iso8859-1",
   * //   "ru|RU|iso8859-5",
   * //   "ko|KR|unicode" ],
   * @function toJsObject
   * @memberof Wsh.OS.WMI
   * @param {sWbemObjectSet} sWbemObjSet - The SWbemObjectSet.
   * @returns {object} - The converted object.
   */
  os.WMI.toJsObject = function (sWbemObjSet) {
    // var FN = 'os.WMI.toJsObject';
    var wmiObj = {};

    // Store Properties
    var enumedProps = new Enumerator(sWbemObjSet.Properties_);

    while (!enumedProps.atEnd()) {
      var propItem = enumedProps.item();

      var val = propItem.Value;
      if (propItem.IsArray) {
        /**
         * @function toArray
         * @memberof VBArray
         * Returns the standard JavaScript array converted from a VBArray. {@link https://developer.mozilla.org/en-US/docs/Web/JavaScript/Microsoft_Extensions/VBArray/toArray|MDN}
         * @param {VBArray} safeArray
         * @returns {Array} - The standard JavaScript array
         */
        if (val !== null) val = new VBArray(val).toArray();
      }

      wmiObj[propItem.Name] = val;
      enumedProps.moveNext();
    }

    // Store Methods
    /**
     * @todo Not work.
    var enumedMethods = new Enumerator(sWbemObjSet.Methods_);

    while (!enumedMethods.atEnd()) {
      var methodItem = enumedMethods.item();
      var methdName = methodItem.Name;
      // Property name and method name do not overlap. maybe...
      wmiObj[methdName] = sWbemObjSet[methdName];
      enumedMethods.moveNext();
    }
    */

    return wmiObj;
  }; // }}}

  // os.WMI.toJsObjects {{{
  /**
   * Converts Enumerated SWbemObjectSets to JScript Objects. {@link https://qiita.com/tnakagawa/items/ae579f19d74dd86e40c6|Ref.2}
   *
   * @example
   * var wmi = Wsh.OS.WMI;
   *
   * var sWbemObjSets = wmi.execQuery('SELECT * FROM CIM_BIOSElement');
   * var wmiObjs = wmi.toJsObjects(sWbemObjSets);
   * console.dir(wmiObjs[0].ListOfLanguages);
   * // Outputs: [
   * //   "en|US|iso8859-1",
   * //   "fr|FR|iso8859-1",
   * //   "zh|TW|unicode",
   * //   "zh|CN|unicode",
   * //   "ja|JP|unicode",
   * //   "de|DE|iso8859-1",
   * //   "es|ES|iso8859-1",
   * //   "ru|RU|iso8859-5",
   * //   "ko|KR|unicode" ],
   * @function toJsObjects
   * @memberof Wsh.OS.WMI
   * @param {sWbemObjectSet[]} sWbemObjSets - The SWbemObjectSets.
   * @returns {object[]} - The array of converted objects.
   */
  os.WMI.toJsObjects = function (sWbemObjSets) {
    // var FN = 'os.WMI.toJsObjects';

    var wmiObjs = sWbemObjSets.map(function (sWbemObjSet) {
      return os.WMI.toJsObject(sWbemObjSet);
    });

    return wmiObjs;
  }; // }}}

  // os.WMI.getProcesses {{{
  /**
   * @typedef {object} typeGetProcessesOptions
   * @property {string[]} [selects=[]] - The field names to output. 'Caption', 'CommandLine', 'ExecutablePath', 'ProcessID'. default '*'.
   * @property {string[]} [matchWords=[]] - The matching filter words.
   * @property {string[]} [excludingWords=[]] - The excluding filter words.
   */

  /**
   * Returns Enumerated SWbemObjectSets of the specified process. {@link https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-process|Microsoft Docs}
   *
   * @example
   * var wmi = Wsh.OS.WMI;
   *
   * // Ex.1 Specifing a process name (Do not omit the extension)
   * var sWbemObjSets = wmi.getProcesses('chrome.exe');
   *
   * sWbemObjSets.forEach(function (sWbemObjSet) {
   *   console.log('ProcessID: ' + sWbemObjSet.ProcessId);
   *   console.log('Handle: ' + sWbemObjSet.Handle);
   *   console.log('CreationDate: ' + sWbemObjSet.CreationDate);
   *
   *   try {
   *     var iRetVal = sWbemObjSet.Terminate();
   *     console.log('Terminated it with returning value ' + iRetVal);
   *   } catch (e) {
   *     console.error('Failed to terminate. Already terminated?');
   *   }
   * });
   * // Outputs:
   * //   ProcessID: 8356
   * //   Handle: 18012
   * //   CreationDate: 20200217042344.517814+540
   * //   Terminated it with returning value 0
   * //   ....
   *
   * // Ex.2 Specifing a process ID
   * var sWbemObjSets = wmi.getProcesses(8356);
   *
   * // Ex.3 Specifing a full path
   * var sWbemObjSets = wmi.getProcesses('D:\\Apps\\Firefox\\firefox.exe');
   *
   * // Ex.4 Specifing options
   * var sWbemObjSets = wmi.getProcesses('node.exe', {
   *   matchWords: ['nodemon'], // Compare with CommandLine
   *   excludingWords: ['80', '8080'] // Compare with CommandLine
   * });
   * @function getProcesses
   * @memberof Wsh.OS.WMI
   * @param {(number|string)} [processName] - The PID or the process name or the full path. If empty to specify all processes.
   * @param {typeGetProcessesOptions} [options] - Optional parameters.
   * @returns {sWbemObjectSet[]} - The array of enumerated SWbem-Object-Sets.
   */
  os.WMI.getProcesses = function (processName, options) {
    var FN = 'os.WMI.getProcesses';

    var query = 'SELECT';

    // SELECT phrase
    var selects = obtain(options, 'selects', []);
    var selectQuery = '';
    if (isSolidArray(selects)) {
      // The essential selectors
      if (!includes(selects, 'Caption')) selects.push('Caption');
      if (!includes(selects, 'CommandLine')) selects.push('CommandLine');
      if (!includes(selects, 'ExecutablePath')) selects.push('ExecutablePath');
      if (!includes(selects, 'ProcessID')) selects.push('ProcessID');

      // Join
      for (var i = 0, len = selects.length - 1; i < len; i++) {
        selectQuery += selects[i] + ', ';
      }
      selectQuery += util.last(selects);
    } else {
      selectQuery = '*';
    }

    query += ' ' + selectQuery + ' FROM Win32_Process';

    // WHERE phrase
    if (!isPureNumber(processName) && !isString(processName)) {
      throwErrNonStr(FN, processName);
    }
    var comparesFullPath = false;

    if (isPureNumber(processName)) { // PID
      query += ' WHERE ProcessID=' + String(processName);
    } else if (isSolidString(processName)) {
      query += ' WHERE Caption="' + path.basename(processName) + '"';
      comparesFullPath = path.isAbsolute(processName);
    }

    // Execute
    var sWbemObjSets = os.WMI.execQuery(query);
    if (sWbemObjSets.length === 0) return [];

    // Filter
    var matchWords = obtain(options, 'matchWords', []);
    // Convert matchWords to Array, if it is string
    if (isSolidString(matchWords)) {
      matchWords = [matchWords];
    } else if (!isArray(matchWords)) {
      matchWords = [];
    }

    var excludingWords = obtain(options, 'excludingWords', []);
    // Convert excludedWords to Array, if it is string
    if (isSolidString(excludingWords)) {
      excludingWords = [excludingWords];
    } else if (!isArray(excludingWords)) {
      excludingWords = [];
    }

    return sWbemObjSets.filter(function (sWbemObjSet) {
      // Filter with ExecutablePath
      if (comparesFullPath
          && !isSameMeaning(sWbemObjSet.ExecutablePath, processName)) {
        return false;
      }

      // Filter with CommandLine
      /** CommandLine sometimes does not contain a full path */
      var notMatched = matchWords.some(function (word) {
        if (!includes(sWbemObjSet.CommandLine, word, 'i')) return true;
      });
      if (notMatched) return false;

      var matchedExcluding = excludingWords.some(function (word) {
        if (includes(sWbemObjSet.CommandLine, word, 'i')) return true;
      });
      if (matchedExcluding) return false;

      return true;
    });
  }; // }}}

  // os.WMI.getProcess {{{
  /**
   * Returns the SWbemObjectSet of the specified process
   *
   * @example
   * var wmi = Wsh.OS.WMI;
   *
   * var sWbemObjSet = wmi.getProcess('excel.exe');
   *
   * console.log('ProcessID: ' + sWbemObjSet.ProcessId);
   * console.log('ParentProcessId: ' + sWbemObjSet.ParentProcessId);
   * var iRetVal = sWbemObjSet.Terminate();
   * console.log('Terminated it with returning value ' + iRetVal);
   * // Outputs:
   * //   ProcessID: 19220
   * //   ParentProcessId: 160624
   * //   Terminated it with returning value 0
   * @function getProcess
   * @memberof Wsh.OS.WMI
   * @param {(number|string)} [processName] - The PID or the process name or the full path. If empty to specify all processes.
   * @param {typeGetProcessesOptions} [options] - Optional parameters.
   * @returns {sWbemObjectSet|null} - The enumerated SWbem-Object-Set.
   */
  os.WMI.getProcess = function (processName, options) {
    var FN = 'os.WMI.getProcess';
    if (!isPureNumber(processName) && !isString(processName)) {
      throwErrNonStr(FN, processName);
    }

    var sWbemObjSets = os.WMI.getProcesses(processName, options);

    if (sWbemObjSets.length === 0) return null;
    return sWbemObjSets[0];
  }; // }}}

  // os.WMI.getWithSWbemPath {{{
  /**
   * Returns the Enumerated SWbem-Object-Sets from the SWbem path (e.g. 'Win32_LogicalDisk.DeviceID="C:"').
   *
   * @example
   * var wmi = Wsh.OS.WMI;
   *
   * var sWbemPath = 'Win32_LogicalDisk.DeviceID="C:"';
   * var sWbemObjSet = wmi.getWithSWbemPath(sWbemPath);
   *
   * console.log('DeviceID: ' + sWbemObjSet.DeviceID);
   * console.log('Caption: ' + sWbemObjSet.Caption);
   * console.log('FileSystem: ' + sWbemObjSet.FileSystem);
   * console.log('FreeSpace: ' + sWbemObjSet.FreeSpace);
   * // Outputs:
   * //   DeviceID: C:
   * //   Caption: C:
   * //   FileSystem: NTFS
   * //   FreeSpace: 193899817343
   *
   * // The below code does not work. Because the return object is not JS objects.
   * // Use Wsh.OS.WMI.toJsObject to convert to A JS Object
   * Object.keys(sWbemObjSet).forEach(function (propName) {
   *   console.log(propName + ': ' + sWbemObjSet[propName]);
   * });
   * // No display
   *
   * // Ex.2 @todo Not work. Please tell me why
   * var sWbemPath = 'Win32_Process.ProcessId=11888';
   * var sWbemObjSet = os.WMI.getWithSWbemPath(sWbemPath); // Error [-2147217350]
   * @function getWithSWbemPath
   * @memberof Wsh.OS.WMI
   * @param {string} sWbemPath - The SWbem path.
   * @param {typeExecWmiQueryOptions} [options] - Optional parameters.
   * @returns {sWbemObjectSet|null} - The enumerated SWbem-Object-Set.
   */
  os.WMI.getWithSWbemPath = function (sWbemPath, options) {
    var FN = 'os.WMI.getWithSWbemPath';
    if (!isSolidString(sWbemPath)) throwErrNonStr(FN, sWbemPath);

    var compName = obtain(options, 'compName', null);
    var domainName = obtain(options, 'domainName', '.');
    var namespace = obtain(options, 'namespace', '\\\\' + domainName + '\\root\\CIMV2');
    var wbem = WScript.CreateObject('WbemScripting.SWbemLocator');

    try {
      var swbemServices = wbem.ConnectServer(compName, namespace);
      var sWbemObjSet = swbemServices.Get(sWbemPath);
      if (!sWbemObjSet) return null;
      return sWbemObjSet;
    } catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + FN + ' (' + MODULE_TITLE + ')\n'
        + '  query: ' + sWbemPath + '\n  options: ' + insp(options));
    }
  }; // }}}

  // os.WMI.getWindowsUserAccounts {{{
  /**
   * Returns Enumerated SWbemObjectSets of Windows user accounts. {@link https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-useraccount|Microsoft Docs}
   *
   * @example
   * var sWbemObjSets = os.WMI.getWindowsUserAccounts();
   *
   * sWbemObjSets.forEach(function (sWbemObjSet) {
   *   console.log('AccountType: ' + sWbemObjSet.AccountType);
   *   console.log('Caption: ' + sWbemObjSet.Caption);
   *   console.log('Name: ' + sWbemObjSet.Name);
   *   console.log('Domain: ' + sWbemObjSet.Domain);
   *   console.log('Status: ' + sWbemObjSet.Status);
   * });
   * // Outputs:
   * //   AccountType: 512
   * //   Caption: COMPNAME\Administrator
   * //   Name: Administrator
   * //   Domain: COMPNAME
   * //   Status: Degraded
   * //   AccountType: 512
   * //   Caption: COMPNAME\DefaultAccount
   * //   ....
   * @function getWindowsUserAccounts
   * @memberof Wsh.OS.WMI
   * @returns {sWbemObjectSet[]} - The array of enumerated SWbem-Object-Sets.
   */
  os.WMI.getWindowsUserAccounts = function () {
    // var FN = 'os.WMI.getWindowsUserAccounts';
    var sWbemObjSets = os.WMI.execQuery('SELECT * FROM Win32_UserAccount');
    return sWbemObjSets;
  }; // }}}

  os._thisProcessID = undefined;
  os._thisProcessSWbemObjSet = undefined;
  os._thisParentProcessID = undefined;

  // os.WMI.getThisProcess {{{
  /**
   * Returns the enumerated SWbem-Object-Set of the self process.
   *
   * @example
   * var wmi = Wsh.OS.WMI;
   *
   * var thisProcess = wmi.getThisProcess();
   *
   * console.log('ProcessID: ' + thisProcess.ProcessId);
   * console.log('Caption: ' + thisProcess.Caption);
   * console.log('Name: ' + thisProcess.Name);
   * console.log('CommandLine: ' + thisProcess.CommandLine);
   * console.log('ParentProcessId: ' + thisProcess.ParentProcessId);
   * // Outputs:
   * //   ProcessID: 23576
   * //   Caption: cscript.exe
   * //   Name: cscript.exe
   * //   CommandLine: cscript.exe  //nologo OS.test.wsf -t WMI_getThisProcess
   * //   ParentProcessId: 300
   * @function getThisProcess
   * @memberof Wsh.OS.WMI
   * @returns {sWbemObjectSet} - The enumerated SWbem-Object-Set.
   */
  os.WMI.getThisProcess = function () {
    var FN = 'os.WMI.getThisProcess';

    if (os._thisProcessSWbemObjSet !== undefined) {
      return os._thisProcessSWbemObjSet;
    }

    // var FN = 'os.WMI.getThisProcess';
    var tmpJsCode = 'while(true) WScript.Sleep(1000);';
    var tmpJsPath = os.writeTempText(tmpJsCode, '.js');

    os.run(os.exefiles.wscript, [tmpJsPath]);

    var sWbemObjSet = os.WMI.getProcess(os.exefiles.wscript, {
      matchWords: [tmpJsPath] // tmpJsPath is the unique
    });

    // @FIXME sometime, ParentProcessId is null?
		try {
			os._thisProcessID = Number(sWbemObjSet.ParentProcessId);
		} catch (e) {
      throw new Error(insp(e) + '\n'
        + '  at ' + FN + ' (' + MODULE_TITLE + ')');
		}

    sWbemObjSet.Terminate();
    fso.DeleteFile(tmpJsPath, CD.fso.force.yes);

    os._thisProcessSWbemObjSet = os.WMI.getProcess(os._thisProcessID);
    os._thisParentProcessID = os._thisProcessSWbemObjSet.ParentProcessId;

    return os._thisProcessSWbemObjSet;
  }; // }}}
})();

// vim:set foldmethod=marker commentstring=//%s :
