﻿!function(){var CD,util,fso,path,insp,isArray,isPureNumber,isString,isSolidArray,isSolidString,isSameMeaning,obtain,includes,os,throwErrNonStr;Wsh&&Wsh.OS||(Wsh.ShellApplication||(Wsh.ShellApplication=WScript.CreateObject("Shell.Application")),Wsh.OS={},CD=Wsh.Constants,util=Wsh.Util,fso=Wsh.FileSystemObject,path=Wsh.Path,insp=util.inspect,isArray=util.isArray,isPureNumber=util.isPureNumber,isString=util.isString,isSolidArray=util.isSolidArray,isSolidString=util.isSolidString,isSameMeaning=util.isSameMeaning,obtain=util.obtainPropVal,includes=util.includes,os=Wsh.OS,throwErrNonStr=function(functionName,typeErrVal){util.throwTypeError("string","WshOS/WMI.js",functionName,typeErrVal)},os.WMI={},os.WMI.execQuery=function(query,options){isSolidString(query)||throwErrNonStr("os.WMI.execQuery",query);var compName=obtain(options,"compName",null),domainName=obtain(options,"domainName","."),namespace=obtain(options,"namespace","\\\\"+domainName+"\\root\\CIMV2"),wbem=WScript.CreateObject("WbemScripting.SWbemLocator");try{var swbemSet=wbem.ConnectServer(compName,namespace).ExecQuery(query),enumeSet=new Enumerator(swbemSet),sWbemObjSets=[];if(!sWbemObjSets)return[];for(;!enumeSet.atEnd();)sWbemObjSets.push(enumeSet.item()),enumeSet.moveNext();return sWbemObjSets}catch(e){throw new Error(insp(e)+"\n  at os.WMI.execQuery (WshOS/WMI.js)\n  query: "+query+"\n  options: "+insp(options))}},os.WMI.toJsObject=function(sWbemObjSet){for(var wmiObj={},enumedProps=new Enumerator(sWbemObjSet.Properties_);!enumedProps.atEnd();){var propItem=enumedProps.item(),val=propItem.Value;propItem.IsArray&&null!==val&&(val=new VBArray(val).toArray()),wmiObj[propItem.Name]=val,enumedProps.moveNext()}return wmiObj},os.WMI.toJsObjects=function(sWbemObjSets){return sWbemObjSets.map(function(sWbemObjSet){return os.WMI.toJsObject(sWbemObjSet)})},os.WMI.getProcesses=function(processName,options){var query="SELECT",selects=obtain(options,"selects",[]),selectQuery="";if(isSolidArray(selects)){includes(selects,"Caption")||selects.push("Caption"),includes(selects,"CommandLine")||selects.push("CommandLine"),includes(selects,"ExecutablePath")||selects.push("ExecutablePath"),includes(selects,"ProcessID")||selects.push("ProcessID");for(var i=0,len=selects.length-1;i<len;i++)selectQuery+=selects[i]+", ";selectQuery+=util.last(selects)}else selectQuery="*";query+=" "+selectQuery+" FROM Win32_Process",isPureNumber(processName)||isString(processName)||throwErrNonStr("os.WMI.getProcesses",processName);var comparesFullPath=!1;isPureNumber(processName)?query+=" WHERE ProcessID="+String(processName):isSolidString(processName)&&(query+=' WHERE Caption="'+path.basename(processName)+'"',comparesFullPath=path.isAbsolute(processName));var sWbemObjSets=os.WMI.execQuery(query);if(0===sWbemObjSets.length)return[];var matchWords=obtain(options,"matchWords",[]);isSolidString(matchWords)?matchWords=[matchWords]:isArray(matchWords)||(matchWords=[]);var excludingWords=obtain(options,"excludingWords",[]);return isSolidString(excludingWords)?excludingWords=[excludingWords]:isArray(excludingWords)||(excludingWords=[]),sWbemObjSets.filter(function(sWbemObjSet){return!(comparesFullPath&&!isSameMeaning(sWbemObjSet.ExecutablePath,processName))&&(!matchWords.some(function(word){if(!includes(sWbemObjSet.CommandLine,word,"i"))return!0})&&!excludingWords.some(function(word){if(includes(sWbemObjSet.CommandLine,word,"i"))return!0}))})},os.WMI.getProcess=function(processName,options){isPureNumber(processName)||isString(processName)||throwErrNonStr("os.WMI.getProcess",processName);var sWbemObjSets=os.WMI.getProcesses(processName,options);return 0===sWbemObjSets.length?null:sWbemObjSets[0]},os.WMI.getWithSWbemPath=function(sWbemPath,options){isSolidString(sWbemPath)||throwErrNonStr("os.WMI.getWithSWbemPath",sWbemPath);var compName=obtain(options,"compName",null),domainName=obtain(options,"domainName","."),namespace=obtain(options,"namespace","\\\\"+domainName+"\\root\\CIMV2"),wbem=WScript.CreateObject("WbemScripting.SWbemLocator");try{var sWbemObjSet=wbem.ConnectServer(compName,namespace).Get(sWbemPath);return sWbemObjSet?sWbemObjSet:null}catch(e){throw new Error(insp(e)+"\n  at os.WMI.getWithSWbemPath (WshOS/WMI.js)\n  query: "+sWbemPath+"\n  options: "+insp(options))}},os.WMI.getWindowsUserAccounts=function(){return os.WMI.execQuery("SELECT * FROM Win32_UserAccount")},os._thisProcessID=undefined,os._thisProcessSWbemObjSet=undefined,os._thisParentProcessID=undefined,os.WMI.getThisProcess=function(){if(os._thisProcessSWbemObjSet!==undefined)return os._thisProcessSWbemObjSet;var tmpJsPath=os.writeTempText("while(true) WScript.Sleep(1000);",".js");os.run(os.exefiles.wscript,[tmpJsPath]);var sWbemObjSet=os.WMI.getProcess(os.exefiles.wscript,{matchWords:[tmpJsPath]});return os._thisProcessID=Number(sWbemObjSet.ParentProcessId),sWbemObjSet.Terminate(),fso.DeleteFile(tmpJsPath,CD.fso.force.yes),os._thisProcessSWbemObjSet=os.WMI.getProcess(os._thisProcessID),os._thisParentProcessID=os._thisProcessSWbemObjSet.ParentProcessId,os._thisProcessSWbemObjSet})}();
!function(){var sh=Wsh.Shell,CD=Wsh.Constants,util=Wsh.Util,fso=Wsh.FileSystemObject,path=Wsh.Path,insp=util.inspect,isString=util.isString,isSolidString=util.isSolidString,includes=util.includes,os=Wsh.OS,throwErrNonStr=function(functionName,typeErrVal){util.throwTypeError("string","WshOS/System.js",functionName,typeErrVal)};os.EOL="\r\n",os.platform=function(){return"win32"},os._cwd=sh.CurrentDirectory;var SYSDIR=String(fso.GetSpecialFolder(CD.folderSpecs.system));os.exefiles={cmd:path.join(SYSDIR,"cmd.exe"),certutil:path.join(SYSDIR,"certutil.exe"),cscript:path.join(SYSDIR,"cscript.exe"),wscript:path.join(SYSDIR,"wscript.exe"),net:path.join(SYSDIR,"net.exe"),netsh:path.join(SYSDIR,"netsh.exe"),notepad:path.join(SYSDIR,"notepad.exe"),ping:path.join(SYSDIR,"PING.EXE"),schtasks:path.join(SYSDIR,"schtasks.exe"),timeout:path.join(SYSDIR,"timeout.exe"),xcopy:path.join(SYSDIR,"xcopy.exe")},os.getEnvVars=function(){for(var itm,environmentVars={},enmSet=new Enumerator(sh.Environment("PROCESS")),i=0;!enmSet.atEnd();enmSet.moveNext())itm=enmSet.item(),!/^=/.test(itm)&&includes(itm,"=")&&(i=itm.indexOf("="),environmentVars[itm.slice(0,i)]=itm.slice(i+1));return environmentVars},os.envVars=os.getEnvVars();var _osArchitecture=undefined;os.arch=function(){return _osArchitecture!==undefined||"x86"===(_osArchitecture=os.envVars.PROCESSOR_ARCHITECTURE.toLowerCase())&&isSolidString(os.envVars.PROCESSOR_ARCHITEW6432)&&(_osArchitecture=os.envVars.PROCESSOR_ARCHITEW6432.toLowerCase()),_osArchitecture};var _osIs64Arch=undefined;os.is64arch=function(){if(_osIs64Arch!==undefined)return _osIs64Arch;var arch=os.arch();return _osIs64Arch="amd64"===arch||"ia64"===arch||"x86"!==arch},os.tmpdir=function(){return os.envVars.TMP},os.makeTmpPath=function(prefix,postfix){var tmpName=fso.GetTempName();return isSolidString(prefix)||(prefix=""),isSolidString(postfix)||(postfix=""),path.join(os.tmpdir(),prefix+tmpName+postfix)},os.homedir=function(){return path.join(os.envVars.HOMEDRIVE,os.envVars.HOMEPATH)},os.hostname=function(){return os.envVars.COMPUTERNAME},os.type=function(){return os.envVars.OS},os.userInfo=function(){return{uid:-1,gid:-1,username:os.envVars.USERNAME,homedir:os.envVars.USERPROFILE,shell:null}},os.writeTempText=function(txtData,ext){isString(txtData)||throwErrNonStr("os.writeTempText",txtData),isSolidString(ext)||(ext="");var tmpTxtPath=os.makeTmpPath("osWriteTempWsh_",ext),tmpTxtObj=fso.CreateTextFile(tmpTxtPath,CD.fso.creates.yes);if(tmpTxtObj.Write(txtData),tmpTxtObj.Close(),!fso.FileExists(tmpTxtPath))throw new Error("Error: [Create TmpFile] "+tmpTxtPath+"\n  at os.writeTempText (WshOS/System.js)");return tmpTxtPath};var _osSysInfo=undefined;os.sysInfo=function(){if(_osSysInfo!==undefined)return _osSysInfo;var sWbemObjSets=os.WMI.execQuery("SELECT * FROM Win32_OperatingSystem"),wmiObj=os.WMI.toJsObject(sWbemObjSets[0]);return _osSysInfo=wmiObj},os.cmdCodeset=function(){return CD.codePageIdentifiers[os.sysInfo().CodeSet]},os.freemem=function(){return 1024*Number(os.sysInfo().FreePhysicalMemory)},os.release=function(){return os.sysInfo().Version},os.totalmem=function(){return 1024*Number(os.sysInfo().TotalVisibleMemorySize)};var _hasUAC=undefined;function _writeLogEvent(type,text){isString(text)||throwErrNonStr("_writeLogEvent",text);try{sh.LogEvent(type,text)}catch(e){throw new Error(insp(e)+"\n  at _writeLogEvent (WshOS/System.js)\n  Failed to write the log event.\n  text: "+text)}}os.hasUAC=function(){return _hasUAC!==undefined?_hasUAC:_hasUAC=(parseFloat(os.sysInfo().Version.slice(0,3)),!0)},os.isUacDisable=function(){var regUac="HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System";return sh.RegRead(regUac+"\\ConsentPromptBehaviorAdmin")+sh.RegRead(regUac+"\\PromptOnSecureDesktop")+sh.RegRead(regUac+"\\EnableLUA")===0},os.writeLogEvent={success:function(text){_writeLogEvent(CD.logLevels.success,text)},error:function(text){_writeLogEvent(CD.logLevels.error,text)},warn:function(text){_writeLogEvent(CD.logLevels.warning,text)},info:function(text){_writeLogEvent(CD.logLevels.infomation,text)}},os.setStrToClipboardAO=function(str){var dynwrap=new ActiveXObject("DynamicWrapper");dynwrap.register("ole32","OleInitialize","f=s","i=l","r=l"),dynwrap.OleInitialize(0),new ActiveXObject("htmlfile").parentWindow.clipboardData.setData("text",str)},os.getStrFromClipboard=function(){return new ActiveXObject("htmlfile").parentWindow.clipboardData.getData("text")}}();
!function(){var CD=Wsh.Constants,util=Wsh.Util,sh=Wsh.Shell,shApp=Wsh.ShellApplication,objAssign=Object.assign,insp=util.inspect,isArray=util.isArray,isNumber=util.isNumber,isPureNumber=util.isPureNumber,isString=util.isString,isSolidArray=util.isSolidArray,isSolidString=util.isSolidString,obtain=util.obtainPropVal,os=Wsh.OS,throwErrNonStr=function(functionName,typeErrVal){util.throwTypeError("string","WshOS/Exec.js",functionName,typeErrVal)};os.surroundPath=function(str){return isString(str)||throwErrNonStr("os.surroundPath",str),""===str?"":util.isASCII(str)&&!/[\s"&<>^|]/.test(str)||/^".*"$/.test(str)?str:'"'+str+'"'};var _srrPath=os.surroundPath;os.escapeForCmd=function(str){return isNumber(str)||isString(str)||throwErrNonStr("os.escapeForCmd",str),""===str?'""':isNumber(str)?String(str):util.isASCII(str)&&!/[\s"&<>^|]/.test(str)?str:'"'+str.replace(/(["])/g,"\\$1").replace(/([&<>^|])/g,"^$1")+'"'};var _escapeForCmd=os.escapeForCmd;os.joinCmdArgs=function(args,options){if(isSolidString(args))return String(args);if(!isSolidArray(args))return"";var escapes=obtain(options,"escapes",!0);return args.reduce(function(acc,arg){return!escapes||/^(1|2)?>{1,2}(&(1|2))?$/.test(arg)||"<"===arg||"|"===arg?acc+" "+arg:acc+" "+_escapeForCmd(arg)},"").trim()};var _joinCmdArgs=os.joinCmdArgs;function _shRun(command,options){isSolidString(command)||throwErrNonStr("_shRun",command);var winStyle=obtain(options,"winStyle","activeDef"),winStyleCode=isPureNumber(winStyle)?winStyle:CD.windowStyles[winStyle],waits=obtain(options,"waits",CD.waits.yes);try{return sh.Run(command,winStyleCode,waits)}catch(e){throw new Error(insp(e)+"\n  at _shRun (WshOS/Exec.js)\n  command: "+command)}}os.convToCmdCommand=function(cmdStr,args,options){isSolidString(cmdStr)||throwErrNonStr("os.convToCmdCommand",cmdStr);var command=_srrPath(cmdStr),argsStr=isArray(args)?_joinCmdArgs(args,options):isString(args)?args:"";""!==argsStr&&(command+=" "+argsStr);var shell=obtain(options,"shell",!1),closes=obtain(options,"closes",!0);return shell&&(command=closes?_srrPath(os.exefiles.cmd)+' /S /C"'+command+'"':_srrPath(os.exefiles.cmd)+' /S /K"'+command+'"'),command},os.exec=function(cmdStr,args,options){isString(cmdStr)||throwErrNonStr("os.exec",cmdStr);var command=os.convToCmdCommand(cmdStr,args,options);try{return sh.Exec(command)}catch(e){throw new Error(insp(e)+"\n  at os.exec (WshOS/Exec.js)\n  command: "+command)}},os.execSync=function(cmdStr,args,options){isString(cmdStr)||throwErrNonStr("os.execSync",cmdStr);var oExec,command=os.convToCmdCommand(cmdStr,args,options);isSolidString(command)||throwErrNonStr("os.execSync",command);var stdout="",stderr="";try{for(oExec=sh.Exec(command);0===oExec.Status;)stdout+=oExec.StdOut.Read(4096),stderr+=oExec.StdErr.Read(4096),WScript.Sleep(300);for(;!oExec.StdOut.AtEndOfStream;)stdout+=oExec.StdOut.Read(4096);for(;!oExec.StdOut.AtEndOfStream;)stdout+=oExec.StdErr.Read(4096)}catch(e){throw new Error(insp(e)+"\n  at os.execSync (WshOS/Exec.js)\n  command: "+command)}return{exitCode:oExec.ExitCode,stdout:stdout,stderr:stderr,error:isSolidString(stderr)}},os.run=function(cmdStr,args,options){return isString(cmdStr)||throwErrNonStr("os.run",cmdStr),_shRun(os.convToCmdCommand(cmdStr,args,options),objAssign({},options,{waits:!1}))},os.runSync=function(cmdStr,args,options){return isString(cmdStr)||throwErrNonStr("os.runSync",cmdStr),_shRun(os.convToCmdCommand(cmdStr,args,options),objAssign({},options,{waits:!0}))};var _isAdmin=undefined;os.isAdmin=function(){if(_isAdmin!==undefined)return _isAdmin;try{var iRetVal=os.runSync(os.exefiles.net,["session"],{winStyle:"hidden"});_isAdmin=iRetVal===CD.runs.ok}catch(e){_isAdmin=!1}return _isAdmin},os.runAsAdmin=function(cmdStr,args,options){if(os.isAdmin())return os.run(cmdStr,args,options);isString(cmdStr)||throwErrNonStr("os.runAsAdmin",cmdStr);var exePath=cmdStr,argsStr=isArray(args)?_joinCmdArgs(args,options):isString(args)?args:"";obtain(options,"shell",!1)&&(exePath=os.exefiles.cmd,argsStr='/S /C"'+_srrPath(cmdStr)+" "+argsStr+'"');var winStyle=obtain(options,"winStyle","activeDef"),winStyleCode=isPureNumber(winStyle)?winStyle:CD.windowStyles[winStyle];try{return shApp.ShellExecute(exePath,argsStr,"open","runas",winStyleCode)}catch(e){throw new Error(insp(e)+"\n  at _execFileAsAdmin (WshOS/Exec.js)\n  file: "+insp(exePath)+"\n  argsStr: "+insp(argsStr))}}}();
!function(){var CD=Wsh.Constants,util=Wsh.Util,fso=Wsh.FileSystemObject,path=Wsh.Path,objAssign=Object.assign,insp=util.inspect,isPureNumber=util.isPureNumber,isString=util.isString,os=Wsh.OS,MODULE_TITLE="WshOS/TaskScheduler.js",throwErrNonStr=function(functionName,typeErrVal){util.throwTypeError("string",MODULE_TITLE,functionName,typeErrVal)};os.Task={},os.Task.create=function(taskName,cmdStr,args,options){isString(taskName)||throwErrNonStr("os.Task.create",taskName);var mainCmd=os.exefiles.schtasks,argsStr='/Create /F /TN "'+taskName+'" /SC ONCE /ST 23:59 /IT /RL LIMITED /TR "'+os.convToCmdCommand(cmdStr,args,options).replace(/"/g,'\\"')+'"';try{var iRetVal=os.runSync(mainCmd,argsStr,objAssign({},options,{shell:!1,escapes:!1,winStyle:"hidden"}));if(iRetVal===CD.runs.ok)return;throw new Error('Error [ExitCode is not Ok] "'+iRetVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at os.Task.create ("+MODULE_TITLE+')\n  mainCmd: "'+mainCmd+'"\n  argsStr: "'+argsStr+'"')}},os.Task.exists=function(taskName){isString(taskName)||throwErrNonStr("os.Task.exists",taskName);var mainCmd=os.exefiles.schtasks,args=["/Query","/XML","/TN",taskName];try{return os.runSync(mainCmd,args,{shell:!1,winStyle:"hidden"})===CD.runs.ok}catch(e){throw new Error(insp(e)+"\n  at os.Task.exists ("+MODULE_TITLE+")")}},os.Task.run=function(taskName){isString(taskName)||throwErrNonStr("os.Task.run",taskName);var mainCmd=os.exefiles.schtasks,args=["/Run","/I","/TN",taskName];try{os.run(mainCmd,args,{shell:!1,winStyle:"hidden"})}catch(e){throw new Error(insp(e)+"\n  at os.Task.run ("+MODULE_TITLE+")")}},os.Task.del=function(taskName){if(isString(taskName)||throwErrNonStr("os.Task.del",taskName),os.Task.exists(taskName)){var mainCmd=os.exefiles.schtasks,args=["/Delete","/F","/TN",taskName];try{var iRetVal=os.runSync(mainCmd,args,{shell:!1,winStyle:"hidden"});if(iRetVal===CD.runs.ok)return;throw new Error('Error [ExitCode is not Ok] "'+iRetVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at os.Task.del ("+MODULE_TITLE+")")}}},os.Task.ensureToDelete=function(taskName,msecTimeOut){if(isString(taskName)||throwErrNonStr("os.Task.ensureToDelete",taskName),os.Task.exists(taskName)){msecTimeOut=isPureNumber(msecTimeOut)?msecTimeOut:1e4;do{try{return void os.Task.del(taskName)}catch(e){WScript.Sleep(100),msecTimeOut-=100}}while(0<msecTimeOut);throw new Error('Error [Delete Task] "'+taskName+'"\n  at os.Task.ensureToDelete ('+MODULE_TITLE+")")}},os.Task.ensureToCreate=function(taskName,mainCmd,args,options,msecTimeOut){isString(taskName)||throwErrNonStr("os.Task.ensureToCreate",taskName),os.Task.exists(taskName)&&os.Task.ensureToDelete(taskName),os.Task.create(taskName,mainCmd,args,options),msecTimeOut=isPureNumber(msecTimeOut)?msecTimeOut:1e4;do{try{if(os.Task.exists(taskName))return}catch(e){WScript.Sleep(100),msecTimeOut-=100}}while(0<msecTimeOut);throw new Error("Error: [Create the task(TimeOut)] "+taskName+"\n  at os.Task.ensureToCreate ("+MODULE_TITLE+")\n  mainCmd: "+mainCmd+"\n  args: "+insp(args))},os.Task.runTemporary=function(cmdStr,args,options){isString(cmdStr)||throwErrNonStr("os.Task.runTemporary",cmdStr);var tmpCode="WScript.CreateObject('WScript.Shell').Run('"+os.convToCmdCommand(cmdStr,args,options).replace(/\\{1}/g,"\\\\")+"', "+String(CD.windowStyles.hidden)+", "+String(CD.notWait)+");",tmpJsPath=os.writeTempText(tmpCode,".js"),taskName="Task_"+path.basename(os.makeTmpPath())+"_"+util.createDateString();os.Task.ensureToCreate(taskName,os.exefiles.wscript,[tmpJsPath],options),os.Task.run(taskName),os.Task.ensureToDelete(taskName),fso.DeleteFile(tmpJsPath,CD.fso.force.yes)},os.Task.createAsAdmin=function(){}}();
!function(){var sh=Wsh.Shell,CD=Wsh.Constants,util=Wsh.Util,fso=Wsh.FileSystemObject,insp=util.inspect,isPureNumber=util.isPureNumber,isString=util.isString,isSolidString=util.isSolidString,os=Wsh.OS,MODULE_TITLE="WshOS/Handler.js",throwErrNonStr=function(functionName,typeErrVal){util.throwTypeError("string",MODULE_TITLE,functionName,typeErrVal)};os.isExistingDrive=function(driveLetter){return isString(driveLetter)||throwErrNonStr("os.isExistingDrive",driveLetter),fso.DriveExists(driveLetter)},os.getDrivesInfo=function(){for(var drv,drvs=new Enumerator(fso.Drives),rtnDrvs=[];!drvs.atEnd();)(drv=drvs.item()).IsReady&&rtnDrvs.push({letter:drv.DriveLetter,name:drv.VolumeName,type:drv.DriveType}),drvs.moveNext();return rtnDrvs},os.getTheLastFreeDriveLetter=function(){for(var driveLetter,freeDriveLetter="",i=26;2<=i;i--)if(driveLetter=String.fromCharCode(64+i),!os.isExistingDrive(driveLetter)){freeDriveLetter=driveLetter;break}return freeDriveLetter},os.getHardDriveLetters=function(){return os.getDrivesInfo().filter(function(val){return val.type===CD.driveTypes.fixed})},os.assignDriveLetter=function(dirPath,letter,userName,password){isString(dirPath)||throwErrNonStr("os.assignDriveLetter",dirPath),isString(letter)||letter||throwErrNonStr("os.assignDriveLetter",letter),isSolidString(letter)||(letter=os.getTheLastFreeDriveLetter());var uncPath=dirPath;/^[a-z]:/i.test(dirPath)&&(uncPath=dirPath.replace(/^([a-z]):/i,"\\\\"+os.hostname()+"\\$1$"));var mainCmd=os.exefiles.net,args=["use",letter+":",uncPath];isSolidString(userName)&&isSolidString(password)&&(args=args.concat(["/user:"+userName,password]));try{var iRetVal=os.runSync(mainCmd,args,{winStyle:"hidden"});if(iRetVal!==CD.runs.ok)throw new Error('Error [ExitCode is not Ok] "'+iRetVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at os.assignDriveLetter ("+MODULE_TITLE+")")}return letter},os.removeAssignedDriveLetter=function(letter){isString(letter)||throwErrNonStr("os.removeAssignedDriveLetter",letter);var mainCmd=os.exefiles.net,args=["use",letter+":","/delete"];try{var iRetVal=os.runSync(mainCmd,args,{shell:!1,winStyle:"hidden"});if(iRetVal===CD.runs.ok)return;throw new Error('Error [ExitCode is not Ok] "'+iRetVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at os.removeAssignedDriveLetter ("+MODULE_TITLE+")")}},os.getProcessIDs=function(processName,options){return isPureNumber(processName)||isString(processName)||throwErrNonStr("os.getProcessIDs",processName),os.WMI.getProcesses(processName,options).map(function(sWbemObjSet){return sWbemObjSet.ProcessId})},os.getProcessObjs=function(processName,options){return isPureNumber(processName)||isString(processName)||throwErrNonStr("os.getProcessObjs",processName),os.WMI.getProcesses(processName,options).map(function(sWbemObjSet){return os.WMI.toJsObject(sWbemObjSet)})},os.getProcessObj=function(processName,options){isPureNumber(processName)||isString(processName)||throwErrNonStr("os.getProcessObj",processName);var sWbemObjSet=os.WMI.getProcess(processName,options);return os.WMI.toJsObject(sWbemObjSet)},os.activateProcess=function(processName,options){isPureNumber(processName)||isSolidString(processName)||throwErrNonStr("os.activateProcess",processName);var sWbemObjSet=os.WMI.getProcess(processName,options);return sWbemObjSet?(sh.AppActivate(sWbemObjSet.ProcessId),sWbemObjSet.ProcessId):null},os.terminateProcesses=function(processName,options){isPureNumber(processName)||isSolidString(processName)||throwErrNonStr("os.terminateProcesses",processName);var sWbemObjSets=os.WMI.getProcesses(processName,options);0!==sWbemObjSets.length&&sWbemObjSets.forEach(function(sWbemObjSet){sWbemObjSet.Terminate()})},os.getThisProcessID=function(){return os._thisProcessID!==undefined?os._thisProcessID:os.WMI.getThisProcess().ProcessId},os.getThisParentProcessID=function(){return os._thisParentProcessID!==undefined?os._thisParentProcessID:os.WMI.getThisProcess().ParentProcessId},os.addUser=function(userName,password){isString(userName)||throwErrNonStr("os.addUser",userName);var mainCmd=os.exefiles.net,args=["user",userName,password,"/ADD","/Y"];try{var iRetVal=os.runSync(mainCmd,args,{shell:!1,winStyle:"hidden"});if(iRetVal===CD.runs.ok)return;throw new Error('Error [ExitCode is not Ok] "'+iRetVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at os.addUser ("+MODULE_TITLE+")")}},os.attachAdminAuthorityToUser=function(userName){isString(userName)||throwErrNonStr("os.attachAdminAuthorityToUser",userName);var mainCmd=os.exefiles.net,args=["localgroup","Administrators",userName,"/ADD","/Y"];try{var iRetVal=os.runSync(mainCmd,args,{shell:!1,winStyle:"hidden"});if(iRetVal===CD.runs.ok)return;throw new Error('Error [ExitCode is not Ok] "'+iRetVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at os.attachAdminAuthorityToUser ("+MODULE_TITLE+")")}},os.deleteUser=function(userName){isString(userName)||throwErrNonStr("os.deleteUser",userName);var mainCmd=os.exefiles.net,args=["user",userName,"/DELETE","/Y"];try{var iRetVal=os.runSync(mainCmd,args,{shell:!1,winStyle:"hidden"});if(iRetVal===CD.runs.ok)return;throw new Error('Error [ExitCode is not Ok] "'+iRetVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at os.deleteUser ("+MODULE_TITLE+")")}}}();
