﻿!function(){var CD,util,fso,path,insp,isArray,isPureNumber,isString,isSolidArray,isSolidString,isSameMeaning,obtain,includes,os,throwErrNonStr;Wsh&&Wsh.OS||(Wsh.ShellApplication||(Wsh.ShellApplication=WScript.CreateObject("Shell.Application")),Wsh.OS={},CD=Wsh.Constants,util=Wsh.Util,fso=Wsh.FileSystemObject,path=Wsh.Path,insp=util.inspect,isArray=util.isArray,isPureNumber=util.isPureNumber,isString=util.isString,isSolidArray=util.isSolidArray,isSolidString=util.isSolidString,isSameMeaning=util.isSameMeaning,obtain=util.obtainPropVal,includes=util.includes,os=Wsh.OS,throwErrNonStr=function(functionName,typeErrVal){util.throwTypeError("string","WshOS/WMI.js",functionName,typeErrVal)},os.WMI={},os.WMI.execQuery=function(query,options){var FN="os.WMI.execQuery",compName=(isSolidString(query)||throwErrNonStr(FN,query),obtain(options,"compName",null)),domainName=obtain(options,"domainName","."),domainName=obtain(options,"namespace","\\\\"+domainName+"\\root\\CIMV2"),wbem=WScript.CreateObject("WbemScripting.SWbemLocator");try{var swbemSet=wbem.ConnectServer(compName,domainName).ExecQuery(query),enumeSet=new Enumerator(swbemSet),sWbemObjSets=[];if(!sWbemObjSets)return[];for(;!enumeSet.atEnd();)sWbemObjSets.push(enumeSet.item()),enumeSet.moveNext();return sWbemObjSets}catch(e){throw new Error(insp(e)+"\n  at "+FN+" (WshOS/WMI.js)\n  query: "+query+"\n  options: "+insp(options))}},os.WMI.toJsObject=function(sWbemObjSet){for(var wmiObj={},enumedProps=new Enumerator(sWbemObjSet.Properties_);!enumedProps.atEnd();){var propItem=enumedProps.item(),val=propItem.Value;propItem.IsArray&&null!==val&&(val=new VBArray(val).toArray()),wmiObj[propItem.Name]=val,enumedProps.moveNext()}return wmiObj},os.WMI.toJsObjects=function(sWbemObjSets){return sWbemObjSets.map(function(sWbemObjSet){return os.WMI.toJsObject(sWbemObjSet)})},os.WMI.getProcesses=function(processName,options){var query="SELECT",selects=obtain(options,"selects",[]),selectQuery="";if(isSolidArray(selects)){includes(selects,"Caption")||selects.push("Caption"),includes(selects,"CommandLine")||selects.push("CommandLine"),includes(selects,"ExecutablePath")||selects.push("ExecutablePath"),includes(selects,"ProcessID")||selects.push("ProcessID");for(var i=0,len=selects.length-1;i<len;i++)selectQuery+=selects[i]+", ";selectQuery+=util.last(selects)}else selectQuery="*";query+=" "+selectQuery+" FROM Win32_Process",isPureNumber(processName)||isString(processName)||throwErrNonStr("os.WMI.getProcesses",processName);var comparesFullPath=!1,query=(isPureNumber(processName)?query+=" WHERE ProcessID="+String(processName):isSolidString(processName)&&(query+=' WHERE Caption="'+path.basename(processName)+'"',comparesFullPath=path.isAbsolute(processName)),os.WMI.execQuery(query));if(0===query.length)return[];var matchWords=obtain(options,"matchWords",[]),excludingWords=(isSolidString(matchWords)?matchWords=[matchWords]:isArray(matchWords)||(matchWords=[]),obtain(options,"excludingWords",[]));return isSolidString(excludingWords)?excludingWords=[excludingWords]:isArray(excludingWords)||(excludingWords=[]),query.filter(function(sWbemObjSet){return!(comparesFullPath&&!isSameMeaning(sWbemObjSet.ExecutablePath,processName))&&(!matchWords.some(function(word){if(!includes(sWbemObjSet.CommandLine,word,"i"))return!0})&&!excludingWords.some(function(word){if(includes(sWbemObjSet.CommandLine,word,"i"))return!0}))})},os.WMI.getProcess=function(processName,options){isPureNumber(processName)||isString(processName)||throwErrNonStr("os.WMI.getProcess",processName);processName=os.WMI.getProcesses(processName,options);return 0===processName.length?null:processName[0]},os.WMI.getWithSWbemPath=function(sWbemPath,options){var FN="os.WMI.getWithSWbemPath",compName=(isSolidString(sWbemPath)||throwErrNonStr(FN,sWbemPath),obtain(options,"compName",null)),domainName=obtain(options,"domainName","."),domainName=obtain(options,"namespace","\\\\"+domainName+"\\root\\CIMV2"),wbem=WScript.CreateObject("WbemScripting.SWbemLocator");try{var sWbemObjSet=wbem.ConnectServer(compName,domainName).Get(sWbemPath);return sWbemObjSet?sWbemObjSet:null}catch(e){throw new Error(insp(e)+"\n  at "+FN+" (WshOS/WMI.js)\n  query: "+sWbemPath+"\n  options: "+insp(options))}},os.WMI.getWindowsUserAccounts=function(){return os.WMI.execQuery("SELECT * FROM Win32_UserAccount")},os._thisProcessID=undefined,os._thisProcessSWbemObjSet=undefined,os._thisParentProcessID=undefined,os.WMI.getThisProcess=function(){if(os._thisProcessSWbemObjSet!==undefined)return os._thisProcessSWbemObjSet;var tmpJsPath=os.writeTempText("while(true) WScript.Sleep(1000);",".js"),sWbemObjSet=(os.run(os.exefiles.wscript,[tmpJsPath]),os.WMI.getProcess(os.exefiles.wscript,{matchWords:[tmpJsPath]}));try{os._thisProcessID=Number(sWbemObjSet.ParentProcessId)}catch(e){throw new Error(insp(e)+"\n  at os.WMI.getThisProcess (WshOS/WMI.js)")}return sWbemObjSet.Terminate(),fso.DeleteFile(tmpJsPath,CD.fso.force.yes),os._thisProcessSWbemObjSet=os.WMI.getProcess(os._thisProcessID),os._thisParentProcessID=os._thisProcessSWbemObjSet.ParentProcessId,os._thisProcessSWbemObjSet})}();
!function(){var sh=Wsh.Shell,CD=Wsh.Constants,util=Wsh.Util,fso=Wsh.FileSystemObject,path=Wsh.Path,insp=util.inspect,isString=util.isString,isSolidString=util.isSolidString,includes=util.includes,os=Wsh.OS,throwErrNonStr=function(functionName,typeErrVal){util.throwTypeError("string","WshOS/System.js",functionName,typeErrVal)},SYSDIR=(os.EOL="\r\n",os.platform=function(){return"win32"},os._cwd=sh.CurrentDirectory,String(fso.GetSpecialFolder(CD.folderSpecs.system))),_osArchitecture=(os.exefiles={cmd:path.join(SYSDIR,"cmd.exe"),certutil:path.join(SYSDIR,"certutil.exe"),cscript:path.join(SYSDIR,"cscript.exe"),wscript:path.join(SYSDIR,"wscript.exe"),net:path.join(SYSDIR,"net.exe"),netsh:path.join(SYSDIR,"netsh.exe"),notepad:path.join(SYSDIR,"notepad.exe"),ping:path.join(SYSDIR,"PING.EXE"),schtasks:path.join(SYSDIR,"schtasks.exe"),timeout:path.join(SYSDIR,"timeout.exe"),xcopy:path.join(SYSDIR,"xcopy.exe")},os.getEnvVars=function(){for(var itm,i,environmentVars={},enmSet=new Enumerator(sh.Environment("PROCESS"));!enmSet.atEnd();enmSet.moveNext())itm=enmSet.item(),!/^=/.test(itm)&&includes(itm,"=")&&(i=itm.indexOf("="),environmentVars[itm.slice(0,i)]=itm.slice(i+1));return environmentVars},os.envVars=os.getEnvVars(),undefined),_osIs64Arch=(os.arch=function(){return _osArchitecture!==undefined?_osArchitecture:_osArchitecture="x86"===(_osArchitecture=os.envVars.PROCESSOR_ARCHITECTURE.toLowerCase())&&isSolidString(os.envVars.PROCESSOR_ARCHITEW6432)?os.envVars.PROCESSOR_ARCHITEW6432.toLowerCase():_osArchitecture},undefined),_osSysInfo=(os.is64arch=function(){if(_osIs64Arch!==undefined)return _osIs64Arch;var arch=os.arch();return _osIs64Arch="amd64"===arch||"ia64"===arch||"x86"!==arch},os.tmpdir=function(){return os.envVars.TMP},os.makeTmpPath=function(prefix,postfix){var tmpName=fso.GetTempName();return isSolidString(prefix)||(prefix=""),isSolidString(postfix)||(postfix=""),path.join(os.tmpdir(),prefix+tmpName+postfix)},os.homedir=function(){return path.join(os.envVars.HOMEDRIVE,os.envVars.HOMEPATH)},os.hostname=function(){return os.envVars.COMPUTERNAME},os.type=function(){return os.envVars.OS},os.userInfo=function(){return{uid:-1,gid:-1,username:os.envVars.USERNAME,homedir:os.envVars.USERPROFILE,shell:null}},os.writeTempText=function(txtData,ext){var FN="os.writeTempText",ext=(isString(txtData)||throwErrNonStr(FN,txtData),isSolidString(ext)||(ext=""),os.makeTmpPath("osWriteTempWsh_",ext)),tmpTxtObj=fso.CreateTextFile(ext,CD.fso.creates.yes);if(tmpTxtObj.Write(txtData),tmpTxtObj.Close(),fso.FileExists(ext))return ext;throw new Error("Error: [Create TmpFile] "+ext+"\n  at "+FN+" (WshOS/System.js)")},undefined),_hasUAC=(os.sysInfo=function(){if(_osSysInfo!==undefined)return _osSysInfo;var sWbemObjSets=os.WMI.execQuery("SELECT * FROM Win32_OperatingSystem"),sWbemObjSets=os.WMI.toJsObject(sWbemObjSets[0]);return _osSysInfo=sWbemObjSets},os.cmdCodeset=function(){return CD.codePageIdentifiers[os.sysInfo().CodeSet]},os.freemem=function(){return 1024*Number(os.sysInfo().FreePhysicalMemory)},os.release=function(){return os.sysInfo().Version},os.totalmem=function(){return 1024*Number(os.sysInfo().TotalVisibleMemorySize)},undefined);function _writeLogEvent(type,text){var FN="_writeLogEvent";isString(text)||throwErrNonStr(FN,text);try{sh.LogEvent(type,text)}catch(e){throw new Error(insp(e)+"\n  at "+FN+" (WshOS/System.js)\n  Failed to write the log event.\n  text: "+text)}}os.hasUAC=function(){return _hasUAC!==undefined?_hasUAC:(parseFloat(os.sysInfo().Version.slice(0,3)),_hasUAC=!0)},os.isUacDisable=function(){var regUac="HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System";return sh.RegRead(regUac+"\\ConsentPromptBehaviorAdmin")+sh.RegRead(regUac+"\\PromptOnSecureDesktop")+sh.RegRead(regUac+"\\EnableLUA")===0},os.writeLogEvent={success:function(text){_writeLogEvent(CD.logLevels.success,text)},error:function(text){_writeLogEvent(CD.logLevels.error,text)},warn:function(text){_writeLogEvent(CD.logLevels.warning,text)},info:function(text){_writeLogEvent(CD.logLevels.infomation,text)}},os.setStrToClipboardAO=function(str){var dynwrap=new ActiveXObject("DynamicWrapper");dynwrap.register("ole32","OleInitialize","f=s","i=l","r=l"),dynwrap.OleInitialize(0),new ActiveXObject("htmlfile").parentWindow.clipboardData.setData("text",str)},os.getStrFromClipboard=function(){return new ActiveXObject("htmlfile").parentWindow.clipboardData.getData("text")}}();
!function(){var CD=Wsh.Constants,util=Wsh.Util,sh=Wsh.Shell,shApp=Wsh.ShellApplication,objAdd=Object.assign,insp=util.inspect,isArray=util.isArray,isNumber=util.isNumber,isPureNumber=util.isPureNumber,isString=util.isString,isSolidArray=util.isSolidArray,isSolidString=util.isSolidString,includes=util.includes,obtain=util.obtainPropVal,os=Wsh.OS,throwErrNonStr=function(functionName,typeErrVal){util.throwTypeError("string","WshOS/Exec.js",functionName,typeErrVal)},_srrPath=(os.surroundCmdArg=function(str){return isString(str)||throwErrNonStr("os.surroundCmdArg",str),""===str?"":/^".*"$/.test(str)||/^[A-Za-z0-9_>/=-]+(".+")?[A-Za-z0-9_&/=-]*$/.test(str)||util.isASCII(str)&&!includes(str," ")?str:'"'+str+'"'},os.surroundCmdArg),_escapeForCmd=(os.escapeForCmd=function(str){return isNumber(str)||isString(str)||throwErrNonStr("os.escapeForCmd",str),isNumber(str)?String(str):""===str||/^(1|2)?>{1,2}(&(1|2))?$/.test(str)||"<"===str||"|"===str?str:str.replace(/(["])/g,"\\$1").replace(/([&<>^|])/g,"^$1")},os.escapeForCmd),_joinCmdArgs=(os.joinCmdArgs=function(args,options){if(isSolidString(args))return String(args);if(!isSolidArray(args))return"";var escapes=obtain(options,"escapes",!0);return args.reduce(function(acc,arg){arg=escapes?_escapeForCmd(arg):arg;return acc+" "+os.surroundCmdArg(arg)},"").trim()},os.joinCmdArgs);function _shRun(command,options){var FN="_shRun";if(isSolidString(command)||throwErrNonStr(FN,command),obtain(options,"isDryRun",!1))return"dry-run [_shRun]: "+command;var winStyle=obtain(options,"winStyle","activeDef"),winStyle=isPureNumber(winStyle)?winStyle:CD.windowStyles[winStyle],options=obtain(options,"waits",CD.waits.yes);try{return sh.Run(command,winStyle,options)}catch(e){throw new Error(insp(e)+"\n  at "+FN+" (WshOS/Exec.js)\n  command: "+command)}}os.convToCmdlineStr=function(cmdStr,args,options){isSolidString(cmdStr)||throwErrNonStr("os.convToCmdlineStr",cmdStr);cmdStr=_srrPath(cmdStr),args=isArray(args)?_joinCmdArgs(args,options):isString(args)?args:"",""!==args&&(cmdStr+=" "+args),args=obtain(options,"shell",!1),options=obtain(options,"closes",!0);return cmdStr=args?options?_srrPath(os.exefiles.cmd)+' /S /C"'+cmdStr+'"':_srrPath(os.exefiles.cmd)+' /S /K"'+cmdStr+'"':cmdStr},os.exec=function(cmdStr,args,options){var FN="os.exec",cmdStr=(isSolidString(cmdStr)||throwErrNonStr(FN,cmdStr),os.convToCmdlineStr(cmdStr,args,options));if(isSolidString(cmdStr)||throwErrNonStr(FN,cmdStr),obtain(options,"isDryRun",!1))return"dry-run ["+FN+"]: "+cmdStr;try{return sh.Exec(cmdStr)}catch(e){throw new Error(insp(e)+"\n  at "+FN+" (WshOS/Exec.js)\n  command: "+cmdStr)}},os.execSync=function(cmdStr,args,options){var oExec,FN="os.execSync",cmdStr=(isSolidString(cmdStr)||throwErrNonStr(FN,cmdStr),os.convToCmdlineStr(cmdStr,args,options));if(isSolidString(cmdStr)||throwErrNonStr(FN,cmdStr),obtain(options,"isDryRun",!1))return"dry-run ["+FN+"]: "+cmdStr;var stdout="",stderr="";try{for(oExec=sh.Exec(cmdStr);0===oExec.Status;)stdout+=oExec.StdOut.Read(4096),stderr+=oExec.StdErr.Read(4096),WScript.Sleep(300);for(;!oExec.StdOut.AtEndOfStream;)stdout+=oExec.StdOut.Read(4096);for(;!oExec.StdOut.AtEndOfStream;)stdout+=oExec.StdErr.Read(4096)}catch(e){throw new Error(insp(e)+"\n  at "+FN+" (WshOS/Exec.js)\n  command: "+cmdStr)}return{exitCode:oExec.ExitCode,stdout:stdout,stderr:stderr,error:isSolidString(stderr)}},os.run=function(cmdStr,args,options){return _shRun(os.convToCmdlineStr(cmdStr,args,options),objAdd({},options,{waits:!1}))},os.runSync=function(cmdStr,args,options){return _shRun(os.convToCmdlineStr(cmdStr,args,options),objAdd({},options,{waits:!0}))};var _isAdmin=undefined;os.isAdmin=function(){if(_isAdmin!==undefined)return _isAdmin;try{var iRetVal=os.runSync(os.exefiles.net,["session"],{winStyle:"hidden"});_isAdmin=iRetVal===CD.runs.ok}catch(e){_isAdmin=!1}return _isAdmin},os.runAsAdmin=function(cmdStr,args,options){if(os.isAdmin())return os.run(cmdStr,args,options);var FN="os.runAsAdmin",exePath=(isSolidString(cmdStr)||throwErrNonStr(FN,cmdStr),cmdStr),args=isArray(args)?_joinCmdArgs(args,options):isString(args)?args:"";if(obtain(options,"shell",!1)&&(exePath=os.exefiles.cmd,args='/S /C"'+_srrPath(cmdStr)+" "+args+'"'),obtain(options,"isDryRun",!1))return"dry-run ["+FN+"]: "+exePath+" "+args;cmdStr=obtain(options,"winStyle","activeDef"),FN=isPureNumber(cmdStr)?cmdStr:CD.windowStyles[cmdStr];try{return shApp.ShellExecute(exePath,args,"open","runas",FN)}catch(e){throw new Error(insp(e)+"\n  at _execFileAsAdmin (WshOS/Exec.js)\n  file: "+insp(exePath)+"\n  argsStr: "+insp(args))}}}();
!function(){var CD=Wsh.Constants,util=Wsh.Util,fso=Wsh.FileSystemObject,path=Wsh.Path,objAdd=Object.assign,insp=util.inspect,isPureNumber=util.isPureNumber,isString=util.isString,obtain=util.obtainPropVal,includes=util.includes,os=Wsh.OS,WSCRIPT=os.exefiles.wscript,MODULE_TITLE="WshOS/TaskScheduler.js",throwErrNonStr=function(functionName,typeErrVal){util.throwTypeError("string",MODULE_TITLE,functionName,typeErrVal)};os.Task={},os.Task.create=function(taskName,cmdStr,args,options){var FN="os.Task.create",mainCmd=(isString(taskName)||throwErrNonStr(FN,taskName),os.exefiles.schtasks),cmdStr=os.convToCmdlineStr(cmdStr,args,options),args='/Create /F /TN "'+taskName+'" /SC ONCE /ST 23:59 /IT',taskName=obtain(options,"runsWithHighest",!1),cmdStr=(args=(args+=taskName?" /RL HIGHEST":" /RL LIMITED")+(' /TR "'+cmdStr.replace(/"/g,'\\"')+'"'),objAdd({},options,{shell:!1,escapes:!1,winStyle:"hidden"}));try{var retVal=taskName?os.runAsAdmin(mainCmd,args,cmdStr):os.runSync(mainCmd,args,cmdStr);if(obtain(options,"isDryRun",!1))return"dry-run ["+FN+"]: "+retVal;if(taskName||retVal===CD.runs.ok)return;throw new Error('Error [ExitCode is not Ok] "'+retVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at "+FN+" ("+MODULE_TITLE+')\n  mainCmd: "'+mainCmd+'"\n  argsStr: "'+args+'"')}},os.Task.exists=function(taskName,options){var FN="os.Task.exists",mainCmd=(isString(taskName)||throwErrNonStr(FN,taskName),os.exefiles.schtasks),taskName=["/Query","/XML","/TN",'"'+taskName+'"'],runOptions=objAdd({},options,{shell:!1,escapes:!1,winStyle:"hidden"});try{var retVal=os.runSync(mainCmd,taskName,runOptions);return obtain(options,"isDryRun",!1)?"dry-run ["+FN+"]: "+retVal:retVal===CD.runs.ok}catch(e){throw new Error(insp(e)+"\n  at "+FN+" ("+MODULE_TITLE+")")}},os.Task.xmlString=function(taskName,options){var FN="os.Task.xmlString",mainCmd=(isString(taskName)||throwErrNonStr(FN,taskName),os.exefiles.schtasks),taskName=["/Query","/XML","/TN",'"'+taskName+'"'],runOptions=objAdd({},options,{shell:!1,escapes:!1,winStyle:"hidden"});try{var retVal=os.execSync(mainCmd,taskName,runOptions);if(obtain(options,"isDryRun",!1))return"dry-run ["+FN+"]: "+retVal;if(!1===retVal.error)return retVal.stdout;throw new Error("Error: "+insp(retVal)+'"\n')}catch(e){throw new Error(insp(e)+"\n  at "+FN+" ("+MODULE_TITLE+")")}},os.Task.run=function(taskName,options){var FN="os.Task.run",mainCmd=(isString(taskName)||throwErrNonStr(FN,taskName),os.exefiles.schtasks),taskName=["/Run","/I","/TN",'"'+taskName+'"'],runOptions=objAdd({},options,{shell:!1,escapes:!1,winStyle:"hidden"});try{var retVal=os.run(mainCmd,taskName,runOptions);return obtain(options,"isDryRun",!1)?"dry-run ["+FN+"]: "+retVal:void 0}catch(e){throw new Error(insp(e)+"\n  at "+FN+" ("+MODULE_TITLE+")")}},os.Task.del=function(taskName,options){var FN="os.Task.del",isDryRun=(isString(taskName)||throwErrNonStr(FN,taskName),obtain(options,"isDryRun",!1));if(isDryRun||os.Task.exists(taskName)){var mainCmd=os.exefiles.schtasks,args=["/Delete","/F","/TN",'"'+taskName+'"'],options=objAdd({},options,{shell:!1,escapes:!1,winStyle:"hidden"}),taskName=os.Task.xmlString(taskName,{isDryRun:isDryRun}),taskName=includes(taskName,"<RunLevel>HighestAvailable</RunLevel>","i");try{var retVal=taskName?os.runAsAdmin(mainCmd,args,options):os.runSync(mainCmd,args,options);if(isDryRun)return"dry-run ["+FN+"]: "+retVal;if(taskName||retVal===CD.runs.ok)return;throw new Error('Error [ExitCode is not Ok] "'+retVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at "+FN+" ("+MODULE_TITLE+")")}}},os.Task.ensureToDelete=function(taskName,msecTimeOut,options){var FN="os.Task.ensureToDelete";if(isString(taskName)||throwErrNonStr(FN,taskName),os.Task.exists(taskName)){msecTimeOut=isPureNumber(msecTimeOut)?msecTimeOut:1e4;do{try{var retVal=os.Task.del(taskName,options);return obtain(options,"isDryRun",!1)?"dry-run ["+FN+"]: "+retVal:void 0}catch(e){WScript.Sleep(100),msecTimeOut-=100}}while(0<msecTimeOut);throw new Error('Error [Delete Task] "'+taskName+'"\n  at '+FN+" ("+MODULE_TITLE+")")}},os.Task.ensureToCreate=function(taskName,mainCmd,args,options,msecTimeOut){var FN="os.Task.ensureToCreate";isString(taskName)||throwErrNonStr(FN,taskName),os.Task.exists(taskName)&&os.Task.ensureToDelete(taskName),os.Task.create(taskName,mainCmd,args,options),msecTimeOut=isPureNumber(msecTimeOut)?msecTimeOut:1e4;do{try{var retVal=os.Task.exists(taskName,options);if(obtain(options,"isDryRun",!1))return"dry-run ["+FN+"]: "+retVal;if(retVal)return}catch(e){WScript.Sleep(100),msecTimeOut-=100}}while(0<msecTimeOut);throw new Error("Error: [Create the task(TimeOut)] "+taskName+"\n  at "+FN+" ("+MODULE_TITLE+")\n  mainCmd: "+mainCmd+"\n  args: "+insp(args))},os.Task.runTemporary=function(cmdStr,args,options){var FN="os.Task.runTemporary";isString(cmdStr)||throwErrNonStr(FN,cmdStr);var cmdStr="WScript.CreateObject('WScript.Shell').Run('"+os.convToCmdlineStr(cmdStr,args,options).replace(/\\{1}/g,"\\\\")+"', "+String(CD.windowStyles.hidden)+", "+String(CD.notWait)+");",args=os.writeTempText(cmdStr,".js"),cmdStr="Task_"+path.basename(os.makeTmpPath())+"_"+util.createDateString(),isDryRun=obtain(options,"isDryRun",!1),retLog="",retVal=os.Task.ensureToCreate(cmdStr,WSCRIPT,[args],options);if(isDryRun&&(retLog="dry-run ["+FN+"]: "+retVal),retVal=os.Task.run(cmdStr,options),isDryRun&&(retLog+="\n"+retVal),retVal="\n"+os.Task.ensureToDelete(cmdStr,options),isDryRun&&(retLog+="\n"+retVal),fso.DeleteFile(args,CD.fso.force.yes),isDryRun)return retLog}}();
!function(){var sh=Wsh.Shell,CD=Wsh.Constants,util=Wsh.Util,fso=Wsh.FileSystemObject,insp=util.inspect,isPureNumber=util.isPureNumber,isString=util.isString,isSolidString=util.isSolidString,os=Wsh.OS,MODULE_TITLE="WshOS/Handler.js",throwErrNonStr=function(functionName,typeErrVal){util.throwTypeError("string",MODULE_TITLE,functionName,typeErrVal)};os.isExistingDrive=function(driveLetter){return isString(driveLetter)||throwErrNonStr("os.isExistingDrive",driveLetter),fso.DriveExists(driveLetter)},os.getDrivesInfo=function(){for(var drv,drvs=new Enumerator(fso.Drives),rtnDrvs=[];!drvs.atEnd();)(drv=drvs.item()).IsReady&&rtnDrvs.push({letter:drv.DriveLetter,name:drv.VolumeName,type:drv.DriveType}),drvs.moveNext();return rtnDrvs},os.getTheLastFreeDriveLetter=function(){for(var driveLetter,freeDriveLetter="",i=26;2<=i;i--)if(driveLetter=String.fromCharCode(64+i),!os.isExistingDrive(driveLetter)){freeDriveLetter=driveLetter;break}return freeDriveLetter},os.getHardDriveLetters=function(){return os.getDrivesInfo().filter(function(val){return val.type===CD.driveTypes.fixed})},os.assignDriveLetter=function(dirPath,letter,userName,password){var FN="os.assignDriveLetter",uncPath=(isString(dirPath)||throwErrNonStr(FN,dirPath),isString(letter)||letter||throwErrNonStr(FN,letter),isSolidString(letter)||(letter=os.getTheLastFreeDriveLetter()),dirPath),dirPath=(/^[a-z]:/i.test(dirPath)&&(uncPath=dirPath.replace(/^([a-z]):/i,"\\\\"+os.hostname()+"\\$1$")),os.exefiles.net),uncPath=["use",letter+":",uncPath];isSolidString(userName)&&isSolidString(password)&&(uncPath=uncPath.concat(["/user:"+userName,password]));try{var iRetVal=os.runSync(dirPath,uncPath,{winStyle:"hidden"});if(iRetVal!==CD.runs.ok)throw new Error('Error [ExitCode is not Ok] "'+iRetVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at "+FN+" ("+MODULE_TITLE+")")}return letter},os.removeAssignedDriveLetter=function(letter){var FN="os.removeAssignedDriveLetter",mainCmd=(isString(letter)||throwErrNonStr(FN,letter),os.exefiles.net),letter=["use",letter+":","/delete"];try{var iRetVal=os.runSync(mainCmd,letter,{shell:!1,winStyle:"hidden"});if(iRetVal===CD.runs.ok)return;throw new Error('Error [ExitCode is not Ok] "'+iRetVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at "+FN+" ("+MODULE_TITLE+")")}},os.getProcessIDs=function(processName,options){return isPureNumber(processName)||isString(processName)||throwErrNonStr("os.getProcessIDs",processName),os.WMI.getProcesses(processName,options).map(function(sWbemObjSet){return sWbemObjSet.ProcessId})},os.getProcessObjs=function(processName,options){isPureNumber(processName)||isString(processName)||throwErrNonStr("os.getProcessObjs",processName);processName=os.WMI.getProcesses(processName,options);return os.WMI.toJsObjects(processName)},os.getProcessObj=function(processName,options){isPureNumber(processName)||isString(processName)||throwErrNonStr("os.getProcessObj",processName);processName=os.WMI.getProcess(processName,options);try{return os.WMI.toJsObject(processName)}catch(e){return null}},os.activateProcess=function(processName,options){isPureNumber(processName)||isSolidString(processName)||throwErrNonStr("os.activateProcess",processName);processName=os.WMI.getProcess(processName,options);return processName?(sh.AppActivate(processName.ProcessId),processName.ProcessId):null},os.terminateProcesses=function(processName,options){isPureNumber(processName)||isSolidString(processName)||throwErrNonStr("os.terminateProcesses",processName);processName=os.WMI.getProcesses(processName,options);0!==processName.length&&processName.forEach(function(sWbemObjSet){sWbemObjSet.Terminate()})},os.getThisProcessID=function(){return os._thisProcessID!==undefined?os._thisProcessID:os.WMI.getThisProcess().ProcessId},os.getThisParentProcessID=function(){return os._thisParentProcessID!==undefined?os._thisParentProcessID:os.WMI.getThisProcess().ParentProcessId},os.addUser=function(userName,password){var FN="os.addUser",mainCmd=(isString(userName)||throwErrNonStr(FN,userName),os.exefiles.net),userName=["user",userName,password,"/ADD","/Y"];try{var iRetVal=os.runSync(mainCmd,userName,{shell:!1,winStyle:"hidden"});if(iRetVal===CD.runs.ok)return;throw new Error('Error [ExitCode is not Ok] "'+iRetVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at "+FN+" ("+MODULE_TITLE+")")}},os.attachAdminAuthorityToUser=function(userName){var FN="os.attachAdminAuthorityToUser",mainCmd=(isString(userName)||throwErrNonStr(FN,userName),os.exefiles.net),userName=["localgroup","Administrators",userName,"/ADD","/Y"];try{var iRetVal=os.runSync(mainCmd,userName,{shell:!1,winStyle:"hidden"});if(iRetVal===CD.runs.ok)return;throw new Error('Error [ExitCode is not Ok] "'+iRetVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at "+FN+" ("+MODULE_TITLE+")")}},os.deleteUser=function(userName){var FN="os.deleteUser",mainCmd=(isString(userName)||throwErrNonStr(FN,userName),os.exefiles.net),userName=["user",userName,"/DELETE","/Y"];try{var iRetVal=os.runSync(mainCmd,userName,{shell:!1,winStyle:"hidden"});if(iRetVal===CD.runs.ok)return;throw new Error('Error [ExitCode is not Ok] "'+iRetVal+'"\n')}catch(e){throw new Error(insp(e)+"\n  at "+FN+" ("+MODULE_TITLE+")")}}}();
