# WshOS

Adds useful functions that handle Windows OS into WSH (Windows Script Host).

## tuckn/Wsh series dependency

[WshModeJs](https://github.com/tuckn/WshModeJs)  
└─ [WshZLIB](https://github.com/tuckn/WshZLIB)  
&emsp;└─ [WshNet](https://github.com/tuckn/WshNet)  
&emsp;&emsp;└─ [WshChildProcess](https://github.com/tuckn/WshChildProcess)  
&emsp;&emsp;&emsp;└─ [WshProcess](https://github.com/tuckn/WshProcess)  
&emsp;&emsp;&emsp;&emsp;&emsp;└─ [WshFileSystem](https://github.com/tuckn/WshFileSystem)  
&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;└─ WshOS - This repository  
&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;└─ [WshPath](https://github.com/tuckn/WshPath)  
&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;└─ [WshUtil](https://github.com/tuckn/WshUtil)  
&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;└─ [WshPolyfill](https://github.com/tuckn/WshPolyfill)

The upper layer module can use all the functions of the lower layer module.

## Operating environment

Works on JScript in Windows.

## Installation

(1) Create a directory of your WSH project.

```console
D:\> mkdir MyWshProject
D:\> cd MyWshProject
```

(2) Download this ZIP and unzip or Use the following `git` command.

```console
> git clone https://github.com/tuckn/WshOS.git ./WshModules/WshOS
or
> git submodule add https://github.com/tuckn/WshOS.git ./WshModules/WshOS
```

(3) Create your JScript (.js) file. For Example,

```console
D:\MyWshProject\
├─ MyScript.js <- Your JScript code will be written in this.
└─ WshModules\
    └─ WshOS\
        └─ dist\
          └─ bundle.js
```

I recommend JScript (.js) file encoding to be UTF-8 [BOM, CRLF].

(4) Create your WSF packaging scripts file (.wsf).

```console
D:\MyWshProject\
├─ Run.wsf <- WSH entry file
├─ MyScript.js
└─ WshModules\
    └─ WshOS\
        └─ dist\
          └─ bundle.js
```

And you should include _.../dist/bundle.js_ into the WSF file.
For Example, The content of the above _Run.wsf_ is

```xml
<package>
  <job id = "run">
    <script language="JScript" src="./WshModules/WshOS/dist/bundle.js"></script>
    <script language="JScript" src="./MyScript.js"></script>
  </job>
</package>
```

I recommend this WSH file (.wsf) encoding to be UTF-8 [BOM, CRLF].

Awesome! This WSH configuration allows you to use the following functions in JScript (_.\\MyScript.js_).

## Usage

Now your JScript (_.\\MyScript.js_ ) can use helper functions to handle paths.
For example,

```js
var os = Wsh.OS; // Shorthand

os.is64arch(); // true
os.tmpdir(); // 'C:\\Users\\YourUserName\\AppData\\Local\\Temp'
os.makeTmpPath(); // 'C:\\Users\\UserName\\AppData\\Local\\Temp\\rad6E884.tmp'
os.homedir(); // 'C:\\Users\\%UserName%'
os.hostname(); // 'MYPC0123'
os.cmdCodeset(); // 'shift_jis'

os.writeLogEvent.error('Error Log');
// Logs the error event in Windows Event Log.

os.surroundCmdArg('C:\\Program Files');
// Returns: '"C:\\Program Files"'
os.surroundCmdArg('C:\\Windows\\System32\\notepad.exe');
// Returns: 'C:\\Windows\\System32\\notepad.exe'

os.escapeForCmd('/RegExp="^(A|The) $"');
// Returns: '"/RegExp=\\"^^(A^|The) $\\""'

// Executer1: Asynchronously run.
os.shRun('notepad.exe', ['D:\\Test.txt'], { winStyle: 'activeMax' });
// You can control a window style. but can not receive a return value.

// Executer2: Synchronously run.
os.shRunSync('notepad.exe', ['D:\\Test.txt'], { winStyle: 'activeMin' });
// You can receive a return number but can not receive StdOut.

// Executer3: Not wait for finished the process
var oExec1 = os.shExec('dir', 'D:\\Temp1', { shell: true });
// You can control the application later.
while (oExec1.Status == 0) WScript.Sleep(300); // Waiting the finished
var result = oExec2.StdOut.ReadAll();
console.log(result); // Outputs the result of dir

// Executer4: Wait for finished the process
var retObj2 = os.shExecSync('ping.exe', ['127.0.0.1']);
console.dir(retObj2);
// You can receive a return number, command and, also StdOut.
// Outputs: {
//   exitCode: 0,
//   command: "ping.exe 127.0.0.1",
//   stdout: <The result of ping 127.0.0.1>,
//   stderr: ""
//   error: false };

os.isAdmin(); // false
os.runAsAdmin('mklink', 'D:\\Temp-Symlink D:\\Temp', { shell: true });

os.Task.create('MyTask', 'wscript.exe', '//job:run my-task.wsf');
os.Task.exists('MyTask'); // true
os.Task.xmlString('MyTask'); // Returns: The task XML string
os.Task.run('MyTask');
os.Task.del('MyTask');

os.assignDriveLetter('\\\\MyNAS\\MultiMedia', 'M', 'myname', 'mY-p@ss');
// Return 'M'

var pIDs = os.getProcessIDs('Chrome.exe');
// Returns: [33221, 22044, 43113, 42292, 17412]
var pIDs = os.getProcessIDs('C:\\Program Files\\Git\\bin\\git.exe');
// Returns: [1732, 4316]

os.addUser('MyUserName', 'mY-P@ss');
os.attachAdminAuthorityToUser('MyUserName');
os.deleteUser('MyUserName');

// and so on...
```

Many other functions will be added.
See the [documentation](https://tuckn.net/docs/WshOS/) for more details.

### Dependency Modules

You can also use the following helper functions in your JScript (_.\\MyScript.js_).

- [tuckn/WshPolyfill](https://github.com/tuckn/WshPolyfill)
- [tuckn/WshUtil](https://github.com/tuckn/WshUtil)
- [tuckn/WshPath](https://github.com/tuckn/WshPath)

## Documentation

See all specifications [here](https://tuckn.net/docs/WshOS/) and also below.

- [WshPolyfill](https://tuckn.net/docs/WshPolyfill/)
- [WshUtil](https://tuckn.net/docs/WshUtil/)
- [WshPath](https://tuckn.net/docs/WshPath/)

## License

MIT

Copyright (c) 2020 [Tuckn](https://github.com/tuckn)
