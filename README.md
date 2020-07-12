# WshOS

Adds useful functions that handles Windows OS into WSH (Windows Script Host).

## tuckn/Wsh series dependency

[WshModeJs](https://github.com/tuckn/WshModeJs)  
└─ [WshNet](https://github.com/tuckn/WshNet)  
&emsp;└─ [WshChildProcess](https://github.com/tuckn/WshChildProcess)  
&emsp;&emsp;└─ [WshProcess](https://github.com/tuckn/WshProcess)  
&emsp;&emsp;&emsp;&emsp;└─ [WshFileSystem](https://github.com/tuckn/WshFileSystem)  
&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;└─ WshOS - This repository  
&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;└─ [WshPath](https://github.com/tuckn/WshPath)  
&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;└─ [WshUtil](https://github.com/tuckn/WshUtil)  
&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;└─ [WshPolyfill](https://github.com/tuckn/WshPolyfill)

The upper layer module can use all the functions of the lower layer module.

## Operating environment

Works on JScript in Windows.

## Installation

(1) Create a directory of your WSH project.

```console
D:\> mkdir MyWshProject
D:\> cd MyWshProject
```

(2) Download this ZIP and unzipping or Use bellowing `git` command.

```console
> git clone https://github.com/tuckn/WshOS.git ./WshModules/WshOS
or
> git submodule add https://github.com/tuckn/WshOS.git ./WshModules/WshOS
```

(3) Include _.\WshOS\dist\bundle.js_ into your .wsf file.
For Example, if your file structure is

```console
D:\MyWshProject\
├─ Run.wsf
├─ MyScript.js
└─ WshModules\
    └─ WshOS\
        └─ dist\
          └─ bundle.js
```

The content of above _Run.wsf_ is

```xml
<package>
  <job id = "run">
    <script language="JScript" src="./WshModules/WshOS/dist/bundle.js"></script>
    <script language="JScript" src="./MyScript.js"></script>
  </job>
</package>
```

I recommend this .wsf file encoding to be UTF-8 [BOM, CRLF].
This allows the following functions to be used in _.\MyScript.js_.

## Usage

Now _.\MyScript.js_ (JScript) can use the useful functions to handle paths.
for example,

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

os.surroundPath('C:\\Program Files');
// Returns: '"C:\\Program Files"'

os.escapeForCmd('/RegExp="^(A|The) $"');
// Returns: '"/RegExp=\\"^^(A^|The) $\\""'

var retObj2 = os.execSync('ping.exe', ['127.0.0.1']);
console.dir(retObj2);
// Outputs: {
//   exitCode: 0,
//   stdout: <The result of ping 127.0.0.1>,
//   stderr: ""
//   error: false };

os.run('notepad.exe', ['D:\\Test.txt'], { winStyle: 'activeMax' });

os.isAdmin(); // false
os.runAsAdmin('mklink', 'D:\\Temp-Symlink D:\\Temp', { shell: true });

os.Task.create('MyTask', 'wscript.exe', '//job:run my-task.wsf');
os.Task.exists('MyTask'); // true
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

Many other functions are added.
See the [documentation](https://docs.tuckn.net/WshOS) for more details.

And you can also use all functions of [tuckn/WshPolyfill](https://github.com/tuckn/WshPolyfill), [tuckn/WshUtil](https://github.com/tuckn/WshUtil) and [tuckn/WshPath](https://github.com/tuckn/WshPath).

## Documentation

See all specifications [here](https://docs.tuckn.net/WshOS) and also below.

- [WshPolyfill](https://docs.tuckn.net/WshPolyfill)
- [WshUtil](https://docs.tuckn.net/WshUtil)
- [WshPath](https://docs.tuckn.net/WshPath)

## License

MIT

Copyright (c) 2020 [Tuckn](https://github.com/tuckn)
