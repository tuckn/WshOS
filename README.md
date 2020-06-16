# WshOS

Adds some useful functions that handles Windows OS into WSH (Windows Script Host).

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

Now _.\MyScript.js_ (JScript ) can use the useful functions to handle paths.
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
console.log(retObj2);
// Returns: {
//   exitCode: 0,
//   stdout: <The result of ping 127.0.0.1>,
//   stderr: '',
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

And you can also use all functions of [WshPolyfill](https://github.com/tuckn/WshPolyfill), [WshUtil](https://github.com/tuckn/WshUtil) and [WshPath](https://github.com/tuckn/WshPath).
for example,

### WshPolyfill

```js
var array1 = [1, 4, 9, 16];
var map1 = array1.map(function (x) {
  return x * 2;
});

console.dir(map1);
// Output: [2, 8, 18, 32]

var strJson = JSON.stringify({ from: array1, to: map1 });
console.log(strJson);
// Output: '{"from":[1,4,9,16],"to":[2,8,18,32]}'

// and so on...
```

### WshUtil

```js
var _ = Wsh.Util; // Shorthand

// Check deep strict equality
_.isEqual({ a: 'A', b: ['B'] }, { a: 'A', b: ['B'] }); // true
_.isEqual({ a: 'A', b: ['B'] }, { a: 'A', b: ['b'] }); // false

// Create a unique ID
_.uuidv4(); // '9f1e53ba-3f08-4c9d-91c7-ad4226312f40'

// Create a date string
_.createDateString(); // '20200528T065424+0900'
_.createDateString('yyyy-MM'); // '2020-05'

// 半角カナを全角に変換
_.toZenkakuKana('もぅﾏﾁﾞ無理。'); // 'もぅマヂ無理'

// and so on...
```

## WshPath

```js
var path = Wsh.Path; // Shorthand

path.dirname('C:\\My Data\\image.jpg'); // 'C:\\My Data'
path.basename('C:\\foo\\bar\\baz\\quux.html'); // 'quux.html'
path.extname('index.coffee.md'); // '.md'

path.parse('C:\\home\\user\\dir\\file.txt');
// Returns:
// { root: 'C:\\',
//   dir: 'C:\\home\\user\\dir',
//   base: 'file.txt',
//   ext: '.txt',
//   name: 'file' };

path.isAbsolute('C:\\My Data\\hoge.png'); // true
path.isAbsolute('bar\\baz'); // false
path.isAbsolute('.'); // false

path.normalize('C:\\Git\\mingw64\\lib\\..\\etc\\.gitconfig');
// Returns: 'C:\\Git\\mingw64\\etc\\.gitconfig'

path.join(['mingw64\\lib', '..\\etc', '.gitconfig']);
// Returns: 'mingw64\\etc\\.gitconfig'

path.resolve(['mingw64\\lib', '..\\etc', '.gitconfig']);
// Returns: '<Current Working Directory>\\mingw64\\etc\\.gitconfig'

path.toUNC('C:\\foo\\bar.baz');
// Returns: '\\\\<Your CompName>\\C$\\foo\\bar.baz'

var from = 'C:\\MyApps\\Paint\\Gimp';
var to = 'C:\\MyApps\\Converter\\ImageMagick\\convert.exe';
path.relative(from, to);
// Returns: '..\\..\\Converter\\ImageMagick\\convert.exe'

// and so on...
```

## Documentation

See all specifications [here](https://docs.tuckn.net/WshOS) and also below.

- [WshPolyfill](https://docs.tuckn.net/WshPolyfill)
- [WshUtil](https://docs.tuckn.net/WshUtil).
- [WshPath](https://docs.tuckn.net/WshPath).

## License

MIT

Copyright (c) 2020 [Tuckn](https://github.com/tuckn)
