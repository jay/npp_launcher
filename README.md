npp launcher
============

npp launcher - Execute Notepad++ in the context of the user's explorer

### Usage

Usage: `npp [[arg]...]`  
Execute
[Notepad++](https://notepad-plus-plus.org/)
in the context of the user's explorer.

npp is useful if you are running as administrator and you want to execute
Notepad++ in the context of the user's explorer (eg it will de-elevate the
program if explorer isn't running as admin).

The execution takes place asynchronously by dispatch. The exit code returned
by npp is 0 if the command line to be executed was *dispatched* successfully or
otherwise an HRESULT indicating why it wasn't. There is no other interaction.

Other
-----

### Resources

This project uses
[my ExecInExplorer fork](https://github.com/jay/ExecInExplorer).

### Incorrect exit code when run from Windows Command Prompt?

npp uses the Windows subsystem not the console subsystem, which means that when
you execute it from the Windows command interpreter the interpreter runs it
asynchronously and therefore does not update %ERRORLEVEL% with the exit code.
However if you execute it from a batch file the interpreter will run it
synchronously and once it terminates its exit code will be put in %ERRORLEVEL%.
None of this has any bearing on how npp executes Notepad++, which is always
asynchronous.

### License

[MIT license](https://github.com/jay/npp_launcher/blob/master/LICENSE)

### Source

The source can be found on
[GitHub](https://github.com/jay/npp_launcher).
Since you're reading this maybe you're already there?

### Send me any questions you have

Jay Satiro `<raysatiro$at$yahoo{}com>` and put 'npp launcher' in the subject.
