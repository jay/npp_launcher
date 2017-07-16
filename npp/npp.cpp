/* LICENSE: MIT License
Copyright (C) 2017 Jay Satiro <raysatiro@yahoo.com>
https://github.com/jay/RunAsDesktopUser/blob/master/LICENSE
*/

#include <windows.h>

#include <sstream>
#include <vector>

#include "ExecInExplorer_Util.h"
#include "ExecInExplorer_Implementation.h"

using namespace std;

#define WHILE_FALSE \
  __pragma(warning(push)) \
  __pragma(warning(disable:4127)) \
  while(0) \
  __pragma(warning(pop))

#define USE_DEBUGMSG

#ifdef USE_DEBUGMSG
#define DEBUGMSG(streamargs) \
  do { \
    wstringstream ss_d_; \
    ss_d_ << "npp launcher: " << streamargs << "\n"; \
    OutputDebugStringW(ss_d_.str().c_str()); \
  } WHILE_FALSE;
#else
#define DEBUGMSG(streamargs)
#endif

void SwitchToNotepadPlusPlusWindow(void)
{
  VOID (WINAPI *SwitchToThisWindow)(HWND hWnd, BOOL fAltTab) =
    (VOID (WINAPI *)(HWND, BOOL))GetProcAddress(GetModuleHandleW(L"user32"),
                                                "SwitchToThisWindow");

  if(!SwitchToThisWindow)
    return;

  int interval = 100;
  int max_wait = 5000;

  /* Try to set Notepad++ as the foreground window */
  for(int i = 0; i < (max_wait / interval); ++i) {
    Sleep((DWORD)interval);

    HWND hFore = GetForegroundWindow();
    if(!hFore)
      continue;

    /* Get the class name of the foreground window. The received class name is
       truncated if the buffer is not big enough, so to ensure we don't receive
       a truncated class name that compares equal to classname the buffer must
       be sized 1 more than expected for countof classname. */
    const wchar_t classname[] = L"Notepad++";
    wchar_t buf[_countof(classname) + 1];
    if(GetClassNameW(hFore, buf, (int)_countof(buf)) &&
       !wcscmp(classname, buf))
      break;

    HWND hwnd = FindWindowW(classname, NULL);
    if(!hwnd)
      continue;

    Sleep(1);

    // if the foreground window has changed and is Notepad++ then nothing to do
    if(hwnd == GetForegroundWindow())
      break;

    // if the foreground window has changed then try again
    if(hFore != GetForegroundWindow())
      continue;

    DEBUGMSG("GetForegroundWindow: " << hFore);
    DEBUGMSG("Notepad++ window: " << hwnd);

    // if the foreground window sets properly then nothing to do
    if(SetForegroundWindow(hwnd)) {
      DEBUGMSG("SetForegroundWindow: TRUE");
      break;
    }
    else
      DEBUGMSG("SetForegroundWindow: FALSE");

    // steal focus: minimize and then restore in alt+tab style
    DEBUGMSG("Attempting to steal focus");
    ShowWindow(hwnd, SW_MINIMIZE);
    SwitchToThisWindow(hwnd, TRUE);
    /* there's no way to know for sure whether that method is going to succeed.
       the foreground window may be in transition due to above so it's possible
       GetForegroundWindow returns 0. or it's possible it just doesn't work at
       all. in any case to avoid flickering don't try it again. */
    break;
  }
}

int WINAPI wWinMain(HINSTANCE, HINSTANCE, const PWSTR cmdline, int)
{
  HRESULT hr;
  WCHAR *dir;

  if(!GetCurrentDir_ThreadUnsafe(&dir))
    return -1;

  hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
  if(SUCCEEDED(hr)) {
    hr = ShellExecInExplorerProcess(L"notepad++", cmdline, dir);
    CoUninitialize();
  }

  free(dir);
  dir = NULL;

  if(FAILED(hr))
    return (int)hr;

  /* When Notepad++ starts it should make itself the foreground window, even if
     it was already running in a mono instance. Due to the way we're executing
     via dispatch that may not be possible due to a lack of permissions. We'll
     try for a few seconds to steal focus and make Notepad++ the foreground
     window if it's not already. Any failure is not considered an error. */
  SwitchToNotepadPlusPlusWindow();
  return 0;
}
