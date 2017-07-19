/* LICENSE: MIT License
Copyright (C) 2017 Jay Satiro <raysatiro@yahoo.com>
https://github.com/jay/npp_launcher/blob/master/LICENSE
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
  } WHILE_FALSE
#else
#define DEBUGMSG(streamargs)
#endif


/*
Attempt to set the foreground window for a short period of time.

This function calls SetForegroundWindow one or more times for up to
dwMilliseconds. Each call is separated by a small interval of sleep.

Return TRUE (success) as soon as SetForegroundWindow returns TRUE; otherwise
return FALSE when dwMilliseconds has passed. Behavior may be modified by the
flags below.

Flags
-----

SFW_CONFIRM:
After a successful call to SetForegroundWindow do not consider it successful
or return early unless GetForegroundWindow == hWnd.

SFW_ENFORCE:
Set and enforce hWnd as the foreground window by calling SetForegroundWindow
when GetForegroundWindow != hwnd, for the entire dwMilliseconds.
Return the result of the last GetForegroundWindow == hWnd.
*/
BOOL SetForegroundWindowTryHarder(HWND hWnd, DWORD dwMilliseconds,
                                  DWORD dwFlags = 0)
{
  const DWORD interval = 100;

#define SFW_CONFIRM  (1<<0)
#define SFW_ENFORCE  (1<<1)

  DWORD start = GetTickCount();
  BOOL ret = 0;
  DWORD gle = 0;

  SetLastError(0);

  for(DWORD elapsed = 0; elapsed <= dwMilliseconds;
      elapsed = GetTickCount() - start) {
    /* GetTickCount operates at the resolution of the system clock and it may
       return the same result on subsequent calls. */
    DWORD remaining = dwMilliseconds - elapsed;

    if(!ret) {
      ret = SetForegroundWindow(hWnd);
      gle = GetLastError();
    }

    for(int i = 0; i < 2; ++i) {
      if((dwFlags & (SFW_CONFIRM | SFW_ENFORCE)))
        ret = (GetForegroundWindow() == hWnd);

      if(ret && !(dwFlags & SFW_ENFORCE))
        return TRUE;

      if(!i)
        Sleep(remaining < interval ? remaining : interval);
    }
  }

  SetLastError(gle);
  return ret;
}

/* Try to switch to Notepad++ for up to approximately dwMilliseconds + several
   seconds. */
void SwitchToNotepadPlusPlusWindow(void)
{
  const DWORD interval = 100;
  const DWORD dwMilliseconds = 5000;

  VOID (WINAPI *SwitchToThisWindow)(HWND hWnd, BOOL fAltTab) =
    (VOID (WINAPI *)(HWND, BOOL))GetProcAddress(GetModuleHandleW(L"user32"),
                                                "SwitchToThisWindow");

  DWORD start = GetTickCount();

  for(DWORD elapsed = 0; elapsed <= dwMilliseconds;
      elapsed = GetTickCount() - start) {
    /* GetTickCount operates at the resolution of the system clock and it may
       return the same result on subsequent calls. */
    DWORD remaining = dwMilliseconds - elapsed;

    Sleep(remaining < interval ? remaining : interval);

    HWND hFore = GetForegroundWindow();
    if(!hFore)
      continue;

    HWND hForeRootOwner = GetAncestor(hFore, GA_ROOTOWNER);
    if(!hForeRootOwner)
      continue;

    // Find the Notepad++ window as hwnd
    HWND hwnd;
    const wchar_t classname[] = L"Notepad++";
    /* The buffer that receives the class name must be 1 more than countof
       classname to ensure that a truncated class name won't compare equal. */
    wchar_t buf[_countof(classname) + 1];
    if(GetClassNameW(hForeRootOwner, buf, (int)_countof(buf)) &&
       !wcscmp(classname, buf)) {
      // Notepad++ has the foreground so stop if it's enabled and not minimized
      if(IsWindowEnabled(hFore) && !IsIconic(hFore) &&
         !IsIconic(hForeRootOwner))
        break;
      hwnd = hForeRootOwner;
    }
    else {
      hwnd = FindWindowW(classname, NULL);
      if(!hwnd)
        continue;

      hwnd = GetAncestor(hwnd, GA_ROOTOWNER);
      if(!hwnd)
        continue;
    }

    DEBUGMSG("GetForegroundWindow: " << hFore <<
             " (GA_ROOTOWNER: " << hForeRootOwner << ")");

    HWND hwndPop = GetWindow(hwnd, GW_ENABLEDPOPUP);

    DEBUGMSG("Notepad++ window: " << hwnd <<
             " (GW_ENABLEDPOPUP: " << hwndPop << ")");

    /* GW_ENABLEDPOPUP is documented to return hwnd if there's no popup, but
       I've found it can return NULL even if hwnd is enabled. */
    if(!hwndPop && IsWindowEnabled(hwnd))
      hwndPop = hwnd;

    /* Target an enabled Notepad++ window, if possible:
       Some enabled popups such as the "Create it?" message box disable hwnd.
       Since SetForegroundWindow can assign the focus to a disabled window we
       want to avoid that if possible by targeting an enabled window, but if
       not we'll use hwnd as a fallback: Notepad++ calls SetForegroundWindow on
       its main window (hwnd) even if it's disabled. */
    HWND hTarget = (IsWindowEnabled(hwnd) ? hwnd :
                    (hwndPop && IsWindowEnabled(hwndPop)) ? hwndPop :
                    hwnd);

    DEBUGMSG("Target window: " << hTarget);

    if(!hTarget)
      break;

    DEBUGMSG("IsWindowEnabled: " <<
             (IsWindowEnabled(hTarget) ? "TRUE" : "FALSE"));

    /* Enforce hTarget as the foreground window for 300ms.
       This is to remedy a race condition with Notepad++ where it may set the
       disabled main window of a previously existing mono-instance to the
       foreground while at the same time this program is trying to set the
       foreground to the Notepad++ enabled popup that has disabled the main
       window. */
    if(SetForegroundWindowTryHarder(hTarget, 300, SFW_ENFORCE)) {
      DEBUGMSG("SetForegroundWindowTryHarder SFW_ENFORCE: TRUE");
      if(!IsIconic(hwnd) && !IsIconic(hTarget))
        break;
    }
    else
      DEBUGMSG("SetForegroundWindowTryHarder SFW_ENFORCE: FALSE");

    // steal focus: minimize and then restore in alt+tab style.
    DEBUGMSG("Attempting to steal focus");

    if(!IsIconic(hwnd)) {
      BOOL b = ShowWindow(hwnd, SW_MINIMIZE);
      DEBUGMSG("ShowWindow SW_MINIMIZE: " << (b ? "TRUE" : "FALSE"));
    }

    if(SwitchToThisWindow) {
      SwitchToThisWindow(hwnd, TRUE);
      DEBUGMSG("SwitchToThisWindow");
    }

    /* Restore the window if necessary. SwitchToThisWindow should have already
       done this, however it's possible that it can't restore the window for
       some reason. For example I observe it doesn't always restore if hwnd is
       disabled (likely due to a popup), but SW_RESTORE does work. For details
       refer to "npp NOTES.txt" */
    if(!SwitchToThisWindow || IsIconic(hwnd)) {
      BOOL b = ShowWindow(hwnd, SW_RESTORE);
      DEBUGMSG("ShowWindow SW_RESTORE: " << (b ? "TRUE" : "FALSE"));
    }

    DEBUGMSG("IsIconic Notepad++ window: " <<
             (IsIconic(hwnd)? "TRUE" : "FALSE"));
    DEBUGMSG("IsIconic target window: " <<
             (IsIconic(hTarget)? "TRUE" : "FALSE"));
    DEBUGMSG("GetForegroundWindow: " << GetForegroundWindow());

    // Try for up to 2 seconds to make hTarget the foreground window
    if(SetForegroundWindowTryHarder(hTarget, 2000, SFW_CONFIRM))
      DEBUGMSG("SetForegroundWindowTryHarder SFW_CONFIRM: TRUE");
    else
      DEBUGMSG("SetForegroundWindowTryHarder SFW_CONFIRM: FALSE");

    // to avoid flickering don't try again
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
