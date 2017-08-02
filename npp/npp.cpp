/* LICENSE: MIT License
Copyright (C) 2017 Jay Satiro <raysatiro@yahoo.com>
https://github.com/jay/npp_launcher/blob/master/LICENSE
*/

#ifndef _CRT_SECURE_NO_WARNINGS
#define _CRT_SECURE_NO_WARNINGS
#endif

#include <windows.h>

#include <io.h>
#include <stdio.h>
#include <time.h>

#include <fstream>
#include <iomanip>
#include <sstream>
#include <vector>

#include "ExecInExplorer_Util.h"
#include "ExecInExplorer_Implementation.h"

using namespace std;

DWORD pid;
wofstream logfile;

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
    const wstring& str_d_ = ss_d_.str(); \
    OutputDebugStringW(str_d_.c_str()); \
    if(logfile) { \
      SYSTEMTIME st_d_ = { 0, }; \
      GetLocalTime(&st_d_); \
      logfile << "[" << SystemTimeStr(&st_d_).c_str() << "]: " \
              << "[" << pid << "] " << ss_d_.str(); \
      logfile.flush(); \
    } \
  } WHILE_FALSE
#else
#define DEBUGMSG(streamargs)
#endif


/* system time in format: Tue May 16 03:24:31.123 PM */
string SystemTimeStr(const SYSTEMTIME *t)
{
  const char *dow[] = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
  const char *mon[] = { NULL, "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
  stringstream ss;

  unsigned t_12hr = (t->wHour > 12 ? t->wHour - 12 : t->wHour ? t->wHour : 12);
  const char *t_ampm = (t->wHour < 12 ? "AM" : "PM");

  ss.fill('0');
  ss << dow[t->wDayOfWeek] << " "
     << mon[t->wMonth] << " "
     << setw(2) << t->wDay << " "
     << setw(2) << t_12hr << ":"
     << setw(2) << t->wMinute << ":"
     << setw(2) << t->wSecond << "."
     << setw(3) << t->wMilliseconds << " "
     << t_ampm;

  return ss.str();
}

void DebugDumpWindowInfo(HWND hWnd)
{
#ifndef USE_DEBUGMSG
  return;
#else
#define PRINTPOINT(point) \
  DEBUGMSG("HWND " << hWnd << " " #point << ": " << \
           "(" << point.x << ", " << point.y << ")");

#define PRINTRECT(rect) \
  DEBUGMSG("HWND " << hWnd << " " #rect << ": " << \
           "(" << rect.left << ", " << rect.top << ")-" << \
           "(" << rect.right << ", " << rect.bottom << ")");

  WINDOWPLACEMENT wp = {sizeof(wp),};
  if(GetWindowPlacement(hWnd, &wp)) {
    const char *showstr;
    switch(wp.showCmd) {
    case SW_SHOWMAXIMIZED: showstr = "SW_SHOWMAXIMIZED"; break;
    case SW_SHOWMINIMIZED: showstr = "SW_SHOWMINIMIZED"; break;
    case SW_SHOWNORMAL: showstr = "SW_SHOWNORMAL"; break;
    default: showstr = "<unknown>"; break;
    }
    DEBUGMSG("HWND " << hWnd << " wp.showCmd: " << showstr);
    PRINTPOINT(wp.ptMaxPosition);
    PRINTPOINT(wp.ptMinPosition);
    PRINTRECT(wp.rcNormalPosition);
  }

  RECT window_rect;
  if(GetWindowRect(hWnd, &window_rect))
    PRINTRECT(window_rect);

  RECT client_rect;
  if(GetClientRect(hWnd, &client_rect))
    PRINTRECT(client_rect);

  DEBUGMSG("HWND " << hWnd << " MonitorFromWindow: " <<
           MonitorFromWindow(hWnd, MONITOR_DEFAULTTONULL));
#endif // USE_DEBUGMSG
}

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
  BOOL beep = FALSE;
  const DWORD interval = 100;
  const DWORD dwMilliseconds = 5000;

  /* 'verification_delay' is the number of milliseconds to wait for the most
     recently found Notepad++ owner window to stay the same, so that Notepad++
     has a chance to load and put itself in the foreground on its own. This
     also makes focusing on an initially found window from an earlier
     multi-instance less likely. The delay should be short because if Notepad++
     doesn't put itself in the foreground then we want to step in before the
     user thinks they have to. */
  const DWORD verification_delay = 1000;
  HWND prev_hwnd = NULL;  /* The most recently found Notepad++ owner window */
  UINT prev_hwnd_showstate = 0;  /* The SW_xxx state when prev_hwnd was set */
  DWORD prev_hwnd_tickcount = 0;  /* The tickcount when prev_hwnd was set */

  /* The maximum wait time is a rough estimate and there are various waits that
     may push it over a little. At a minimum enough time is needed so that the
     loop can complete a full iteration to set the foreground window. */
  const DWORD max_wait = dwMilliseconds + (verification_delay * 2) + interval;

  DWORD start = GetTickCount();

  for(DWORD elapsed = 0; elapsed <= max_wait;
      elapsed = GetTickCount() - start) {
    /* GetTickCount operates at the resolution of the system clock and it may
       return the same result on subsequent calls. */
    DWORD remaining = max_wait - elapsed;

    /* for an iteration of this loop to complete successfully at the very least
       we'll need time to complete the verification delay, otherwise there's no
       point in continuing. */
    DWORD needed;

    if(prev_hwnd) {
      DWORD x = (GetTickCount() - prev_hwnd_tickcount);
      needed = (verification_delay > x ? verification_delay - x : 0);
    }
    else
      needed = verification_delay;

    if(remaining < needed)
      break;

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
      if(IsWindowEnabled(hFore) &&
         !IsIconic(hFore) && IsWindowVisible(hFore) &&
         MonitorFromWindow(hFore, MONITOR_DEFAULTTONULL) &&
         !IsIconic(hForeRootOwner) && IsWindowVisible(hForeRootOwner) &&
         MonitorFromWindow(hForeRootOwner, MONITOR_DEFAULTTONULL))
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

    WINDOWPLACEMENT wp = { sizeof(wp), };
    UINT hwnd_showstate = GetWindowPlacement(hwnd, &wp) ? wp.showCmd : 0;

    if(IsHungAppWindow(hwnd) || !IsWindowVisible(hwnd) ||
       !MonitorFromWindow(hwnd, MONITOR_DEFAULTTONULL) ||
       (hwnd == prev_hwnd && hwnd_showstate != prev_hwnd_showstate)) {
      prev_hwnd = NULL;
      prev_hwnd_showstate = 0;
      prev_hwnd_tickcount = 0;
      continue;
    }

    if(hwnd != prev_hwnd) {
      prev_hwnd = hwnd;
      prev_hwnd_showstate = hwnd_showstate;
      prev_hwnd_tickcount = GetTickCount();
    }

    if(verification_delay > (GetTickCount() - prev_hwnd_tickcount))
      continue;

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
       its main window (hwnd) even if it's disabled.

       At this point we already have checked that hwnd is visible, associated
       with a monitor and not hung, but hwndPop could be any of those
       things so don't assign it to target unless it meets the same. */
    HWND hTarget = (IsWindowEnabled(hwnd) ? hwnd :
                    (hwndPop && !IsHungAppWindow(hwndPop) &&
                     IsWindowEnabled(hwndPop) &&
                     MonitorFromWindow(hwndPop, MONITOR_DEFAULTTONULL) &&
                     IsWindowVisible(hwndPop)) ? hwndPop :
                    hwnd);

    DEBUGMSG("Target window: " << hTarget);

    if(!hTarget)
      break;

    DEBUGMSG("IsWindowEnabled: " <<
             (IsWindowEnabled(hTarget) ? "TRUE" : "FALSE"));

    DWORD lt = (DWORD)-1;
    if(SystemParametersInfo(SPI_GETFOREGROUNDLOCKTIMEOUT, 0, &lt, 0))
      DEBUGMSG("SPI_GETFOREGROUNDLOCKTIMEOUT: " << lt <<
               " (" << lt / 1000 << " seconds)");
    else
      DEBUGMSG("SPI_GETFOREGROUNDLOCKTIMEOUT: <unavailable>");

    /* Enforce hTarget as the foreground window for 100ms.
       This is to remedy a race condition with Notepad++ where it may set the
       disabled main window of a previously existing mono-instance to the
       foreground while at the same time this program is trying to set the
       foreground to the Notepad++ enabled popup that has disabled the main
       window. */
    if(SetForegroundWindowTryHarder(hTarget, 100, SFW_ENFORCE)) {
      DEBUGMSG("SetForegroundWindowTryHarder SFW_ENFORCE: TRUE");
      if(!IsIconic(hwnd) && !IsIconic(hTarget))
        break;
    }
    else
      DEBUGMSG("SetForegroundWindowTryHarder SFW_ENFORCE: FALSE");

    // steal focus: minimize and then restore in alt+tab style.
    DEBUGMSG("Attempting to steal focus");

    DebugDumpWindowInfo(hwnd);

    if(!IsIconic(hwnd)) {
      BOOL b = ShowWindow(hwnd, SW_MINIMIZE);
      DEBUGMSG("ShowWindow SW_MINIMIZE: " << (b ? "TRUE" : "FALSE"));
      DebugDumpWindowInfo(hwnd);
    }

    SwitchToThisWindow(hwnd, TRUE);
    DEBUGMSG("SwitchToThisWindow");
    DebugDumpWindowInfo(hwnd);

    /* Recover from race condition. At the same time we minimize the window
       it's possible Notepad++ restored itself, and the window manager gets
       confused and thinks it's not minimized even though it somewhat is since
       it's not visible on any monitor (-32000, -32000). In that case minimize
       the window again so that it can be restored properly. For details refer
       to "npp NOTES.txt" */
    if(!IsIconic(hwnd) && !MonitorFromWindow(hwnd, MONITOR_DEFAULTTONULL)) {
      DEBUGMSG("Minimizing again to recover from race condition");
      BOOL b = ShowWindow(hwnd, SW_MINIMIZE);
      DEBUGMSG("ShowWindow SW_MINIMIZE: " << (b ? "TRUE" : "FALSE"));
      DebugDumpWindowInfo(hwnd);
      //beep = TRUE;
    }

    /* Restore the window if necessary. SwitchToThisWindow should have already
       done this, however it's possible that it can't restore the window for
       some reason. For example I observe it doesn't always restore if hwnd is
       disabled (likely due to a popup), but SW_RESTORE does work. For details
       refer to "npp NOTES.txt" */
    if(IsIconic(hwnd) || !MonitorFromWindow(hwnd, MONITOR_DEFAULTTONULL)) {
      BOOL b = ShowWindow(hwnd, SW_RESTORE);
      DEBUGMSG("ShowWindow SW_RESTORE: " << (b ? "TRUE" : "FALSE"));
      DebugDumpWindowInfo(hwnd);
    }

    DEBUGMSG("IsIconic Notepad++ window: " <<
             (IsIconic(hwnd)? "TRUE" : "FALSE"));
    DEBUGMSG("IsIconic target window: " <<
             (IsIconic(hTarget)? "TRUE" : "FALSE"));
    DEBUGMSG("GetForegroundWindow: " << GetForegroundWindow());

    lt = (DWORD)-1;
    if(SystemParametersInfo(SPI_GETFOREGROUNDLOCKTIMEOUT, 0, &lt, 0))
      DEBUGMSG("SPI_GETFOREGROUNDLOCKTIMEOUT: " << lt <<
               " (" << lt / 1000 << " seconds)");
    else
      DEBUGMSG("SPI_GETFOREGROUNDLOCKTIMEOUT: <unavailable>");

    // Try for up to 2 seconds to make hTarget the foreground window
    if(SetForegroundWindowTryHarder(hTarget, 2000))
      DEBUGMSG("SetForegroundWindowTryHarder: TRUE");
    else {
      DEBUGMSG("SetForegroundWindowTryHarder: FALSE");

      /* 2 seconds have passed so check that the windows are still valid and
         then try a straight minimize and restore as a last resort */
      if(IsWindow(hTarget) && IsWindowVisible(hTarget) &&
         !IsHungAppWindow(hTarget) &&
         hwnd == GetAncestor(hTarget, GA_ROOTOWNER) &&
         IsWindowVisible(hwnd) && !IsHungAppWindow(hwnd)) {
        BOOL b = ShowWindow(hwnd, SW_MINIMIZE);
        DEBUGMSG("ShowWindow SW_MINIMIZE: " << (b ? "TRUE" : "FALSE"));
        DebugDumpWindowInfo(hwnd);

        if(b || IsIconic(hwnd)) {
          b = ShowWindow(hwnd, SW_RESTORE);
          DEBUGMSG("ShowWindow SW_RESTORE: " << (b ? "TRUE" : "FALSE"));
          DebugDumpWindowInfo(hwnd);
        }

        if(SetForegroundWindow(hTarget))
          DEBUGMSG("SetForegroundWindow: TRUE");
        else
          DEBUGMSG("SetForegroundWindow: FALSE");

        //beep = TRUE;
      }
    }

    DEBUGMSG("GetForegroundWindow: " << GetForegroundWindow());

    // to avoid flickering don't try again
    break;
  }

  if(beep)
    Beep(750, 300);
}

int WINAPI wWinMain(HINSTANCE, HINSTANCE, const PWSTR cmdline, int)
{
  HRESULT hr;
  WCHAR *dir;

  pid = GetCurrentProcessId();

  WCHAR tempdir[MAX_PATH + 1];
  DWORD rc = GetEnvironmentVariableW(L"TEMP", tempdir, _countof(tempdir));
  if(rc && rc < _countof(tempdir)) {
    wstring path(tempdir);
    if(*path.rbegin() != '\\' && *path.rbegin() != '/')
      path += '\\';
    path += L"npp_launcher.log";
    // if the logfile exists, open it for append
    if(path.length() < MAX_PATH && !_waccess_s(path.c_str(), 2))
      logfile.open(path, wofstream::app);
  }

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
