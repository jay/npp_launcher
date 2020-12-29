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


string HResultStr(HRESULT hr)
{
  switch(hr) {
  case 0x00000000: return "S_OK";
  case 0x00000001: return "S_FALSE";
  case 0x80004001: return "E_NOTIMPL";
  case 0x80004002: return "E_NOINTERFACE";
  case 0x80004003: return "E_POINTER";
  case 0x80004004: return "E_ABORT";
  case 0x80004005: return "E_FAIL";
  case 0x8000FFFF: return "E_UNEXPECTED";
  case 0x80070005: return "E_ACCESSDENIED";
  case 0x80070006: return "E_HANDLE";
  case 0x8007000E: return "E_OUTOFMEMORY";
  case 0x80070057: return "E_INVALIDARG";
  }

  stringstream ss;
  ss << hex << "0x" << hr;
  return ss.str();
}

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

HWND InitHelperWindow()
{
  const TCHAR *window_class_name =
    TEXT("npp launcher window {84C74316-2E90-44BE-808E-C27F57BDF77B}");

  WNDCLASS wc;
  wc.style         = CS_NOCLOSE;
  wc.lpfnWndProc   = DefWindowProc;
  wc.cbClsExtra    = 0;
  wc.cbWndExtra    = 0;
  wc.hInstance     = GetModuleHandle(NULL);
  wc.hIcon         = NULL;
  wc.hCursor       = NULL;
  wc.hbrBackground = NULL;
  wc.lpszMenuName  = NULL;
  wc.lpszClassName = window_class_name;

  ATOM atomWindowClass = RegisterClass(&wc);
  if(!atomWindowClass) {
    DWORD gle = GetLastError();
    DEBUGMSG("RegisterClass() failed to make window class " <<
             "\"" << window_class_name << "\"" <<
             " with error code " << gle << ".");
    return NULL;
  }

  HWND hwnd = CreateWindowEx(0, window_class_name, window_class_name,
                             WS_POPUP | WS_VISIBLE, 0, 0, 0, 0,
                             NULL, NULL, NULL, NULL);
  if(!hwnd) {
    DWORD gle = GetLastError();
    DEBUGMSG("CreateWindowEx() failed to make window " <<
             "\"" << window_class_name << "\"" <<
             " with error code " << gle << ".");
    return NULL;
  }

  return hwnd;
}

// StealFocus parameters
struct sfparam {
  // The window to switch to via ALT+TAB simulation
  HWND hSwitchToThis;

  /* If hSwitchToThis is disabled and is the root owner of an enabled popup
     then set the foreground to that popup instead of hSwitchToThis. */
#define SFP_IF_DISABLED_PREFER_ENABLED_POPUP (1<<0)
  DWORD dwFlags;

  // How many milliseconds to attempt stealing the focus before giving up
  DWORD dwMilliseconds;
};

string sfparam_flags_to_str(DWORD flags)
{
  stringstream ss;

#define EXTRACT_SFP_FLAG(f) \
  if((flags & f)) { \
    if(ss.tellp() > 0) \
      ss << " | "; \
    ss << #f; \
    flags &= ~f; \
  }

  EXTRACT_SFP_FLAG(SFP_IF_DISABLED_PREFER_ENABLED_POPUP);

  if(flags) {
    if(ss.tellp() > 0)
      ss << " | ";
    ss << "0x" << hex << flags;
  }

  if(ss.tellp() <= 0)
    ss << "<none>";

  return ss.str();
}

/*
Attempt to steal the focus.
Switch to window hSwitchToThis in struct sfparam and set the foreground window
to either that window or elsewhere depending on the flags.
Return nonzero if target was successfully set as or is the foreground window.
*/
DWORD WINAPI StealFocus(LPVOID lpParameter)
{
#define IF_FALSE_GOTO_CLEANUP(f) \
  if(!(f)) { \
    DEBUGMSG(#f ": FALSE"); \
    goto cleanup; \
  }

#define IF_HWND_BAD_GOTO_CLEANUP(h) \
  do { \
    IF_FALSE_GOTO_CLEANUP(IsWindow(h)); \
    IF_FALSE_GOTO_CLEANUP(IsWindowVisible(h)); \
    IF_FALSE_GOTO_CLEANUP(!IsHungAppWindow(h)); \
  } WHILE_FALSE

#define IF_VALIDATION_FAILS_GOTO_CLEANUP() \
  do { \
    IF_HWND_BAD_GOTO_CLEANUP(w->hSwitchToThis); \
    if(hForegroundToThis) { \
      IF_HWND_BAD_GOTO_CLEANUP(hForegroundToThis); \
      if(w->hSwitchToThis != GetAncestor(hForegroundToThis, GA_ROOTOWNER)) { \
        DEBUGMSG("w->hSwitchToThis no longer owns hForegroundToThis"); \
        goto cleanup; \
      } \
    } \
  } WHILE_FALSE

#define PROCESS_WINDOW_MESSAGES_NOSLEEP() \
  do { \
    for(MSG msg; PeekMessage(&msg, NULL, 0, 0, PM_REMOVE);) { \
      TranslateMessage(&msg); \
      DispatchMessage(&msg); \
    } \
  } WHILE_FALSE

#define PROCESS_WINDOW_MESSAGES() \
  do { \
    PROCESS_WINDOW_MESSAGES_NOSLEEP(); \
    Sleep(1); \
    PROCESS_WINDOW_MESSAGES_NOSLEEP(); \
  } WHILE_FALSE

  BOOL ret = FALSE;

  const struct sfparam *w = (struct sfparam *)lpParameter;

  if(!w) {
    DEBUGMSG("SetFocus invalid parameter");
    return FALSE;
  }

  DEBUGMSG("Attempting to steal focus");
  DEBUGMSG("w->hSwitchToThis: " << w->hSwitchToThis);
  DEBUGMSG("w->dwFlags: " << sfparam_flags_to_str(w->dwFlags).c_str());
  DEBUGMSG("w->dwMilliseconds: " << w->dwMilliseconds);

  HWND hHelper = InitHelperWindow();

  if(!hHelper) {
    DEBUGMSG("SetFocus helper window failed initialization");
    return FALSE;
  }

  DEBUGMSG("hHelper: " << hHelper);

  DWORD start = GetTickCount();
  DWORD max_wait = w->dwMilliseconds;

  // Sleep(interval) at the beginning of each iteration
  DWORD interval = 100;

  // The window to set to the foreground, unknown at this point
  HWND hForegroundToThis = NULL;

  IF_VALIDATION_FAILS_GOTO_CLEANUP();

  DebugDumpWindowInfo(w->hSwitchToThis);

  BOOL firstrun = TRUE;
  BOOL restored = FALSE;

  for(DWORD elapsed = 0; elapsed <= max_wait;
      elapsed = GetTickCount() - start) {
    /* GetTickCount operates at the resolution of the system clock and it may
       return the same result on subsequent calls. */
    DWORD remaining = max_wait - elapsed;

    if(firstrun) {
      firstrun = FALSE;
      PROCESS_WINDOW_MESSAGES();
    }
    else {
      PROCESS_WINDOW_MESSAGES_NOSLEEP();
      Sleep(remaining < interval ? remaining : interval);
      PROCESS_WINDOW_MESSAGES_NOSLEEP();
    }

    IF_VALIDATION_FAILS_GOTO_CLEANUP();

    // Check if the windows are in a temporarily bad/transition state
    if(!MonitorFromWindow(w->hSwitchToThis, MONITOR_DEFAULTTONULL) ||
       (hForegroundToThis &&
        !MonitorFromWindow(hForegroundToThis, MONITOR_DEFAULTTONULL)))
      continue;

    /* Restore the Notepad++ window if it is minimized. In practice this
       shouldn't be necessary because when Notepad++ is launched it restores
       its window. Typically if a restoration is needed SwitchToThisWindow will
       do it but that doesn't seem to work if the window is disabled, and in
       our case sometimes it is. */
    if(!restored) {
      if(IsIconic(w->hSwitchToThis)) {
        ShowWindow(w->hSwitchToThis, SW_RESTORE);
        DEBUGMSG("ShowWindow(w->hSwitchToThis, SW_RESTORE)");
        DebugDumpWindowInfo(w->hSwitchToThis);

        while(IsIconic(w->hSwitchToThis) ||
              !MonitorFromWindow(w->hSwitchToThis, MONITOR_DEFAULTTONULL)) {
          elapsed = GetTickCount() - start;

          if(elapsed > max_wait) {
            DEBUGMSG("Ran out of time while waiting for window to restore");
            goto cleanup;
          }

          remaining = max_wait - elapsed;

          PROCESS_WINDOW_MESSAGES_NOSLEEP();
          Sleep(remaining < interval ? remaining : interval);
          PROCESS_WINDOW_MESSAGES_NOSLEEP();

          IF_VALIDATION_FAILS_GOTO_CLEANUP();
        }

        DebugDumpWindowInfo(w->hSwitchToThis);
      }

      restored = TRUE;
      PROCESS_WINDOW_MESSAGES();
      IF_VALIDATION_FAILS_GOTO_CLEANUP();
    }

    if(!hForegroundToThis) {
      SwitchToThisWindow(w->hSwitchToThis, TRUE);
      DEBUGMSG("SwitchToThisWindow(w->hSwitchToThis, TRUE)");
      DebugDumpWindowInfo(w->hSwitchToThis);

      /* GW_ENABLEDPOPUP is documented to return the passed in hwnd if there's
         no popup, but I've found it can return NULL even if hwnd is enabled,
         or minimized and the popups are missing visibility bit. */
      HWND hPopup = GetWindow(w->hSwitchToThis, GW_ENABLEDPOPUP);
      DEBUGMSG("w->hSwitchToThis: " << w->hSwitchToThis <<
                " (GW_ENABLEDPOPUP: " << hPopup << ")");

      /* Target an enabled Notepad++ window, if possible:
         Some enabled popups such as the "Create it?" message box disable hwnd.
         Since SetForegroundWindow can assign the focus to a disabled window we
         want to avoid that if possible by targeting an enabled window. */
      if((w->dwFlags & SFP_IF_DISABLED_PREFER_ENABLED_POPUP) &&
         !IsWindowEnabled(w->hSwitchToThis) &&
         hPopup && !IsHungAppWindow(hPopup) &&
         IsWindowEnabled(hPopup) && IsWindowVisible(hPopup) &&
         MonitorFromWindow(hPopup, MONITOR_DEFAULTTONULL) &&
         w->hSwitchToThis == GetAncestor(hPopup, GA_ROOTOWNER)) {
        hForegroundToThis = hPopup;
      }
      else
        hForegroundToThis = w->hSwitchToThis;

      DEBUGMSG("hForegroundToThis: " << hForegroundToThis);
      DebugDumpWindowInfo(hForegroundToThis);

      PROCESS_WINDOW_MESSAGES();
      IF_VALIDATION_FAILS_GOTO_CLEANUP();
    }

    HWND hFore = GetForegroundWindow();
    DEBUGMSG("GetForegroundWindow: " << hFore <<
             (hFore == hHelper ? " (The foreground is hHelper)" :
                                 " (The foreground is NOT hHelper)"));

    DWORD lt;
    if(SystemParametersInfo(SPI_GETFOREGROUNDLOCKTIMEOUT, 0, &lt, 0)) {
      DEBUGMSG("SPI_GETFOREGROUNDLOCKTIMEOUT: " << lt <<
               " (" << lt / 1000 << " seconds)");
    }
    else
      DEBUGMSG("SPI_GETFOREGROUNDLOCKTIMEOUT: <unavailable>");

    ret = SetForegroundWindow(hForegroundToThis);
    DEBUGMSG("SetForegroundWindow(hForegroundToThis): " <<
              (ret ? "TRUE" : "FALSE"));
    if(ret)
      break;

    // Attempt to steal the foreground
    PROCESS_WINDOW_MESSAGES();
    ShowWindow(hHelper, SW_SHOWMINNOACTIVE);
    PROCESS_WINDOW_MESSAGES();
    ShowWindow(hHelper, SW_RESTORE);
    IF_VALIDATION_FAILS_GOTO_CLEANUP();
  }

cleanup:
  PROCESS_WINDOW_MESSAGES();
  DestroyWindow(hHelper);
  PROCESS_WINDOW_MESSAGES();

  return ret || GetForegroundWindow() == hForegroundToThis;
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
  const DWORD verification_delay = 500;
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

    DEBUGMSG("Notepad++ window: " << hwnd);

    // steal focus: attempt to bypass setforegroundwindow restrictions
    sfparam w = {};
    w.hSwitchToThis = hwnd;
    w.dwFlags = SFP_IF_DISABLED_PREFER_ENABLED_POPUP;
    w.dwMilliseconds = 2000;
    HANDLE hThread = CreateThread(NULL, 0, StealFocus, (LPVOID)&w, 0, NULL);
    if(!hThread) {
      DWORD gle = GetLastError();
      DEBUGMSG("CreateThread failed, error " << gle);
      break;
    }
    WaitForSingleObject(hThread, INFINITE);
    DWORD trc;
    if(!GetExitCodeThread(hThread, &trc))
      trc = FALSE;
    CloseHandle(hThread);
    DEBUGMSG("StealFocus: " << (trc ? "TRUE" : "FALSE"));

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

  if(!GetCurrentDir_ThreadUnsafe(&dir)) {
    DWORD gle = GetLastError();
    DEBUGMSG("GetCurrentDir_ThreadUnsafe failed, error " << gle);
    return -1;
  }

  hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
  if(SUCCEEDED(hr)) {
    hr = ShellExecInExplorerProcess(L"notepad++", cmdline, dir);
    if(FAILED(hr))
      DEBUGMSG("ShellExecInExplorerProcess notepad++ failed, error " <<
               HResultStr(hr).c_str() << ", dir: " << dir <<
               ", args: " << cmdline);
    CoUninitialize();
  }
  else
    DEBUGMSG("CoInitializeEx failed, error " << HResultStr(hr).c_str());

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
