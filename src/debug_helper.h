#pragma once

//
// debug_helper.h
// 
// Input #defines:
//   [Optional] PROJECT_NAME 
//   [Optional] _DEBUG
// 
// Output #defines:
//   PROJECT_NAME (if not already specified) 
//   DEBUGBREAK
//   HRCALL_NORET
//   HRCALL
// 
// Output #includes:
//   common.h
//   dbghelp.h
//
// Links:
//   dbghelp.lib
//
// Output symbols:
//   [Static] capture_callstack
//   [Static] msgbox_hresult

#include "common.h"
#include <stdio.h>

#ifndef PROJECT_NAME
#define PROJECT_NAME ""
#endif 

// Stack trace 
#include <dbghelp.h>
#pragma comment(lib, "dbghelp.lib")

#ifdef _DEBUG
#define DEBUGBREAK __debugbreak();
#else 
#define DEBUGBREAK 
#endif 

static void capture_callstack(wchar_t *callstack, size_t size) {
  void *stack[64];
  // Skip 2 -> Skip capture_callstack and msgbox
  unsigned short frames = RtlCaptureStackBackTrace(2, 64, stack, 0);

  HANDLE process = GetCurrentProcess();
  if (!SymInitialize(process, NULL, TRUE)) {
    wcscpy_s(callstack, size, L"Failed to initialize symbols.");
    return;
  }

  wchar_t buffer[256];
  size_t remaining = size;
  size_t symbol_size = sizeof(SYMBOL_INFOW) + 256 * sizeof(wchar_t);
  SYMBOL_INFOW *symbol = (SYMBOL_INFOW *)malloc(symbol_size);

  if (!symbol) {
    wcscpy_s(callstack, size, L"Failed to allocate memory for symbol info.");
    SymCleanup(process);
    return;
  }

  memset(symbol, 0, symbol_size);
  symbol->SizeOfStruct = sizeof(SYMBOL_INFOW);
  symbol->MaxNameLen = 255;

  callstack[0] = 0;

  for (unsigned short i = 0; i < frames; ++i) {
    DWORD64 address = (DWORD64)(stack[i]);
    DWORD64 displacement = 0;

    if (SymFromAddrW(process, address, &displacement, symbol)) {
      if (wcscmp(L"capture_callstack", symbol->Name) == 0 || wcscmp(L"msgbox_hresult", symbol->Name) == 0) {
        continue;
      }

      _snwprintf_s(buffer, sizeof(buffer) / sizeof(wchar_t), _TRUNCATE, L"Frame %d: %s (0x%p)\n", i, symbol->Name, stack[i]);
    }
    else {
      _snwprintf_s(buffer, sizeof(buffer) / sizeof(wchar_t), _TRUNCATE, L"Frame %d: Unknown symbol (0x%p)\n", i, stack[i]);
    }

    size_t len = wcslen(buffer);
    if (len + 1 > remaining) {
      _snwprintf_s(buffer, sizeof(buffer) / sizeof(wchar_t), _TRUNCATE, L"... Callstack truncated\n");
      wcscat_s(callstack, remaining, buffer);
      break;
    }

    wcscat_s(callstack, remaining, buffer);
    remaining -= len;
  }

  free(symbol);
  SymCleanup(process);
}

static void msgbox_hresult(HRESULT hr, const wchar_t *format, ...) {
	wchar_t error_msg[2048] = {};
	wchar_t system_message[512] = {};
	wchar_t custom_message[512] = {};
	wchar_t callstack[4096] = {};

	// Get the system error message for the HRESULT
	FormatMessageW(
		FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
		0,
		hr,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		system_message,
		sizeof(system_message) / sizeof(wchar_t),
		0);

	capture_callstack(callstack, sizeof(callstack) / sizeof(wchar_t));

	va_list args;
	va_start(args, format);
	vswprintf_s(custom_message, sizeof(custom_message) / sizeof(wchar_t), format, args);
	va_end(args);

	swprintf_s(
		error_msg,
		sizeof(error_msg) / sizeof(wchar_t),
		L"%s\nHRESULT: 0x%08X\n%s\nCallstack:\n%s",
		custom_message, hr, system_message, callstack
	);

	MessageBoxW(0, error_msg, L"[Error] " PROJECT_NAME, MB_ICONERROR | MB_OK);
}

#define HRCALL_NORET(x) { \
  HRESULT hr = x; \
  if (FAILED(hr)) { \
    msgbox_hresult(hr, L"File: %s\nLine: %d\nStatement: %s\n", __FILEW__, __LINE__, L#x); \
    DEBUGBREAK; \
  } \
}

#define HRCALL(x) { \
  HRESULT hr = x; \
  if (FAILED(hr)) { \
    msgbox_hresult(hr, L"File: %s\nLine: %d\nStatement: %s\n", __FILEW__, __LINE__, L#x); \
    DEBUGBREAK; \
    return hr; \
  } \
}
