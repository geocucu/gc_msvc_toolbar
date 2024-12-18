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
	SymInitialize(process, 0, TRUE);

	wchar_t buffer[256];
	SYMBOL_INFOW *symbol = (SYMBOL_INFOW *)malloc(sizeof(SYMBOL_INFOW) + 256 * sizeof(wchar_t));
	if (!symbol) {
		wcscpy_s(callstack, size, L"Failed to allocate memory for symbol info.");
		return;
	}

	symbol->SizeOfStruct = sizeof(SYMBOL_INFOW);
	symbol->MaxNameLen = 255;

	for (unsigned short i = 0; i < frames; ++i) {
		DWORD64 address = (DWORD64)(stack[i]);
		DWORD64 displacement = 0;

		if (SymFromAddrW(process, address, &displacement, symbol)) {
			// Skip capture_callstack and msgbox
			if (wcscmp(L"capture_callstack", symbol->Name) == 0 ||
				wcscmp(L"msgbox_hresult", symbol->Name) == 0) {
				continue;
			}

			swprintf_s(buffer, sizeof(buffer) / sizeof(wchar_t), L"Frame %d: %s (0x%p)\n", i, symbol->Name, stack[i]);
		}
		else {
			swprintf_s(buffer, sizeof(buffer) / sizeof(wchar_t), L"Frame %d: Unknown symbol (0x%p)\n", i, stack[i]);
		}

		wcscat_s(callstack, size, buffer);
	}

	free(symbol);
	SymCleanup(process);
}

static void msgbox_hresult(HRESULT hr, const wchar_t *format, ...) {
	wchar_t error_msg[2048] = {};
	wchar_t system_message[512] = {};
	wchar_t custom_message[512] = {};
	wchar_t callstack[2048] = {};

	// Get the system error message for the HRESULT
	FormatMessageW(
		FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
		0,
		hr,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		system_message,
		sizeof(system_message) / sizeof(wchar_t),
		0);

	// Capture the call stack
	capture_callstack(callstack, sizeof(callstack) / sizeof(wchar_t));

	// Format the custom message
	va_list args;
	va_start(args, format);
	vswprintf_s(custom_message, sizeof(custom_message) / sizeof(wchar_t), format, args);
	va_end(args);

	// Create the final error message
	swprintf_s(
		error_msg,
		sizeof(error_msg) / sizeof(wchar_t),
		L"%s\nHRESULT: 0x%08X\n%s\nCallstack:\n%s",
		custom_message, hr, system_message, callstack
	);

	// Display the error message
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
