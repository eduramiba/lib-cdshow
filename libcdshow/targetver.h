#pragma once

// Target Windows 7 and newer. The version must be selected before SDKDDKVer.h
// applies its defaults, otherwise the SDK defines a newer value first.
#include <WinSDKVer.h>
#define _WIN32_WINNT _WIN32_WINNT_WIN7
#include <SDKDDKVer.h>
