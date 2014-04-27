/// @file xlcall.hpp
///
/// @brief Wrapper for XLCALL.H so that we don't need \#include <WINDOWS.H>
///
/// XLCALL.H is the only required file from the Excel XLL Software
/// Development Kit
///

// Copyright (c) 2014 Edward Lam
//
// All rights reserved. This software is distributed under the
// Mozilla Public License, v. 2.0 ( http://www.mozilla.org/MPL/2.0/ ).
//
// Redistributions of source code must retain the above copyright
// and license notice and the following restrictions and disclaimer.
//
// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
// "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
// LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
// A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
// OWNER OR CONTRIBUTORS BE LIABLE FOR ANY INDIRECT, INCIDENTAL,
// SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
// LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
// DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
// THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
// (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
// OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

#ifndef XLKIT_XLCALL_HPP
#define XLKIT_XLCALL_HPP

#include <xlkit/xlversion.hpp>
#include <stdint.h>

namespace xlkit {
XLKIT_USE_VERSION_NAMESPACE
namespace XLKIT_VERSION_NAME {
namespace detail {

struct xlPOINT {
	int32_t x;
	int32_t y;
};

} // namespace detail
} // namespace XLKIT_VERSION_NAME
} // namespace xlkit

// If WINDOWS.H has not been included, then define the types needed by XLCALL.H
#if !defined(_WINDOWS_H_) && !defined(_INC_WINDOWS)
#define XLCALL_HPP_WINDEFS
#define WINAPI		__stdcall
#define CALLBACK	__stdcall
#define pascal		__stdcall
#define VOID		void
#define INT32		int32_t
#define WCHAR		wchar_t
#define BYTE		uint8_t
#define WORD		uint16_t
#define SHORT		int16_t
#define DWORD		uint32_t
#define DWORD_PTR	uint32_t*
#define LONG		int32_t
#define LPSTR		char*
#define LPCSTR		const char*
#define HANDLE		void*
#define HWND		void*
#define POINT		xlkit::detail::xlPOINT
#endif

#include <XLCALL.H>

// Undo defines that were used just to get XLCALL.H included properly
#ifdef XLCALL_HPP_WINDEFS
#undef XLCALL_HPP_WINDEFS
#undef WINAPI
#undef CALLBACK
#undef pascal
#undef VOID
#undef INT32
#undef WCHAR
#undef BYTE
#undef WORD
#undef SHORT
#undef DWORD
#undef DWORD_PTR
#undef LONG
#undef LPSTR
#undef LPCSTR
#undef HANDLE
#undef HWND
#undef POINT
#endif

#endif // XLKIT_XLCALL_HPP
