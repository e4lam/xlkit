/// @file xldebug.hpp
///
/// @brief XLKit debugging facility for the XLDBG() macro
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

#ifndef XLKIT_XLDEBUG_HPP
#define XLKIT_XLDEBUG_HPP

#include <xlkit/xlutil.hpp>
#include <xlkit/xlversion.hpp>
#include <string.h>

namespace xlkit {
XLKIT_USE_VERSION_NAMESPACE
namespace XLKIT_VERSION_NAME {

namespace detail {

void outputDebugString(const char *msg);

std::string
debugMsgV(const char* file, int n, const char* func, const char *fmt, va_list args) {
	const char* base = strrchr(file, '\\');
	if (!base)
		base = strrchr(file, '/');
	if (base)
		file = base + 1;
	std::string msg = strprintf("%s(%d) [%s]: ", file, n, func);
	msg.append(strprintfV(fmt, args));
	msg.append("\n");
	return msg;
}
std::string
debugMsg(const char* file, int n, const char* func, const char *fmt, ...) {
	va_list args;
	va_start(args, fmt);
	std::string msg = debugMsgV(file, n, func, fmt, args);
	va_end(args);
	return msg;
}
std::string
debugMsgS(const char* file, int n, const char* func, const std::string& msg) {
	return debugMsg(file, n, func, "%s", msg.c_str());
}
void
debugOut(const char* file, int n, const char* func, const char *fmt, ...) {
	va_list args;
	va_start(args, fmt);
	std::string msg = debugMsgV(file, n, func, fmt, args);
	va_end(args);
	detail::outputDebugString(msg.c_str());
}

} // namespace detail

/// @def XLDBG 
/// Provides a printf style debug output. When run inside Visual
/// Studio, it will print to the Output window. When run outside the debugger,
/// it will allocate a text console and output to it.
#ifdef _DEBUG
#define XLDBG(FORMAT, ...) \
				xlkit::detail::debugOut(__FILE__, __LINE__, __FUNCTION__, FORMAT, __VA_ARGS__) \
				/**/
#else
#define XLDBG(FORMAT, ...)
#endif

} // namespace XLKIT_VERSION_NAME
} // namespace xlkit

#endif // XLKIT_XLDEBUG_HPP
