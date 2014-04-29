/// @file xlutil.hpp
///
/// @brief XLKit utilities
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

#ifndef XLKIT_UTIL_HPP
#define XLKIT_UTIL_HPP

#include <xlkit/xlversion.hpp>

#include <string>

#include <stdio.h>
#include <stdarg.h>

namespace xlkit {
XLKIT_USE_VERSION_NAMESPACE
namespace XLKIT_VERSION_NAME {

#ifdef _MSC_VER
#define XLKIT_PUSH_DISABLE_WARN_DEPRECATION \
			__pragma(warning(push)) \
			__pragma(warning(disable:4996)) \
			/**/
#define XLKIT_POP_DISABLE_WARN_DEPRECATION \
			__pragma(warning(pop)) \
			/**/
#else
#define XLKIT_PUSH_DISABLE_WARN_DEPRECATION 
#define XLKIT_POP_DISABLE_WARN_DEPRECATION 
#endif

/// vsprintf() analog that returns an std::string
inline std::string
strprintfV(const char *fmt, va_list args) {
	static const size_t NBUF = 2048;
	std::string str(NBUF, 0);
	while (true) {
		XLKIT_PUSH_DISABLE_WARN_DEPRECATION
		int n = vsnprintf(&str[0], str.size(), fmt, args);
		XLKIT_POP_DISABLE_WARN_DEPRECATION
		if (n > 0 && n < (int)str.size()) {
			str.resize(n);
			break;
		}
		str.append(NBUF, 0);
	}
	return str;
}
/// sprintf() analog that returns an std::string
inline std::string
strprintf(const char *fmt, ...) {
	va_list args;
	va_start(args, fmt);
	std::string str = strprintfV(fmt, args);
	va_end(args);
	return str;
}

} // namespace XLKIT_VERSION_NAME
} // namespace xlkit

#endif // XLKIT_UTIL_HPP
