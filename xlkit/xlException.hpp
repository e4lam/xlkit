/// @file xlException.hpp
///
/// @brief Provides XLKIT_THROW() macro and xlException class
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

#ifndef XLKIT_XLEXCEPTON_HPP
#define XLKIT_XLEXCEPTON_HPP

#include <xlkit/xldebug.hpp>
#include <xlkit/xlversion.hpp>
#include <stdexcept>
#include <string>

namespace xlkit {
XLKIT_USE_VERSION_NAMESPACE
namespace XLKIT_VERSION_NAME {

/// A std::runtime_error subclass for exceptions thrown by the libary,
/// via XLKIT_THROW()
class xlException : public std::runtime_error {
  public:
	xlException(const std::string& what) : std::runtime_error(what) {
	}
};

/// Throws an xlException with the given string literal
#define XLKIT_THROW(MSG) \
			throw xlkit::xlException( \
					xlkit::detail::debugMsgS(__FILE__,__LINE__,__FUNCTION__,MSG)) \
			/**/

} // namespace XLKIT_VERSION_NAME
} // namespace xlkit

#endif // XLKIT_XLEXCEPTON_HPP
