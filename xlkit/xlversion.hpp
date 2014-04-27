///
/// @file xlversion.hpp
///
/// @brief XLKit version defines

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

#ifndef XLKIT_VERSION_HPP

/// @addtogroup other_macros Other Macros
/// Other macros defined by xlkit
/// @{

/// Version namespace for this library
#define XLKIT_VERSION_NAME v0_1_0

/// Major version number
#define XLKIT_MAJOR_VERSION 0
/// Minor version number
#define XLKIT_MINOR_VERSION 1
/// Patch version number
#define XLKIT_PATCH_VERSION 0

/// Library version as packed integer: "%02x%02x%04x" % (major, minor, patch)
#define XLKIT_VERSION \
			(((XLKIT_MAJOR_VERSION & 0x00FF) << 24) | \
			 ((XLKIT_MINOR_VERSION & 0x00FF) << 16) | \
			 ((XLKIT_PATCH_VERSION & 0xFFFF)      ) )

/// Macro used to pull the versioned namespace into the main xlkit namepsace
///
/// @note The empty namespace clause below ensures that XLKIT_VERSION_NAME is
/// recognized as a namespace name.
#define XLKIT_USE_VERSION_NAMESPACE \
			namespace XLKIT_VERSION_NAME {} \
			using namespace XLKIT_VERSION_NAME;

/// @}

#endif // XLKIT_VERSION_HPP
