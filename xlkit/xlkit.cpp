/// @file xlkit.cpp
///
/// @brief Main program for XLKit.
/// This can either be compiled into its own library for use by an XLL or be
/// simply \#include'd into a .cpp file already used in building the XLL.

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

#include <xlkit/xlkit.hpp>

#include <xlkit/xldebug.hpp>
#include <xlkit/xlutil.hpp>

#include <boost/lexical_cast.hpp>
#include <boost/algorithm/string/predicate.hpp>

#include <io.h>
#include <stdio.h>
#include <fcntl.h>
#include <assert.h>
#include <ios>

// Include this last to avoid windows.h contaimination
#include <xlkit/xlwindows.hpp>


namespace xlkit {
XLKIT_USE_VERSION_NAMESPACE
namespace XLKIT_VERSION_NAME {

namespace detail {

static bool theHasConsole = false;

// Implementation for xldebug.hpp
void
outputDebugString(const char *msg) {
	if (::IsDebuggerPresent()) {
		::OutputDebugStringA(msg);
	} else {
		if (_isatty(_fileno(stderr)) && !theHasConsole) {
			theHasConsole = true;
			(void) AllocConsole();
			int h = _open_osfhandle((intptr_t)::GetStdHandle(STD_ERROR_HANDLE), _O_TEXT);
			*stderr = *_fdopen(h, "wt");
			setvbuf(stderr, NULL, _IONBF, 0);
			std::ios::sync_with_stdio();
		}
		fprintf(stderr, "%s", msg);
	}
}

typedef int (__cdecl *ExcelProc4)(int xlfn, LPXLOPER operRes,
								  int count, ...);
typedef int (__stdcall *ExcelProc4v)(int xlfn, LPXLOPER operRes,
									 int count, LPXLOPER opers[]);

// Global functions to call into Excel
static ExcelProc4	Excel4_;
static ExcelProc4v	Excel4v_;

XLOPER*
xloperCast(xlOperand* ptr) {
	return reinterpret_cast<XLOPER*>(ptr);
}
const XLOPER*
xloperCast(const xlOperand* ptr) {
	return reinterpret_cast<const XLOPER*>(ptr);
}
xlOperand*
xlOperandCast(LPXLOPER ptr) {
	return reinterpret_cast<xlOperand*>(ptr);
}

class ExcelHost {
	static const int MAX_XL4_STR_LEN	= 255u;
	static const int MAX_XL11_ROWS		= 65536;
	static const int MAX_XL11_COLS		= 256;
	static const int MAX_XL11_UDF_ARG	= 30;
	static const int MAX_XL12_ROWS		= 1048576;
	static const int MAX_XL12_COLS		= 16384;
	static const int MAX_XL12_STR_LEN	= 32767u;
	static const int MAX_XL12_UDF_ARG	= 255;

  public:
	/// Get the singleton instance
	static ExcelHost&
	instance() {
		if (!theInstance)
			theInstance = new ExcelHost;
		return *theInstance;
	}

	/// Set the addin label
	void setAddinLabel(const char* label) {
		myAddinLabel = label;
	}
	/// Set the addin label
	const std::string& addinLabel() const {
		return myAddinLabel;
	}

	/// Result type for call...() functions
	class ExcelResult : public xlOperand  {
	  public:
		~ExcelResult() {
			Excel4_(xlFree, NULL, /*count*/1, this);
		}
	  private:
	};


	/// Attach to host
	void
	attach() {
		Progress progress("ExcelHost: Attaching");

		std::vector<xlOperand> args;

		xlOperand dll_name;
		Excel4_(xlGetName, xloperCast(&dll_name), 0);

		const Registry::NameMap& functions = Registry::instance().functions();
		for (const auto& it : functions) {

			args.clear();

			args.emplace_back(dll_name);			// pxModuleText
			args.emplace_back(it.first);			// pxProcedure
			args.emplace_back(it.second.myTypes);	// pxTypeText
			args.emplace_back(it.second.myFuncName);// pxFunctionText
			args.emplace_back(it.second.myArgNames);// pxArgumentText
			args.emplace_back();					// pxMacroType
													// (default: from anywhere)
			args.emplace_back(addinLabel());		// pxCategory
			args.emplace_back();					// pxShortcutText (none)
			args.emplace_back();					// pxHelpTopic (none)
			args.emplace_back(it.second.myFuncHelp);// pxFunctionHelp

			// pxArgumentHelp...
			const std::vector<std::string>& helps = it.second.myParmHelp;
			for (int j = 0, n = helps.size(); j < n; ++j) {
				// See http://msdn.microsoft.com/en-us/library/bb687841.aspx
				// for _Argument Description String Truncation in the
				// Function Wizard_ for why we need to do this.  In
				// reality, it looks lik Excel actually avoids
				// truncation by specifically looking for ". ".
				if (j < n - 1) {
					args.emplace_back(helps[j]);
				} else {
					args.emplace_back(helps[j] + ". ");
				}
			}

			ExcelResult func_id;
			callV(xlfRegister, func_id, args);
			if (func_id.isError()) {
				std::string error_str = func_id.get<std::string>();
				XLDBG("Failed to register %s (%s) in %s: Error %s (%d)",
					  it.first.c_str(),
					  it.second.myTypes.c_str(),
					  dll_name.get<const char*>(),
					  error_str.c_str(),
					  func_id.get<xlError>().num);
			} else {
				XLDBG("Register %s (%s) in %s as %f",
					  it.first.c_str(),
					  it.second.myTypes.c_str(),
					  dll_name.get<const char*>(),
					  func_id.get<double>());
			}
		}
	}

	/// Detach from host
	void
	detach() {
		Progress progress("ExcelHost: Detaching");
	}

	template <typename... PARMS>
	bool call(int xlfn, const PARMS&... parms) {
		const size_t n_parms = sizeof...(parms);
		std::vector<xlOperand> args;
		args.reserve(n_parms);
		getArgs(args, parms...);
		return callV(xlfn, args);
	}
	template <typename... PARMS>
	bool evalCall(int xlfn, ExcelResult& result, const PARMS&... parms) {
		const size_t n_parms = sizeof...(parms);
		std::vector<xlOperand> args;
		args.reserve(n_parms);
		getArgs(args, parms...);
		return callV(xlfn, xloperCast(&result), args);
	}

	void setStatusV(const char *fmt, va_list args) {
		call(xlcMessage, true, strprintfV(fmt, args).c_str());
	}
	void clearStatus() {
		call(xlcMessage, false, "");
	}

	class Progress {
	  public:
		Progress(const char *fmt, ...) {
			va_list args;
			va_start(args, fmt);
			ExcelHost::instance().setStatusV(fmt, args);
			va_end(args);
		}
		~Progress() {
			ExcelHost::instance().clearStatus();
		}
	};

  private:

	ExcelHost()
		: myAddinLabel("Generic XLKit Addin") {
		HMODULE handle = LoadLibraryA("XLCALL32.DLL");
		if (!handle)
			XLKIT_THROW("Failed to load XLCALL32.DLL");
		Excel4_ = (ExcelProc4) ::GetProcAddress(handle, "Excel4");
		if (!Excel4_)
			XLKIT_THROW("Failed to get Excel4 function address");
		Excel4v_ = (ExcelProc4v) ::GetProcAddress(handle, "Excel4v");
		if (!Excel4v_)
			XLKIT_THROW("Failed to get Excel4v function address");
	}

	bool
	callV(int xlfn, xlOperand& result, const std::vector<xlOperand>& args) {
#ifdef _DEBUG
		const char *xlfn_type = "<unknown xlfn type>";
		if (xlfn & xlCommand)
			xlfn_type = "xlCommand";
		else if (xlfn & xlSpecial)
			xlfn_type = "xlSpecial";
		else if (xlfn & xlIntl)
			xlfn_type = "xlIntl";
		else if (xlfn & xlPrompt)
			xlfn_type = "xlPrompt";
		//XLDBG("callV %s %d with %d args", xlfn_type, xlfn & 0x0FFF, args.size());
#endif
		static_assert(sizeof(XLOPER) == sizeof(xlOperand),
					  "Operand has the wrong size!");
		std::vector<LPXLOPER> parms;
		parms.reserve(args.size());
		for (int i = 0, n = args.size(); i < n; i++)
			parms.push_back(const_cast<LPXLOPER>(xloperCast(&args[i])));
		int xlret = Excel4v_(xlfn, xloperCast(&result), parms.size(), parms.data());
#ifdef _DEBUG
		if (xlret != xlretSuccess) {
			XLDBG("callV %s %d with %d args", xlfn_type, xlfn & 0x0FFF, args.size());
			// Multiple error bits might be on
			const char *xlret_type = "<unknown xlret type>";
			if (xlret & xlretAbort)
				xlret_type = "xlretAbort";
			if (xlret & xlretInvXlfn)
				xlret_type = "xlretInvXlfn";
			if (xlret & xlretInvCount)
				xlret_type = "xlretInvCount";
			if (xlret & xlretInvXloper)
				xlret_type = "xlretInvXloper";
			if (xlret & xlretStackOvfl)
				xlret_type = "xlretStackOvfl";
			if (xlret & xlretFailed)
				xlret_type = "xlretFailed";
			if (xlret & xlretUncalced)
				xlret_type = "xlretUncalced";
			XLDBG("-> FAILED with %s", xlret_type);
		}
#endif
		return (xlret == xlretSuccess);
	}
	bool callV(int xlfn, const std::vector<xlOperand>& args) {
		ExcelResult unused_result;
		return callV(xlfn, unused_result, args);
	}

	void getArgs(std::vector<xlOperand>&) {
	}
	template <typename T>
	void getArgs(std::vector<xlOperand>& args, const T& parm) {
		args.push_back(xlOperand(parm));
	}
	template <typename T, typename... REST>
	void getArgs(std::vector<xlOperand>& args, const T& parm, const REST&... rest) {
		args.push_back(xlOperand(parm));
		getArgs(args, rest...);
	}

  private:
	std::string myAddinLabel;

	static ExcelHost* theInstance;
};

//
// Initialize global data for namespace xlkit::detail
//
ExcelHost* ExcelHost::theInstance = NULL;

} // namespace detail

//
// ResultOperandPtr
//
__declspec(thread) static XLOPER theTLSOperand;
ResultOperandPtr::ResultOperandPtr()
	: myOperand(detail::xlOperandCast(&theTLSOperand)) {
	// Reset it
	theTLSOperand.xltype = xltypeMissing;
	theTLSOperand.val.num = 0;;
}
ResultOperandPtr::ResultOperandPtr(const xlOperand& copy)
	: myOperand(detail::xlOperandCast(&theTLSOperand)) {
	// Reset it
	theTLSOperand.xltype = xltypeMissing;
	theTLSOperand.val.num = 0;;
	// Copy
	*myOperand = copy;
}

//
// Registry
//
Registry* Registry::theInstance = NULL;

void
Registry::setAddinLabel(const char* label) {
	detail::ExcelHost::instance().setAddinLabel(label);
}

} // namespace XLKIT_VERSION_NAME
} // namespace xlkit


// Force export of functions required by Excel so that for msvc we don't need
// to use a def file. The number after the @ sign is the number of bytes for
// the sum of all parameters and return value.
#pragma comment (linker, "/export:_xlAddInManagerInfo=_xlAddInManagerInfo@4")
#pragma comment (linker, "/export:_xlAutoOpen=_xlAutoOpen@0")
#pragma comment (linker, "/export:_xlAutoClose=_xlAutoClose@0")
#pragma comment (linker, "/export:_xlAutoRemove=_xlAutoRemove@0")
#pragma comment (linker, "/export:_xlAutoFree=_xlAutoFree@4")

// Define entry points for .xll
extern "C" {

	LPXLOPER WINAPI
	xlAddInManagerInfo(LPXLOPER xAction) {

		using namespace xlkit;
		using namespace xlkit::detail;

		ResultOperandPtr result;
		result->set(xlError(xlerrValue));

		try {
			if (xlOperandCast(xAction)->get<int>() == 1)
				result->set(ExcelHost::instance().addinLabel());
		} catch(std::exception &err) {
			XLDBG("Exception caught: %s", err.what());
		} catch(...) {
			XLDBG("Unknown EXCEPTION!");
		}

		return xloperCast(result);
	}

	int WINAPI
	xlAutoOpen() {

		using namespace xlkit::detail;

		try {
			// Register our hooks
			ExcelHost::instance().attach();
			XLDBG("Opened.");
		} catch(std::exception &err) {
			XLDBG("Exception caught: %s", err.what());
		} catch(...) {
			XLDBG("Unknown EXCEPTION!");
		}
		return 1; // must return 1
	}

	static bool theAutoRemoveCalled = false;

	int WINAPI
	xlAutoClose() {

		using namespace xlkit::detail;

		try {

			if (theAutoRemoveCalled) {
				// we can safely unregister the functions here as the user has
				// unloaded the xll and so won't expect to be able to use the
				// functions
				ExcelHost::instance().detach();
			} else {
				// note that we don't unregister the functions here
				// excel has some strange behaviour when exiting and can
				// call xlAutoClose before the user has been asked about the close
			}

			XLDBG("Closed.");
		} catch(std::exception &err) {
			XLDBG("Exception caught: %s", err.what());
		} catch(...) {
			XLDBG("Unknown EXCEPTION!");
		}
		return 1; // must return 1
	}

	int WINAPI
	xlAutoRemove() {

		try {
			// tell auto close we've been called so that we can call deregister
			theAutoRemoveCalled = true;
			XLDBG("Removed.");
		} catch(std::exception &err) {
			XLDBG("Exception caught: %s", err.what());
		} catch(...) {
			XLDBG("Unknown EXCEPTION!");
		}
		return 1; // must return 1
	}

	void WINAPI
	xlAutoFree(LPXLOPER pxFree) {
		using namespace xlkit::detail;
		xlOperandCast(pxFree)->reset();
	}

} // extern "C"

