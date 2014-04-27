/// @file xlkit.hpp
///
/// @brief Main XLKit header
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

#ifndef XLKIT_HPP
#define XLKIT_HPP

#include <xlkit/xldebug.hpp>
#include <xlkit/xlException.hpp>
#include <xlkit/xlOperand.hpp>
#include <xlkit/xlversion.hpp>

#define BOOST_FT_AUTODETECT_CALLING_CONVENTIONS
#include <boost/function_types/components.hpp>
#include <boost/function_types/is_nonmember_callable_builtin.hpp>

#include <boost/mpl/begin_end.hpp>
#include <boost/mpl/deref.hpp>
#include <boost/unordered_map.hpp>
#include <boost/utility/enable_if.hpp>

#include <string>
#include <vector>
#include <stdio.h>

namespace xlkit {
XLKIT_USE_VERSION_NAMESPACE
namespace XLKIT_VERSION_NAME {

namespace mpl = boost::mpl;
namespace ft = boost::function_types;

/// Return value for XLL functions
class ResultOperandPtr {
  public:

	/// Get pointer to TLS copy and default initialize it
	ResultOperandPtr();

	/// Get pointer to TLS copy and default initialize it with given
	/// xlOperand.
	ResultOperandPtr(const xlOperand& copy);

	operator xlOperand*()	{
		return myOperand;
	}

	xlOperand* operator->() {
		return myOperand;
	}
	xlOperand& operator*()  {
		return *myOperand;
	}

  private:
	xlOperand* myOperand;
};

namespace detail {

// Empty help
struct empty_ { };

// Info for type T
template <typename T>
struct TypeInfo;

#define XLKIT_TYPEINFO(TYPE, CODE, HELP) \
			template <> \
			struct TypeInfo<TYPE> { \
				static size_t		size()	{ return sizeof(TYPE); } \
				static char			code()	{ return CODE; } \
				static const char*	name()	{ return #TYPE; } \
				static const char*	help()	{ return HELP; } \
			}; \
			/**/

XLKIT_TYPEINFO(double,				'B', "Number")
XLKIT_TYPEINFO(const char*,			'C', "String")
XLKIT_TYPEINFO(uint16_t,			'H', "Unsigned Integer")
XLKIT_TYPEINFO(int16_t,				'I', "Signed Integer")
XLKIT_TYPEINFO(int32_t,				'J', "Signed Integer")
XLKIT_TYPEINFO(xlOperand*,			'P', "Cell or Cell Range")
XLKIT_TYPEINFO(const xlOperand*,	'P', "Cell or Cell Range")
XLKIT_TYPEINFO(ResultOperandPtr,	'P', "Cell or Cell Range")

#undef XLKIT_TYPEINFO

// Label for type T, defaults to nothing for unknown types.
template <typename T>
struct ParmHelp {
	static const char* name() {
		return NULL;
	}
	static const char* help() {
		return NULL;
	}
};

} // namespace detail

/// Function parameter with associated help
template <typename T, typename PARM_HELP = detail::empty_>
class Parm {
  public:
	typedef T			type;
	typedef PARM_HELP	parm_help_tag;

	/// Name of the parameter for Excel
	static const char *name() {
		return detail::ParmHelp<PARM_HELP>::name();
	}
	/// Help of the parameter for Excel
	static const char *help() {
		return detail::ParmHelp<PARM_HELP>::help();
	}

	/// Value of the parameter
	/// @{
	const T& value() const {
		return myValue;
	}
	operator const T&() const {
		return myValue;
	}
	operator T&() {
		return myValue;
	}
	/// @}

  private:
	T myValue;
};

namespace detail {

template <typename T, typename PARM_HELP>
struct TypeInfo< Parm<T, PARM_HELP> > {
	static size_t size() {
		return TypeInfo<T>::size();
	}
	static char code() {
		return TypeInfo<T>::code();
	}
	static const char* name() {
		const char* name = Parm<T, PARM_HELP>::name();
		if (!name)
			return TypeInfo<T>::name();
		return name;
	}
	static const char* help() {
		const char* help = Parm<T, PARM_HELP>::help();
		if (!help)
			return TypeInfo<T>::help();
		return help;
	}
};

class ExcelHost;

} // namespace detail

/// Stores all registry of functions for the XLL
class Registry {

  public:

	/// Get the singleton instance
	static Registry& instance() {
		if (!theInstance)
			theInstance = new Registry;
		return *theInstance;
	}

	void setAddinLabel(const char* label);

	/// Register a new function.
	/// @{
	/// Name of the Excel function is the same as the C++ function name
	template <typename F>
	void addFunction(const std::string& name, F f, const char* help) {
		myFunctions[name] = makeWrapper<F>(name, f, help);
	}
	/// Name of the Excel function is different from the the C++ function name
	template <typename F>
	void addFunction(const std::string& excel_name,
			const std::string& name, F f, const char* help) {
		myFunctions[name] = makeWrapper<F>(excel_name, f, help);
	}
	/// @}

	/// Print a list of all registered functions for debugging
	void dump() {
		for (const auto& i : myFunctions) {
			printf("'%s' -> '%s' [", i.first.c_str(), i.second.myTypes.c_str());
			for (int j = 0, n = i.second.myParmHelp.size(); j < n; ++j) {
				if (j > 0)
					printf(",");
				printf("%s", i.second.myParmHelp[j].c_str());
			}
			printf("]\n");
		}
	}

	/// Information for a registered function
	struct Wrapper {
		Wrapper() {
		}
		Wrapper(const std::string& func_name,
				const std::string& sig,
				const std::string& func_help,
				const std::string& arg_names,
				const std::vector<std::string>& parm_help)
			: myFuncName(func_name)
			, myTypes(sig)
			, myFuncHelp(func_help)
			, myArgNames(arg_names)
			, myParmHelp(parm_help) {
		}

		std::string myFuncName;
		std::string myTypes;
		std::string myFuncHelp;
		std::string myArgNames;
		std::vector<std::string> myParmHelp;
	};

	typedef boost::unordered_map<std::string, Wrapper> NameMap;

	/// List of registered functions
	const NameMap& functions() const {
		return myFunctions;
	}

  private:

	Registry() {}

	// Make a wrapper object for function F
	template <typename F>
	Wrapper
	makeWrapper(const std::string&name, F f, const char* func_help,
				typename boost::enable_if< ft::is_nonmember_callable_builtin<F, ft::cdecl_cc>
				>::type *dummy = 0) {

		typedef Func<F> FuncT;

		std::string types;
		Func<F>::getTypes(types);

		// Skip first type, which is the return type
		typedef typename mpl::next<mpl::begin< ft::components<F> >::type>::type Second;
		typedef typename mpl::end< ft::components<F> >::type End;

		std::string arg_names;
		Func<F, Second, End>::getArgNames(arg_names);

		std::vector<std::string> parm_help;
		Func<F, Second, End>::getParmHelp(parm_help);

		return Wrapper(name, types, func_help, arg_names, parm_help);
	}

	// Get the information for function F (first arg is return type)
	template
	< typename F
	  , class Beg = typename mpl::begin< ft::components<F> >::type
	  , class End = typename mpl::end< ft::components<F> >::type
	  >
	struct Func;

  private:

	NameMap myFunctions;

	static Registry* theInstance;
};

inline void
dumpRegistry() {
	Registry::instance().dump();
}

//
// Specializations for function introspection
//
template <typename F, typename Curr, typename End>
struct Registry::Func {
	// Type string signature
	static void getTypes(std::string& types) {
		typedef typename mpl::next<Curr>::type Next;
		typedef typename mpl::deref<Curr>::type Type;

		types += detail::TypeInfo<Type>::code();
		Func<F, Next, End>::getTypes(types);
	}
	// Names of the arguments
	static void getArgNames(std::string& arg_names) {
		typedef typename mpl::next<Curr>::type Next;
		typedef typename mpl::deref<Curr>::type Type;

		if (arg_names.size() > 0)
			arg_names += ", ";
		arg_names += detail::TypeInfo<Type>::name();
		Func<F, Next, End>::getArgNames(arg_names);
	}
	// Help description for individual arguments
	static void getParmHelp(std::vector<std::string>& parm_help) {
		typedef typename mpl::next<Curr>::type Next;
		typedef typename mpl::deref<Curr>::type Type;

		parm_help.push_back(detail::TypeInfo<Type>::help());
		Func<F, Next, End>::getParmHelp(parm_help);
	}
};

template <typename F, typename End>
struct Registry::Func<F, End, End> {
	static void getTypes(std::string&) { }
	static void getArgNames(std::string&) { }
	static void getParmHelp(std::vector<std::string>&) { }
};

} // namespace XLKIT_VERSION_NAME
} // namespace xlkit

/// @defgroup aliases Common Types
/// These types are pulled outside of the xlkit namespace for convenience.
/// @{

/// An xlkit function parameter of type T, with optional help for convenience.
/// See @ref xlkit::XLKIT_VERSION_NAME::Parm "Parm".
template <typename T, typename PARM_HELP = xlkit::detail::empty_>
using xlParm = xlkit::Parm<T, PARM_HELP>;

/// Return value for XLL functions.
/// See @ref xlkit::XLKIT_VERSION_NAME::ResultOperandPtr "xlResultOperandPtr".
typedef xlkit::ResultOperandPtr xlResultOperandPtr;

/// @}

/// @addtogroup macros Main Macros
/// Macros for interfacing with Excel
/// @{

/// Macro to register the name of the add-in. This shows up as the category in
/// Excel's Function Wizard for all registered functions.
#define XLKIT_INIT_ADDIN_LABEL(LABEL) \
			struct XLInitAddinLabel { \
				XLInitAddinLabel() { \
					xlkit::Registry::instance().setAddinLabel(LABEL); \
				} \
			}; \
			static XLInitAddinLabel theInitAddinLabel; \
			/**/


/// Macro to create `xlParm<NAME>` typedef as `xlParm<VALUE_TYPE, HELP>` proxy type for a variable of VALUE_TYPE.
/// @code typedef of @pre xlParm<VALUE_TYPE, HELP> xlParm\#\#NAME @endcode
#define XLKIT_PARM(VALUE_TYPE, NAME, HELP) \
			struct HELP_FOR_##NAME { }; \
			namespace xlkit { \
			XLKIT_USE_VERSION_NAMESPACE \
			namespace XLKIT_VERSION_NAME { \
			namespace detail { \
			template <> struct ParmHelp<HELP_FOR_##NAME> { \
				static const char* name() { return #NAME; } \
				static const char* help() { return HELP; } \
			}; \
			} } } \
			typedef xlParm<VALUE_TYPE, HELP_FOR_##NAME> xlParm##NAME;
			/**/

/// Macro to register the given function with xlkit
#define XLKIT_REGISTER(FUNC, HELP) \
			struct FUNC##Registrar { \
				FUNC##Registrar() { \
					xlkit::Registry::instance().addFunction(#FUNC,FUNC,HELP); \
				} \
			}; \
			static FUNC##Registrar the##FUNC##Registrar; \
			/**/

/// Macro to register the given function with xlkit with a different name from
/// the C++ function name.
#define XLKIT_REGISTER_AS(XLNAME, FUNC, HELP) \
			struct FUNC##Registrar { \
				FUNC##Registrar() { \
					xlkit::Registry::instance().addFunction(XLNAME,#FUNC,FUNC,HELP); \
				} \
			}; \
			static FUNC##Registrar the##FUNC##Registrar; \
			/**/

/// All registered functions must have this
#define XLKIT_EXPORT	extern "C" __declspec(dllexport)

/// All Excel functions begin with this macro
#define XLKIT_BEGIN_FUNCTION \
			try { \
			/**/

/// All Excel functions end with this macro
/// @note Return 0 will be interpreted by Excel as \#NULL!.
#define XLKIT_END_FUNCTION \
			} catch (xlkit::xlException& err) { \
				XLDBG("Exception caught: %s", err.what()); \
				return 0; \
			} catch (std::exception& err){ \
				XLDBG("Exception caught: %s", err.what()); \
				xlResultOperandPtr result; \
				result->set(err.what()); \
				return result; \
			} catch (xlkit::xlError& err){ \
				XLDBG("Exception caught: %s", err.str().c_str()); \
				xlResultOperandPtr result; \
				result->set(err); \
				return result; \
			} catch (...) { \
				XLDBG("Unknown exception caught"); \
				xlResultOperandPtr result; \
				result->set(xlkit::xlError(xlerrValue)); \
				return result; \
			} \
			/**/

/// @}

#endif // XLKIT_HPP
