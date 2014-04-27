/// @file xlOperand.hpp
///
/// @brief xlkit::xlOperand class
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

#ifndef XLKIT_XLOPERAND_HPP
#define XLKIT_XLOPERAND_HPP

#include <xlkit/xlcall.hpp>
#include <xlkit/xlException.hpp>
#include <xlkit/xlutil.hpp>
#include <xlkit/xlversion.hpp>

#include <boost/lexical_cast.hpp>

#include <utility>
#include <string>

#include <stddef.h>
#include <stdint.h>
#include <string.h>

namespace xlkit {
XLKIT_USE_VERSION_NAMESPACE
namespace XLKIT_VERSION_NAME {

namespace detail {

template<typename T>
struct unimplemented : std::false_type {};


} // namespace detail

/// An error number value
struct xlError {
	xlError() : num(xlerrNull) { }
	/// Construct from xlerr* enum
	xlError(int x) : num(x) { }
	/// Return underlying xlerr* enum
	operator int() {
		return num;
	}
	/// Return enum value as a string
	std::string str() const {
		if (num == xlerrNull)
			return std::string("xlerrNull");
		if (num == xlerrDiv0)
			return std::string("xlerrDiv0");
		if (num == xlerrValue)
			return std::string("xlerrValue");
		if (num == xlerrRef)
			return std::string("xlerrRef");
		if (num == xlerrName)
			return std::string("xlerrName");
		if (num == xlerrNum)
			return std::string("xlerrNum");
		if (num == xlerrNA)
			return std::string("xlerrNA");
		if (num == xlerrGettingData)
			return std::string("xlerrGettingData");
		return boost::lexical_cast<std::string>(num);
	}
	int num;
};

std::string
xltypeString(unsigned int xltype) {

	std::string xlfree;
	if (xltype & xlbitXLFree)
		xlfree += "|xlbitXLFree";
	if (xltype & xlbitDLLFree)
		xlfree += "|xlbitDLLFree";

	// Do big data first as it has both Str and Int bits enabled
	if (xltype & xltypeBigData)
		return std::string("xltypeBigData") + xlfree;

	if (xltype & xltypeNum)
		return std::string("xltypeNum") + xlfree;
	if (xltype & xltypeStr)
		return std::string("xltypeStr") + xlfree;
	if (xltype & xltypeBool)
		return std::string("xltypeBool") + xlfree;
	if (xltype & xltypeRef)
		return std::string("xltypeRef") + xlfree;
	if (xltype & xltypeErr)
		return std::string("xltypeErr") + xlfree;
	if (xltype & xltypeFlow)
		return std::string("xltypeFlow") + xlfree;
	if (xltype & xltypeMulti)
		return std::string("xltypeMulti") + xlfree;
	if (xltype & xltypeMissing)
		return std::string("xltypeMissing") + xlfree;
	if (xltype & xltypeNil)
		return std::string("xltypeNil") + xlfree;
	if (xltype & xltypeSRef)
		return std::string("xltypeSRef") + xlfree;
	if (xltype & xltypeInt)
		return std::string("xltypeInt") + xlfree;

	return std::string("Unknown xltype") + xlfree;
}

/// C++ methods that operate on top of an XLOPER struct
/// @note xlOper4 must *not* contain any data so that its memory layout is
/// identical to an XLOPER. The two direcly inherited members are the union
/// val and the uint16_t xltype.
class xlOper4 : private XLOPER {

  public:

	/// Proxy class into an operand's cell matrix (mutable)
	/// @{
	class CellMatrixRef {
	  public:

		/// Rows in matrix
		int rows() const {
			return myOperand->val.array.rows;
		}
		/// Columns in matrix
		int cols() const {
			return myOperand->val.array.columns;
		}

		/// (row,col) value in matrix 
		/// @{
		const xlOper4& operator()(int i, int j) const {
			return *((xlOper4*)myOperand->val.array.lparray
					 + (i * myOperand->val.array.columns) + j);
		}
		xlOper4& operator()(int i, int j) {
			return *((xlOper4*)myOperand->val.array.lparray
					 + (i * myOperand->val.array.columns) + j);
		}
		/// @}

	  private:
		explicit CellMatrixRef(xlOper4* operand)
			: myOperand(operand) { }

		xlOper4* myOperand;
		friend class xlOper4;
		friend class ConstCellMatrixRef;
	};
	/// Proxy class into an operand's cell matrix (non-mutable)
	class ConstCellMatrixRef {
	  public:
		/// Construct from non-const CellMatrixRef
		explicit ConstCellMatrixRef(CellMatrixRef ref)
			: myOperand(ref.myOperand) {
		}

		/// Rows in matrix
		int rows() const {
			return myOperand->val.array.rows;
		}
		/// Columns in matrix
		int cols() const {
			return myOperand->val.array.columns;
		}

		/// (row,col) value in matrix 
		const xlOper4& operator()(int i, int j) const {
			return *((xlOper4*)myOperand->val.array.lparray
					 + (i * myOperand->val.array.columns) + j);
		}

	  private:
		explicit ConstCellMatrixRef(const xlOper4* operand)
			: myOperand(operand) { }

		const xlOper4* myOperand;
		friend class xlOper4;
	};
	friend class xlOper4::CellMatrixRef;
	friend class xlOper4::ConstCellMatrixRef;
	/// @}

	/// Default constructor, initializes as xltypeMissing
	xlOper4() {
		init();
	}
	~xlOper4() {
		reset();
	}
	xlOper4(const xlOper4& other) {
		init();
		*this = other;
	}
	xlOper4(xlOper4&& other) {
		init();
		*this = std::move(other);
	}

	/// Construct a double precision floating point number
	explicit xlOper4(double v) {
		init();
		set(v);
	}
	/// Construct an integer
	explicit xlOper4(int v) {
		init();
		set(v);
	}
	/// Construct from an std::string
	explicit xlOper4(const std::string& v) {
		init();
		set(v);
	}
	/// Construct from a C-style null-terminated string
	explicit xlOper4(const char* v) {
		init();
		set(v);
	}
	/// Construct a bool
	explicit xlOper4(bool v) {
		init();
		set(v);
	}
	/// Construct an error number (xlerrValue, xlerrNA, etc..)
	/// @note See xlcall.h
	explicit xlOper4(xlError num) {
		init();
		set(num);
	}
	/// Construct a cell matrix of given size
	/// @note If init_val, is not given, it will be set to a matrix of
	/// of all xltypeMissing elements.
	explicit xlOper4(int rows, int cols, xlOper4* init_val = NULL) {
		init();
		setMatrix(rows, cols, init_val);
	}
	/// Construct a copy of a cell matrix
	explicit xlOper4(ConstCellMatrixRef cell_ref) {
		init();
		set(cell_ref);
	}

	/// Free allocated memory and reset to initial state (xltypeMissing)
	inline void reset() {
		if (xltype & xltypeStr) {
			if (xltype & xlbitXLFree)
				XLKIT_THROW("Cannot reset memory allocated by Excel!");
			else if (xltype & xlbitDLLFree)
				::free(val.str);
		} else if (xltype & xltypeMulti) {
			if (xltype & xlbitXLFree)
				XLKIT_THROW("Cannot reset memory allocated by Excel!");
			else if (xltype & xlbitDLLFree)
				::free(reinterpret_cast<void*>(val.array.lparray));
		}
		init();
	}

	/// Assignment operator
	xlOper4& operator=(const xlOper4& other) {
		if (this != &other) {
			reset();
			if (other.isString()) {
				set(other.get<const char *>());
			} else if (other.isCellMatrix()) {
				set(other.get<ConstCellMatrixRef>());
			} else {
				::memcpy(this, &other, sizeof(*this));
			}
		}
		return *this;
	}
	/// Assignment move operator
	xlOper4& operator=(xlOper4&& other) {
		if (this != &other) {
			reset();
			::memcpy(this, &other, sizeof(*this));
			other.init();
		}
		return *this;
	}

	/// Test operand type
	/// @{
	inline bool isDouble() const {
		return (xltype == xltypeNum);
	}
	inline bool isInteger() const {
		return (xltype == xltypeInt);
	}
	inline bool isString() const {
		return (   xltype ==  xltypeStr
				   || xltype == (xltypeStr|xlbitXLFree)
				   || xltype == (xltypeStr|xlbitDLLFree));
	}
	inline bool isBool() const {
		return (xltype == xltypeBool);
	}
	inline bool isError() const {
		return (xltype == xltypeErr);
	}
	inline bool isMissing() const {
		return (xltype == xltypeMissing);
	}
	inline bool isCellMatrix() const {
		return (   xltype ==  xltypeMulti
				   || xltype == (xltypeMulti|xlbitXLFree)
				   || xltype == (xltypeMulti|xlbitDLLFree));
	}
	/// @}

	template <typename T> T get() const {
		static_assert( detail::unimplemented<T>::value
					   , "Only specializations of get<>() const may be used" );
	}

	template <>
	double get<double>() const {
		if (!isDouble())
			return castValue<double>();
		return val.num;
	}
	template <>
	int get<int>() const {
		if (!isInteger())
			return castValue<int>();
		return val.w;
	}
	template <>
	std::string get<std::string>() const {
		if (!isString())
			return castValue<std::string>();
		return std::string(reinterpret_cast<char *>(val.str + 1));
	}
	template <>
	const char* get<const char*>() const {
		// This method is for efficiency only, does not convert.
		if (!isString())
			XLKIT_THROW("Cannot cast to const char* from " + xltypeString(xltype));
		return reinterpret_cast<char *>(val.str + 1);
	}
	template <>
	bool get<bool>() const {
		if (!isBool())
			return castValue<bool>();
		return (val.xbool != 0);
	}
	template <>
	xlError get<xlError>() const {
		if (!isError())
			XLKIT_THROW("Cannot cast to xlError from " + xltypeString(xltype));
		return xlError(val.err);
	}
	template <>
	ConstCellMatrixRef get<ConstCellMatrixRef>() const {
		if (!isCellMatrix())
			XLKIT_THROW("Cannot cast to ConstCellMatrixRef from " + xltypeString(xltype));
		return ConstCellMatrixRef(this);
	}

	/// Make a matrix of the given size and return a ref to it.
	/// @note If init_val, is not given, all elements will be xltypeMissing.
	CellMatrixRef
	setMatrix(int rows, int cols, xlOper4* init_val = NULL) {
		reset();
		xltype = xltypeMulti | xlbitDLLFree;
		val.array.rows = rows;
		val.array.columns = cols;
		val.array.lparray = reinterpret_cast<XLOPER*>(
								::malloc(rows * cols * sizeof(xlOper4)));
		CellMatrixRef dst(this);
		if (init_val) {
			for (int i = 0; i < rows; ++i) {
				for (int j = 0; j < cols; ++j) {
					dst(i, j) = *init_val;
				}
			}
		} else {
			for (int i = 0; i < rows; ++i) {
				for (int j = 0; j < cols; ++j) {
					dst(i, j).init();
				}
			}
		}
		return CellMatrixRef(this);
	}
	/// Obtain cell matrix reference
	/// @{
	CellMatrixRef asCellMatrixRef() {
		if (!isCellMatrix())
			XLKIT_THROW("Cannot cast to CellMatrixRef from " + xltypeString(xltype));
		return CellMatrixRef(this);
	}
	ConstCellMatrixRef asCellMatrixRef() const {
		if (!isCellMatrix())
			XLKIT_THROW("Cannot cast to ConstCellMatrixRef from " + xltypeString(xltype));
		return ConstCellMatrixRef(this);
	}
	/// @}

	/// For a string, return its length
	int stringLength() const {
		if (!isString())
			XLKIT_THROW("Not a string");
		return *(reinterpret_cast<uint8_t*>(&val.str[0]));
	}
	/// For a cell matrix, return its number of rows
	int cellMatrixRows() const {
		if (!isCellMatrix())
			XLKIT_THROW("Not a cell matrix");
		return val.array.rows;
	}
	/// For a cell matrix, return its number of columns
	int cellMatrixCols() const {
		if (!isCellMatrix())
			XLKIT_THROW("Not a cell matrix");
		return val.array.columns;
	}

	/// Set a new value
	/// @{
	void set(double v) {
		reset();
		xltype = xltypeNum;
		val.num = v;
	}
	void set(int v) {
		reset();
		xltype = xltypeInt;
		val.w = (int16_t)v;
	}
	void set(const std::string& v) {
		reset();
		xltype = xltypeStr | xlbitDLLFree;
		size_t len = v.size() + 2;
		val.str = reinterpret_cast<char*>(::malloc(len * sizeof(uint8_t)));
		uint8_t* blen = reinterpret_cast<uint8_t*>(&val.str[0]);
		*blen = uint8_t(len < 255 ? len : 255);
		XLKIT_PUSH_DISABLE_WARN_DEPRECATION
		::strcpy(reinterpret_cast<char *>(val.str + 1), v.c_str());
		XLKIT_POP_DISABLE_WARN_DEPRECATION
	}
	void set(const char* v) {
		reset();
		xltype = xltypeStr | xlbitDLLFree;
		size_t len = ::strlen(v) + 2;
		val.str = reinterpret_cast<char*>(::malloc(len * sizeof(uint8_t)));
		uint8_t* blen = reinterpret_cast<uint8_t*>(&val.str[0]);
		*blen = uint8_t(len < 255 ? len : 255);
		XLKIT_PUSH_DISABLE_WARN_DEPRECATION
		::strcpy(reinterpret_cast<char *>(val.str + 1), v);
		XLKIT_POP_DISABLE_WARN_DEPRECATION
	}
	void set(bool v) {
		reset();
		xltype = xltypeBool;
		val.xbool = v;
	}
	void set(xlError num) {
		reset();
		xltype = xltypeErr;
		val.err = num;
	}
	void set(ConstCellMatrixRef src) {
		CellMatrixRef dst(setMatrix(src.rows(), src.cols()));
		for (int i = 0, rows = src.rows(); i < rows; ++i) {
			for (int j = 0, cols = src.cols(); j < cols; ++j) {
				dst(i, j) = src(i, j);
			}
		}
	}
	/// @}

  private: // methods

	// Mimic default ctor behaviour, assumes we're uninitialized
	inline void init() {
		xltype = xltypeMissing;
		val.num = 0;
	}

	template <typename T>
	T castValue() const {
		// Casting to a number
		if (isDouble())
			return (T)(get<double>());
		if (isInteger())
			return (T)(get<int>());
		if (isString())
			return boost::lexical_cast<T>(get<const char*>());
		if (isBool())
			return (T)(get<bool>());
		XLKIT_THROW("Unsupported conversion from " + xltypeString(xltype));
	}
	template <>
	std::string castValue<std::string>() const {
		// Casting to a string
		if (isDouble())
			return boost::lexical_cast<std::string>(get<double>());
		if (isInteger())
			return boost::lexical_cast<std::string>(get<int>());
		if (isString())
			return std::string(get<const char*>());
		if (isBool())
			return boost::lexical_cast<std::string>(get<bool>());
		if (isError())
			return xlError(val.err).str();
		if (isMissing())
			return std::string("xltypeMissing");
		XLKIT_THROW("Cannot cast to string from " + xltypeString(xltype));
	}
	template <>
	bool castValue<bool>() const {
		// Casting to a bool
		if (isDouble())
			return (get<double>() != 0);
		if (isInteger())
			return (get<int>() != 0);
		if (isString())
			return (stringLength() != 0);
		if (isBool())
			return get<bool>();
		XLKIT_THROW("Cannot cast to bool from " + xltypeString(xltype));
	}

};
typedef xlOper4						xlOperand;
typedef xlOper4::CellMatrixRef		xlCellMatrixRef;
typedef xlOper4::ConstCellMatrixRef	xlConstCellMatrixRef;

} // namespace XLKIT_VERSION_NAME
} // namespace xlkit

/// @addtogroup aliases
/// @{

/// C++ wrapper for XLOPER struct. See @ref xlkit::XLKIT_VERSION_NAME::xlOper4 "xlOper4"
typedef xlkit::xlOperand xlOperand;

/// Proxy class into an operand's cell matrix (mutable). See @ref xlkit::XLKIT_VERSION_NAME::xlOper4::CellMatrixRef "CellMatrixRef"
typedef xlkit::xlCellMatrixRef xlCellMatrixRef;

/// Proxy class into an operand's cell matrix (non-mutable).  See @ref xlkit::XLKIT_VERSION_NAME::xlOper4::ConstCellMatrixRef "ConstCellMatrixRef"
typedef xlkit::xlConstCellMatrixRef xlConstCellMatrixRef;

/// @}

#endif // XLKIT_XLOPERAND_HPP
