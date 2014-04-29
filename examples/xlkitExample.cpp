///
/// @example xlkitExample.cpp
///
/// @brief Example of how to use XLKit
///
/// @addtogroup HOWTO 
/// @brief How to use XLKit
/// @{
///
/// @section setup Setup
///
/// Here are the steps that were used to create the xlkitExample project using
/// Visual Studio 2013+:
///   - From the main menu, choose File > New Project > Templates > Visual C++
///     > Win32.
///   - Configure your name and project location and hit OK.
///	  - From the main menum choose Project > Add Existing Item 
///   - Add this file to the project.
///	  - From the main menu, choose Project > *SolutionName* Properties.
///   - At the top of this dialog, change the Configuration menu to All
///     Configurations so that we modify both the debug and release
///     configurations at the same time.
///	  - Under Configuration Properties, go to VC++ Directories and adjust the
///	    Include Directories to include to the *PARENT* directory for where
///	    the following relative paths can be resolved. Use the drop down menu
///	    on the right to edit it and hit Apply to save your changes when done.
///			-# `<xlkit/xlkit.hpp>`	(from xlkit library)
///			-# `<boost/config.hpp>`	(from boost library)
///			-# `<xlcall.h>`			(from Excel XLL Software Development Kit)
///	  - Under the Configuration Properties > Build Events > Post-Build Event
///	    page, enter this command and then hit Apply to save your changes:
///		@code
///			copy /Y "$(TargetDir)$(ProjectName).dll" "$(SolutionDir)\$(ProjectName).xll"
///		@endcode
///   - Now change the Configuration menu at the top to Debug only and modify
///     the command line so that the debug XLL is set to a different name, eg.
///		@code
///			copy /Y "$(TargetDir)$(ProjectName).dll" "$(SolutionDir)\$(ProjectName)_debug.xll"
///		@endcode
///
/// Once you've compiled the XLL, you need to tell Excel to load it from the
/// Add-In Manager.
///
/// @}

/*
This file is released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or
distribute this software, either in source code form or as a compiled
binary, for any purpose, commercial or non-commercial, and by any
means.

In jurisdictions that recognize copyright laws, the author or authors
of this software dedicate any and all copyright interest in the
software to the public domain. We make this dedication for the benefit
of the public at large and to the detriment of our heirs and
successors. We intend this dedication to be an overt act of
relinquishment in perpetuity of all present and future rights to this
software under copyright law.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.

For more information, please refer to <http://unlicense.org>
*/

// Include xlkit.hpp 
#include <xlkit/xlkit.hpp>

// Includes used by example function code
#include <stdio.h>


//////////////////////////////////////////////////////////////////////////////
//
// Simplest example of an Excel function to create.
//

// Write your function, using the XLKIT_API calling convention
double XLKIT_API
xlCirc(double diameter)
{
	// Being the function with the XLKIT_BEGIN_FUNCTION macro to export this
	// function to the DLL and to start a try block.
	XLKIT_BEGIN_FUNCTION

	return diameter * 3.14159;

	// End the function with XLKIT_END_FUNCTION(return_type). The return type
	// allows it to return something reasonable under error conditions.
	XLKIT_END_FUNCTION(double)
}

// Register your function for the XLL
XLKIT_REGISTER(xlCirc, "Compute circle circumference")


//////////////////////////////////////////////////////////////////////////////
//
// Example to input and output cell arrays
//

// Set up help to be used for parms that is displayed in Excel's Function
// Wizard. In concept, the example XLKIT_PARM() below expands to:
//
//	typedef xlParm<const xlOperand*, "Cell range of numbers"> Diameter;
//
XLKIT_PARM(const xlOperand*, DataRange, "Cell range of numbers")

// To return a cell, which can be any type, return an xlOperand*.
// In this example, we input rectangular range of cells and output
// an 1x2 cell range of the mean and variance.
// NOTE: Use CTRL+SHIFT+ENTER in the formula field to set the output cell
//       range to your cell selection.
xlOperand* XLKIT_API
xlStats(xlParmDataRange cells)
{
	XLKIT_BEGIN_FUNCTION

	// All xlOperand* return values must be an xlResultOperandPtr which
	// represents a pointer to a thread-local copy of an xlOperand for return
	// to Excel.
	xlResultOperandPtr result;

	double sum = 0.0;
	double sum_of_squares = 0.0;

	// Iterate over the cells in the incoming matrix.
	xlConstCellMatrixRef src(cells.value()->get<xlConstCellMatrixRef>());
	for (int i = 0; i < src.rows(); ++i)
	{
		for (int j = 0; j < src.cols(); ++j)
		{
			double x = src(i, j).get<double>();
			sum += x;
			sum_of_squares += x*x;
		}
	}

	// Avoid divide by zero
	int num_items = src.rows() * src.cols();
	if (num_items == 0)
		XLKIT_THROW("Can't calculate stats on empty range");

	// Create reference to an output results matrix of size 1x2
	xlCellMatrixRef mat(result->setMatrix(1, 2));

	double average = sum / num_items;
	mat(0, 0).set(average);
	mat(0, 1).set(sum_of_squares / num_items - average * average);

	return result;

	XLKIT_END_FUNCTION(xlResultOperandPtr)
}
XLKIT_REGISTER(xlStats, "Compute mean and variance as 1x2 cell range")

//////////////////////////////////////////////////////////////////////////////
//
// Simple example to do a cell range reference.
// Use CTRL+SHIFT+ENTER in the formula field to set the output cell range to
// your cell selection.
//
xlOperand* XLKIT_API
xlMatrixRef(xlParmDataRange cells)
{
	XLKIT_BEGIN_FUNCTION

	xlResultOperandPtr result;
	result->set(cells.value()->get<xlConstCellMatrixRef>());
	return result;

	XLKIT_END_FUNCTION(xlResultOperandPtr)
}
XLKIT_REGISTER(xlMatrixRef, "Reference a cell range")
