// Include xlkit.cpp into *ONE* of your .cpp files to define the Excel entry
// points and initialization of static variables used by XLKit. NOTE: This will
// bring in Windows.h! Alternatively, you may compile xlkit.cpp into its own
// library and link to it from your application.
#include <xlkit/xlkit.cpp>

// Set the label that shows up in the Add-in Manager. You should only do this
// once for your XLL.
XLKIT_INIT_ADDIN_LABEL("XLKit Test Addin")
