#include <iostream>
#include <string>

#include <OpenXLSX.hpp>

using namespace OpenXLSX;

int main(int argc, char** argv)
{
    std::cout << "********************************************************************************\n";
    std::cout << "DEMO PROGRAM #11: Shared Formulas (Read & Expand)\n";
    std::cout << "********************************************************************************\n";

    if (argc < 2) {
        std::cout << "Usage: Demo11 <xlsx_with_shared_formulas>\n";
        std::cout << "Hint : Open an Excel file that contains shared formulas (e.g. filled-down formulas).\n";
        return 1;
    }

    const std::string path = argv[1];
    std::cout << "Opening: " << path << "\n";

    try {
        XLDocument doc;
        doc.open(path);

        auto ws = doc.workbook().worksheet(1);
        auto used = ws.range();

        for (auto& cell : used) {
            if (!cell || !cell.hasFormula()) continue;

            auto ref = cell.cellReference().address();

            // Expanded formula string (will expand shared formulas)
            std::string expanded = cell.formula().get();

            // Raw formula object with metadata
            auto raw = cell.formula().getRawFormula();

            std::cout << ref << ":\n";
            std::cout << "  expanded: " << expanded << "\n";
            if (raw.isShared()) {
                std::cout << "  [shared] si=" << raw.sharedIndex();
                if (!raw.sharedRange().empty()) std::cout << " ref=" << raw.sharedRange();
                if (!raw.get().empty()) std::cout << " master=\"" << raw.get() << "\"";
                std::cout << "\n";
            }

            // Calculated value (if present)
            try {
                auto val = cell.calculatedValue();
                // print as string for simplicity
                std::cout << "  value   : " << val.get<std::string>() << "\n";
            }
            catch (...) {
                // fall back to double
                try {
                    auto vd = cell.calculatedValue().get<double>();
                    std::cout << "  value   : " << vd << "\n";
                }
                catch (...) {
                    std::cout << "  value   : (unavailable)\n";
                }
            }
        }

        doc.close();
    }
    catch (const std::exception& ex) {
        std::cerr << "Error: " << ex.what() << "\n";
        return 2;
    }

    return 0;
}

