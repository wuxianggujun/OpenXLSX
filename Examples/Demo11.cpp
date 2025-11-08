#include <iostream>
#include <string>

#include <OpenXLSX.hpp>

using namespace OpenXLSX;

int main(int argc, char** argv)
{
    std::cout << "********************************************************************************\n";
    std::cout << "DEMO PROGRAM #11: Shared Formulas (Read & Expand)\n";
    std::cout << "********************************************************************************\n";

    // Modes:
    // 1) Demo11 --generate [out.xlsx]  -> Generate a sample workbook with shared formulas
    // 2) Demo11 <xlsx>                 -> Read and print formulas

    try {
        if (argc >= 2 && std::string(argv[1]) == "--generate") {
            std::string outPath = (argc >= 3) ? std::string(argv[2]) : std::string("./Demo11-Shared.xlsx");
            std::cout << "Generating shared formula workbook: " << outPath << "\n";

            XLDocument doc;
            doc.create(outPath, XLForceOverwrite);
            auto ws = doc.workbook().worksheet(1);

            // Fill some values in column B (B2..B10)
            for (uint32_t r = 2; r <= 10; ++r) {
                ws.cell(r, 2).value() = static_cast<int64_t>(r); // simple integers
            }

            // Set a shared formula in column A (A2..A10) using master formula relative to A2
            // masterFormula uses the top-left (A2) as reference base: "B2*2"
            ws.setSharedFormula("A2:A10", "=B2*2", true);
            // Pre-fill cached values (<v>) so numbers show even before recalculation
            for (uint32_t r = 2; r <= 10; ++r) {
                auto bval = ws.cell(r, 2).value().get<int64_t>();
                ws.cell(r, 1).value() = static_cast<int64_t>(bval * 2);
            }
            // Force full recalculation on next open
            doc.workbook().setFullCalculationOnLoad();

            doc.save();
            doc.close();
            std::cout << "Generated. You can run: Demo11 " << outPath << " to inspect.\n";
            return 0;
        }

        if (argc < 2) {
            // No args → auto-generate a sample with shared formulas
            std::string outPath = "./Demo11-Shared.xlsx";
            std::cout << "No args provided. Generating default sample: " << outPath << "\n";

            XLDocument doc;
            doc.create(outPath, XLForceOverwrite);
            auto ws = doc.workbook().worksheet(1);

            // Fill some values in column B (B2..B10)
            for (uint32_t r = 2; r <= 10; ++r) ws.cell(r, 2).value() = static_cast<int64_t>(r);

            // Set shared formula in A2..A10 with master at A2 using "B2*2"
            ws.setSharedFormula("A2:A10", "=B2*2", true);
            // Pre-fill cached values (<v>) so numbers show even before recalculation
            for (uint32_t r = 2; r <= 10; ++r) {
                auto bval = ws.cell(r, 2).value().get<int64_t>();
                ws.cell(r, 1).value() = static_cast<int64_t>(bval * 2);
            }
            // Force full recalculation on next open
            doc.workbook().setFullCalculationOnLoad();

            doc.save();
            doc.close();
            std::cout << "Generated. Re-run Demo11 " << outPath << " to inspect formulas.\n";
            return 0;
        }

        const std::string path = argv[1];
        std::cout << "Opening: " << path << "\n";

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
    } catch (const std::exception& ex) {
        std::cerr << "Error: " << ex.what() << "\n";
        return 2;
    }

    return 0;
}
