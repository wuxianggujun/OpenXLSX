#include <iostream>
#include <string>

#include <OpenXLSX.hpp>
#include <filesystem>
#if defined(_WIN32)
#  include <nowide/convert.hpp>
#endif

using namespace OpenXLSX;

// Convert Excel column letters (e.g. "A", "Z", "AA") to 1-based column index
static uint16_t colLettersToIndex(const std::string& s)
{
    if (s.empty()) return 0;
    uint32_t idx = 0;
    for (char ch : s) {
        if (ch >= 'a' && ch <= 'z') ch = static_cast<char>(ch - 'a' + 'A');
        if (ch < 'A' || ch > 'Z') return 0;
        idx = idx * 26 + (static_cast<uint32_t>(ch - 'A') + 1);
        if (idx > std::numeric_limits<uint16_t>::max()) return 0;
    }
    return static_cast<uint16_t>(idx);
}

int main(int argc, char** argv)
{
    std::cout << "********************************************************************************\n";
    std::cout << "DEMO PROGRAM #12: Print Shared Formulas From 3rd Sheet (Hardcoded Path)\n";
    std::cout << "********************************************************************************\n";
    std::cout << "Usage: Demo12 [COLUMN|ALL]  (default COLUMN=G)\n";

    // Build path with std::filesystem + universal character names to avoid codepage pitfalls
#if defined(_WIN32)
    std::filesystem::path p = L"C:\\Users\\wuxianggujun\\CodeSpace\\CMakeProjects\\IntegratedPower";
    p /= L"\u81EA\u52A8\u5316";
    p /= L"\u5C0F\u6BB5-20251107";
    p /= L"HTDDPSD3.0251024001-0130C00124\u751F\u4EA7\u8981\u6C42(\u7532\u65B9\u5BA2\u6237\u62A5\u8868)-1.xlsx";
    const std::string  path  = nowide::narrow(p.native());
#else
    std::filesystem::path p = 
        L"C:/Users/wuxianggujun/CodeSpace/CMakeProjects/IntegratedPower";
    p /= L"\u81EA\u52A8\u5316";
    p /= L"\u5C0F\u6BB5-20251107";
    p /= L"HTDDPSD3.0251024001-0130C00124\u751F\u4EA7\u8981\u6C42(\u7532\u65B9\u5BA2\u6237\u62A5\u8868)-1.xlsx";
    const std::string  path  = p.u8string();
#endif

    // Optional column selector
    bool scanAll = false;
    uint16_t targetCol = 7; // default 'G'
    if (argc >= 2) {
        std::string arg = argv[1];
        if (arg == "ALL" || arg == "all") {
            scanAll = true;
        } else {
            auto idx = colLettersToIndex(arg);
            if (idx == 0) {
                std::cerr << "Invalid column identifier: " << arg << "\n";
                return 2;
            }
            targetCol = idx;
        }
    }

    try {
        XLDocument doc;
        doc.open(path);

        // 第3张工作表（1-based index）
        auto ws = doc.workbook().worksheet(3);

        const uint32_t rowCount = ws.rowCount();
        const uint16_t colCount = ws.columnCount();

        std::size_t totalFormula = 0;
        std::size_t sharedFormula = 0;
        std::size_t arrayFormula  = 0;
        std::size_t normalFormula = 0;

        for (uint16_t c = 1; c <= colCount; ++c) {
            if (!scanAll && c != targetCol) continue;
            for (uint32_t r = 1; r <= rowCount; ++r) {
                auto cell = ws.findCell(r, c);
                if (!cell || cell.empty() || !cell.hasFormula()) continue;

                bool isShared = false;
                bool isArray = false;
                std::string expanded;
                std::string info;

                // 获取展开公式
                try {
                    expanded = cell.formula().get();
                }
                catch (const std::exception& ex) {
                    expanded = std::string("[error: ") + ex.what() + "]";
                }

                // 获取原始元信息
                try {
                    auto raw = cell.formula().getRawFormula();
                    isShared = raw.isShared();
                    if (isShared) {
                        info += " si=" + std::to_string(raw.sharedIndex());
                        if (!raw.sharedRange().empty()) info += " ref=" + raw.sharedRange();
                        if (!raw.get().empty()) info += " master=\"" + raw.get() + "\"";
                    }
                }
                catch (const std::exception& ex) {
                    isArray = true;
                    info += std::string(" type=array err=") + ex.what();
                }

                const auto addr = cell.cellReference().address();
                std::cout << addr << ":\n";
                std::cout << "  expanded: " << expanded << "\n";
                if (isArray)
                    std::cout << "  [array]" << (info.empty() ? "" : (" " + info)) << "\n";
                else if (isShared)
                    std::cout << "  [shared]" << (info.empty() ? "" : (" " + info)) << "\n";
                else
                    std::cout << "  [normal]" << "\n";

                // 计算结果（若存在 <v>）
                try {
                    auto dv = cell.calculatedValue().get<double>();
                    std::cout << "  value   : " << dv << "\n";
                }
                catch (...) {
                    try {
                        auto sv = cell.calculatedValue().get<std::string>();
                        std::cout << "  value   : " << sv << "\n";
                    }
                    catch (...) {
                        std::cout << "  value   : (unavailable)\n";
                    }
                }

                ++totalFormula;
                if (isShared) ++sharedFormula;
                else if (isArray) ++arrayFormula;
                else ++normalFormula;
            }
        }

        if (scanAll) {
            std::cout << "Total formulas in sheet:  " << totalFormula  << "\n";
            std::cout << "Shared formulas in sheet: " << sharedFormula << "\n";
            std::cout << "Array formulas in sheet:  " << arrayFormula  << "\n";
            std::cout << "Normal formulas in sheet: " << normalFormula << "\n";
        } else {
            std::cout << "Total formulas in column:  " << totalFormula  << "\n";
            std::cout << "Shared formulas in column: " << sharedFormula << "\n";
            std::cout << "Array formulas in column:  " << arrayFormula  << "\n";
            std::cout << "Normal formulas in column: " << normalFormula << "\n";
        }
        doc.close();
    }
    catch (const std::exception& ex) {
        std::cerr << "Error: " << ex.what() << "\n";
        return 2;
    }

    return 0;
}
