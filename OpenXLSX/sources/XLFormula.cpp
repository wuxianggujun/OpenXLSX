//
// Created by Kenneth Balslev on 27/08/2021.
//

// ===== External Includes ===== //
#include <cassert>
#include <regex>
#include <cctype>
#include <string>
#include <vector>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLFormula.hpp"
#include "XLException.hpp"
#include "XLCellReference.hpp"
#include "XLCell.hpp"

using namespace OpenXLSX;

/**
 * @details Constructor. Default implementation.
 */
XLFormula::XLFormula() = default;

/**
 * @details Copy constructor. Default implementation.
 */
XLFormula::XLFormula(const XLFormula& other) = default;

/**
 * @details Move constructor. Default implementation.
 */
XLFormula::XLFormula(XLFormula&& other) noexcept = default;

/**
 * @details Destructor. Default implementation.
 */
XLFormula::~XLFormula() = default;

/**
 * @details Copy assignment operator. Default implementation.
 */
XLFormula& XLFormula::operator=(const XLFormula& other) = default;

/**
 * @details Move assignment operator. Default implementation.
 */
XLFormula& XLFormula::operator=(XLFormula&& other) noexcept = default;

/**
 * @details Return the m_formulaString member variable.
 */
std::string XLFormula::get() const { return m_formulaString; }

/**
 * @details Set the m_formulaString member to an empty string.
 */
XLFormula& XLFormula::clear()
{
    m_formulaString = "";
    return *this;
}

/**
 * @details
 */
XLFormula::operator std::string() const { return get(); }

/**
 * @details Constructor. Set the m_cell and m_cellNode objects.
 */
XLFormulaProxy::XLFormulaProxy(XLCell* cell, XMLNode* cellNode) : m_cell(cell), m_cellNode(cellNode)
{
    assert(cell);    // NOLINT
}

/**
 * @details Destructor. Default implementation.
 */
XLFormulaProxy::~XLFormulaProxy() = default;

/**
 * @details Copy constructor. default implementation.
 */
XLFormulaProxy::XLFormulaProxy(const XLFormulaProxy& other) = default;

/**
 * @details Move constructor. Default implementation.
 */
XLFormulaProxy::XLFormulaProxy(XLFormulaProxy&& other) noexcept = default;

/**
 * @details Calls the templated string assignment operator.
 */
XLFormulaProxy& XLFormulaProxy::operator=(const XLFormulaProxy& other)
{
    if (&other != this) {
        *this = other.getFormula();
    }

    return *this;
}

/**
 * @details Move assignment operator. Default implementation.
 */
XLFormulaProxy& XLFormulaProxy::operator=(XLFormulaProxy&& other) noexcept = default;

/**
 * @details
 */
XLFormulaProxy::operator std::string() const { return get(); }

/**
 * @details Returns the underlying XLFormula object, by calling getFormula().
 */
XLFormulaProxy::operator XLFormula() const { return getFormula(); }

/**
 * @details Call the .get() function in the underlying XLFormula object.
 */
std::string XLFormulaProxy::get() const { return getFormula().get(); }

/**
 * @details If a formula node exists, it will be erased.
 */
XLFormulaProxy& XLFormulaProxy::clear()
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    // ===== Remove the value node.
    if (not m_cellNode->child("f").empty()) m_cellNode->remove_child("f");
    return *this;
}

// -------------------- Helpers for Shared Formula expansion --------------------
namespace {

    // find quoted string ranges to avoid touching refs inside strings
    std::vector<std::pair<size_t,size_t>> findQuotedRanges(const std::string& s)
    {
        std::vector<std::pair<size_t,size_t>> ranges;
        bool in = false;
        size_t start = 0;
        for (size_t i = 0; i < s.size(); ++i) {
            if (s[i] == '"') {
                if (!in) { in = true; start = i; }
                else {
                    // check escaped double quote
                    if (i + 1 < s.size() && s[i+1] == '"') { ++i; continue; }
                    ranges.emplace_back(start, i + 1);
                    in = false;
                }
            }
        }
        return ranges;
    }

    bool inRanges(const std::vector<std::pair<size_t,size_t>>& ranges, size_t pos)
    {
        for (auto const& r : ranges) {
            if (pos >= r.first && pos < r.second) return true;
        }
        return false;
    }

    // convert column letters to number (A=1)
    uint16_t columnLettersToNumber(const std::string& col)
    {
        uint32_t colNo = 0;
        for (char c : col) {
            if (c >= 'A' && c <= 'Z') colNo = colNo * 26 + (c - 'A' + 1);
            else break;
        }
        return static_cast<uint16_t>(colNo);
    }

    // convert number to column letters (1=A)
    std::string numberToColumnLetters(uint16_t column)
    {
        // reuse XLCellReference to avoid re-implementing; strip row afterwards
        OpenXLSX::XLCellReference tmp(1, column);
        std::string addr = tmp.address(); // e.g. A1
        // strip trailing digits
        size_t i = addr.size();
        while (i > 0 && std::isdigit(static_cast<unsigned char>(addr[i-1]))) --i;
        return addr.substr(0, i);
    }

    std::string expandSharedFormulaString(const std::string& masterFormula,
                                          const OpenXLSX::XLCellReference& masterCell,
                                          const OpenXLSX::XLCellReference& targetCell)
    {
        if (masterFormula.empty()) return masterFormula;
        int rowOffset = static_cast<int>(targetCell.row()) - static_cast<int>(masterCell.row());
        int colOffset = static_cast<int>(targetCell.column()) - static_cast<int>(masterCell.column());

        // pattern: optional sheet ('..' or bare)! then optional $ col letters, optional $ row digits
        std::regex refRe(R"(((?:'[^']+'|[A-Za-z_][\w\.]*)!)?(\$?)([A-Z]{1,3})(\$?)([0-9]{1,7}))");
        auto quoted = findQuotedRanges(masterFormula);

        std::string result;
        result.reserve(masterFormula.size());
        std::sregex_iterator it(masterFormula.begin(), masterFormula.end(), refRe);
        std::sregex_iterator end;
        size_t last = 0;
        for (; it != end; ++it) {
            size_t pos = static_cast<size_t>((*it).position());
            size_t len = static_cast<size_t>((*it).length());
            // skip if inside quoted string
            if (inRanges(quoted, pos)) continue;
            // append text before match
            if (pos > last) result.append(masterFormula, last, pos - last);

            std::smatch m = *it;
            std::string sheetPart = m[1].str();
            bool colAbs = !m[2].str().empty();
            std::string colLetters = m[3].str();
            bool rowAbs = !m[4].str().empty();
            uint32_t row = static_cast<uint32_t>(std::stoul(m[5].str()));

            uint16_t col = columnLettersToNumber(colLetters);
            uint16_t newCol = static_cast<uint16_t>(colAbs ? col : static_cast<int>(col) + colOffset);
            uint32_t newRow = static_cast<uint32_t>(rowAbs ? row : static_cast<int>(row) + rowOffset);

            if (newCol < 1) newCol = 1; // basic guard
            if (newRow < 1) newRow = 1;

            std::string newColLetters = numberToColumnLetters(newCol);
            if (colAbs) newColLetters = "$" + newColLetters;
            std::string newRowStr = std::to_string(newRow);
            if (rowAbs) newRowStr = "$" + newRowStr;

            result += sheetPart + newColLetters + newRowStr;
            last = pos + len;
        }
        // append remaining tail
        if (last < masterFormula.size()) result.append(masterFormula, last, std::string::npos);
        if (result.empty()) return masterFormula; // in case every match fell into quoted ranges
        return result;
    }

    bool findMasterSharedFormulaForIndex(const pugi::xml_node& sheetData,
                                         uint32_t si,
                                         OpenXLSX::XLCellReference& masterOut,
                                         std::string& masterFormulaOut,
                                         std::string& rangeOut)
    {
        for (auto row = sheetData.child("row"); !row.empty(); row = row.next_sibling("row")) {
            for (auto c = row.child("c"); !c.empty(); c = c.next_sibling("c")) {
                auto f = c.child("f");
                if (f.empty()) continue;
                auto tAttr = f.attribute("t");
                if (tAttr.empty() || std::string(tAttr.value()) != "shared") continue;
                auto siAttr = f.attribute("si");
                if (siAttr.empty() || siAttr.as_uint() != si) continue;
                auto text = std::string(f.text().get());
                if (!text.empty()) {
                    masterFormulaOut = text;
                    rangeOut = f.attribute("ref").as_string("");
                    masterOut = OpenXLSX::XLCellReference{ c.attribute("r").value() };
                    return true;
                }
            }
        }
        return false;
    }
}

/**
 * @details Convenience function for setting the formula. This method is called from the templated
 * string assignment operator.
 */
void XLFormulaProxy::setFormulaString(const char* formulaString, bool resetValue) // NOLINT
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    if (formulaString[0] == 0) {    // if formulaString is empty
        m_cellNode->remove_child("f");    // clear the formula node
        return;                           // and exit
    }

    // ===== If the cell node doesn't have formula or value child nodes, create them.
    if (m_cellNode->child("f").empty()) m_cellNode->append_child("f");
    if (m_cellNode->child("v").empty()) m_cellNode->append_child("v");

    // ===== Remove the formula type and shared index attributes, if they exist.
    m_cellNode->child("f").remove_attribute("t");
    m_cellNode->child("f").remove_attribute("si");

    // ===== Set the text of the formula and value nodes.
    m_cellNode->child("f").text().set(formulaString);
    if (resetValue) m_cellNode->child("v").text().set(0);

    // BEGIN pull request #189
    // ===== Remove cell type attribute so that it can be determined by Office Suite when next calculating the formula.
    m_cellNode->remove_attribute("t");

    // ===== Remove inline string <is> tag, in case previous type was "inlineStr".
    m_cellNode->remove_child("is");

    // ===== Ensure that the formula node <f> is the first child, listed before the value <v> node.
    m_cellNode->prepend_move(m_cellNode->child("f"));
    // END pull request #189
}

/**
 * @details Creates and returns an XLFormula object, based on the formula string in the underlying
 * XML document. 若遇到共享公式，将尝试展开为当前单元格对应的实际公式。
 */
XLFormula XLFormulaProxy::getFormula() const
{
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    const auto formulaNode = m_cellNode->child("f");

    // ===== If the formula node doesn't exist, return an empty XLFormula object.
    if (formulaNode.empty()) return XLFormula();

    // ===== Handle specific types
    std::string t = formulaNode.attribute("t").empty() ? std::string() : std::string(formulaNode.attribute("t").value());
    if (t == "array")
        throw XLFormulaError("Array formulas not supported.");

    if (t == "shared") {
        // If this is the master (has text), return it directly and mark metadata
        XLFormula f;
        f.setType(XLFormulaType::Shared);
        f.setSharedIndex(formulaNode.attribute("si").as_uint());
        auto text = std::string(formulaNode.text().get());
        auto refAttr = std::string(formulaNode.attribute("ref").as_string(""));
        if (!text.empty()) {
            f = text.c_str();
            if (!refAttr.empty()) f.setSharedRange(refAttr);
            return f;
        }

        // Non-master: find master and expand
        const auto sheetData = m_cellNode->parent().parent();
        OpenXLSX::XLCellReference masterRef;
        std::string masterText;
        std::string range;
        if (findMasterSharedFormulaForIndex(sheetData, f.sharedIndex(), masterRef, masterText, range)) {
            auto expanded = expandSharedFormulaString(masterText, masterRef, m_cell->cellReference());
            f = expanded.c_str();
            if (!range.empty()) f.setSharedRange(range);
            return f;
        }

        // fallback: return empty normal formula
        return XLFormula();
    }

    // normal
    return XLFormula(formulaNode.text().get());
}

/**
 * @brief 获取原始公式，不做共享展开（若非主单元格则只返回元信息）
 */
XLFormula XLFormulaProxy::getRawFormula() const
{
    assert(m_cellNode != nullptr);
    assert(not m_cellNode->empty());
    const auto formulaNode = m_cellNode->child("f");
    if (formulaNode.empty()) return XLFormula();

    std::string t = formulaNode.attribute("t").empty() ? std::string() : std::string(formulaNode.attribute("t").value());
    if (t == "array")
        throw XLFormulaError("Array formulas not supported.");

    if (t == "shared") {
        XLFormula f;
        f.setType(XLFormulaType::Shared);
        f.setSharedIndex(formulaNode.attribute("si").as_uint());
        auto text = std::string(formulaNode.text().get());
        auto refAttr = std::string(formulaNode.attribute("ref").as_string(""));
        if (!text.empty()) f = text.c_str();
        if (!refAttr.empty()) f.setSharedRange(refAttr);
        return f;
    }

    return XLFormula(formulaNode.text().get());
}
