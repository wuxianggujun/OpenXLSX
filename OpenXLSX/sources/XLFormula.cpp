//
// Created by Kenneth Balslev on 27/08/2021.
//

// ===== External Includes ===== //
#include <algorithm>
#include <cassert>
#include <cctype>
#include <regex>
#include <string>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLCell.hpp"
#include "XLCellReference.hpp"
#include "XLFormula.hpp"
#include "XLException.hpp"
#include "XLXmlParser.hpp"              // pugixml wrapper

using namespace OpenXLSX;

namespace
{
    std::vector<std::pair<size_t, size_t>> findQuotedRanges(const std::string& formula)
    {
        std::vector<std::pair<size_t, size_t>> ranges;
        bool insideString = false;
        size_t start = 0;

        for (size_t i = 0; i < formula.size(); ++i) {
            if (formula[i] != '"') continue;

            if (!insideString) {
                insideString = true;
                start = i;
                continue;
            }

            if (i + 1 < formula.size() && formula[i + 1] == '"') {
                ++i;
                continue;
            }

            ranges.emplace_back(start, i + 1);
            insideString = false;
        }

        return ranges;
    }

    bool isInRanges(const std::vector<std::pair<size_t, size_t>>& ranges, size_t position)
    {
        for (const auto& range : ranges) {
            if (position >= range.first && position < range.second) return true;
        }
        return false;
    }

    uint16_t columnLettersToNumber(const std::string& columnLetters)
    {
        uint32_t columnNumber = 0;
        for (char ch : columnLetters) {
            const auto upper = static_cast<char>(std::toupper(static_cast<unsigned char>(ch)));
            if (upper < 'A' || upper > 'Z') break;
            columnNumber = columnNumber * 26 + static_cast<uint32_t>(upper - 'A' + 1);
        }
        return static_cast<uint16_t>(columnNumber);
    }

    std::string numberToColumnLetters(uint16_t columnNumber)
    {
        XLCellReference reference(1, columnNumber);
        std::string address = reference.address();
        size_t pos = address.size();
        while (pos > 0 && std::isdigit(static_cast<unsigned char>(address[pos - 1]))) --pos;
        return address.substr(0, pos);
    }

    std::string expandSharedFormulaString(const std::string& masterFormula,
                                          const XLCellReference& masterCell,
                                          const XLCellReference& targetCell)
    {
        if (masterFormula.empty()) return masterFormula;

        const int rowOffset = static_cast<int>(targetCell.row()) - static_cast<int>(masterCell.row());
        const int columnOffset = static_cast<int>(targetCell.column()) - static_cast<int>(masterCell.column());
        const std::regex referenceRegex(R"(((?:'[^']+'|[A-Za-z_][\w\.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7}))");
        const auto quotedRanges = findQuotedRanges(masterFormula);

        std::string result;
        result.reserve(masterFormula.size());
        size_t last = 0;

        for (std::sregex_iterator it(masterFormula.begin(), masterFormula.end(), referenceRegex), end; it != end; ++it) {
            const size_t position = static_cast<size_t>(it->position());
            const size_t length = static_cast<size_t>(it->length());
            if (isInRanges(quotedRanges, position)) continue;

            if (position > last) result.append(masterFormula, last, position - last);

            const std::smatch match = *it;
            const std::string sheetPart = match[1].str();
            const bool columnAbsolute = !match[2].str().empty();
            const bool rowAbsolute = !match[4].str().empty();
            const uint16_t column = columnLettersToNumber(match[3].str());
            const uint32_t row = static_cast<uint32_t>(std::stoul(match[5].str()));

            const int adjustedColumn = columnAbsolute ? static_cast<int>(column) : static_cast<int>(column) + columnOffset;
            const int adjustedRow = rowAbsolute ? static_cast<int>(row) : static_cast<int>(row) + rowOffset;

            std::string columnText = numberToColumnLetters(static_cast<uint16_t>(std::max(1, adjustedColumn)));
            if (columnAbsolute) columnText = "$" + columnText;

            std::string rowText = std::to_string(std::max(1, adjustedRow));
            if (rowAbsolute) rowText = "$" + rowText;

            result += sheetPart + columnText + rowText;
            last = position + length;
        }

        if (last < masterFormula.size()) result.append(masterFormula, last, std::string::npos);
        return result.empty() ? masterFormula : result;
    }

    bool findSharedFormulaMaster(XMLNode sheetDataNode,
                                 uint32_t sharedIndex,
                                 XLCellReference& masterCell,
                                 std::string& masterFormula)
    {
        for (XMLNode row = sheetDataNode.child("row"); !row.empty(); row = row.next_sibling("row")) {
            for (XMLNode cell = row.child("c"); !cell.empty(); cell = cell.next_sibling("c")) {
                const XMLNode formulaNode = cell.child("f");
                if (formulaNode.empty()) continue;

                const auto typeAttribute = formulaNode.attribute("t");
                if (typeAttribute.empty() || std::string(typeAttribute.value()) != "shared") continue;

                const auto sharedIndexAttribute = formulaNode.attribute("si");
                if (sharedIndexAttribute.empty() || sharedIndexAttribute.as_uint() != sharedIndex) continue;

                const std::string formulaText = formulaNode.text().get();
                if (formulaText.empty()) continue;

                masterFormula = formulaText;
                masterCell = XLCellReference(cell.attribute("r").value());
                return true;
            }
        }

        return false;
    }
}

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
 * @details [private] Constructor. Set the m_cell and m_cellNode objects.
 */
XLFormulaProxy::XLFormulaProxy(XLCell* cell, XMLNode* cellNode)
 : m_cell(cell),
   m_cellNode(cellNode)
{
    assert(cell);    // NOLINT
}

/**
 * @details [private] Copy constructor
 */
XLFormulaProxy::XLFormulaProxy(const XLFormulaProxy& other)
 : m_cell(other.m_cell),
   m_cellNode(other.m_cellNode)
{}

/**
 * @details [private] Move constructor. Default implementation.
 */
XLFormulaProxy::XLFormulaProxy(XLFormulaProxy&& other) noexcept = default;

/**
 * @details Destructor. Default implementation.
 */
XLFormulaProxy::~XLFormulaProxy() = default;

/**
 * @details Copy assignment operator. Calls the templated string assignment operator.
 */
XLFormulaProxy& XLFormulaProxy::operator=(const XLFormulaProxy& other)
{
    if (&other != this) {
        *this = other.getFormula();
    }

    return *this;
}

/**
 * @details [private] Move assignment operator. Default implementation.
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
    m_cellNode->child("f").remove_attribute("ref");

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
 * XML document.
 */
XLFormula XLFormulaProxy::getFormula() const
{
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    const auto formulaNode = m_cellNode->child("f");

    // ===== If the formula node doesn't exist, return an empty XLFormula object.
    if (formulaNode.empty()) return XLFormula();

    // ===== If the formula type is 'shared', expand it from the master formula when needed.
    if (not formulaNode.attribute("t").empty() ) {    // 2024-05-28: de-duplicated check (only relevant for performance,
                                                      //  xml_attribute::value() returns an empty string for empty attributes)
        if (std::string(formulaNode.attribute("t").value()) == "shared") {
            const std::string formulaText = formulaNode.text().get();
            if (!formulaText.empty()) return XLFormula(formulaText);

            const auto sharedIndexAttribute = formulaNode.attribute("si");
            if (sharedIndexAttribute.empty()) return XLFormula();

            XLCellReference masterCell;
            std::string masterFormula;
            if (!findSharedFormulaMaster(m_cellNode->parent().parent(), sharedIndexAttribute.as_uint(), masterCell, masterFormula))
                return XLFormula();

            return XLFormula(expandSharedFormulaString(masterFormula, masterCell, m_cell->cellReference()));
        }
        if (std::string(formulaNode.attribute("t").value()) == "array")
            throw XLFormulaError("Array formulas not supported.");
    }

    return XLFormula(formulaNode.text().get());
}
