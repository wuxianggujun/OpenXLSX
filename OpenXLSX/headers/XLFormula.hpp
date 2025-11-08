/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2018, Kenneth Troldal Balslev

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

#ifndef OPENXLSX_XLFORMULA_HPP
#define OPENXLSX_XLFORMULA_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#   pragma warning(push)
#   pragma warning(disable : 4251)
#   pragma warning(disable : 4275)
#endif // _MSC_VER

// ===== External Includes ===== //
#include <iostream>
#include <string>
#include <variant>
#include <cstdint>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"

// ========== CLASS AND ENUM TYPE DEFINITIONS ========== //
namespace OpenXLSX
{
    constexpr bool XLResetValue    = true;
    constexpr bool XLPreserveValue = false;

    // Formula type (enhanced reading metadata)
    enum class XLFormulaType {
        Normal,
        Shared,
        Array,      // not yet supported
        DataTable   // not yet supported
    };

    //---------- Forward Declarations ----------//
    class XLFormulaProxy;
    class XLCell;

    /**
     * @brief The XLFormula class encapsulates the concept of an Excel formula. The class is essentially
     * a wrapper around a std::string.
     * @warning This class currently only supports simple formulas. Array formulas and shared formulas are
     * not supported. Unfortunately, many spreadsheets have shared formulas, so this class is probably
     * best used for adding formulas, not reading them from an existing spreadsheet.
     * @todo Enable handling of shared and array formulas.
     */
    class OPENXLSX_EXPORT XLFormula
    {
        //---------- Friend Declarations ----------//

        friend bool          operator==(const XLFormula& lhs, const XLFormula& rhs);
        friend bool          operator!=(const XLFormula& lhs, const XLFormula& rhs);
        friend std::ostream& operator<<(std::ostream& os, const XLFormula& value);

    public:
        /**
         * @brief Constructor
         */
        XLFormula();

        /**
         * @brief Constructor, taking a string-type argument
         * @tparam T Type of argument used. Must be string-type.
         * @param formula The formula to initialize the object with.
         */
        template<typename T,
                 typename = std::enable_if_t<
                     std::is_same_v<std::decay_t<T>, std::string> || std::is_same_v<std::decay_t<T>, std::string_view> ||
                     std::is_same_v<std::decay_t<T>, const char*> || std::is_same_v<std::decay_t<T>, char*>>>
        explicit XLFormula(T formula)
        {
            // ===== If the argument is a const char *, use the argument directly; otherwise, assume it has a .c_str() function.
            if constexpr (std::is_same_v<std::decay_t<T>, const char*> || std::is_same_v<std::decay_t<T>, char*>)
                m_formulaString = formula;
            else if constexpr (std::is_same_v<std::decay_t<T>, std::string_view>)
                m_formulaString = std::string(formula);
            else
                m_formulaString = formula.c_str();
        }

        /**
         * @brief Copy constructor.
         * @param other Object to be copied.
         */
        XLFormula(const XLFormula& other);

        /**
         * @brief Move constructor.
         * @param other Object to be moved.
         */
        XLFormula(XLFormula&& other) noexcept;

        /**
         * @brief Destructor.
         */
        ~XLFormula();

        /**
         * @brief Copy assignment operator.
         * @param other Object to be copied.
         * @return Reference to copied-to object.
         */
        XLFormula& operator=(const XLFormula& other);

        /**
         * @brief Move assignment operator.
         * @param other Object to be moved.
         * @return Reference to moved-to object.
         */
        XLFormula& operator=(XLFormula&& other) noexcept;

        /**
         * @brief Templated assignment operator, taking a string-type object as an argument.
         * @tparam T Type of argument (only string-types are allowed).
         * @param formula String containing the formula.
         * @return Reference to the assigned-to object.
         */
        template<typename T,
                 typename = std::enable_if_t<
                     std::is_same_v<std::decay_t<T>, std::string> || std::is_same_v<std::decay_t<T>, std::string_view> ||
                     std::is_same_v<std::decay_t<T>, const char*> || std::is_same_v<std::decay_t<T>, char*>>>
        XLFormula& operator=(T formula)
        {
            XLFormula temp(formula);
            std::swap(*this, temp);
            return *this;
        }

        /**
         * @brief Templated setter function, taking a string-type object as an argument.
         * @tparam T Type of argument (only string-types are allowed).
         * @param formula String containing the formula.
         */
        template<typename T,
                 typename = std::enable_if_t<
                     std::is_same_v<std::decay_t<T>, std::string> || std::is_same_v<std::decay_t<T>, std::string_view> ||
                     std::is_same_v<std::decay_t<T>, const char*> || std::is_same_v<std::decay_t<T>, char*>>>
        void set(T formula)
        {
            *this = formula;
        }

        /**
         * @brief Get the formula as a std::string.
         * @return A std::string with the formula.
         */
        std::string get() const;

        // ==== Shared formula metadata (enhanced API) ====
        XLFormulaType type() const { return m_type; }
        void          setType(XLFormulaType type) { m_type = type; }
        uint32_t      sharedIndex() const { return m_sharedIndex; }
        void          setSharedIndex(uint32_t index) { m_sharedIndex = index; }
        std::string   sharedRange() const { return m_sharedRange; }
        void          setSharedRange(const std::string& range) { m_sharedRange = range; }
        bool          isShared() const { return m_type == XLFormulaType::Shared; }

        /**
         * @brief Conversion operator, for converting object to a std::string.
         * @return The formula as a std::string.
         */
        operator std::string() const;    // NOLINT

        /**
         * @brief Clear the formula.
         * @return Return a reference to the cleared object.
         */
        XLFormula& clear();

    private:
        std::string  m_formulaString;                 /**< formula string */
        XLFormulaType m_type { XLFormulaType::Normal };/**< formula type */
        uint32_t      m_sharedIndex { 0 };             /**< shared formula index (si) */
        std::string   m_sharedRange;                   /**< shared formula range (master only) */
    };

    /**
     * @brief The XLFormulaProxy serves as a placeholder for XLFormula objects. This enable
     * getting and setting formulas through the same interface.
     */
    class OPENXLSX_EXPORT XLFormulaProxy
    {
        friend class XLCell;
        friend class XLFormula;

    public:
        /**
         * @brief Destructor
         */
        ~XLFormulaProxy();

        /**
         * @brief Copy assignment operator.
         * @param other Object to be copied.
         * @return A reference to the copied-to object.
         */
        XLFormulaProxy& operator=(const XLFormulaProxy& other);

        /**
         * @brief Templated assignment operator, taking a string-type argument.
         * @tparam T Type of argument.
         * @param formula The formula string to be assigned.
         * @return A reference to the copied-to object.
         */
        template<
            typename T,
            typename = std::enable_if_t<std::is_same_v<std::decay_t<T>, XLFormula> || std::is_same_v<std::decay_t<T>, std::string> ||
                                        std::is_same_v<std::decay_t<T>, std::string_view> || std::is_same_v<std::decay_t<T>, const char*> ||
                                        std::is_same_v<std::decay_t<T>, char*>>>
        XLFormulaProxy& operator=(T formula)
        {
            if constexpr (std::is_same_v<std::decay_t<T>, XLFormula>)
                setFormulaString(formula.get().c_str());
            else if constexpr (std::is_same_v<std::decay_t<T>, std::string>)
                setFormulaString(formula.c_str());
            else if constexpr (std::is_same_v<std::decay_t<T>, std::string_view>)
                setFormulaString(std::string(formula).c_str());
            else
                setFormulaString(formula);

            return *this;
        }

        /**
         * @brief Templated setter, taking a string-type argument.
         * @tparam T Type of argument.
         * @param formula The formula string to be assigned.
         */
        template<typename T,
                 typename = std::enable_if_t<
                     std::is_same_v<std::decay_t<T>, std::string> || std::is_same_v<std::decay_t<T>, std::string_view> ||
                     std::is_same_v<std::decay_t<T>, const char*> || std::is_same_v<std::decay_t<T>, char*>>>
        void set(T formula)
        {
            *this = formula;
        }

        /**
         * @brief Get the formula as a std::string.
         * @return A std::string with the formula.
         */
        std::string get() const;

        /**
         * @brief Clear the formula.
         * @return Return a reference to the cleared object.
         */
        XLFormulaProxy& clear();

        // Get the raw formula object (no shared expansion); keeps type/index metadata
        XLFormula getRawFormula() const;

        /**
         * @brief 设置主共享公式（仅供高级用法）。
         *        在当前单元格创建 <f t="shared" si=".." ref="..">masterFormula</f>
         * @param sharedIndex 共享索引（工作表内唯一）
         * @param rangeRef    共享范围（如 "A1:A10"）
         * @param masterFormula 主公式文本（以当前单元格为基准）
         * @param resetValue 若为 true，则将 <v> 节点值设为 0
         * @return true 表示已写入
         */
        bool setSharedMaster(uint32_t sharedIndex,
                             const std::string& rangeRef,
                             const std::string& masterFormula,
                             bool resetValue = XLResetValue);

        /**
         * @brief 设置共享公式引用单元格（仅供高级用法）。
         *        在当前单元格创建 <f t="shared" si=".."/>
         * @param sharedIndex 共享索引
         * @param resetValue 若为 true，则将 <v> 节点值设为 0
         * @return true 表示已写入
         */
        bool setSharedRef(uint32_t sharedIndex, bool resetValue = XLResetValue);

        /**
         * @brief Conversion operator, for converting the object to a std::string.
         * @return The formula as a std::string.
         */
        operator std::string() const;    // NOLINT

        /**
         * @brief Implicit conversion to XLFormula object.
         * @return Returns the corresponding XLFormula object.
         */
        operator XLFormula() const;    // NOLINT

    private:
        /**
         * @brief Constructor, taking pointers to the cell and cell node objects.
         * @param cell Pointer to the associated cell object.
         * @param cellNode Pointer to the associated cell node object.
         */
        XLFormulaProxy(XLCell* cell, XMLNode* cellNode);

        /**
         * @brief Copy constructor.
         * @param other Object to be copied.
         */
        XLFormulaProxy(const XLFormulaProxy& other);

        /**
         * @brief Move constructor.
         * @param other Object to be moved.
         */
        XLFormulaProxy(XLFormulaProxy&& other) noexcept;

        /**
         * @brief Move assignment operator.
         * @param other Object to be moved.
         * @return A reference to the moved-to object.
         */
        XLFormulaProxy& operator=(XLFormulaProxy&& other) noexcept;

        /**
         * @brief Set the formula to the given string.
         * @param formulaString String holding the formula.
         * @param resetValue if true (XLResetValue), the cell value will be set to 0, if false (XLPreserveValue), it will remain unchanged
         */
        void setFormulaString(const char* formulaString, bool resetValue = XLResetValue);

        /**
         * @brief Get the underlying XLFormula object.
         *        Shared formulas are expanded to the actual formula for this cell.
         */
        XLFormula getFormula() const;

        //---------- Private Member Variables ---------- //
        XLCell*  m_cell;     /**< Pointer to the owning XLCell object. */
        XMLNode* m_cellNode; /**< Pointer to corresponding XML cell node. */
    };
}    // namespace OpenXLSX

// ========== FRIEND FUNCTION IMPLEMENTATIONS ========== //
namespace OpenXLSX
{
    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator==(const XLFormula& lhs, const XLFormula& rhs) { return lhs.m_formulaString == rhs.m_formulaString; }

    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator!=(const XLFormula& lhs, const XLFormula& rhs) { return lhs.m_formulaString != rhs.m_formulaString; }

    /**
     * @brief send a formula string to an ostream
     * @param os the output destination
     * @param value the formula to send to os
     * @return
     */
    inline std::ostream& operator<<(std::ostream& os, const XLFormula& value) { return os << value.m_formulaString; }

    /**
     * @brief send a formula string to an ostream
     * @param os the output destination
     * @param value the formula proxy whose formula to send to os
     * @return
     */
    inline std::ostream& operator<<(std::ostream& os, const XLFormulaProxy& value) { return os << value.get(); }
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#   pragma warning(pop)
#endif // _MSC_VER

#endif    // OPENXLSX_XLFORMULA_HPP
