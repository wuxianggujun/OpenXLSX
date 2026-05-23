//
// Created by Kenneth Balslev on 27/08/2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>
#include <string>

using namespace OpenXLSX;

namespace
{
    std::string replaceSheetData(std::string xml, const std::string& sheetData)
    {
        const std::string sheetDataOpen = "<sheetData";
        const auto begin = xml.find(sheetDataOpen);
        REQUIRE(begin != std::string::npos);

        const auto openEnd = xml.find('>', begin);
        REQUIRE(openEnd != std::string::npos);

        size_t end = openEnd + 1;
        if (openEnd == 0 || xml[openEnd - 1] != '/') {
            const std::string sheetDataClose = "</sheetData>";
            const auto closeBegin = xml.find(sheetDataClose, openEnd);
            REQUIRE(closeBegin != std::string::npos);
            end = closeBegin + sheetDataClose.size();
        }

        xml.replace(begin, end - begin, sheetData);
        return xml;
    }

    void createWorkbookWithSheetData(const std::string& fileName, const std::string& sheetData)
    {
        XLDocument doc;
        doc.create(fileName, XLForceOverwrite);
        doc.workbook().worksheet("Sheet1").cell("A1").value() = 1;
        doc.save();
        doc.close();

        XLZipArchive archive;
        archive.open(fileName);
        const std::string worksheetXml = archive.getEntry("xl/worksheets/sheet1.xml");
        archive.addEntry("xl/worksheets/sheet1.xml", replaceSheetData(worksheetXml, sheetData));
        archive.save();
        archive.close();
    }
}

TEST_CASE("XLFormula Tests", "[XLFormula]")
{
    SECTION("Default Constructor")
    {
        XLFormula formula;

        REQUIRE(formula.get().empty());
    }

    SECTION("String Constructor")
    {
        XLFormula formula1("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        std::string s = "BLAH2";
        XLFormula formula2(s);
        REQUIRE(formula2.get() == "BLAH2");
    }

    SECTION("Copy Constructor")
    {
        XLFormula formula1("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        XLFormula formula2 = formula1;
        REQUIRE(formula2.get() == "BLAH1");

    }

    SECTION("Move Constructor")
    {
        XLFormula formula1("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        XLFormula formula2 = std::move(formula1);
        REQUIRE(formula2.get() == "BLAH1");

    }

    SECTION("Copy assignment")
    {
        XLFormula formula1("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        XLFormula formula2;
        formula2 = formula1;
        REQUIRE(formula2.get() == "BLAH1");
    }

    SECTION("Move assignment")
    {
        XLFormula formula1("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        XLFormula formula2;
        formula2 = std::move(formula1);
        REQUIRE(formula2.get() == "BLAH1");

    }

    SECTION("Clear")
    {
        XLFormula formula1("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        formula1.clear();
        REQUIRE(formula1.get().empty());
    }

    SECTION("String assignment")
    {
        XLFormula formula1;
        formula1 = "BLAH1";
        REQUIRE(formula1.get() == "BLAH1");

        XLFormula formula2;
        formula2 = std::string("BLAH2");
        REQUIRE(formula2.get() == "BLAH2");
    }

    SECTION("String setter")
    {
        XLFormula formula1;
        formula1.set("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        XLFormula formula2;
        formula2.set(std::string("BLAH2"));
        REQUIRE(formula2.get() == "BLAH2");
    }

    SECTION("Implicit conversion")
    {
        XLFormula formula1;
        formula1.set("BLAH1");
        REQUIRE(formula1.get() == "BLAH1");

        auto result = std::string(formula1);
        REQUIRE(result == "BLAH1");
    }

    SECTION("FormulaProxy")
    {
        XLDocument doc;
        doc.create("./testXLFormula.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        wks.cell("A1").formula() = "=1+1";
        wks.cell("B2").formula() = wks.cell("A1").formula();
        REQUIRE(wks.cell("B2").formula() == XLFormula("=1+1"));

        XLFormula form = wks.cell("B2").formula();
        REQUIRE(form == XLFormula("=1+1"));

        REQUIRE(wks.cell("A1").hasFormula());
        wks.cell("A1").formula().clear();
        REQUIRE_FALSE(wks.cell("A1").hasFormula());
        REQUIRE(wks.cell("B2").formula() == XLFormula("=1+1"));

    }

    SECTION("Shared Formula Expansion")
    {
        const std::string fileName = "./testXLFormulaShared.xlsx";
        createWorkbookWithSheetData(fileName, R"xml(<sheetData>
<row r="1">
<c r="A1"><v>1</v></c>
<c r="B1"><v>2</v></c>
<c r="C1"><v>3</v></c>
</row>
<row r="2">
<c r="B2"><f t="shared" ref="B2:D4" si="0">A1+$A1+A$1+$A$1+SUM(A1:B2)+'Other Sheet'!A1+'O''Brien'!A1+LOG10(A1)+"A1"+Table1[Column1]+XFE1+A1048577</f><v>0</v></c>
<c r="C2"><f t="shared" si="0"/><v>0</v></c>
</row>
<row r="3">
<c r="B3"><f t="shared" si="0"/><v>0</v></c>
<c r="C3"><f t="shared" si="0"/><v>0</v></c>
</row>
<row r="4">
<c r="D4"><f t="shared" si="0"/><v>0</v></c>
</row>
</sheetData>)xml");

        XLDocument doc;
        doc.open(fileName);
        auto wks = doc.workbook().worksheet("Sheet1");

        REQUIRE(wks.cell("B2").formula().get()
                == "A1+$A1+A$1+$A$1+SUM(A1:B2)+'Other Sheet'!A1+'O''Brien'!A1+LOG10(A1)+\"A1\"+Table1[Column1]+XFE1+A1048577");
        REQUIRE(wks.cell("C2").formula().get()
                == "B1+$A1+B$1+$A$1+SUM(B1:C2)+'Other Sheet'!B1+'O''Brien'!B1+LOG10(B1)+\"A1\"+Table1[Column1]+XFE1+A1048577");
        REQUIRE(wks.cell("B3").formula().get()
                == "A2+$A2+A$1+$A$1+SUM(A2:B3)+'Other Sheet'!A2+'O''Brien'!A2+LOG10(A2)+\"A1\"+Table1[Column1]+XFE1+A1048577");
        REQUIRE(wks.cell("C3").formula().get()
                == "B2+$A2+B$1+$A$1+SUM(B2:C3)+'Other Sheet'!B2+'O''Brien'!B2+LOG10(B2)+\"A1\"+Table1[Column1]+XFE1+A1048577");
        REQUIRE(wks.cell("D4").formula().get()
                == "C3+$A3+C$1+$A$1+SUM(C3:D4)+'Other Sheet'!C3+'O''Brien'!C3+LOG10(C3)+\"A1\"+Table1[Column1]+XFE1+A1048577");

        doc.close();
    }
}
