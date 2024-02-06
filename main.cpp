
#include <iostream>
#include <vector>
#include <string>
#include <algorithm>
#include <OpenXLSX.hpp>

std::vector<std::string> check_target;

const std::string input_file_path = "input.xlsx";
const std::string output_file_path = "output.xlsx";

// È¥³ý×Ö·û´®Ç°ÃæµÄ¿Õ°××Ö·û
std::string ltrim(const std::string& str) {
    auto it = std::find_if(str.begin(), str.end(), [](unsigned char ch) {
        return !std::isspace(ch);
    });
    return std::string(it, str.end());
}

// È¥³ý×Ö·û´®ºóÃæµÄ¿Õ°××Ö·û
std::string rtrim(const std::string& str) {
    auto it = std::find_if(str.rbegin(), str.rend(), [](unsigned char ch) {
        return !std::isspace(ch);
    });
    return std::string(str.begin(), it.base());
}

// È¥³ý×Ö·û´®Ç°ºóµÄ¿Õ°××Ö·û
std::string trim(const std::string& str) {
    return rtrim(ltrim(str));
}

void split_string(const std::string s) {
    int start = 0;
    int end = s.find_first_of(";");
    while(end != std::string::npos) {
        std::string sub_string = trim(s.substr(start, end - start));

        if(sub_string.length() == 0) {
            continue;
        }

        check_target.push_back(sub_string);

        start = end + 1;
        end = s.find_first_of(";", start);
    }
    std::string s1 = trim(s.substr(start));
    if(s1.length() != 0) {
        check_target.push_back(s1);
    }

}

void init_check_target(const std::string file_path = input_file_path) {
    OpenXLSX::XLDocument doc;
    doc.open(file_path);
    OpenXLSX::XLWorksheet sheet = doc.workbook().sheet(1);

    for(int i = 2; i <= sheet.rowCount(); ++i) {
        OpenXLSX::XLCellValue value  = sheet.cell(i, 1).value();
        std::string s = value.get<std::string>();
        split_string(s);
    }

    doc.close();
}


void query(const std::string file_path = output_file_path) {
    OpenXLSX::XLDocument doc;
    doc.open(file_path);
    OpenXLSX::XLWorksheet sheet = doc.workbook().sheet(1);
    int target_column = sheet.columnCount() + 1;


    for(int i = 1; i <= sheet.rowCount(); ++i) {
        OpenXLSX::XLCellValue value = sheet.cell(i, 1).value();
        std::string query_name = value.get<std::string>();

        auto exist = std::find(check_target.cbegin(), check_target.cend(), query_name);
        if(exist != check_target.cend()) {
            // find it
            std::cout << "find it in the row " << i << std::endl;
            sheet.cell(i, target_column).value() = 1;
        }
    }

    doc.save();
    doc.close();
}

int main(int argc, char** argv) {

    std::cout << "the usage: test.exe [excel name] [excel name]" << std::endl;
    std::cout << "the first excel is used to query if the target exist" << std::endl;
    std::cout << "the second excel is the target file" << std::endl;
    std::cout << "if the target exist, it will add a column at the cell end" << std::endl;

    if(argc != 3) {
        std::cout << "use default excel file name: " << input_file_path << ", " << output_file_path << std::endl;
        init_check_target();
        query();
    } else {
        std::cout << "the first excel file name is " << argv[1] << std::endl;
        std::cout << "the second excel file name is " << argv[2] << std::endl;
        init_check_target(argv[1]);
        init_check_target(argv[2]);
    }

    return 0;
}
