#include "ExcelProcessorCore.h"
#include <fstream>
#include <sstream>
#include <algorithm>
#include <iostream>
#include <iomanip>
// #include <filesystem>  // Disabled for MinGW compatibility
#include <future>
#include <thread>
#include <set>
#include <ctime>
#include <QAxObject>
#include <QVariant>
#include <QDebug>
#include <QFileInfo>
#include <QDir>
#include <QFile>
#include <QDateTime> // Added for QDateTime
#include <QRegExp> // Added for wildcard matching
#include <cctype> // For std::tolower

// namespace fs = std::filesystem;  // Disabled for MinGW compatibility

// ActiveQt Excel Reader for .xlsx support
class ActiveQtExcelReader : public ExcelReader {
private:
    std::function<void(const std::string&)> logger_;

public:
    void setLogger(std::function<void(const std::string&)> logger) override {
        logger_ = logger;
    }

    bool readExcelFile(const std::string& filename, std::vector<DataRow>& data, const std::string& sheetName = "", int maxRows = 0, int offset = 0, bool includeHeader = false) override {
        qDebug() << "DEBUG: Reading Excel File:" << QString::fromStdString(filename) 
                 << "Offset:" << offset 
                 << "IncludeHeader:" << includeHeader;

        QAxObject excel("Excel.Application");
        if (excel.isNull()) {
            qDebug() << "ERROR: Failed to create Excel.Application";
            if (logger_) logger_("ERROR: Failed to create Excel.Application (ActiveQt)");
            return false;
        }
        excel.setProperty("Visible", false);
        excel.setProperty("DisplayAlerts", false);

        QAxObject* workbooks = excel.querySubObject("Workbooks");
        if (!workbooks) {
            excel.dynamicCall("Quit()");
            return false;
        }

        // Convert to absolute path with backslashes
        QString absPath = QFileInfo(QString::fromStdString(filename)).absoluteFilePath();
        absPath.replace("/", "\\");

        QAxObject* workbook = workbooks->querySubObject("Open(const QString&)", absPath);
        if (!workbook) {
            qDebug() << "ERROR: Failed to open workbook:" << absPath;
            if (logger_) logger_("ERROR: Failed to open workbook: " + absPath.toStdString());
            excel.dynamicCall("Quit()");
            return false;
        }

        QAxObject* sheets = workbook->querySubObject("Worksheets");
        if (!sheets) {
            workbook->dynamicCall("Close()");
            excel.dynamicCall("Quit()");
            return false;
        }

        QAxObject* sheet = nullptr;
        if (sheetName.empty()) {
            sheet = sheets->querySubObject("Item(int)", 1);
        } else {
            sheet = sheets->querySubObject("Item(const QString&)", QString::fromStdString(sheetName));
        }

        if (!sheet) {
            qDebug() << "ERROR: Sheet not found:" << QString::fromStdString(sheetName);
            if (logger_) logger_("ERROR: Sheet not found: " + sheetName);
            workbook->dynamicCall("Close()");
            excel.dynamicCall("Quit()");
            return false;
        }

        // Debug: Check A1 directly to verify file content vs read content
        QAxObject* a1 = sheet->querySubObject("Range(const QString&)", "A1");
        if (a1) {
             QString a1Val = a1->dynamicCall("Value").toString();
             qDebug() << "DEBUG: Direct A1 Cell Check:" << a1Val;
             if (logger_) {
                 logger_("DEBUG: Direct A1 Cell Value: [" + a1Val.toStdString() + "]");
             }
        }

        QAxObject* usedRange = sheet->querySubObject("UsedRange");
        if (!usedRange) {
            workbook->dynamicCall("Close()");
            excel.dynamicCall("Quit()");
            return false;
        }

        // Fix for "Header Reading Error" where UsedRange starts at a later row (e.g., 1093)
        // We must read based on absolute row numbers (starting from 1).
        
        int usedRowStart = usedRange->property("Row").toInt();
        int usedRowsCount = usedRange->querySubObject("Rows")->property("Count").toInt();
        int lastRow = usedRowStart + usedRowsCount - 1;
        
        int usedColStart = usedRange->property("Column").toInt();
        int usedColsCount = usedRange->querySubObject("Columns")->property("Count").toInt();
        int lastCol = usedColStart + usedColsCount - 1;

        qDebug() << "DEBUG: UsedRange: Row" << usedRowStart << "Count" << usedRowsCount << "(Last:" << lastRow << ")"
                 << "Col" << usedColStart << "Count" << usedColsCount;

        // Determine absolute start row (1-based)
        int absStartRow;
        if (offset == 0) {
            absStartRow = 1;
        } else {
            // If offset > 0, we have already processed some data rows.
            // If includeHeader is true, it means there IS a header row (Row 1) that is NOT data.
            // So absolute row = 1 (Header) + offset (Data Rows) + 1 (Next Data Row) = offset + 2.
            // If includeHeader is false, it means Row 1 is Data.
            // So absolute row = offset (Data Rows) + 1 (Next Data Row) = offset + 1.
            
            if (includeHeader) {
                absStartRow = offset + 2;
            } else {
                absStartRow = offset + 1;
            }
        }
        
        if (logger_) {
             std::string dbg = "[DEBUG] ReadExcel: Offset=" + std::to_string(offset) 
                             + " | IncludeHeader=" + (includeHeader?"YES":"NO")
                             + " | StartRow=" + std::to_string(absStartRow)
                             + " | LastRow=" + std::to_string(lastRow);
             logger_(dbg); 
        }

        if (absStartRow > lastRow) {
            workbook->dynamicCall("Close()");
            excel.dynamicCall("Quit()");
            return true; // Nothing to read
        }

        int rowsToRead = lastRow - absStartRow + 1;
        if (maxRows > 0) {
            rowsToRead = std::min(rowsToRead, maxRows);
        }

        // Use Cells to get absolute range
        // We read from Column 1 (A) to the last used column to ensure we cover all potential data
        // even if UsedRange starts at Column C.
        QAxObject* cells = sheet->querySubObject("Cells");
        QAxObject* startCell = cells->querySubObject("Item(int, int)", absStartRow, 1);
        
        // Ensure we don't read 0 columns
        int colsToRead = std::max(lastCol, 1);
        
        QAxObject* range = startCell->querySubObject("Resize(int, int)", rowsToRead, colsToRead);

        if (!range) {
            workbook->dynamicCall("Close()");
            excel.dynamicCall("Quit()");
            return false;
        }

        QVariant var = range->dynamicCall("Value");
        
        QList<QVariant> rawRows = var.toList();
        
        // Don't clear data if appending (offset > 0), but usually caller handles clearing.
        // Here we just append to the vector provided.
        // data.clear(); // REMOVED: Let caller manage data vector

        int rowIdx = absStartRow - 1; // Start row index (will be incremented to absStartRow in loop)
        
        // Skip header logic
        // We have already handled skipping via absStartRow calculation.
        bool skipHeader = false; 
        
        if (skipHeader) {
             qDebug() << "DEBUG: Will skip first row as header (includeHeader=false)";
        }

        for (const auto& rowVar : rawRows) {
            rowIdx++; // Increment absolute row index first

            QList<QVariant> colVars = rowVar.toList();
            
            // Check if row is completely empty (all cells are null/empty string)
            bool isAllEmpty = true;
            if (!colVars.isEmpty()) {
                for(const auto& cell : colVars) {
                    if(!cell.toString().isEmpty()) {
                        isAllEmpty = false;
                        break;
                    }
                }
            }
            
            // Special handling for Header Row (rowIdx == 1)
            // If includeHeader is true, we MUST include it even if it's empty, 
            // to ensure data[0] corresponds to the header row.
            bool isHeaderRow = (rowIdx == 1);

            if (colVars.isEmpty()) {
                 if (isHeaderRow && includeHeader) {
                     // Force include empty header
                     qDebug() << "DEBUG: Found empty row at absolute index 1, but keeping it as Header.";
                 } else {
                     // Skip empty row
                     continue;
                 }
            }
            
            if (skipHeader) {
                // If we are here, it means this is the first row (rowIdx==1) and we want to skip it.
                // But wait, rowIdx is already incremented.
                // If offset=0, rowIdx starts at 0. First loop rowIdx=1.
                // skipHeader is true.
                
                skipHeader = false;
                
                // Log what we skipped
                QString skippedContent;
                for(const auto& cell : colVars) skippedContent += cell.toString() + "|";
                qDebug() << "DEBUG: Skipped Header Row:" << skippedContent;
                
                continue;
            }

            if (isAllEmpty && !(isHeaderRow && includeHeader)) {
                // Skip empty data rows (but not empty header if requested)
                // qDebug() << "DEBUG: Found empty row at absolute index" << rowIdx;
                continue;
            }

            DataRow row;
            row.rowNumber = rowIdx;
            row.sheetName = sheet->property("Name").toString().toStdString();
            
            for (const auto& cellVar : colVars) {
                // Convert QVariant to our std::variant
                if (cellVar.userType() == QMetaType::Double) {
                    row.data.push_back(cellVar.toDouble());
                } else if (cellVar.userType() == QMetaType::Int || cellVar.userType() == QMetaType::LongLong) {
                    row.data.push_back(cellVar.toInt());
                } else if (cellVar.userType() == QMetaType::Bool) {
                    row.data.push_back(cellVar.toBool());
                } else if (cellVar.userType() == QMetaType::QDate || cellVar.userType() == QMetaType::QDateTime) {
                     // Convert to std::tm
                     QDateTime dt = cellVar.toDateTime();
                     std::tm tm = {};
                     tm.tm_year = dt.date().year() - 1900;
                     tm.tm_mon = dt.date().month() - 1;
                     tm.tm_mday = dt.date().day();
                     row.data.push_back(tm);
                } else {
                    row.data.push_back(cellVar.toString().toStdString());
                }
            }
            row.isValid = true;
            data.push_back(row);
        }

        workbook->dynamicCall("Close()");
        excel.dynamicCall("Quit()");
        return true;
    }

    bool getSheetNames(const std::string& filename, std::vector<std::string>& sheetNames) const override {
        QAxObject excel("Excel.Application");
        if (excel.isNull()) return false;
        excel.setProperty("Visible", false);
        excel.setProperty("DisplayAlerts", false);

        QString absPath = QFileInfo(QString::fromStdString(filename)).absoluteFilePath();
        absPath.replace("/", "\\");

        QAxObject* workbooks = excel.querySubObject("Workbooks");
        if (!workbooks) {
            excel.dynamicCall("Quit()");
            return false;
        }

        QAxObject* workbook = workbooks->querySubObject("Open(const QString&)", absPath);
        if (!workbook) {
            excel.dynamicCall("Quit()");
            return false;
        }

        QAxObject* sheets = workbook->querySubObject("Worksheets");
        int count = sheets->property("Count").toInt();
        sheetNames.clear();
        
        for (int i = 1; i <= count; ++i) {
            QAxObject* sheet = sheets->querySubObject("Item(int)", i);
            if (sheet) {
                sheetNames.push_back(sheet->property("Name").toString().toStdString());
            }
        }

        workbook->dynamicCall("Close()");
        excel.dynamicCall("Quit()");
        return true;
    }

    int getRowCount(const std::string& sheetName) const override { return 0; }
    int getColumnCount(const std::string& sheetName) const override { return 0; }
};

// ActiveQt Excel Writer for .xlsx support
class ActiveQtExcelWriter : public ExcelWriter {
private:
    QAxObject* excelApp_ = nullptr;
    std::map<std::string, QAxObject*> openWorkbooks_; // Key: Absolute Path

    void initApp() {
        if (!excelApp_) {
            excelApp_ = new QAxObject("Excel.Application");
            if (!excelApp_->isNull()) {
                excelApp_->setProperty("Visible", false);
                excelApp_->setProperty("DisplayAlerts", false);
            } else {
                delete excelApp_;
                excelApp_ = nullptr;
            }
        }
    }

public:
    ActiveQtExcelWriter() = default;

    ~ActiveQtExcelWriter() {
        closeAll();
    }

    ExcelWriterType getType() const override { return ExcelWriterType::ActiveQt; }

    void closeAll() override {
        for (auto& pair : openWorkbooks_) {
            if (pair.second) {
                pair.second->dynamicCall("Close()");
                delete pair.second;
            }
        }
        openWorkbooks_.clear();

        if (excelApp_) {
            excelApp_->dynamicCall("Quit()");
            delete excelApp_;
            excelApp_ = nullptr;
        }
    }

    bool writeExcelFile(const std::string& filename,
                       const std::vector<DataRow>& data,
                       const std::string& sheetName = "Sheet1") override {
        initApp();
        if (!excelApp_) return false;

        QString absPath = QFileInfo(QString::fromStdString(filename)).absoluteFilePath();
        absPath.replace("/", "\\");
        std::string key = absPath.toStdString();

        // If already open, close it (we are overwriting)
        if (openWorkbooks_.count(key)) {
            openWorkbooks_[key]->dynamicCall("Close()");
            delete openWorkbooks_[key];
            openWorkbooks_.erase(key);
        }

        QAxObject* workbooks = excelApp_->querySubObject("Workbooks");
        if (!workbooks) return false;

        QAxObject* workbook = workbooks->querySubObject("Add");
        if (!workbook) return false;

        QAxObject* sheets = workbook->querySubObject("Worksheets");
        QAxObject* sheet = sheets->querySubObject("Item(int)", 1);
        if (sheet) {
            sheet->setProperty("Name", QString::fromStdString(sheetName));
            writeDataToSheet(sheet, data);
        }

        workbook->dynamicCall("SaveAs(const QString&)", absPath);
        openWorkbooks_[key] = workbook; // Keep open
        return true;
    }

    bool writeMultipleSheets(const std::string& filename,
                              const std::map<std::string, std::vector<DataRow>>& sheetData) override {
        initApp();
        if (!excelApp_) return false;

        QString absPath = QFileInfo(QString::fromStdString(filename)).absoluteFilePath();
        absPath.replace("/", "\\");
        std::string key = absPath.toStdString();

        // If already open, close it
        if (openWorkbooks_.count(key)) {
            openWorkbooks_[key]->dynamicCall("Close()");
            delete openWorkbooks_[key];
            openWorkbooks_.erase(key);
        }

        QAxObject* workbooks = excelApp_->querySubObject("Workbooks");
        QAxObject* workbook = workbooks->querySubObject("Add");
        if (!workbook) return false;

        QAxObject* sheets = workbook->querySubObject("Worksheets");
        
        // Remove default sheets except one
        int count = sheets->dynamicCall("Count").toInt();
        while (count > 1) {
            QAxObject* s = sheets->querySubObject("Item(int)", count);
            s->dynamicCall("Delete()");
            count--;
        }

        bool first = true;
        for (const auto& [name, data] : sheetData) {
            QAxObject* sheet = nullptr;
            if (first) {
                sheet = sheets->querySubObject("Item(int)", 1);
                first = false;
            } else {
                sheet = sheets->querySubObject("Add()");
            }
            
            if (sheet) {
                sheet->setProperty("Name", QString::fromStdString(name));
                writeDataToSheet(sheet, data);
            }
        }
        
        workbook->dynamicCall("SaveAs(const QString&)", absPath);
        openWorkbooks_[key] = workbook; // Keep open
        return true;
    }

    bool appendToSheet(const std::string& filename,
                        const std::vector<DataRow>& data,
                        const std::string& sheetName,
                        bool overwrite = false) override {
        initApp();
        if (!excelApp_) return false;

        QString absPath = QFileInfo(QString::fromStdString(filename)).absoluteFilePath();
        absPath.replace("/", "\\");
        std::string key = absPath.toStdString();
        
        QAxObject* workbook = nullptr;
        QAxObject* workbooks = excelApp_->querySubObject("Workbooks");

        if (openWorkbooks_.count(key)) {
            workbook = openWorkbooks_[key];
        } else {
            if (QFile::exists(absPath)) {
                workbook = workbooks->querySubObject("Open(const QString&)", absPath);
            } else {
                workbook = workbooks->querySubObject("Add");
                workbook->dynamicCall("SaveAs(const QString&)", absPath);
            }
            if (workbook) openWorkbooks_[key] = workbook;
        }

        if (!workbook) return false;

        QAxObject* sheets = workbook->querySubObject("Worksheets");
        QAxObject* sheet = nullptr;
        
        int count = sheets->dynamicCall("Count").toInt();
        for(int i=1; i<=count; ++i) {
            QAxObject* s = sheets->querySubObject("Item(int)", i);
            if (s->property("Name").toString() == QString::fromStdString(sheetName)) {
                sheet = s;
                break;
            }
        }
        
        if (!sheet) {
            sheet = sheets->querySubObject("Add()");
            sheet->setProperty("Name", QString::fromStdString(sheetName));
        } else if (overwrite) {
            sheet->querySubObject("Cells")->dynamicCall("Clear()");
        }
        
        int startRow = 1;
        if (!overwrite) {
            QAxObject* usedRange = sheet->querySubObject("UsedRange");
            if (usedRange) {
                 QAxObject* a1 = sheet->querySubObject("Range(const QString&)", "A1");
                 if (a1 && !a1->dynamicCall("Value").toString().isEmpty()) {
                     startRow = usedRange->querySubObject("Rows")->property("Count").toInt() + 1;
                 }
            }
        }
        
        writeDataToSheet(sheet, data, startRow);

        workbook->dynamicCall("Save()");
        return true;
    }

    bool isSheetEmpty(const std::string& filename, const std::string& sheetName) {
        initApp();
        if (!excelApp_) return false;

        QString absPath = QFileInfo(QString::fromStdString(filename)).absoluteFilePath();
        if (!QFile::exists(absPath)) return true;

        absPath.replace("/", "\\");
        std::string key = absPath.toStdString();
        
        QAxObject* workbook = nullptr;
        
        if (openWorkbooks_.count(key)) {
            workbook = openWorkbooks_[key];
        } else {
             QAxObject* workbooks = excelApp_->querySubObject("Workbooks");
             if (!workbooks) return false;
             workbook = workbooks->querySubObject("Open(const QString&)", absPath);
             if (workbook) openWorkbooks_[key] = workbook;
        }
        
        if (!workbook) return true;
        
        QAxObject* sheets = workbook->querySubObject("Worksheets");
        if (!sheets) return false;
        
        QAxObject* sheet = nullptr;
        int count = sheets->dynamicCall("Count").toInt();
        for(int i=1; i<=count; ++i) {
            QAxObject* s = sheets->querySubObject("Item(int)", i);
            if (s->property("Name").toString() == QString::fromStdString(sheetName)) {
                sheet = s;
                break;
            }
        }
        
        if (!sheet) return true;
        
        QAxObject* usedRange = sheet->querySubObject("UsedRange");
        if (!usedRange) return true;
        
        QAxObject* a1 = sheet->querySubObject("Range(const QString&)", "A1");
        if (a1 && a1->dynamicCall("Value").toString().isEmpty()) return true;
        
        return false;
    }

private:
    void writeDataToSheet(QAxObject* sheet, const std::vector<DataRow>& data, int startRow = 1) {
        if (data.empty()) return;
        
        QList<QVariant> rows;
        int cols = 0;
        for (const auto& row : data) {
            QList<QVariant> colVars;
            for (const auto& cell : row.data) {
                        std::visit([&colVars](const auto& val) {
                            using T = std::decay_t<decltype(val)>;
                            if constexpr (std::is_same_v<T, std::string>) {
                                colVars << QVariant::fromValue(QString::fromStdString(val));
                            } else if constexpr (std::is_same_v<T, tm>) {
                                QDate date(val.tm_year + 1900, val.tm_mon + 1, val.tm_mday);
                                QTime time(val.tm_hour, val.tm_min, val.tm_sec);
                                colVars << QVariant::fromValue(QDateTime(date, time));
                            } else {
                                colVars << QVariant::fromValue(val);
                            }
                        }, cell);
                    }
            if (colVars.size() > cols) cols = colVars.size();
            rows << QVariant(colVars);
        }
        
        QString startCell = QString("A%1").arg(startRow);
        QAxObject* range = sheet->querySubObject("Range(const QString&)", startCell);
        if (range) {
             range = range->querySubObject("Resize(int, int)", (int)data.size(), cols);
             range->setProperty("Value", QVariant(rows));
        }
    }
};

// CSV Excel Reader
class CSVExcelReader : public ExcelReader {
public:
    bool readExcelFile(const std::string& filename, std::vector<DataRow>& data, const std::string& sheetName = "", int maxRows = 0, int offset = 0, bool includeHeader = false) override {
        std::ifstream file(filename);
        if (!file.is_open()) {
            return false;
        }
        
        std::string line;
        
        // Skip offset lines + 1 (Header)
        // Note: This logic assumes offset 0 means "start after header"
        // and offset N means "start after N rows of data"
        // If includeHeader is true (and offset is 0), we don't skip anything.
        int skipLines = offset + 1;
        if (offset == 0 && includeHeader) {
            skipLines = 0;
        }

        for(int i=0; i < skipLines; ++i) {
             if (!std::getline(file, line)) return true;
        }
        
        int count = 0;
        while(maxRows == 0 || count < maxRows) {
            if (!std::getline(file, line)) break;
            
            DataRow row;
            std::istringstream lineStream(line);
            std::string cell;
            bool isValid = true;

            while (std::getline(lineStream, cell, ',')) {
                try {
                    auto value = parseCellValue(cell);
                    row.data.push_back(value);
                } catch (...) {
                    row.data.push_back(std::string(""));
                    isValid = false;
                }
            }

            row.rowNumber = offset + count + 1; // 1-based
            row.isValid = isValid;
            row.sheetName = "Sheet1";
            data.push_back(std::move(row));
            
            count++;
        }

        return true;
    }

    bool getSheetNames(const std::string& filename, std::vector<std::string>& sheetNames) const override {
        sheetNames.clear();
        sheetNames.push_back("Sheet1"); // CSV has only one sheet
        return true;
    }

    int getRowCount(const std::string& sheetName) const override {
        std::ifstream file(sheetName);
        if (!file.is_open()) return 0;

        int rowCount = -1; // Subtract header
        std::string line;
        while (std::getline(file, line)) {
            rowCount++;
        }
        return rowCount > 0 ? rowCount : 0;
    }

    int getColumnCount(const std::string& sheetName) const override {
        std::ifstream file(sheetName);
        if (!file.is_open()) return 0;

        std::string line;
        if (!std::getline(file, line)) return 0;

        std::istringstream lineStream(line);
        std::string cell;
        int columnCount = 0;
        while (std::getline(lineStream, cell, ',')) {
            columnCount++;
        }
        return columnCount;
    }

private:
    std::variant<std::string, int, double, bool, std::tm> parseCellValue(const std::string& cell) {
        // Remove quotes
        std::string trimmedCell = cell;
        if (trimmedCell.length() >= 2 && trimmedCell.front() == '"' && trimmedCell.back() == '"') {
            trimmedCell = trimmedCell.substr(1, trimmedCell.length() - 2);
        }

        // Handle empty
        if (trimmedCell.empty()) {
            return std::string("");
        }

        // Handle boolean
        std::string lowerCell = trimmedCell;
        std::transform(lowerCell.begin(), lowerCell.end(), lowerCell.begin(), 
            [](unsigned char c){ return std::tolower(c); });
            
        if (lowerCell == "true") return true;
        if (lowerCell == "false") return false;

        // Handle integer
        try {
            if (trimmedCell.find('.') == std::string::npos &&
                trimmedCell.find_first_not_of("0123456789-+") == std::string::npos) {
                return std::stoi(trimmedCell);
            }
        } catch (...) {}

        // Handle double
        try {
            if (trimmedCell.find_first_of("0123456789.") != std::string::npos) {
                return std::stod(trimmedCell);
            }
        } catch (...) {}

        // Handle date
        try {
            std::tm tm = {};
            std::istringstream ss(trimmedCell);
            ss >> std::get_time(&tm, "%Y-%m-%d");
            if (!ss.fail()) return tm;

            ss.str(trimmedCell);
            ss.clear();
            ss >> std::get_time(&tm, "%Y/%m/%d");
            if (!ss.fail()) return tm;
        } catch (...) {}

        // Default to string
        return trimmedCell;
    }
};

// CSV Excel Writer
class CSVExcelWriter : public ExcelWriter {
public:
    ExcelWriterType getType() const override { return ExcelWriterType::CSV; }

    bool writeExcelFile(const std::string& filename,
                        const std::vector<DataRow>& data,
                        const std::string& sheetName = "Sheet1") override {
        std::ofstream file(filename);
        if (!file.is_open()) {
            return false;
        }

        // Write data
        for (const auto& row : data) {
            std::string line;
            for (size_t i = 0; i < row.data.size(); ++i) {
                if (i > 0) line += ",";
                line += formatCellValue(row.data[i]);
            }
            file << line << std::endl;
        }

        return true;
    }

    bool writeMultipleSheets(const std::string& filename,
                               const std::map<std::string, std::vector<DataRow>>& sheetData) override {
        // CSV only supports single sheet, write first one
        if (!sheetData.empty()) {
            return writeExcelFile(filename + "_" + sheetData.begin()->first + ".csv",
                              sheetData.begin()->second,
                              sheetData.begin()->first);
        }
        return true;
    }

    bool appendToSheet(const std::string& filename,
                        const std::vector<DataRow>& data,
                        const std::string& sheetName,
                        bool overwrite = false) override {
        std::ios::openmode mode = std::ios::app;
        if (overwrite) mode = std::ios::trunc;
        std::ofstream file(filename, mode);
        if (!file.is_open()) {
            return false;
        }

        // Write data
        for (const auto& row : data) {
            std::string line;
            for (size_t i = 0; i < row.data.size(); ++i) {
                if (i > 0) line += ",";
                line += formatCellValue(row.data[i]);
            }
            file << line << std::endl;
        }

        return true;
    }

    bool isSheetEmpty(const std::string& filename, const std::string& sheetName) {
        std::ifstream file(filename);
        return file.peek() == std::ifstream::traits_type::eof();
    }

private:
    std::string formatCellValue(const std::variant<std::string, int, double, bool, std::tm>& value) {
        std::ostringstream oss;
        std::visit([&oss](const auto& val) {
            using ValType = std::decay_t<decltype(val)>;
            if constexpr (std::is_same_v<ValType, std::string>) {
                oss << "\"" << val << "\"";
            } else if constexpr (std::is_same_v<ValType, bool>) {
                oss << (val ? "true" : "false");
            } else if constexpr (std::is_same_v<ValType, std::tm>) {
                oss << std::put_time(&val, "%Y-%m-%d");
            } else {
                oss << val;
            }
        }, value);
        return oss.str();
    }
};

// High Performance Data Processor
class HighPerformanceDataProcessor : public DataProcessor {
public:
    void setProgressCallback(std::function<void(int, const std::string&)> callback) override {
        progressCallback_ = callback;
    }

    ProcessingResult processData(std::vector<DataRow>& data,
                                const std::vector<Rule>& rules,
                                const std::map<int, std::vector<int>>& combinations) override {
        ProcessingResult result;
        auto startTime = std::chrono::high_resolution_clock::now();

        result.totalRows = data.size();

        // If no combinations, check if we should apply all enabled rules as default filter
        // Or if we should just return original data.
        // For backward compatibility/simplicity, if combinations are empty, we might assume NO filtering (pass-through).
        // However, if the user expects filtering, they should probably have a combination or task.
        // Given the "Empty Output" issue, let's assume we proceed with whatever combinations provided.
        // If empty, nothing happens, data remains.

        // Parallel processing optimization
        const size_t chunkSize = 10000; 
        const size_t numThreads = std::thread::hardware_concurrency();

        std::vector<std::future<std::vector<size_t>>> futures;

        // We need to determine if we are in "Keep" mode (Filter) or "Delete" mode.
        // This can be complex with multiple combinations.
        // Simplified Logic: 
        // 1. Calculate set of rows that match ANY combination (Union of matches).
        // 2. If the matched rules are FILTER type, we KEEP these rows.
        // 3. If the matched rules are DELETE type, we REMOVE these rows.
        
        // For now, let's assume the primary use case is FILTER (Keep matches).
        // If any rule in a combination is DELETE, we treat that combination as a "Delete Signal".

        // Map comboId to type (true=Filter/Keep, false=Delete)
        std::map<int, bool> comboTypes;
        for (const auto& [comboId, ruleIds] : combinations) {
             bool isDelete = false;
             for(int rid : ruleIds) {
                 if (rid > 0 && rid <= (int)rules.size()) {
                     if (rules[rid-1].type == RuleType::DELETE_ROW) {
                         isDelete = true;
                         break;
                     }
                 }
             }
             comboTypes[comboId] = !isDelete; // true = Keep, false = Delete
        }

        for (const auto& [comboId, ruleIds] : combinations) {
            futures.push_back(std::async(std::launch::async, [this, &data, &rules, ruleIds, chunkSize]() -> std::vector<size_t> {
                std::vector<size_t> matchedIndices;

                // Process in chunks
                for (size_t start = 0; start < data.size(); start += chunkSize) {
                    size_t end = std::min(start + chunkSize, data.size());

                    // Process current chunk
                    for (size_t i = start; i < end; ++i) {
                        if (evaluateRuleCombination(data[i], rules, ruleIds)) {
                            matchedIndices.push_back(i);
                        }
                    }
                }

                return matchedIndices;
            }));
        }

        // Collect results
        std::set<size_t> rowsToKeep;
        std::set<size_t> rowsToDelete;
        
        bool hasFilterRules = false;
        bool hasDeleteRules = false;
        
        // Check what kind of rules we have active
        for(const auto& [id, isKeep] : comboTypes) {
            if(isKeep) hasFilterRules = true;
            else hasDeleteRules = true;
        }

        int futureIdx = 0;
        int totalCombos = combinations.size();
        for (const auto& [comboId, ruleIds] : combinations) {
            auto indices = futures[futureIdx++].get();
            bool isKeep = comboTypes[comboId];
            
            for (size_t index : indices) {
                if (isKeep) rowsToKeep.insert(index);
                else rowsToDelete.insert(index);
            }

            if (progressCallback_) {
                int p = (futureIdx * 100) / (totalCombos > 0 ? totalCombos : 1);
                progressCallback_(p, "\xE5\xB7\xB2\xE5\xA4\x84\xE7\x90\x86\xE8\xA7\x84\xE5\x88\x99\xE7\xBB\x84\xE5\x90\x88 " + std::to_string(comboId));
            }
        }

        // Construct new data
        std::vector<DataRow> newData;
        newData.reserve(data.size());

        for (size_t i = 0; i < data.size(); ++i) {
            bool keep = true;

            // Logic:
            // 1. If there are Delete rules, and row matches any Delete rule -> Drop.
            if (hasDeleteRules && rowsToDelete.count(i)) {
                keep = false;
            }
            
            // 2. If there are Filter rules, and row matches NO Filter rule -> Drop.
            //    (Only if we haven't already dropped it)
            if (keep && hasFilterRules && !rowsToKeep.count(i)) {
                keep = false;
            }

            // 3. If no rules at all? Keep.
            // 4. If only Delete rules? Keep unless deleted.
            // 5. If only Filter rules? Keep only if filtered.

            if (keep) {
                newData.push_back(std::move(data[i]));
                if (rowsToKeep.count(i)) {
                    result.matchedRows++;
                }
            } else {
                result.deletedRows++;
            }
        }

        data = std::move(newData);
        result.processedRows = data.size();

        auto endTime = std::chrono::high_resolution_clock::now();
        result.processingTime = std::chrono::duration<double, std::milli>(endTime - startTime).count() / 1000.0;

        return result;
    }

    bool optimizeProcessing(std::vector<DataRow>& data) override {
        if (data.empty()) return true;

        // Sort optimization
        std::sort(data.begin(), data.end(), [](const DataRow& a, const DataRow& b) {
            if (a.data.size() > 0 && b.data.size() > 0) {
                try {
                    return std::visit([](const auto& val1, const auto& val2) -> bool {
                        using Type1 = std::decay_t<decltype(val1)>;
                        using Type2 = std::decay_t<decltype(val2)>;

                        if constexpr (std::is_same_v<Type1, Type2> &&
                                     std::is_arithmetic_v<Type1>) {
                            return val1 < val2;
                        }
                        return false;
                    }, a.data[0], b.data[0]);
                } catch (...) {
                    return false;
                }
            }
            return false;
        });

        return true;
    }

    bool validateData(const std::vector<DataRow>& data, std::vector<std::string>& errors) const override {
        errors.clear();

        if (data.empty()) {
            errors.push_back("Data is empty");
            return false;
        }

        // Check consistency
        size_t columnCount = 0;
        for (size_t i = 0; i < data.size(); ++i) {
            if (i == 0) {
                columnCount = data[i].data.size();
            } else if (data[i].data.size() != columnCount) {
                errors.push_back("Row " + std::to_string(i + 1) + " column count mismatch");
            }

            if (!data[i].isValid) {
                errors.push_back("Row " + std::to_string(i + 1) + " invalid data");
            }
        }

        return errors.empty();
    }

private:
    std::function<void(int, const std::string&)> progressCallback_;

    bool evaluateRuleCombination(const DataRow& row,
                               const std::vector<Rule>& rules,
                               const std::vector<int>& ruleIds) {
        // Get rule engine
        auto ruleEngine = createRuleEngine();

        // AND logic
        for (int ruleId : ruleIds) {
            if (ruleId > 0 && ruleId <= static_cast<int>(rules.size())) {
                const Rule& rule = rules[ruleId - 1];
                if (!ruleEngine->evaluateRule(rule, row, &rules)) {
                    return false;
                }
            }
        }
        return true;
    }
};

// Core implementation
ExcelProcessorCore::ExcelProcessorCore() {
    reader_ = std::make_unique<CSVExcelReader>();
    writer_ = std::make_unique<CSVExcelWriter>();
    ruleEngine_ = createRuleEngine();
    dataProcessor_ = std::make_unique<HighPerformanceDataProcessor>();

    stats_.startTime = std::chrono::high_resolution_clock::now();
}

ExcelProcessorCore::~ExcelProcessorCore() = default;

bool ExcelProcessorCore::loadConfiguration(const std::string& configFile) {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);

    std::ifstream file(configFile);
    if (!file.is_open()) {
        addError("Unable to open config file: " + configFile);
        return false;
    }

    loadedConfigFilename_ = configFile;
    rules_.clear();
    tasks_.clear(); // Clear existing tasks
    std::string line;
    int lineNumber = 0;

    // Skip header
    if (std::getline(file, line)) {
        lineNumber++;
    }
    
    bool isTaskSection = false;

    // Parse rules
    while (std::getline(file, line)) {
        lineNumber++;
        
        // Trim whitespace
        line.erase(0, line.find_first_not_of(" \t\r\n"));
        line.erase(line.find_last_not_of(" \t\r\n") + 1);
        
        if (line.empty()) continue;

        if (line == "TASKS_SECTION") {
            isTaskSection = true;
            // Skip task header
            if (std::getline(file, line)) lineNumber++;
            continue;
        }

        try {
            if (isTaskSection) {
                // Parse task
                // ID,TaskName,OutputMode,OutputWorkbookName,IncludeRuleIds,ExcludeRuleIds,OverwriteSheet,InputFilenamePattern,InputSheetName
                ProcessingTask task;
                std::vector<std::string> parts;
                std::stringstream ss(line);
                std::string item;
                
                while (std::getline(ss, item, ',')) {
                    // Handle quoted strings (simple version)
                    if (!item.empty() && item.front() == '"' && item.back() != '"') {
                        std::string next;
                        while (std::getline(ss, next, ',')) {
                            item += "," + next;
                            if (!next.empty() && next.back() == '"') break;
                        }
                    }
                    // Remove quotes
                    if (item.size() >= 2 && item.front() == '"' && item.back() == '"') {
                        item = item.substr(1, item.size() - 2);
                    }
                    parts.push_back(item);
                }
                
                if (parts.size() >= 6) {
                    task.id = std::stoi(parts[0]);
                    task.taskName = parts[1];
                    task.outputMode = static_cast<OutputMode>(std::stoi(parts[2]));
                    task.outputWorkbookName = parts[3];
                    
                    // Parse Include Rules (with Granular Exclusions)
                    std::stringstream incSs(parts[4]);
                    std::string entryStr;
                    while (std::getline(incSs, entryStr, '|')) {
                        if (entryStr.empty()) continue;
                        
                        // Format: ruleId:ex1,ex2... or just ruleId
                        size_t colonPos = entryStr.find(':');
                        if (colonPos != std::string::npos) {
                            int ruleId = std::stoi(entryStr.substr(0, colonPos));
                            TaskRuleEntry entry(ruleId);
                            
                            std::string exPart = entryStr.substr(colonPos + 1);
                            std::stringstream exSs(exPart);
                            std::string exIdStr;
                            while (std::getline(exSs, exIdStr, ',')) {
                                if (!exIdStr.empty()) entry.excludeRuleIds.push_back(std::stoi(exIdStr));
                            }
                            task.rules.push_back(entry);
                        } else {
                            task.rules.push_back(TaskRuleEntry(std::stoi(entryStr)));
                        }
                    }
                    
                    // Parse Exclude IDs (Global)
                    std::stringstream excSs(parts[5]);
                    std::string idStr;
                    while (std::getline(excSs, idStr, '|')) {
                        if (!idStr.empty()) task.excludeRuleIds.push_back(std::stoi(idStr));
                    }
                    
                    if (parts.size() >= 7) {
                        task.overwriteSheet = (parts[6] == "TRUE");
                    }

                    if (parts.size() >= 8) {
                        task.inputFilenamePattern = parts[7];
                    }

                    if (parts.size() >= 9) {
                        task.inputSheetName = parts[8];
                    }

                    if (parts.size() >= 10) {
                        try {
                            task.ruleLogic = static_cast<RuleLogic>(std::stoi(parts[9]));
                        } catch (...) {
                            task.ruleLogic = RuleLogic::OR; // Default
                        }
                    }
                    
                    if (parts.size() >= 11) {
                        task.useHeader = (parts[10] == "TRUE");
                    }
                    
                    task.enabled = true; // Default to enabled
                    tasks_.push_back(task);
                }
            } else {
                Rule rule = parseRuleLine(line);
                if (validateRule(rule)) {
                    rules_.push_back(rule);
                }
            }
        } catch (const std::exception& e) {
            addError("Config parse error at line " + std::to_string(lineNumber) + ": " + e.what());
        }
    }

    return true;
}

bool ExcelProcessorCore::saveConfiguration(const std::string& configFile) const {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);

    std::ofstream file(configFile);
    if (!file.is_open()) {
        addError("Unable to create config file: " + configFile);
        return false;
    }

    // Write header
    file << "ID,Name,Type,Conditions,TargetSheet,Transform,Enabled,Desc,Priority,OutputMode,OutputName\n";

    // Write rules
    for (const auto& rule : rules_) {
        file << rule.id << ","
             << "\"" << rule.name << "\","
             << static_cast<int>(rule.type) << ",";

        // Write conditions
        std::ostringstream condOss;
        for (size_t i = 0; i < rule.conditions.size(); ++i) {
            if (i > 0) condOss << "|";
            condOss << rule.conditions[i].column << ";;" 
                    << static_cast<int>(rule.conditions[i].oper) << ";;";
            
            std::visit([&condOss](const auto& val) {
                using ValType = std::decay_t<decltype(val)>;
                if constexpr (std::is_same_v<ValType, std::string>) {
                    condOss << val;
                } else {
                    condOss << val;
                }
            }, rule.conditions[i].value);
            
            // Save Split Logic
            condOss << ";;" << rule.conditions[i].splitSymbol << ";;" 
                    << static_cast<int>(rule.conditions[i].splitTarget);
        }
        file << "\"" << condOss.str() << "\",";

        file << "\"" << rule.targetSheet << "\","
             << "Transform," 
             << (rule.enabled ? "TRUE" : "FALSE") << ","
             << "\"" << rule.description << "\","
             << rule.priority << ","
             << static_cast<int>(rule.outputMode) << ","
             << "\"" << rule.outputName << "\"\n";
    }

    file << "TASKS_SECTION\n";
    file << "ID,TaskName,OutputMode,OutputWorkbookName,IncludeRuleIds,ExcludeRuleIds,OverwriteSheet,InputFilenamePattern,InputSheetName,RuleLogic,UseHeader\n";
    
    for (const auto& task : tasks_) {
        file << task.id << ","
             << "\"" << task.taskName << "\","
             << static_cast<int>(task.outputMode) << ","
             << "\"" << task.outputWorkbookName << "\",";

        // Include Rule IDs (with Granular Exclusions)
        std::ostringstream incOss;
        for (size_t i = 0; i < task.rules.size(); ++i) {
            if (i > 0) incOss << "|";
            incOss << task.rules[i].ruleId;
            
            if (!task.rules[i].excludeRuleIds.empty()) {
                incOss << ":";
                for (size_t j = 0; j < task.rules[i].excludeRuleIds.size(); ++j) {
                    if (j > 0) incOss << ",";
                    incOss << task.rules[i].excludeRuleIds[j];
                }
            }
        }
        file << "\"" << incOss.str() << "\",";

        // Exclude Rule IDs (Global)
        std::ostringstream excOss;
        for (size_t i = 0; i < task.excludeRuleIds.size(); ++i) {
            if (i > 0) excOss << "|";
            excOss << task.excludeRuleIds[i];
        }
        file << "\"" << excOss.str() << "\",";
        
        // Overwrite Sheet
        file << (task.overwriteSheet ? "TRUE" : "FALSE") << ",";

        // Input Filename Pattern
        file << "\"" << task.inputFilenamePattern << "\",";

        // Input Sheet Name
        file << "\"" << task.inputSheetName << "\",";

        // Rule Logic
        file << static_cast<int>(task.ruleLogic) << ",";
        
        // Use Header
        file << (task.useHeader ? "TRUE" : "FALSE") << "\n";
    }

    return true;
}

void ExcelProcessorCore::addRule(const Rule& rule) {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);

    if (!validateRule(rule)) {
        return;
    }

    rules_.push_back(rule);
}

void ExcelProcessorCore::removeRule(int ruleId) {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);

    auto it = std::remove_if(rules_.begin(), rules_.end(),
        [ruleId](const Rule& rule) { return rule.id == ruleId; });

    rules_.erase(it, rules_.end());
}

void ExcelProcessorCore::updateRule(const Rule& rule) {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);

    if (!validateRule(rule)) {
        return;
    }

    auto it = std::find_if(rules_.begin(), rules_.end(),
        [ruleId = rule.id](const Rule& r) { return r.id == ruleId; });

    if (it != rules_.end()) {
        *it = rule;
    }
}

std::vector<Rule> ExcelProcessorCore::getRules() const {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    return rules_;
}

Rule ExcelProcessorCore::getRule(int ruleId) const {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);

    auto it = std::find_if(rules_.begin(), rules_.end(),
        [ruleId](const Rule& rule) { return rule.id == ruleId; });

    if (it != rules_.end()) {
        return *it;
    }

    return Rule(); 
}

void ExcelProcessorCore::createRuleCombination(int comboId, const std::vector<int>& ruleIds,
                                            CombinationStrategy strategy) {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);

    ruleCombinations_[comboId] = ruleIds;
}

void ExcelProcessorCore::removeRuleCombination(int comboId) {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);

    ruleCombinations_.erase(comboId);
}

std::map<int, std::vector<int>> ExcelProcessorCore::getRuleCombinations() const {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    return ruleCombinations_;
}

ProcessingResult ExcelProcessorCore::processExcelFile(const std::string& inputFile,
                                                    const std::string& outputFile,
                                                    const std::string& sheetName) {
    std::lock_guard<std::mutex> lock(dataMutex_);

    ProcessingResult result;
    auto processingStartTime = std::chrono::high_resolution_clock::now();

    // Determine reader type based on extension
    std::string ext = inputFile.substr(inputFile.find_last_of(".") + 1);
    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
    
    if (ext == "xlsx" || ext == "xls") {
        reader_ = std::make_unique<ActiveQtExcelReader>();
    } else {
        reader_ = std::make_unique<CSVExcelReader>();
    }

    if (logger_) reader_->setLogger(logger_);

    // Read input
    if (!reader_->readExcelFile(inputFile, currentData_, sheetName)) {
        addError("Unable to read input file: " + inputFile);
        return result;
    }

    // Validate
    if (!dataProcessor_->validateData(currentData_, errors_)) {
        addError("Data validation failed");
        return result;
    }

    // Optimize
    dataProcessor_->optimizeProcessing(currentData_);

    // Process
    if (progressCallback_) {
        dataProcessor_->setProgressCallback(progressCallback_);
    }
    result = dataProcessor_->processData(currentData_, rules_, ruleCombinations_);

    // Determine writer type based on output extension
    std::string outExt = outputFile.substr(outputFile.find_last_of(".") + 1);
    std::transform(outExt.begin(), outExt.end(), outExt.begin(), ::tolower);
    
    if (outExt == "xlsx" || outExt == "xls") {
        writer_ = std::make_unique<ActiveQtExcelWriter>();
    } else {
        writer_ = std::make_unique<CSVExcelWriter>();
    }

    // Write output
    if (!writer_->writeExcelFile(outputFile, currentData_)) {
        addError("Unable to write output file: " + outputFile);
    }
    
    // Ensure writer resources are released immediately
    writer_->closeAll();

    auto processingEndTime = std::chrono::high_resolution_clock::now();
    result.processingTime = std::chrono::duration<double, std::milli>(
        processingEndTime - processingStartTime).count() / 1000.0;

    updatePerformanceStats(result.processingTime);

    return result;
}

bool ExcelProcessorCore::loadFile(const std::string& filename, const std::string& sheetName, int maxRows, bool includeHeader) {
    std::lock_guard<std::mutex> lock(dataMutex_);

    std::string ext = filename.substr(filename.find_last_of(".") + 1);
    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);

    if (ext == "xlsx" || ext == "xls") {
        reader_ = std::make_unique<ActiveQtExcelReader>();
    } else {
        reader_ = std::make_unique<CSVExcelReader>();
    }

    if (logger_) reader_->setLogger(logger_);

    if (!reader_->readExcelFile(filename, currentData_, sheetName, maxRows, 0, includeHeader)) {
        addError("Unable to read file: " + filename);
        return false;
    }

    if (includeHeader && !currentData_.empty() && logger_) {
         std::string headerStr;
         bool isEmpty = true;
         if (!currentData_[0].data.empty()) {
             for (const auto& cell : currentData_[0].data) {
                 std::visit([&headerStr, &isEmpty](const auto& val) {
                     using ValType = std::decay_t<decltype(val)>;
                     if constexpr (std::is_same_v<ValType, std::string>) {
                         headerStr += val + " | ";
                         if (!val.empty()) isEmpty = false;
                     } else if constexpr (std::is_arithmetic_v<ValType>) {
                         headerStr += std::to_string(val) + " | ";
                         isEmpty = false;
                     } else if constexpr (std::is_same_v<ValType, bool>) {
                         headerStr += (val ? "TRUE" : "FALSE") + std::string(" | ");
                         isEmpty = false;
                     }
                 }, cell);
             }
         }
         logger_("Loaded Header Row (Row 1): " + headerStr);
         if (isEmpty) logger_("WARNING: Header Row appears to be empty.");
    }
    return true;
}

std::vector<std::string> ExcelProcessorCore::getSheetNames(const std::string& filename) {
    std::unique_ptr<ExcelReader> tempReader;
    std::string ext = filename.substr(filename.find_last_of(".") + 1);
    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);

    if (ext == "xlsx" || ext == "xls") {
        tempReader = std::make_unique<ActiveQtExcelReader>();
    } else {
        tempReader = std::make_unique<CSVExcelReader>();
    }
    
    if (logger_) tempReader->setLogger(logger_);

    std::vector<std::string> names;
    tempReader->getSheetNames(filename, names);
    return names;
}

void ExcelProcessorCore::setLogger(std::function<void(const std::string&)> logger) {
    logger_ = logger;
    if (reader_) {
        reader_->setLogger(logger);
    }
}

void ExcelProcessorCore::setProgressCallback(std::function<void(int, const std::string&)> callback) {
    progressCallback_ = callback;
}

bool ExcelProcessorCore::previewResults(const std::string& inputFile, const std::string& sheetName, int maxPreviewRows) {
    if (logger_) logger_("Previewing file: " + inputFile + ", Sheet: " + (sheetName.empty() ? "Default" : sheetName) + ", Max Rows: " + std::to_string(maxPreviewRows));
    
    {
        std::lock_guard<std::mutex> lock(dataMutex_);
        currentData_.clear();
    }

    // Load data for preview
    // Force includeHeader=true to ensure we capture the first row (header) in memory
    bool success = loadFile(inputFile, sheetName, maxPreviewRows, true);
    
    if (logger_) logger_("File load result: " + std::string(success ? "Success" : "Failed") + ", Rows loaded: " + std::to_string(currentData_.size()));
    return success;
}

std::vector<DataRow> ExcelProcessorCore::getPreviewData(int maxRows) const {
    std::lock_guard<std::mutex> lock(dataMutex_);

    std::vector<DataRow> preview;
    size_t rowsToCopy = std::min(static_cast<size_t>(maxRows), currentData_.size());

    for (size_t i = 0; i < rowsToCopy; ++i) {
        preview.push_back(currentData_[i]);
    }

    return preview;
}

std::vector<DataRow> ExcelProcessorCore::getProcessedPreviewData(int maxRows) {
    std::lock_guard<std::mutex> lock(dataMutex_);
    std::lock_guard<std::recursive_mutex> rulesLock(rulesMutex_);

    std::vector<DataRow> preview;
    size_t rowsToCopy = std::min(static_cast<size_t>(maxRows), currentData_.size());

    for (size_t i = 0; i < rowsToCopy; ++i) {
        preview.push_back(currentData_[i]);
    }

    // Apply rules (this modifies preview in-place)
    if (dataProcessor_) {
        dataProcessor_->processData(preview, rules_, ruleCombinations_);
    }

    return preview;
}

std::vector<DataRow> ExcelProcessorCore::getProcessedPreviewData(const std::vector<int>& ruleIds, int maxRows) {
    std::lock_guard<std::mutex> lock(dataMutex_);
    std::lock_guard<std::recursive_mutex> rulesLock(rulesMutex_);

    std::vector<DataRow> preview;
    if (currentData_.empty()) return preview;

    size_t rowsToCopy = std::min(static_cast<size_t>(maxRows), currentData_.size());

    for (size_t i = 0; i < rowsToCopy; ++i) {
        preview.push_back(currentData_[i]);
    }

    // Filter rules
    std::vector<Rule> selectedRules;
    for (int id : ruleIds) {
        for (const auto& rule : rules_) {
            if (rule.id == id) {
                selectedRules.push_back(rule);
                break;
            }
        }
    }

    // Apply rules
    if (dataProcessor_) {
        std::map<int, std::vector<int>> emptyCombinations;
        dataProcessor_->processData(preview, selectedRules, emptyCombinations);
    }

    return preview;
}

std::vector<DataRow> ExcelProcessorCore::getTaskPreviewData(int taskId, int maxRows) {
    std::lock_guard<std::mutex> lock(dataMutex_);
    std::lock_guard<std::recursive_mutex> rulesLock(rulesMutex_);

    std::vector<DataRow> preview;
    if (currentData_.empty()) return preview;

    // Find Task
    ProcessingTask task;
    bool found = false;
    for (const auto& t : tasks_) {
        if (t.id == taskId) {
            task = t;
            found = true;
            break;
        }
    }
    if (!found) return preview;

    size_t rowsToCopy = std::min(static_cast<size_t>(maxRows), currentData_.size());
    std::vector<DataRow> sourceData;
    sourceData.reserve(rowsToCopy);
    for (size_t i = 0; i < rowsToCopy; ++i) {
        sourceData.push_back(currentData_[i]);
    }

    // Emulate Task Logic (Subset of processExcelFile logic)
    // We don't need full validation/error handling here, just filtering/transforming
    
    std::vector<DataRow> resultData;
    resultData.reserve(sourceData.size());

    // Helper to find rule
    auto findRule = [this](int id) -> const Rule* {
        for (const auto& r : rules_) {
            if (r.id == id) return &r;
        }
        return nullptr;
    };

    if (logger_) logger_("Generating preview for Task ID: " + std::to_string(taskId));
    
    // Log Header Info
    if (!sourceData.empty()) {
        QString headerContent;
        if (!sourceData[0].data.empty()) {
            // Just peek first few columns
            int cols = std::min((size_t)5, sourceData[0].data.size());
            for(int k=0; k<cols; ++k) {
                if (std::holds_alternative<std::string>(sourceData[0].data[k])) {
                    headerContent += QString::fromStdString(std::get<std::string>(sourceData[0].data[k])) + "|";
                }
            }
        }
        if (logger_) logger_("Source Data Row 0 (Potential Header): " + headerContent.toStdString());
    }

    // Pre-fetch rules for performance
    // ... or just lookup in loop since it's preview (small size)

    int matchCount = 0;
    for (size_t i = 0; i < sourceData.size(); ++i) {
        // Special handling for header row
        if (i == 0 && task.useHeader) {
            resultData.push_back(sourceData[i]);
            if (logger_) logger_("Row 0 preserved as Header.");
            continue;
        }

        const auto& row = sourceData[i];
        bool include = false;
        std::string debugReason;
        
        if (task.rules.empty()) {
            include = true;
            debugReason = "No rules (include all)";
        } else {
            if (task.ruleLogic == RuleLogic::OR) {
                include = false;
                for (const auto& ruleEntry : task.rules) {
                    const Rule* rule = findRule(ruleEntry.ruleId);
                    if (rule && ruleEngine_->evaluateRule(*rule, row, &rules_)) {
                        bool granularExcluded = false;
                        if (!ruleEntry.excludeRuleIds.empty()) {
                            for (int exId : ruleEntry.excludeRuleIds) {
                                const Rule* exRule = findRule(exId);
                                if (exRule && ruleEngine_->evaluateRule(*exRule, row, &rules_)) {
                                    granularExcluded = true;
                                    debugReason = "Excluded by granular rule ID: " + std::to_string(exId);
                                    break;
                                }
                            }
                        }
                        if (!granularExcluded) {
                            include = true;
                            debugReason = "Matched rule ID: " + std::to_string(rule->id);
                            break; 
                        }
                    }
                }
            } else { // AND
                include = true;
                for (const auto& ruleEntry : task.rules) {
                    const Rule* rule = findRule(ruleEntry.ruleId);
                    if (!rule || !ruleEngine_->evaluateRule(*rule, row, &rules_)) {
                         include = false;
                         debugReason = "Failed rule ID: " + std::to_string(ruleEntry.ruleId);
                         break;
                    }
                    if (!ruleEntry.excludeRuleIds.empty()) {
                        for (int exId : ruleEntry.excludeRuleIds) {
                            const Rule* exRule = findRule(exId);
                            if (exRule && ruleEngine_->evaluateRule(*exRule, row, &rules_)) {
                                include = false;
                                debugReason = "Excluded by granular rule ID: " + std::to_string(exId);
                                break;
                            }
                        }
                        if (!include) break;
                    }
                }
                if (include) debugReason = "Matched all rules (AND)";
            }
        }

        // Global Exclusions
        if (include && !task.excludeRuleIds.empty()) {
             for (int exId : task.excludeRuleIds) {
                 const Rule* exRule = findRule(exId);
                 if (exRule && ruleEngine_->evaluateRule(*exRule, row, &rules_)) {
                     include = false;
                     debugReason = "Excluded by global rule ID: " + std::to_string(exId);
                     break;
                 }
             }
        }

        if (include) {
            resultData.push_back(row);
            matchCount++;
        }
        
        if (logger_ && i < 5) {
             logger_("Row " + std::to_string(i+1) + ": " + (include ? "MATCH" : "SKIP") + " - " + debugReason);
        }
    }
    
    if (logger_) logger_("Preview generated. Input Rows: " + std::to_string(sourceData.size()) + ", Output Rows: " + std::to_string(resultData.size()));

    return resultData;
}

bool ExcelProcessorCore::validateRule(const Rule& rule) const {
    auto errors = ruleEngine_->getValidationErrors(rule);
    for (const auto& error : errors) {
        addError("Rule validation error: " + error);
    }
    return errors.empty();
}

void ExcelProcessorCore::addError(const std::string& error) const {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    errors_.push_back(error);
}

void ExcelProcessorCore::addWarning(const std::string& warning) const {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    warnings_.push_back(warning);
}

std::vector<std::string> ExcelProcessorCore::getErrors() const {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    return errors_;
}

std::vector<std::string> ExcelProcessorCore::getWarnings() const {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    return warnings_;
}

void ExcelProcessorCore::clearErrors() {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    errors_.clear();
}

void ExcelProcessorCore::clearWarnings() {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    warnings_.clear();
}

void ExcelProcessorCore::updatePerformanceStats(double processingTime) {
    stats_.cpuTime += processingTime;
    stats_.endTime = std::chrono::high_resolution_clock::now();
}

// Helper for splitting string with custom delimiter
static std::vector<std::string> splitString(const std::string& str, const std::string& delimiter) {
    std::vector<std::string> tokens;
    size_t prev = 0, pos = 0;
    do {
        pos = str.find(delimiter, prev);
        if (pos == std::string::npos) pos = str.length();
        std::string token = str.substr(prev, pos - prev);
        if (!token.empty()) tokens.push_back(token);
        prev = pos + delimiter.length();
    } while (pos < str.length() && prev < str.length());
    return tokens;
}

// Parse rule line
Rule ExcelProcessorCore::parseRuleLine(const std::string& line) {
    std::vector<std::string> tokens;
    std::string token;
    bool inQuotes = false;
    
    // Better CSV parsing handling quotes
    for (char c : line) {
        if (c == '"') {
            inQuotes = !inQuotes;
        } else if (c == ',' && !inQuotes) {
            // Remove surrounding quotes if present
            if (token.length() >= 2 && token.front() == '"' && token.back() == '"') {
                token = token.substr(1, token.length() - 2);
            }
            tokens.push_back(token);
            token.clear();
        } else {
            token += c;
        }
    }
    // Last token
    if (token.length() >= 2 && token.front() == '"' && token.back() == '"') {
        token = token.substr(1, token.length() - 2);
    }
    tokens.push_back(token);

    Rule rule;
    if (tokens.size() >= 1) rule.id = std::stoi(tokens[0]);
    if (tokens.size() >= 2) rule.name = tokens[1];
    if (tokens.size() >= 3) rule.type = static_cast<RuleType>(std::stoi(tokens[2]));

    // Parse conditions
    if (tokens.size() >= 4) {
        std::string condStr = tokens[3];
        std::vector<std::string> conds = splitString(condStr, "|");
        
        for (const auto& cStr : conds) {
            std::vector<std::string> parts = splitString(cStr, ";;");
            if (parts.size() >= 2) {
                RuleCondition condition;
                condition.column = std::stoi(parts[0]);
                condition.oper = static_cast<Operator>(std::stoi(parts[1]));
                if (parts.size() >= 3) {
                    condition.value = parts[2];
                } else {
                    condition.value = std::string("");
                }
                
                // Parse Split Logic
                if (parts.size() >= 5) {
                    condition.splitSymbol = parts[3];
                    condition.splitTarget = static_cast<SplitTarget>(std::stoi(parts[4]));
                }

                rule.conditions.push_back(condition);
            }
        }
    }

    if (tokens.size() >= 5) rule.targetSheet = tokens[4];
    // tokens[5] is Transform (unused/fixed)
    if (tokens.size() >= 7) rule.enabled = (tokens[6] == "TRUE");
    if (tokens.size() >= 8) rule.description = tokens[7];
    if (tokens.size() >= 9) rule.priority = std::stoi(tokens[8]);
    if (tokens.size() >= 10) rule.outputMode = static_cast<OutputMode>(std::stoi(tokens[9]));
    if (tokens.size() >= 11) rule.outputName = tokens[10];

    // Parse excluded rules (REMOVED)
    /*
    if (tokens.size() >= 12) {
        std::string exStr = tokens[11];
        std::vector<std::string> exIds = splitString(exStr, "|");
        for (const auto& idStr : exIds) {
            try {
                // rule.excludedRuleIds.push_back(std::stoi(idStr));
            } catch (...) {}
        }
    }
    */

    return rule;
}

// Performance statistics methods
PerformanceStats ExcelProcessorCore::getPerformanceStats() const {
    return stats_;
}

void ExcelProcessorCore::resetStats() {
    stats_ = PerformanceStats();
    stats_.startTime = std::chrono::high_resolution_clock::now();
}

// Validation methods
std::vector<std::string> ExcelProcessorCore::getValidationMessages() const {
    std::vector<std::string> messages;

    // Copy errors and warnings to validation messages
    {
        std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
        messages.insert(messages.end(), errors_.begin(), errors_.end());
        messages.insert(messages.end(), warnings_.begin(), warnings_.end());
    }

    return messages;
}

// Task management implementation
void ExcelProcessorCore::addTask(const ProcessingTask& task) {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    tasks_.push_back(task);
}

void ExcelProcessorCore::removeTask(int taskId) {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    auto it = std::remove_if(tasks_.begin(), tasks_.end(),
        [taskId](const ProcessingTask& t) { return t.id == taskId; });
    tasks_.erase(it, tasks_.end());
}

void ExcelProcessorCore::updateTask(const ProcessingTask& task) {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    auto it = std::find_if(tasks_.begin(), tasks_.end(),
        [id = task.id](const ProcessingTask& t) { return t.id == id; });
    if (it != tasks_.end()) {
        *it = task;
    }
}

std::vector<ProcessingTask> ExcelProcessorCore::getTasks() const {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    return tasks_;
}

ProcessingTask ExcelProcessorCore::getTask(int taskId) const {
    std::lock_guard<std::recursive_mutex> lock(rulesMutex_);
    auto it = std::find_if(tasks_.begin(), tasks_.end(),
        [taskId](const ProcessingTask& t) { return t.id == taskId; });
    if (it != tasks_.end()) {
        return *it;
    }
    return ProcessingTask();
}

std::vector<ProcessingResult> ExcelProcessorCore::processTasks(const std::string& inputFile, const std::string& defaultOutputFile, const std::string& sheetName) {
    auto tasks = getTasks();
    return processTasksInternal(tasks, inputFile, defaultOutputFile, sheetName);
}

ProcessingResult ExcelProcessorCore::processTask(int taskId) {
    auto tasks = getTasks();
    ProcessingTask task;
    bool found = false;
    for (const auto& t : tasks) {
        if (t.id == taskId) {
            task = t;
            found = true;
            break;
        }
    }

    ProcessingResult result;
    if (!found) {
        addError("Task not found: " + std::to_string(taskId));
        return result;
    }

    std::vector<ProcessingTask> singleTask = { task };
    
    // Use empty inputFile to trigger pattern matching in processTasksInternal
    auto results = processTasksInternal(singleTask, "", "", "");

    if (!results.empty()) {
        // Aggregate results if multiple files were processed
        result = results[0];
        for (size_t i = 1; i < results.size(); ++i) {
            result.totalRows += results[i].totalRows;
            result.processedRows += results[i].processedRows;
            result.matchedRows += results[i].matchedRows;
            result.deletedRows += results[i].deletedRows;
            result.processingTime += results[i].processingTime;
            result.errors.insert(result.errors.end(), results[i].errors.begin(), results[i].errors.end());
        }
    }
    return result;
}

std::vector<ProcessingResult> ExcelProcessorCore::processTasksInternal(const std::vector<ProcessingTask>& tasksToProcess, const std::string& inputFile, const std::string& defaultOutputFile, const std::string& sheetName) {
    // Dispatcher Mode: If inputFile is empty, scan for files based on enabled tasks
    if (inputFile.empty()) {
        std::vector<ProcessingResult> allResults;
        std::set<std::string> uniqueFiles;
        
        // Get tasks (thread-safe internally)
        auto tasks = tasksToProcess; 
        
        for (const auto& task : tasks) {
            if (!task.enabled || task.inputFilenamePattern.empty()) continue;
            
            QString pattern = QString::fromStdString(task.inputFilenamePattern);
            QFileInfo info(pattern);
            QDir dir = info.isAbsolute() ? info.absoluteDir() : QDir::current();
            QString fileNamePattern = info.fileName();
            
            // Handle case where pattern implies current directory
            if (info.isRelative() && !pattern.contains("/") && !pattern.contains("\\")) {
                 dir = QDir::current();
                 fileNamePattern = pattern;
            }
            
            QStringList filters;
            filters << fileNamePattern;
            
            QFileInfoList list = dir.entryInfoList(filters, QDir::Files);
            for (const auto& fileInfo : list) {
                uniqueFiles.insert(fileInfo.absoluteFilePath().toStdString());
            }
        }
        
        if (uniqueFiles.empty()) {
             // No files found to process
             return allResults;
        }
        
        // Process each unique file sequentially
        for (const auto& file : uniqueFiles) {
            // Call processTasks with specific file (will lock internally)
            auto results = processTasksInternal(tasksToProcess, file, defaultOutputFile, sheetName);
            allResults.insert(allResults.end(), results.begin(), results.end());
        }
        return allResults;
    }

    // Single File Processing Mode
    std::lock_guard<std::mutex> lock(dataMutex_);
    std::vector<ProcessingResult> results;
    
    // Determine sheets to process
    std::vector<std::string> sheetsToProcess;
    if (sheetName.empty()) {
        sheetsToProcess = getSheetNames(inputFile);
        if (sheetsToProcess.empty()) {
             // Fallback to default behavior (read first sheet) if getting names fails or empty
             // But for tasks, maybe we should try to read first sheet explicitly?
             // Let's rely on readExcelFile's default if list is empty
             sheetsToProcess.push_back(""); 
        }
    } else {
        sheetsToProcess.push_back(sheetName);
    }

    // Determine reader based on input file extension for robustness
    std::string ext = inputFile.substr(inputFile.find_last_of(".") + 1);
    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
    
    if (ext == "xlsx" || ext == "xls") {
        reader_ = std::make_unique<ActiveQtExcelReader>();
    } else {
        reader_ = std::make_unique<CSVExcelReader>();
    }

    if (logger_) reader_->setLogger(logger_);

    auto tasks = tasksToProcess;
    auto rules = getRules(); // Thread-safe copy

    auto findRule = [&rules](int id) -> const Rule* {
        for (const auto& r : rules) {
            if (r.id == id) return &r;
        }
        return nullptr;
    };

    // Calculate actual active tasks for progress calculation
    int activeTaskCount = 0;
    for (const auto& task : tasks) {
        if (task.enabled) activeTaskCount++;
    }
    if (activeTaskCount == 0) activeTaskCount = 1;

    // Iterate over sheets
    // Check for missing sheets required by tasks and determine which sheets to process
    bool processAllSheets = false;
    std::vector<std::string> requiredSheets; 
    bool anyTaskMatchesFile = false;

    for (const auto& task : tasks) {
        if (!task.enabled) continue;
        
        // Check if task applies to this file (filename pattern)
        bool fileMatches = true;
        if (!task.inputFilenamePattern.empty()) {
            QString qInputFile = QFileInfo(QString::fromStdString(inputFile)).fileName();
            QString qPattern = QString::fromStdString(task.inputFilenamePattern);
            QFileInfo patternInfo(qPattern);

            if (patternInfo.isAbsolute()) {
                 QString absInput = QFileInfo(QString::fromStdString(inputFile)).absoluteFilePath();
                 QString absPattern = patternInfo.absoluteFilePath();
                 if (absInput.compare(absPattern, Qt::CaseInsensitive) != 0) fileMatches = false;
            } else {
                QRegExp rx(qPattern, Qt::CaseInsensitive, QRegExp::Wildcard);
                if (!rx.exactMatch(qInputFile)) fileMatches = false;
            }
        }
        
        if (fileMatches) {
            anyTaskMatchesFile = true;
            if (task.inputSheetName.empty()) {
                processAllSheets = true;
            } else {
                requiredSheets.push_back(task.inputSheetName);
                
                // Check if this required sheet exists (for warning)
                bool sheetFound = false;
                for (const auto& s : sheetsToProcess) {
                    if (QString::fromStdString(s).compare(QString::fromStdString(task.inputSheetName), Qt::CaseInsensitive) == 0) {
                        sheetFound = true;
                        break;
                    }
                }
                if (!sheetFound) {
                    // "Warning: Task '...' specified worksheet '...' not found in file '...'"
                    std::string msg = "\xE8\xAD\xA6\xE5\x91\x8A: \xE4\xBB\xBB\xE5\x8A\xA1 '" + task.taskName + "' \xE6\x8C\x87\xE5\xAE\x9A\xE7\x9A\x84\xE5\xB7\xA5\xE4\xBD\x9C\xE8\xA1\xA8 '" + task.inputSheetName + "' \xE5\x9C\xA8\xE6\x96\x87\xE4\xBB\xB6 '" + inputFile + "' \xE4\xB8\xAD\xE6\x9C\xAA\xE6\x89\xBE\xE5\x88\xB0\xE3\x80\x82"; 
                    if (logger_) logger_(msg);
                }
            }
        }
    }

    // Filter sheetsToProcess if not all sheets are requested
    if (!processAllSheets) {
        std::vector<std::string> filteredSheets;
        for (const auto& sheet : sheetsToProcess) {
            bool kept = false;
            for (const auto& req : requiredSheets) {
                if (QString::fromStdString(sheet).compare(QString::fromStdString(req), Qt::CaseInsensitive) == 0) {
                    kept = true;
                    break;
                }
            }
            if (kept) {
                filteredSheets.push_back(sheet);
            }
        }
        sheetsToProcess = filteredSheets;
    }

    for (const auto& currentSheet : sheetsToProcess) {
        // Store results for this sheet: TaskID -> ProcessingResult
        std::map<int, ProcessingResult> sheetTaskResults;
        std::map<int, bool> taskHasStarted; // Track if task has processed its first chunk

        // Initialize results for applicable tasks
        for (const auto& task : tasks) {
            if (!task.enabled) continue;
            
            // Filter by Input Filename (Multi-workbook support)
            if (!task.inputFilenamePattern.empty()) {
                QString qInputFile = QFileInfo(QString::fromStdString(inputFile)).fileName();
                QString qPattern = QString::fromStdString(task.inputFilenamePattern);
                QFileInfo patternInfo(qPattern);

                if (patternInfo.isAbsolute()) {
                     QString absInput = QFileInfo(QString::fromStdString(inputFile)).absoluteFilePath();
                     QString absPattern = patternInfo.absoluteFilePath();
                     if (absInput.compare(absPattern, Qt::CaseInsensitive) != 0) continue;
                } else {
                    QRegExp rx(qPattern, Qt::CaseInsensitive, QRegExp::Wildcard);
                    if (!rx.exactMatch(qInputFile)) continue;
                }
            }

            // Filter by Input Sheet Name (Multi-workbook support)
            if (!task.inputSheetName.empty()) {
                 if (QString::fromStdString(currentSheet).compare(QString::fromStdString(task.inputSheetName), Qt::CaseInsensitive) != 0) continue;
            }
            
            sheetTaskResults[task.id] = ProcessingResult();
            sheetTaskResults[task.id].processingTime = 0;
            taskHasStarted[task.id] = false;
        }

        if (logger_) {
            std::string displaySheet = currentSheet.empty() ? "[Default Sheet]" : currentSheet;
            std::string msg = "\xE5\xBC\x80\xE5\xA7\x8B\xE5\xA4\x84\xE7\x90\x86\xE6\x96\x87\xE4\xBB\xB6: " + inputFile + " | \xE5\xB7\xA5\xE4\xBD\x9C\xE8\xA1\xA8: " + displaySheet + " | \xE5\x8C\xB9\xE9\x85\x8D\xE4\xBB\xBB\xE5\x8A\xA1\xE6\x95\xB0: " + std::to_string(sheetTaskResults.size()); 
            logger_(msg);
            
            if (sheetTaskResults.empty()) {
                 logger_("\xE8\xAD\xA6\xE5\x91\x8A: \xE5\xBD\x93\xE5\x89\x8D\xE6\x96\x87\xE4\xBB\xB6/\xE5\xB7\xA5\xE4\xBD\x9C\xE8\xA1\xA8\xE6\x9C\xAA\xE5\x8C\xB9\xE9\x85\x8D\xE5\x88\xB0\xE4\xBB\xBB\xE4\xBD\x95\xE5\x90\xAF\xE7\x94\xA8\xE7\x9A\x84\xE4\xBB\xBB\xE5\x8A\xA1\xE3\x80\x82");
                 
                 // Debug Info
                 for (const auto& task : tasks) {
                    if (!task.enabled) continue;
                    std::string taskInfo = "Task '" + task.taskName + "' mismatch: ";
                    
                    if (!task.inputFilenamePattern.empty()) {
                        QString qInputFile = QFileInfo(QString::fromStdString(inputFile)).fileName();
                        QString qPattern = QString::fromStdString(task.inputFilenamePattern);
                        QFileInfo patternInfo(qPattern);

                        if (patternInfo.isAbsolute()) {
                             QString absInput = QFileInfo(QString::fromStdString(inputFile)).absoluteFilePath();
                             QString absPattern = patternInfo.absoluteFilePath();
                             if (absInput.compare(absPattern, Qt::CaseInsensitive) != 0) {
                                 taskInfo += "File Absolute Path Mismatch (Input: " + absInput.toStdString() + " vs Pattern: " + absPattern.toStdString() + "); ";
                             }
                        } else {
                            QRegExp rx(qPattern, Qt::CaseInsensitive, QRegExp::Wildcard);
                            if (!rx.exactMatch(qInputFile)) {
                                taskInfo += "Filename Wildcard Mismatch (Input: " + qInputFile.toStdString() + " vs Pattern: " + qPattern.toStdString() + "); ";
                            }
                        }
                    }

                    if (!task.inputSheetName.empty()) {
                         if (QString::fromStdString(currentSheet).compare(QString::fromStdString(task.inputSheetName), Qt::CaseInsensitive) != 0) {
                             taskInfo += "Sheet Name Mismatch (Current: " + currentSheet + " vs Task: " + task.inputSheetName + "); ";
                         }
                    }
                    
                    logger_(taskInfo);
                 }
            }
        }

        // Chunked Processing Loop
        int offset = 0;
        int chunkSize = 5000; // Chunk size
        bool isFirstChunk = true;
        
        ActiveQtExcelWriter sharedQtWriter; // Persistent writer for this file processing

        auto sheetStartTime = std::chrono::high_resolution_clock::now();

        while (true) {
            currentData_.clear();
            
            // Notify Progress (Before Read)
            if (progressCallback_) {
                progressCallback_(offset, "Reading " + currentSheet + " (Row " + std::to_string(offset) + ")...");
            }

            // Determine if we need to read header
            // Calculate for every chunk to ensure correct offset calculation in reader
            bool includeHeader = false;
            for (const auto& task : tasks) {
                  if (task.enabled && task.useHeader) {
                      includeHeader = true;
                      break;
                  }
             }

             if (logger_) {
                 logger_("[DEBUG] Chunk: Offset=" + std::to_string(offset) + " | HeaderReq=" + (includeHeader ? "YES" : "NO"));
             }

             if (!reader_->readExcelFile(inputFile, currentData_, currentSheet, chunkSize, offset, includeHeader)) {
                std::string err = "Unable to read input file: " + inputFile + " (Sheet: " + (currentSheet.empty() ? "Default" : currentSheet) + ")";
                if (isFirstChunk) {
                    addError(err);
                }
                if (logger_) logger_("ERROR: " + err);
                break; // Stop or error
            }
            
            if (isFirstChunk && includeHeader && !currentData_.empty() && logger_) {
                 std::string headerStr;
                 bool isEmpty = true;
                 if (!currentData_[0].data.empty()) {
                     for (const auto& cell : currentData_[0].data) {
                         std::visit([&headerStr, &isEmpty](const auto& val) {
                             using ValType = std::decay_t<decltype(val)>;
                             if constexpr (std::is_same_v<ValType, std::string>) {
                                 headerStr += val + " | ";
                                 if (!val.empty()) isEmpty = false;
                             } else if constexpr (std::is_arithmetic_v<ValType>) {
                                 headerStr += std::to_string(val) + " | ";
                                 isEmpty = false;
                             } else if constexpr (std::is_same_v<ValType, bool>) {
                                 headerStr += (val ? "TRUE" : "FALSE") + std::string(" | ");
                                 isEmpty = false;
                             }
                         }, cell);
                     }
                 }
                 logger_("Processing Header Row (Row 1): " + headerStr);
                 if (isEmpty) logger_("WARNING: Processing Header Row appears to be empty.");
            }
            
            if (currentData_.empty()) break; // EOF

            // Optimize/Validate Input Chunk
            // dataProcessor_->optimizeProcessing(currentData_); // DISABLED: This sorts data and breaks header position

            // Process this chunk for each applicable task
            for (const auto& task : tasks) {
                if (sheetTaskResults.find(task.id) == sheetTaskResults.end()) continue;

                ProcessingResult& result = sheetTaskResults[task.id];
                bool isTaskFirstChunk = !taskHasStarted[task.id];
                taskHasStarted[task.id] = true;

                if (logger_ && isTaskFirstChunk) {
                    std::string displayTaskName = task.taskName.empty() ? "[Unnamed Task]" : task.taskName;
                    std::string msg = "正在处理任务: " + displayTaskName;
                    logger_(msg);
                    
                    std::string sheetDisp = currentSheet.empty() ? "默认" : currentSheet;
                    if (task.useHeader) {
                        logger_(std::string("\xE6\x8F\x90\xE7\xA4\xBA: \xE4\xBB\xBB\xE5\x8A\xA1 '") + displayTaskName + "' [\xE5\xB7\xB2\xE5\x8B\xBE\xE9\x80\x89]\xE4\xBD\xBF\xE7\x94\xA8\xE5\x8E\x9F\xE8\xA1\xA8\xE5\xA4\xB4\xE4\xBF\xA1\xE6\x81\xAF\xEF\xBC\x8C'" + sheetDisp + "' \xE5\xB7\xA5\xE4\xBD\x9C\xE8\xA1\xA8\xE7\x9A\x84\xE9\xA6\x96\xE8\xA1\x8C\xE6\x95\xB0\xE6\x8D\xAE\xE5\xB0\x86\xE4\xBD\x9C\xE4\xB8\xBA\xE8\xA1\xA8\xE5\xA4\xB4\xE4\xBF\x9D\xE7\x95\x99\xE4\xB8\x94\xE4\xB8\x8D\xE5\x8F\x82\xE4\xB8\x8E\xE6\x95\xB0\xE6\x8D\xAE\xE7\xAD\x9B\xE9\x80\x89\xE3\x80\x82");
                    } else {
                        std::string logMsg = std::string("\xE4\xBB\xBB\xE5\x8A\xA1 '") + displayTaskName + "' [\xE6\x9C\xAA\xE9\x80\x89\xE6\x8B\xA9]\xE4\xBD\xBF\xE7\x94\xA8\xE5\x8E\x9F\xE8\xA1\xA8\xE5\xA4\xB4\xE4\xBF\xA1\xE6\x81\xAF\xEF\xBC\x8C'" + sheetDisp + "' \xE5\xB7\xA5\xE4\xBD\x9C\xE8\xA1\xA8\xE7\x9A\x84\xE9\xA6\x96\xE8\xA1\x8C\xE6\x95\xB0\xE6\x8D\xAE\xE5\xB0\x86\xE8\xA7\x86\xE4\xB8\xBA\xE6\x95\xB0\xE6\x8D\xAE\xE5\x8F\x82\xE4\xB8\x8E\xE7\xAD\x9B\xE9\x80\x89\xE3\x80\x82";
                        if (task.overwriteSheet) {
                            logger_(std::string("\xE8\xAD\xA6\xE5\x91\x8A: ") + logMsg + " (\xE8\xA6\x86\xE7\x9B\x96\xE6\xA8\xA1\xE5\xBC\x8F\xE5\x8F\xAF\xE8\x83\xBD\xE5\xAF\xBC\xE8\x87\xB4\xE8\xA1\xA8\xE5\xA4\xB4\xE4\xB8\xA2\xE5\xA4\xB1)");
                        } else {
                            logger_(std::string("\xE6\x8F\x90\xE7\xA4\xBA: ") + logMsg);
                        }
                    }
                }

                std::vector<DataRow> taskData;
                taskData.reserve(currentData_.size());

                // Determine if we should write header for this task
                bool shouldWriteHeader = false;
                if (task.useHeader && isTaskFirstChunk) {
                    // Fix: If Output Mode is NEW_WORKBOOK, we are creating/overwriting the file, so we MUST write header.
                    // If Overwrite Sheet is true, we also write header.
                    if (task.outputMode == OutputMode::NEW_WORKBOOK || task.overwriteSheet) {
                        shouldWriteHeader = true;
                    } else {
                        // Check if target sheet/file is empty
                        std::string targetFile;
                        std::string targetSheet = "Sheet1";

                        if (task.outputMode == OutputMode::NEW_SHEET) {
                            targetFile = defaultOutputFile;
                            if (targetFile.empty()) targetFile = inputFile;
                            targetSheet = task.outputWorkbookName.empty() ? task.taskName : task.outputWorkbookName;
                        } else {
                            // NEW_WORKBOOK
                            std::string name = task.outputWorkbookName.empty() ? task.taskName : task.outputWorkbookName;
                            QFileInfo fileInfo(QString::fromStdString(name));
                            if (fileInfo.isRelative()) {
                                std::string baseDir;
                                if (!defaultOutputFile.empty()) {
                                    baseDir = QFileInfo(QString::fromStdString(defaultOutputFile)).absolutePath().toStdString();
                                } else {
                                    baseDir = QFileInfo(QString::fromStdString(inputFile)).absolutePath().toStdString();
                                }
                                targetFile = baseDir + "/" + name;
                            } else {
                                targetFile = name;
                            }
                        }

                        // Normalize extension
                        std::string ext = targetFile.substr(targetFile.find_last_of(".") + 1);
                        std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
                        if (ext == targetFile) { targetFile += ".xlsx"; ext = "xlsx"; }

                        if (ext == "xlsx" || ext == "xls") {
                             shouldWriteHeader = sharedQtWriter.isSheetEmpty(targetFile, targetSheet);
                        } else {
                             CSVExcelWriter csv;
                             shouldWriteHeader = csv.isSheetEmpty(targetFile, targetSheet);
                        }
                    }
                }
                
                int processedRowsInChunk = 0;
                
                int rowIdx = 0;
                for (const auto& row : currentData_) {
                    bool isHeaderRow = (isFirstChunk && includeHeader && rowIdx == 0);
                    
                    if (isHeaderRow) {
                         if (shouldWriteHeader) {
                             taskData.push_back(row);
                             if (logger_ && task.useHeader) {
                                  std::string hStr;
                                  for(const auto& c : row.data) {
                                      std::string valStr = std::visit([](auto&& arg) -> std::string {
                                          using T = std::decay_t<decltype(arg)>;
                                          if constexpr (std::is_same_v<T, std::string>) return arg;
                                          else if constexpr (std::is_same_v<T, int>) return std::to_string(arg);
                                          else if constexpr (std::is_same_v<T, double>) return std::to_string(arg);
                                          else if constexpr (std::is_same_v<T, bool>) return arg ? "TRUE" : "FALSE";
                                          else if constexpr (std::is_same_v<T, std::tm>) {
                                              char buf[64];
                                              std::strftime(buf, sizeof(buf), "%Y-%m-%d %H:%M:%S", &arg);
                                              return buf;
                                          }
                                          return "";
                                      }, c);
                                      hStr += valStr + " | ";
                                  }
                                  // "Writing Header"
                                  logger_(std::string("[DEBUG] \xE5\x86\x99\xE5\x85\xA5\xE8\xA1\xA8\xE5\xA4\xB4: ") + hStr);
                                  // "Header Consistency Check: PASSED (Output Header == Input Header)"
                                  logger_(std::string("[DEBUG] \xE8\xA1\xA8\xE5\xA4\xB4\xE4\xB8\x80\xE8\x87\xB4\xE6\x80\xA7\xE6\xA3\x80\xE6\xB5\x8B: \xE9\x80\x9A\xE8\xBF\x87 (\xE8\xBE\x93\xE5\x87\xBA\xE8\xA1\xA8\xE5\xA4\xB4 == \xE8\xBE\x93\xE5\x85\xA5\xE8\xA1\xA8\xE5\xA4\xB4)"));
                              }
                         } else {
                             if (logger_ && task.useHeader) {
                                 // "SKIPPING Header Write. Reason: Target sheet not empty or overwrite disabled."
                                 logger_(std::string("[DEBUG] \xE8\xB7\xB3\xE8\xBF\x87\xE8\xA1\xA8\xE5\xA4\xB4\xE5\x86\x99\xE5\x85\xA5\xE3\x80\x82\xE5\x8E\x9F\xE5\x9B\xA0: \xE7\x9B\xAE\xE6\xA0\x87\xE5\xB7\xA5\xE4\xBD\x9C\xE8\xA1\xA8\xE9\x9D\x9E\xE7\xA9\xBA \xE6\x88\x96 \xE6\x9C\xAA\xE5\x90\xAF\xE7\x94\xA8\xE8\xA6\x86\xE7\x9B\x96\xE6\xA8\xA1\xE5\xBC\x8F\xE3\x80\x82"));
                             }
                         }
                         rowIdx++;
                         continue;
                    }
                    
                    bool include = false;
                    
                    // Check Inclusion (AND/OR logic)
                    if (task.rules.empty()) {
                        include = true; // Default include all
                    } else {
                        if (task.ruleLogic == RuleLogic::OR) {
                            include = false;
                            for (const auto& ruleEntry : task.rules) {
                                const Rule* rule = findRule(ruleEntry.ruleId);
                                if (rule && ruleEngine_->evaluateRule(*rule, row, &rules)) {
                                    bool granularExcluded = false;
                                    if (!ruleEntry.excludeRuleIds.empty()) {
                                        for (int exId : ruleEntry.excludeRuleIds) {
                                            const Rule* exRule = findRule(exId);
                                            if (exRule && ruleEngine_->evaluateRule(*exRule, row, &rules)) {
                                                granularExcluded = true;
                                                break;
                                            }
                                        }
                                    }
                                    if (!granularExcluded) {
                                        include = true;
                                        break; 
                                    }
                                }
                            }
                        } else { // AND logic
                            include = true;
                            for (const auto& ruleEntry : task.rules) {
                                const Rule* rule = findRule(ruleEntry.ruleId);
                                bool matched = (rule && ruleEngine_->evaluateRule(*rule, row, &rules));
                                if (matched) {
                                    for (int exId : ruleEntry.excludeRuleIds) {
                                        const Rule* exRule = findRule(exId);
                                        if (exRule && ruleEngine_->evaluateRule(*exRule, row, &rules)) {
                                            matched = false;
                                            break;
                                        }
                                    }
                                }
                                if (!matched) {
                                    include = false;
                                    break; 
                                }
                            }
                        }
                    }

                    // Check Exclusion (Global Exclusion)
                    if (include && !task.excludeRuleIds.empty()) {
                        for (int ruleId : task.excludeRuleIds) {
                            const Rule* rule = findRule(ruleId);
                            if (rule && ruleEngine_->evaluateRule(*rule, row, &rules)) {
                                include = false;
                                break; 
                            }
                        }
                    }

                    if (include) {
                        taskData.push_back(row);
                        processedRowsInChunk++;
                    }
                    rowIdx++;
                }
                
                // Update stats
                result.totalRows += currentData_.size();
                result.processedRows += processedRowsInChunk;
                result.matchedRows += processedRowsInChunk; // Assuming matched = processed for now
                result.deletedRows += (currentData_.size() - processedRowsInChunk);

                if (logger_) {
                     std::string displayTaskName = task.taskName.empty() ? "[Unnamed Task]" : task.taskName;
                     logger_(displayTaskName + " - \xE5\xBD\x93\xE5\x89\x8D\xE5\x9D\x97\xE5\x8C\xB9\xE9\x85\x8D: " + std::to_string(processedRowsInChunk) + " \xE8\xA1\x8C (Total: " + std::to_string(result.processedRows) + ")");
                }

                // Write Output (Chunk)
                bool writeSuccess = false;
                std::string targetFile;
                std::string targetSheet = "Sheet1";

                if (task.outputMode == OutputMode::NONE) {
                    writeSuccess = true;
                } else if (task.outputMode == OutputMode::NEW_SHEET) {
                    targetFile = defaultOutputFile;
                    if (targetFile.empty()) targetFile = inputFile;

                    if (targetFile.empty()) {
                        if (isTaskFirstChunk) {
                            addError("Task '" + task.taskName + "' requires output file for New Sheet mode");
                            result.errors.push_back("No output file specified for New Sheet mode");
                        }
                    } else {
                        targetSheet = task.outputWorkbookName.empty() ? task.taskName : task.outputWorkbookName;
                        
                        std::string ext = targetFile.substr(targetFile.find_last_of(".") + 1);
                        std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
                        if (ext == targetFile) { targetFile += ".xlsx"; ext = "xlsx"; }

                        bool overwrite = isTaskFirstChunk ? task.overwriteSheet : false;

                        if (ext == "xlsx" || ext == "xls") {
                            writeSuccess = sharedQtWriter.appendToSheet(targetFile, taskData, targetSheet, overwrite);
                        } else {
                            CSVExcelWriter csvWriter;
                            writeSuccess = csvWriter.appendToSheet(targetFile, taskData, targetSheet, overwrite);
                        }
                    }
                } else {
                    // NEW_WORKBOOK
                    std::string name = task.outputWorkbookName.empty() ? task.taskName : task.outputWorkbookName;
                    QFileInfo fileInfo(QString::fromStdString(name));
                    if (fileInfo.isRelative()) {
                        std::string baseDir;
                        if (!defaultOutputFile.empty()) {
                            baseDir = QFileInfo(QString::fromStdString(defaultOutputFile)).absolutePath().toStdString();
                        } else {
                            baseDir = QFileInfo(QString::fromStdString(inputFile)).absolutePath().toStdString();
                        }
                        targetFile = baseDir + "/" + name;
                    } else {
                        targetFile = name;
                    }

                    std::string ext = targetFile.substr(targetFile.find_last_of(".") + 1);
                    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
                    if (ext == targetFile) { targetFile += ".xlsx"; ext = "xlsx"; }

                    if (ext == "xlsx" || ext == "xls") {
                        if (isTaskFirstChunk) {
                            writeSuccess = sharedQtWriter.writeExcelFile(targetFile, taskData);
                        } else {
                            writeSuccess = sharedQtWriter.appendToSheet(targetFile, taskData, "Sheet1", false);
                        }
                    } else {
                        CSVExcelWriter csvWriter;
                        if (isTaskFirstChunk) {
                            writeSuccess = csvWriter.writeExcelFile(targetFile, taskData);
                        } else {
                            writeSuccess = csvWriter.appendToSheet(targetFile, taskData, "Sheet1", false);
                        }
                    }
                }

                if (!writeSuccess) {
                     addError("Unable to write task output: " + targetFile);
                     result.errors.push_back("Write failed: " + targetFile);
                }
            } // End Task Loop
            
            if (isFirstChunk && includeHeader && !currentData_.empty()) {
                offset += (currentData_.size() - 1);
            } else {
                offset += currentData_.size();
            }
            isFirstChunk = false;

            // Notify Progress (After Chunk Processed)
             if (progressCallback_) {
                progressCallback_(offset, "Processed " + std::to_string(offset) + " rows...");
            }
            
            // Check if we read less than chunkSize, meaning EOF
            if (currentData_.size() < chunkSize) break;

        } // End Chunk Loop
        
        sharedQtWriter.closeAll();

        // Finalize results for this sheet
        auto sheetEndTime = std::chrono::high_resolution_clock::now();
        double sheetDuration = std::chrono::duration<double, std::milli>(sheetEndTime - sheetStartTime).count() / 1000.0;
        
        for (const auto& task : tasks) {
            if (sheetTaskResults.find(task.id) != sheetTaskResults.end()) {
                sheetTaskResults[task.id].processingTime = sheetDuration; 
                results.push_back(sheetTaskResults[task.id]);
            }
        }
    }

    return results;
}
