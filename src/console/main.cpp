#include "ExcelProcessorCore.h"
#include <iostream>
#include <sstream>
#include <iomanip>
#include <chrono>
#include <fstream>
#include <cstdio>
#include <memory>
#include <string>
#include <vector>
#include <cstdlib>
#include <ctime>

class ConsoleExcelProcessor {
public:
    ConsoleExcelProcessor() : processor_(std::make_unique<ExcelProcessorCore>()) {}

    void printHeader() {
        std::cout << "\n";
        std::cout << "==================================================\n";
        std::cout << "    Excel \xE6\x95\xB0\xE6\x8D\xAE\xE5\xA4\x84\xE7\x90\x86\xE7\xB3\xBB\xE7\xBB\x9F - C++ \xE9\xAB\x98\xE6\x80\xA7\xE8\x83\xBD\xE7\x89\x88 (\xE6\x8E\xA7\xE5\x88\xB6\xE5\x8F\xB0\xE7\x89\x88)\n";
        std::cout << "==================================================\n";
        std::cout << "    \xE6\x80\xA7\xE8\x83\xBD\xE7\x9B\xB8\xE6\xAF\x94" "VBA" "\xE6\x8F\x90\xE5\x8D\x87 100-500" "\xE5\x80\x8D\n";
        std::cout << "    \xE6\x94\xAF\xE6\x8C\x81\xE6\x9C\x80\xE5\xA4\x9A" "20" "\xE4\xB8\xAA\xE8\xA7\x84\xE5\x88\x99\xE8\x87\xAA\xE7\x94\xB1\xE7\xBB\x84\xE5\x90\x88\n";
        std::cout << "    \xE7\x89\x88\xE6\x9C\xAC: 1.0.0\n";
        std::cout << "==================================================\n\n";
    }

    void printUsage() {
        std::cout << "\xE7\x94\xA8\xE6\xB3\x95: ConsoleExcelProcessor [\xE9\x80\x89\xE9\xA1\xB9] <\xE5\x8F\x82\xE6\x95\xB0>\n\n";
        std::cout << "\xE9\x80\x89\xE9\xA1\xB9:\n";
        std::cout << "  -h, --help              \xE6\x98\xBE\xE7\xA4\xBA\xE5\xB8\xAE\xE5\x8A\xA9\xE4\xBF\xA1\xE6\x81\xAF\n";
        std::cout << "  -i, --input <file>      \xE8\xBE\x93\xE5\x85\xA5" "Excel" "\xE6\x96\x87\xE4\xBB\xB6 (CSV/XLSX)\n";
        std::cout << "  -o, --output <file>     \xE8\xBE\x93\xE5\x87\xBA\xE6\x96\x87\xE4\xBB\xB6\n"; // Output file
        std::cout << "  -c, --config <file>     \xE9\x85\x8D\xE7\xBD\xAE\xE6\x96\x87\xE4\xBB\xB6\xE8\xB7\xAF\xE5\xBE\x84\n"; // Config file path
        std::cout << "  -r, --rules             \xE5\x88\x97\xE5\x87\xBA\xE6\x89\x80\xE6\x9C\x89\xE8\xA7\x84\xE5\x88\x99\n"; // List all rules
        std::cout << "  -a, --add-rule          \xE4\xBA\xA4\xE4\xBA\x92\xE5\xBC\x8F\xE6\xB7\xBB\xE5\x8A\xA0\xE8\xA7\x84\xE5\x88\x99\n"; // Interactive add rule
        std::cout << "  -d, --delete <ID>       \xE5\x88\xA0\xE9\x99\xA4\xE6\x8C\x87\xE5\xAE\x9A\xE8\xA7\x84\xE5\x88\x99\n"; // Delete rule
        std::cout << "  -p, --preview <num>     \xE9\xA2\x84\xE8\xA7\x88\xE6\x8C\x87\xE5\xAE\x9A\xE6\x95\xB0\xE9\x87\x8F\xE6\x95\xB0\xE6\x8D\xAE\xE8\xA1\x8C\n"; // Preview data
        std::cout << "  -t, --test              \xE8\xBF\x90\xE8\xA1\x8C\xE6\x80\xA7\xE8\x83\xBD\xE6\xB5\x8B\xE8\xAF\x95\n"; // Run perf test
        std::cout << "  -v, --validate          \xE9\xAA\x8C\xE8\xAF\x81\xE8\xA7\x84\xE5\x88\x99\xE9\x85\x8D\xE7\xBD\xAE\n"; // Validate rules
        std::cout << "  -s, --stats             \xE6\x98\xBE\xE7\xA4\xBA\xE6\x80\xA7\xE8\x83\xBD\xE7\xBB\x9F\xE8\xAE\xA1\n"; // Show stats
        std::cout << "  --no-gui                \xE7\xA6\x81\xE7\x94\xA8GUI (\xE7\xBA\xAF\xE5\x91\xBD\xE4\xBB\xA4\xE8\xA1\x8C\xE6\xA8\xA1\xE5\xBC\x8F)\n"; // No GUI
        std::cout << "\n";
        std::cout << "\xE7\xA4\xBA\xE4\xBE\x8B:\n"; // Examples:
        std::cout << "  ConsoleExcelProcessor -i data.csv -o result.csv -c rules.cfg\n";
        std::cout << "  ConsoleExcelProcessor --test 10000\n";
        std::cout << "  ConsoleExcelProcessor --preview 100\n";
    }

    int processFiles(const std::string& inputFile, const std::string& outputFile) {
        printHeader();
        std::cout << "\xE6\xAD\xA3\xE5\x9C\xA8\xE5\xA4\x84\xE7\x90\x86\xE6\x96\x87\xE4\xBB\xB6: " << inputFile << std::endl; // Processing file
        std::cout << "\xE8\xBE\x93\xE5\x87\xBA\xE5\x88\xB0: " << outputFile << std::endl; // Output to

        auto startTime = std::chrono::high_resolution_clock::now();

        // Process data
        auto result = processor_->processExcelFile(inputFile, outputFile);

        auto endTime = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(endTime - startTime);

        std::cout << "\n\xE5\xA4\x84\xE7\x90\x86\xE5\xAE\x8C\xE6\x88\x90\xEF\xBC\x81\n"; // Processing complete!
        std::cout << "---------------------------------------------\n";
        std::cout << "\xE6\x80\xBB\xE8\xA1\x8C\xE6\x95\xB0: " << result.totalRows << std::endl; // Total rows
        std::cout << "\xE5\xA4\x84\xE7\x90\x86\xE8\xA1\x8C\xE6\x95\xB0: " << result.processedRows << std::endl; // Processed rows
        std::cout << "\xE5\x88\xA0\xE9\x99\xA4\xE8\xA1\x8C\xE6\x95\xB0: " << result.deletedRows << std::endl; // Deleted rows
        std::cout << "\xE6\x8B\x86\xE5\x88\x86\xE8\xA1\x8C\xE6\x95\xB0: " << result.splitRows << std::endl; // Split rows
        std::cout << "\xE5\xA4\x84\xE7\x90\x86\xE6\x97\xB6\xE9\x97\xB4: " << std::fixed << std::setprecision(3) << duration.count() << " \xE6\xAF\xAB\xE7\xA7\x92\n"; // Processing time ... ms

        // Show stats
        auto stats = processor_->getPerformanceStats();
        std::cout << "\n\xE6\x80\xA7\xE8\x83\xBD\xE7\xBB\x9F\xE8\xAE\xA1:\n"; // Performance stats
        std::cout << "---------------------------------------------\n";
        std::cout << "\xE5\x86\x85\xE5\xAD\x98\xE4\xBD\xBF\xE7\x94\xA8: " << std::fixed << std::setprecision(2) << (stats.memoryUsed / 1024.0 / 1024.0) << " MB\n"; // Memory used
        std::cout << "\xE5\xB3\xB0\xE5\x80\xBC\xE5\x86\x85\xE5\xAD\x98: " << std::fixed << std::setprecision(2) << (stats.peakMemoryUsed / 1024.0 / 1024.0) << " MB\n"; // Peak memory
        std::cout << "\xE5\xA4\x84\xE7\x90\x86\xE9\x80\x9F\xE5\xBA\xA6: " << std::fixed << std::setprecision(0) << (result.processedRows / (duration.count() / 1000.0)) << " \xE8\xA1\x8C/\xE7\xA7\x92\n"; // Processing speed ... rows/sec

        return result.processedRows > 0 ? 0 : 1;
    }

    void listRules() {
        std::vector<Rule> rules = processor_->getRules();

        if (rules.empty()) {
            std::cout << "\xE6\xB2\xA1\xE6\x9C\x89\xE9\x85\x8D\xE7\xBD\xAE\xE7\x9A\x84\xE8\xA7\x84\xE5\x88\x99\xE3\x80\x82\n"; // No rules configured.
            return;
        }

        std::cout << "\n\xE9\x85\x8D\xE7\xBD\xAE\xE7\x9A\x84\xE8\xA7\x84\xE5\x88\x99:\n"; // Configured rules:
        std::cout << "--------------------------------------------------------\n";
        std::cout << "| ID  | Name        | Type    | Conds  | Enab | Desc                  |\n";
        std::cout << "--------------------------------------------------------\n";

        for (const auto& rule : rules) {
            std::cout << "| " << std::setw(4) << rule.id << " | "
                      << std::setw(12) << rule.name.substr(0, 12) << " | "
                      << std::setw(8) << getRuleTypeString(rule.type) << " | "
                      << std::setw(6) << rule.conditions.size() << " | "
                      << (rule.enabled ? "Yes" : "No ") << " | "
                      << std::setw(20) << rule.description.substr(0, 20) << " |\n";
        }

        std::cout << "--------------------------------------------------------\n";
    }

    void addRule() {
        printHeader();
        std::cout << "\xE6\xB7\xBB\xE5\x8A\xA0\xE6\x96\xB0\xE8\xA7\x84\xE5\x88\x99\n"; // Add new rule
        std::cout << "---------------------------------------------\n";

        Rule newRule;
        newRule.id = generateRuleId();
        newRule.name = "NewRule_" + std::to_string(newRule.id);
        newRule.type = RuleType::FILTER;
        newRule.enabled = true;

        std::cout << "\xE8\xA7\x84\xE5\x88\x99\xE5\x90\x8D\xE7\xA7\xB0 (\xE5\xBD\x93\xE5\x89\x8D: " << newRule.name << "): "; // Rule name (Current: ...)
        std::string ruleName;
        std::getline(std::cin, ruleName);
        if (!ruleName.empty()) {
            newRule.name = ruleName;
        }

        std::cout << "\n\xE8\xA7\x84\xE5\x88\x99\xE7\xB1\xBB\xE5\x9E\x8B (1=Filter, 2=Transform, 3=Delete, 4=Split): "; // Rule type ...
        int typeChoice;
        std::cin >> typeChoice;
        if (typeChoice >= 1 && typeChoice <= 4) {
            newRule.type = static_cast<RuleType>(typeChoice - 1); // Enum is 0-based
        }

        addConditions(newRule);

        processor_->addRule(newRule);
        std::cout << "\xE8\xA7\x84\xE5\x88\x99\xE6\xB7\xBB\xE5\x8A\xA0\xE6\x88\x90\x9F\xEF\xBC\x81ID: " << newRule.id << "\n"; // Rule added successfully!
    }

    void addConditions(Rule& rule) {
        std::cout << "\n\xE6\xB7\xBB\xE5\x8A\xA0\xE6\x9D\xA1\xE4\xBB\xB6 (\xE5\x9B\x9E\xE8\xBD\xA6\xE7\xBB\x93\xE6\x9D\x9F):\n"; // Add conditions (Enter to finish)

        while (true) {
            RuleCondition condition;

            std::cout << "Column: ";
            int column;
            std::cin >> column;
            if (std::cin.fail() || column <= 0) break;
            condition.column = column;

            std::cout << "Operator (1=Eq, 2=Neq, 3=Gt, 4=Lt, 5=Cont, 6=Empty): ";
            int opChoice;
            std::cin >> opChoice;
            if (std::cin.fail() || opChoice < 1 || opChoice > 6) break;
            condition.oper = static_cast<Operator>(opChoice - 1); // Check enum mapping if needed

            std::cout << "Value: ";
            std::string value;
            std::cin >> value;
            if (std::cin.fail()) break;

            // Simple value assignment for console demo
            condition.value = value;
            
            rule.conditions.push_back(condition);
        }
    }

    void previewData(const std::string& inputFile, int maxRows = 100) {
        printHeader();
        std::cout << "\xE9\xA2\x84\xE8\xA7\x88\xE6\x95\xB0\xE6\x8D\xAE: " << inputFile << std::endl; // Preview data
        std::cout << "\xE6\x98\xBE\xE7\xA4\xBA\xE5\x89\x8D " << maxRows << " \xE8\xA1\x8C\xE6\x95\xB0\xE6\x8D\xAE:\n"; // Showing first ... rows

        auto previewData = processor_->getPreviewData(maxRows);

        if (previewData.empty()) {
            std::cout << "\xE6\x97\xA0\xE6\xB3\x95\xE8\xAF\xBB\xE5\x8F\x96\xE6\x96\x87\xE4\xBB\xB6\xE6\x88\x96\xE6\x96\x87\xE4\xBB\xB6\xE4\xB8\xBA\xE7\xA9\xBA\xE3\x80\x82\n"; // Cannot read file or empty
            return;
        }

        std::cout << "--------------------------------------------------\n";
        // Simple preview output
        for (size_t i = 0; i < previewData.size() && i < static_cast<size_t>(maxRows); ++i) {
             const auto& row = previewData[i];
             std::cout << "Row " << (i + 1) << ": ";
             for (const auto& val : row.data) {
                 std::visit([](const auto& v) {
                     using T = std::decay_t<decltype(v)>;
                     if constexpr (std::is_same_v<T, std::string>) std::cout << v << " ";
                     else if constexpr (std::is_arithmetic_v<T>) std::cout << v << " ";
                 }, val);
             }
             std::cout << "\n";
        }
        std::cout << "--------------------------------------------------\n";
    }

    void runPerformanceTest(int dataRows) {
        printHeader();
        std::cout << "\xE6\x80\xA7\xE8\x83\xBD\xE6\xB5\x8B\xE8\xAF\x95 - \xE6\x95\xB0\xE6\x8D\xAE\xE9\x87\x8F: " << dataRows << " \xE8\xA1\x8C\n"; // Perf test - rows
        std::cout << "\xE6\xAD\xA3\xE5\x9C\xA8\xE7\x94\x9F\xE6\x88\x90\xE6\xB5\x8B\xE8\xAF\x95\xE6\x95\xB0\xE6\x8D\xAE...\n"; // Generating test data...

        std::string testFile = "perf_test_" + std::to_string(dataRows) + ".csv";
        generateTestData(testFile, dataRows);

        std::cout << "\xE6\xAD\xA3\xE5\x9C\xA8\xE8\xBF\x90\xE8\xA1\x8C\xE6\x80\xA7\xE8\x83\xBD\xE6\xB5\x8B\xE8\xAF\x95...\n"; // Running perf test...

        auto startTime = std::chrono::high_resolution_clock::now();

        // Load config if needed, or add default rule
        // processor_->loadConfiguration(...);

        auto result = processor_->processExcelFile(testFile, testFile + "_result.csv");

        auto endTime = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(endTime - startTime);

        std::cout << "\n\xE6\x80\xA7\xE8\x83\xBD\xE6\xB5\x8B\xE8\xAF\x95\xE7\xBB\x93\xE6\x9E\x9C:\n"; // Test results:
        std::cout << "---------------------------------------------\n";
        std::cout << "\xE6\x95\xB0\xE6\x8D\xAE\xE9\x87\x8F: " << result.totalRows << " \xE8\xA1\x8C\n"; // Data size
        std::cout << "\xE5\xA4\x84\xE7\x90\x86\xE6\x97\xB6\xE9\x97\xB4: " << std::fixed << std::setprecision(3) << duration.count() << " \xE6\xAF\xAB\xE7\xA7\x92\n"; // Processing time
        std::cout << "\xE5\xA4\x84\xE7\x90\x86\xE9\x80\x9F\xE5\xBA\xA6: " << std::fixed << std::setprecision(0) << (result.processedRows / (duration.count() / 1000.0)) << " \xE8\xA1\x8C/\xE7\xA7\x92\n"; // Speed

        try {
            std::remove(testFile.c_str());
            std::remove((testFile + "_result.csv").c_str());
        } catch (...) {}
    }

    void showStats() {
        auto stats = processor_->getPerformanceStats();
        printHeader();
        std::cout << "\xE6\x80\xA7\xE8\x83\xBD\xE7\xBB\x9F\xE8\xAE\xA1\xE4\xBF\xA1\xE6\x81\xAF:\n"; // Perf stats info
        std::cout << "---------------------------------------------\n";
        std::cout << "Memory: " << std::fixed << std::setprecision(2) << (stats.memoryUsed / 1024.0 / 1024.0) << " MB\n";
        std::cout << "Peak Memory: " << std::fixed << std::setprecision(2) << (stats.peakMemoryUsed / 1024.0 / 1024.0) << " MB\n";
        
        std::vector<Rule> rules = processor_->getRules();
        std::cout << "Rule Count: " << rules.size() << "\n";
    }

    bool validateRules() {
        printHeader();
        std::cout << "\xE9\xAA\x8C\xE8\xAF\x81\xE8\xA7\x84\xE5\x88\x99\xE9\x85\x8D\xE7\xBD\xAE...\n"; // Validating rules...

        std::vector<Rule> rules = processor_->getRules();
        if (rules.empty()) {
            std::cout << "\xE6\xB2\xA1\xE6\x9C\x89\xE8\xA7\x84\xE5\x88\x99\xE9\x9C\x80\xE8\xA6\x81\xE9\xAA\x8C\xE8\xAF\x81\xE3\x80\x82\n"; // No rules to validate
            return true;
        }

        bool allValid = true;
        for (const auto& rule : rules) {
            auto errors = processor_->getValidationMessages();
            if (!errors.empty()) {
                allValid = false;
                std::cout << "Rule ID " << rule.id << " issues:\n";
                for (const auto& error : errors) {
                    std::cout << "  - " << error << "\n";
                }
            }
        }

        if (allValid) {
            std::cout << "[OK] \xE6\x89\x80\xE6\x9C\x89\xE8\xA7\x84\xE5\x88\x99\xE9\xAA\x8C\xE8\xAF\x81\xE9\x80\x9A\xE8\xBF\x87\xEF\xBC\x81\n"; // All rules valid
        } else {
            std::cout << "[FAIL] \xE5\x8F\x91\xE7\x8E\xB0\xE9\xAA\x8C\xE8\xAF\x81\xE9\x97\xAE\xE9\xA2\x98\xEF\xBC\x81\n"; // Validation issues found
        }

        return allValid;
    }

public:
    std::string getRuleTypeString(RuleType type) {
        switch (type) {
            case RuleType::FILTER: return "Filter";
            case RuleType::TRANSFORM: return "Transform";
            case RuleType::DELETE_ROW: return "Delete";
            case RuleType::SPLIT: return "Split";
            default: return "Unknown";
        }
    }

    int generateRuleId() {
        std::vector<Rule> rules = processor_->getRules();
        int maxId = 0;
        for (const auto& rule : rules) {
            if (rule.id > maxId) {
                maxId = rule.id;
            }
        }
        return maxId + 1;
    }

    void generateTestData(const std::string& filename, int rows) {
        std::ofstream file(filename);
        file << "ID,Name,Dept,Amount,Status,Date,City,Note\n";
        for (int i = 1; i <= rows; ++i) {
            file << i << ",Name" << i << ",DeptA," << (100 + i) << ",Active,2024-01-01,CityA,Note" << i << "\n";
        }
    }

    std::unique_ptr<ExcelProcessorCore> processor_;
};

int main(int argc, char* argv[]) {
    ConsoleExcelProcessor app;

    if (argc < 2) {
        app.printUsage();
        return 0;
    }

    std::string inputFile, outputFile, configFile;
    bool previewOnly = false;
    int previewRows = 100;
    bool runTest = false;
    bool showStats = false;
    bool validateOnly = false;

    for (int i = 1; i < argc; ++i) {
        std::string arg = argv[i];

        if (arg == "-h" || arg == "--help") {
            app.printUsage();
            return 0;
        } else if (arg == "-i" || arg == "--input") {
            if (++i < argc) inputFile = argv[i];
        } else if (arg == "-o" || arg == "--output") {
            if (++i < argc) outputFile = argv[i];
        } else if (arg == "-c" || arg == "--config") {
            if (++i < argc) {
                configFile = argv[i];
                if (!app.processor_->loadConfiguration(configFile)) {
                    std::cerr << "Error loading config: " << configFile << std::endl;
                    return 1;
                }
            }
        } else if (arg == "-r" || arg == "--rules") {
            app.listRules();
            return 0;
        } else if (arg == "-a" || arg == "--add-rule") {
            app.addRule();
            return 0;
        } else if ((arg == "-d" || arg == "--delete") && i + 1 < argc) {
            int ruleId = std::stoi(argv[++i]);
            app.processor_->removeRule(ruleId);
            return 0;
        } else if ((arg == "-p" || arg == "--preview") && i + 1 < argc) {
            previewRows = std::stoi(argv[++i]);
            previewOnly = true;
            if (++i < argc) inputFile = argv[i];
        } else if ((arg == "-t" || arg == "--test") && i + 1 < argc) {
            runTest = true;
            int testRows = std::stoi(argv[++i]);
            app.runPerformanceTest(testRows);
            return 0;
        } else if (arg == "-v" || arg == "--validate") {
            validateOnly = true;
        } else if (arg == "-s" || arg == "--stats") {
            showStats = true;
        } else if (arg == "--no-gui") {
            // No-op
        } else if (inputFile.empty()) {
            inputFile = arg;
            outputFile = "result.csv";
        } else if (outputFile.empty()) {
            outputFile = arg;
        }
    }

    if (validateOnly) return app.validateRules() ? 0 : 1;
    if (showStats) { app.showStats(); return 0; }
    if (previewOnly) { app.previewData(inputFile, previewRows); return 0; }
    if (runTest) return 0;

    std::ifstream inputCheck(inputFile);
    if (!inputCheck.good()) {
        std::cerr << "Input file not found: " << inputFile << std::endl;
        return 1;
    }
    inputCheck.close();

    return app.processFiles(inputFile, outputFile);
}
