#pragma once
#include <vector>
#include <string>
#include <memory>
#include <map>
#include <variant>
#include <functional>
#include <chrono>
#include <thread>
#include <mutex>
#include <set>
#include <ctime>
#include <iomanip>

// Rule type enumeration
enum class RuleType {
    FILTER,     // Filter rule
    TRANSFORM,   // Transform rule
    DELETE_ROW,      // Delete rule
    SPLIT        // Split rule
};

// Operator enumeration
enum class Operator {
    EQUAL,         // =
    NOT_EQUAL,     // <>
    GREATER,        // >
    LESS,           // <
    GREATER_EQUAL, // >=
    LESS_EQUAL,    // <=
    CONTAINS,      // Contains
    NOT_CONTAINS,   // Does not contain
    EMPTY,          // Is empty
    NOT_EMPTY,      // Is not empty
    STARTS_WITH,    // Starts with
    ENDS_WITH,      // Ends with
    REGEX,          // Regular expression
    CUSTOM          // Custom function
};

// Data type enumeration
enum class DataType {
    STRING,
    INTEGER,
    DOUBLE,
    BOOLEAN,
    DATE
};

// Rule condition structure
struct RuleCondition {
    int column;                    // Column number (1-based)
    Operator oper;                   // Operator
    std::variant<std::string, int, double, bool> value;  // Condition value
    bool case_sensitive = false;    // Case sensitive
    DataType type = DataType::STRING; // Data type

    RuleCondition() : column(0), oper(Operator::EQUAL) {}
};

// Rule logic enumeration
enum class RuleLogic {
    AND,    // All conditions must be met
    OR      // Any condition must be met
};

// Output mode enumeration
enum class OutputMode {
    NONE,           // Do not output independently (process in place or skip)
    NEW_WORKBOOK,   // Save to new workbook
    NEW_SHEET       // Save to new sheet in target workbook
};

// Rule structure
struct Rule {
    int id;                         // Rule ID
    std::string name;               // Rule name
    RuleType type;                   // Rule type
    RuleLogic logic = RuleLogic::AND; // Rule logic
    std::vector<RuleCondition> conditions;  // Condition list
    std::string targetSheet;        // Target sheet (for split rules)
    std::map<int, std::string> transformActions;  // Transform actions (column -> action)
    bool enabled = true;             // Whether enabled
    std::string description;        // Description
    int priority = 0;               // Priority (lower number = higher priority)

    // Output configuration
    OutputMode outputMode = OutputMode::NONE;
    std::string outputName;         // Output name (sheet or workbook name)
    
    // Exclusion configuration REMOVED
    // std::vector<int> excludedRuleIds; // IDs of rules to exclude matches from

    Rule() : id(0), type(RuleType::FILTER) {}
};

// Task Rule Entry for Granular Control
struct TaskRuleEntry {
    int ruleId;
    std::vector<int> excludeRuleIds; // Rules to exclude specifically from this rule's results
    
    TaskRuleEntry() : ruleId(0) {}
    TaskRuleEntry(int id) : ruleId(id) {}
};

// Processing task structure
struct ProcessingTask {
    int id = 0;
    std::string taskName;
    std::string outputWorkbookName; // Name of output workbook or sheet
    std::string inputFilenamePattern; // Pattern to match input filename (empty = all)
    std::string inputSheetName;     // Specific input sheet name (empty = all/current)
    
    std::vector<TaskRuleEntry> rules; // Rules to include with granular exclusions
    
    std::vector<int> excludeRuleIds; // Global Rules to exclude (applies to result of all included rules)
    RuleLogic ruleLogic = RuleLogic::OR; // Logic for combining include rules (AND/OR)
    OutputMode outputMode = OutputMode::NEW_WORKBOOK;
    bool enabled = true;
    bool overwriteSheet = false; // Overwrite existing sheet instead of appending
    bool useHeader = false;      // Copy header from input source (row 1) to output
    
    // Legacy support (to be removed or converted)
    // std::vector<int> includeRuleIds; 
    
    ProcessingTask() = default;
};

// Data row structure
struct DataRow {
    std::vector<std::variant<std::string, int, double, bool, std::tm>> data;
    std::string sheetName;
    int rowNumber;
    bool isValid = true;

    DataRow() : rowNumber(0) {}
    DataRow(int cols) : rowNumber(0) { data.resize(cols); }
};

// Processing result structure
struct ProcessingResult {
    int totalRows = 0;
    int processedRows = 0;
    int matchedRows = 0;
    int deletedRows = 0;
    int splitRows = 0;
    double processingTime = 0.0;  // milliseconds
    std::vector<std::string> errors;
    std::vector<std::string> warnings;

    ProcessingResult() = default;
};

// Performance statistics
struct PerformanceStats {
    size_t memoryUsed = 0;
    size_t peakMemoryUsed = 0;
    double cpuTime = 0.0;
    double ioTime = 0.0;
    std::chrono::high_resolution_clock::time_point startTime;
    std::chrono::high_resolution_clock::time_point endTime;
};

// Rule combination strategy
enum class CombinationStrategy {
    AND,           // All conditions must be met
    OR,            // Any condition can be met
    CUSTOM         // Custom combination
};

// Forward declarations
class ExcelReader;
class ExcelWriter;
class RuleEngine;
class DataProcessor;

// Core processing engine
class ExcelProcessorCore {
public:
    ExcelProcessorCore();
    ~ExcelProcessorCore();

    // Configuration management
    bool loadConfiguration(const std::string& configFile);
    bool saveConfiguration(const std::string& configFile) const;
    void addRule(const Rule& rule);
    void removeRule(int ruleId);
    void updateRule(const Rule& rule);
    std::vector<Rule> getRules() const;
    Rule getRule(int ruleId) const;

    // Task management
    void addTask(const ProcessingTask& task);
    void removeTask(int taskId);
    void updateTask(const ProcessingTask& task);
    std::vector<ProcessingTask> getTasks() const;
    ProcessingTask getTask(int taskId) const;

    // Rule combination management
    void createRuleCombination(int comboId, const std::vector<int>& ruleIds,
                               CombinationStrategy strategy = CombinationStrategy::AND);
    void removeRuleCombination(int comboId);
    std::map<int, std::vector<int>> getRuleCombinations() const;

    // Data processing
    ProcessingResult processExcelFile(const std::string& inputFile,
                                     const std::string& outputFile,
                                     const std::string& sheetName = "");
    std::vector<ProcessingResult> processTasks(const std::string& inputFile, const std::string& defaultOutputFile = "", const std::string& sheetName = "");
    ProcessingResult processTask(int taskId);
    ProcessingResult processMemoryData(std::vector<DataRow>& data);

    // Callbacks
    void setLogger(std::function<void(const std::string&)> logger);
    void setProgressCallback(std::function<void(int, const std::string&)> callback);

    // Data loading
    bool loadFile(const std::string& filename, const std::string& sheetName = "", int maxRows = 0);
    std::vector<std::string> getSheetNames(const std::string& filename);

    // Preview functionality
    bool previewResults(const std::string& inputFile, int maxPreviewRows = 100);
    std::vector<DataRow> getPreviewData(int maxRows = 100) const;
    std::vector<DataRow> getProcessedPreviewData(int maxRows = 100);
    std::vector<DataRow> getProcessedPreviewData(const std::vector<int>& ruleIds, int maxRows = 100);
    std::vector<DataRow> getTaskPreviewData(int taskId, int maxRows = 100);

    // Performance and statistics
    PerformanceStats getPerformanceStats() const;
    void resetStats();

    // Error handling
    std::vector<std::string> getErrors() const;
    std::vector<std::string> getWarnings() const;
    void clearErrors();
    void clearWarnings();

    // Data validation
    bool validateRules() const;
    std::vector<std::string> getValidationMessages() const;

    // Import/export rules
    bool exportRules(const std::string& filename) const;
    bool importRules(const std::string& filename);

private:
    std::unique_ptr<ExcelReader> reader_;
    std::unique_ptr<ExcelWriter> writer_;
    std::unique_ptr<RuleEngine> ruleEngine_;
    std::unique_ptr<DataProcessor> dataProcessor_;

    std::vector<Rule> rules_;
    std::vector<ProcessingTask> tasks_;
    std::map<int, std::vector<int>> ruleCombinations_;
    std::vector<DataRow> currentData_;
    mutable std::vector<std::string> errors_;
    mutable std::vector<std::string> warnings_;
    PerformanceStats stats_;

    mutable std::mutex dataMutex_;
    std::vector<ProcessingResult> processTasksInternal(const std::vector<ProcessingTask>& tasksToProcess, const std::string& overrideInputFile, const std::string& overrideOutputFile, const std::string& overrideSheetName);

    mutable std::recursive_mutex rulesMutex_;

    // Callbacks
    std::function<void(const std::string&)> logger_;
    std::function<void(int, const std::string&)> progressCallback_;
    
    std::string loadedConfigFilename_;

    // Internal methods
    void updatePerformanceStats(double processingTime);
    void addError(const std::string& error) const;
    void addWarning(const std::string& warning) const;
    bool validateRule(const Rule& rule) const;
    bool applyRuleCombination(std::vector<DataRow>& data,
                                const std::vector<int>& ruleIds,
                                CombinationStrategy strategy);
    Rule parseRuleLine(const std::string& line);
};

// Rule engine interface
class RuleEngine {
public:
    virtual ~RuleEngine() = default;
    virtual bool evaluateRule(const Rule& rule, const DataRow& row, const std::vector<Rule>* allRules = nullptr) const = 0;
    virtual bool evaluateCondition(const RuleCondition& condition,
                                   const std::variant<std::string, int, double, bool, std::tm>& value) const = 0;
    virtual std::vector<std::string> getValidationErrors(const Rule& rule) const = 0;
};

// Data reader interface
class ExcelReader {
public:
    virtual ~ExcelReader() = default;
    virtual bool readExcelFile(const std::string& filename, std::vector<DataRow>& data, const std::string& sheetName = "", int maxRows = 0, int offset = 0, bool includeHeader = false) = 0;
    virtual bool getSheetNames(const std::string& filename, std::vector<std::string>& sheetNames) const = 0;
    virtual int getRowCount(const std::string& sheetName) const = 0;
    virtual int getColumnCount(const std::string& sheetName) const = 0;
};

enum class ExcelWriterType {
    ActiveQt,
    CSV
};

// Data writer interface
class ExcelWriter {
public:
    virtual ~ExcelWriter() = default;
    virtual ExcelWriterType getType() const = 0;
    virtual void closeAll() {} // Close persistent resources
    virtual bool writeExcelFile(const std::string& filename,
                               const std::vector<DataRow>& data,
                               const std::string& sheetName = "Sheet1") = 0;
    virtual bool writeMultipleSheets(const std::string& filename,
                                      const std::map<std::string, std::vector<DataRow>>& sheetData) = 0;
    virtual bool appendToSheet(const std::string& filename,
                                const std::vector<DataRow>& data,
                                const std::string& sheetName,
                                bool overwrite = false) = 0;
};

// Data processor interface
class DataProcessor {
public:
    virtual ~DataProcessor() = default;
    virtual ProcessingResult processData(std::vector<DataRow>& data,
                                        const std::vector<Rule>& rules,
                                        const std::map<int, std::vector<int>>& combinations) = 0;
    virtual bool optimizeProcessing(std::vector<DataRow>& data) = 0;
    virtual bool validateData(const std::vector<DataRow>& data,
                              std::vector<std::string>& errors) const = 0;
    virtual void setProgressCallback(std::function<void(int, const std::string&)> callback) = 0;
};

// Factory functions
std::unique_ptr<RuleEngine> createRuleEngine();