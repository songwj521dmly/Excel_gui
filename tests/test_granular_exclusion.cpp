#include "ExcelProcessorCore.h"
#include <iostream>
#include <fstream>
#include <cassert>
#include <vector>
#include <filesystem>

namespace fs = std::filesystem;

void test(bool condition, const std::string& name) {
    if (condition) std::cout << "[PASS] " << name << std::endl;
    else {
        std::cerr << "[FAIL] " << name << std::endl;
        exit(1);
    }
}

int main() {
    // 1. Create test CSV
    std::string inputFile = "test_granular_input.csv";
    std::ofstream out(inputFile);
    out << "Col1,Col2\n";
    out << "A,Keep\n";    // Row 1: Matches A, Not B -> Should be included
    out << "A,Exclude\n"; // Row 2: Matches A, Matches B -> Should be excluded
    out << "B,Keep\n";    // Row 3: Not Matches A -> Should be ignored
    out.close();

    // 2. Setup Processor
    ExcelProcessorCore processor;

    // Rule B: Exclude if Col2 == "Exclude"
    Rule ruleB;
    ruleB.id = 2;
    ruleB.name = "RuleB";
    ruleB.type = RuleType::FILTER;
    ruleB.enabled = true;
    RuleCondition condB;
    condB.column = 2;
    condB.oper = Operator::EQUAL;
    condB.value = std::string("Exclude");
    ruleB.conditions.push_back(condB);
    processor.addRule(ruleB);

    // Rule A: Include if Col1 == "A"
    Rule ruleA;
    ruleA.id = 1;
    ruleA.name = "RuleA";
    ruleA.type = RuleType::FILTER;
    ruleA.enabled = true;
    RuleCondition condA;
    condA.column = 1;
    condA.oper = Operator::EQUAL;
    condA.value = std::string("A");
    ruleA.conditions.push_back(condA);
    processor.addRule(ruleA);

    // Task: Process using Rule A, but exclude Rule B from Rule A's results
    ProcessingTask task;
    task.id = 1;
    task.outputWorkbookName = "test_granular_output.csv";
    task.outputMode = OutputMode::NEW_WORKBOOK;
    
    // New Granular Exclusion Logic
    TaskRuleEntry entryA(ruleA.id);
    entryA.excludeRuleIds.push_back(ruleB.id);
    task.rules.push_back(entryA);

    task.enabled = true;
    processor.addTask(task);

    // 3. Run Process
    std::cout << "Processing tasks..." << std::endl;
    auto results = processor.processTasks(inputFile);

    // 4. Verify Results
    test(results.size() == 1, "Should have 1 result");
    test(results[0].errors.empty(), "Task should succeed (no errors)");
    
    // Check matched rows count
    // Row 1 matches Rule A. Rule B does not match. -> Include
    // Row 2 matches Rule A. Rule B matches. -> Exclude (Granular)
    // Row 3 no match Rule A. -> Ignore
    // So matched rows should be 1.
    std::cout << "Matched rows: " << results[0].matchedRows << std::endl;
    test(results[0].matchedRows == 1, "Should match exactly 1 row (Row 1)");

    // Cleanup
    try {
        fs::remove(inputFile);
        fs::remove("test_granular_output.csv");
    } catch (...) {}

    std::cout << "Granular exclusion test passed!" << std::endl;
    return 0;
}
