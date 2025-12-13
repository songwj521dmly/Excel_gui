#include "ExcelProcessorCore.h"
#include <cassert>
#include <iostream>
#include <variant>

// Helper to print test results
void test(bool result, const std::string& name) {
    if (result) {
        std::cout << "[PASS] " << name << std::endl;
    } else {
        std::cerr << "[FAIL] " << name << std::endl;
        exit(1);
    }
}

int main() {
    auto engine = createRuleEngine();
    
    // Test 1: Double value with CONTAINS string operator
    {
        RuleCondition cond;
        cond.oper = Operator::CONTAINS;
        cond.value = std::string("5.0"); // Rule value is string "5.0"
        cond.type = DataType::STRING; // User treats it as string rule
        
        // Excel cell value is double 5.0
        std::variant<std::string, int, double, bool, std::tm> val = 5.0;
        
        bool result = engine->evaluateCondition(cond, val);
        test(result, "Double 5.0 CONTAINS '5.0'");
    }

    // Test 2: Double value with NOT_CONTAINS string operator
    {
        RuleCondition cond;
        cond.oper = Operator::NOT_CONTAINS;
        cond.value = std::string("15.0");
        cond.type = DataType::STRING;
        
        // Excel cell value is double 5.0 (should NOT contain 15.0)
        std::variant<std::string, int, double, bool, std::tm> val = 5.0;
        
        bool result = engine->evaluateCondition(cond, val);
        test(result, "Double 5.0 NOT_CONTAINS '15.0'");
    }

    // Test 3: Double value 15.0 with NOT_CONTAINS string operator (should fail)
    {
        RuleCondition cond;
        cond.oper = Operator::NOT_CONTAINS;
        cond.value = std::string("15.0");
        cond.type = DataType::STRING;
        
        // Excel cell value is double 15.0
        std::variant<std::string, int, double, bool, std::tm> val = 15.0;
        
        bool result = engine->evaluateCondition(cond, val);
        test(!result, "Double 15.0 NOT_CONTAINS '15.0' should be false");
    }

    // Test 4: Int value with CONTAINS string operator
    {
        RuleCondition cond;
        cond.oper = Operator::CONTAINS;
        cond.value = std::string("5");
        cond.type = DataType::STRING;
        
        // Excel cell value is int 5
        std::variant<std::string, int, double, bool, std::tm> val = 5;
        
        bool result = engine->evaluateCondition(cond, val);
        test(result, "Int 5 CONTAINS '5'");
    }

    // Test 5: Double value with STARTS_WITH
    {
        RuleCondition cond;
        cond.oper = Operator::STARTS_WITH;
        cond.value = std::string("5.");
        cond.type = DataType::STRING;
        
        // Excel cell value is double 5.0
        std::variant<std::string, int, double, bool, std::tm> val = 5.0;
        
        bool result = engine->evaluateCondition(cond, val);
        test(result, "Double 5.0 STARTS_WITH '5.'");
    }

    std::cout << "All tests passed!" << std::endl;
    return 0;
}
