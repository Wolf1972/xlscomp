package com.example.wolf;

import java.util.HashMap;

public class RequirementColumnDescriber {

    public HashMap<RequirementColumnType, Integer> map = new HashMap<>();

    public RequirementColumnDescriber() {
        map.put(RequirementColumnType.RQ_LEVEL,       0);  // A(0): Requirement level
        map.put(RequirementColumnType.RQ_NAME,        1);  // B(1): Requirement
        map.put(RequirementColumnType.RQ_PRIORITY,    2);  // C(2): Requirement priority
        map.put(RequirementColumnType.RQ_DONE,        3);  // D(3): Requirement has realised
        map.put(RequirementColumnType.RQ_OTHER,       4);  // E(4): Requirement from other source (mxWeb)
        map.put(RequirementColumnType.RQ_NEW_REQ,     5);  // F(5): New requirement flag
        map.put(RequirementColumnType.RQ_INTEGRATION, 6);  // G(6): Integration requirement
        map.put(RequirementColumnType.RQ_SERVICE,     7);  // H(7): Integration service requirement
        map.put(RequirementColumnType.RQ_COMMENT,     8);  // I(8): Comment for requirement
        map.put(RequirementColumnType.RQ_LINKED,      9);  // J(9): Linked requirement
        map.put(RequirementColumnType.RQ_CURR_STATUS, 10); // K(10): Current status
        map.put(RequirementColumnType.RQ_TYPE,        11); // L(11): Requirement type
        map.put(RequirementColumnType.RQ_SOURCE,      12); // M(12): Requirement source
        map.put(RequirementColumnType.RQ_FOUNDATION,  13); // N(13): Requirement foundation
        // Private columns
        map.put(RequirementColumnType.RQ_VERSION,     14); // O(14): Plan to realised in version
        map.put(RequirementColumnType.RQ_RELEASE,     15); // P(15): Plan to realized in release
        map.put(RequirementColumnType.RQ_QUESTIONS,   16); // Q(16): Work questions for requirement
        map.put(RequirementColumnType.RQ_OTHER_REL,   17); // R(17): Release in other source (mxWeb)
        map.put(RequirementColumnType.RQ_TT,          18); // S(18): Team track task
        map.put(RequirementColumnType.RQ_TRELLO,      19); // T(19): Trello task
        map.put(RequirementColumnType.RQ_PRIMARY,     20); // U(20): Primary responsible
        map.put(RequirementColumnType.RQ_SECONDARY,   21); // V(21): Secondary responsible
        map.put(RequirementColumnType.RQ_RISK,        22); // W(22): Risk
        map.put(RequirementColumnType.RQ_RISK_DESC,   23); // X(23): Risk description
    }

    public RequirementColumnDescriber(HashMap<RequirementColumnType, Integer> map) {
        this.map = map;
    }

    public Integer get(RequirementColumnType column) {
        return map.get(column);
    }
}
