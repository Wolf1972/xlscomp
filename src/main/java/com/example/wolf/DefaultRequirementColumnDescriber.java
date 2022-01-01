package com.example.wolf;

import java.util.HashMap;

public class DefaultRequirementColumnDescriber extends RequirementColumnDescriber {

    private static HashMap<RequirementFieldType, Integer> describer = new HashMap<>();

    static {
        describer.put(RequirementFieldType.RQ_LEVEL,       0);  // A(0): Requirement level
        describer.put(RequirementFieldType.RQ_NAME,        1);  // B(1): Requirement
        describer.put(RequirementFieldType.RQ_PRIORITY,    2);  // C(2): Requirement priority
        describer.put(RequirementFieldType.RQ_DONE,        3);  // D(3): Requirement has realised
        describer.put(RequirementFieldType.RQ_OTHER,       4);  // E(4): Requirement from other source (mxWeb)
        describer.put(RequirementFieldType.RQ_NEW_REQ,     5);  // F(5): New requirement flag
        describer.put(RequirementFieldType.RQ_INTEGRATION, 6);  // G(6): Integration requirement
        describer.put(RequirementFieldType.RQ_SERVICE,     7);  // H(7): Integration service requirement
        describer.put(RequirementFieldType.RQ_COMMENT,     8);  // I(8): Comment for requirement
        describer.put(RequirementFieldType.RQ_LINKED,      9);  // J(9): Linked requirement
        describer.put(RequirementFieldType.RQ_CURR_STATUS, 10); // K(10): Current status
        describer.put(RequirementFieldType.RQ_TYPE,        11); // L(11): Requirement type
        describer.put(RequirementFieldType.RQ_SOURCE,      12); // M(12): Requirement source
        describer.put(RequirementFieldType.RQ_FOUNDATION,  13); // N(13): Requirement foundation
        // Private columns
        describer.put(RequirementFieldType.RQ_VERSION,     14); // O(14): Plan to realised in version
        describer.put(RequirementFieldType.RQ_RELEASE,     15); // P(15): Plan to realized in release
        describer.put(RequirementFieldType.RQ_QUESTIONS,   16); // Q(16): Work questions for requirement
        describer.put(RequirementFieldType.RQ_OTHER_REL,   17); // R(17): Release in other source (mxWeb)
        describer.put(RequirementFieldType.RQ_TT,          18); // S(18): Team track task
        describer.put(RequirementFieldType.RQ_TRELLO,      19); // T(19): Trello task
        describer.put(RequirementFieldType.RQ_PRIMARY,     20); // U(20): Primary responsible
        describer.put(RequirementFieldType.RQ_SECONDARY,   21); // V(21): Secondary responsible
        describer.put(RequirementFieldType.RQ_RISK,        22); // W(22): Risk
        describer.put(RequirementFieldType.RQ_RISK_DESC,   23); // X(23): Risk description
        // Development & testing columns
        describer.put(RequirementFieldType.RQ_CONSOLE,     24); // Y(24): Console
        describer.put(RequirementFieldType.RQ_CLIENT,      25); // Z(25): Client part
        describer.put(RequirementFieldType.RQ_MOBILE,      26); // AA(26): Mobile application
        describer.put(RequirementFieldType.RQ_NOTE_NEW,    27); // AB(27): Note for new application
        describer.put(RequirementFieldType.RQ_EXIST_OLD,   28); // AC(28): Has is old application
        describer.put(RequirementFieldType.RQ_NOTE_OLD,    29); // AD(29): Note for old application
    }

    public DefaultRequirementColumnDescriber() {
        super(describer);
    }
}
