package com.example.wolf;

import java.util.HashMap;

public class AlternateRequirementColumnDescriber extends RequirementColumnDescriber {

    private static HashMap<RequirementFieldType, Integer> describer = new HashMap<>();

    static {
        describer.put(RequirementFieldType.RQ_LEVEL,       0);  // A(0): Requirement level
        describer.put(RequirementFieldType.RQ_NAME,        1);  // B(1): Requirement
        describer.put(RequirementFieldType.RQ_PRIORITY,    2);  // C(2): Requirement priority
        describer.put(RequirementFieldType.RQ_CONSOLE,     3);  // D(3): Console
        describer.put(RequirementFieldType.RQ_CLIENT,      4);  // E(4): Client part
        describer.put(RequirementFieldType.RQ_MOBILE,      5);  // F(5): Mobile application
        describer.put(RequirementFieldType.RQ_NOTE_NEW,    6);  // G(6): Note for new application
        describer.put(RequirementFieldType.RQ_OTHER,       7);  // H(7): Requirement from other source (mxWeb)
        describer.put(RequirementFieldType.RQ_COMMENT,     8);  // I(8): Comment for requirement
        describer.put(RequirementFieldType.RQ_EXIST_OLD,   9);  // J(9): Has is old application
        describer.put(RequirementFieldType.RQ_NOTE_OLD,    10); // K(10): Note for old application

        describer.put(RequirementFieldType.RQ_TRELLO,      12); // M(12): Trello task
        describer.put(RequirementFieldType.RQ_PRIMARY,     13); // N(13): Primary responsible
        describer.put(RequirementFieldType.RQ_SECONDARY,   14); // O(14): Secondary responsible
        describer.put(RequirementFieldType.RQ_TT,          15); // P(15): Team track task
    }

    public AlternateRequirementColumnDescriber() {
        super(describer);
    }
}
