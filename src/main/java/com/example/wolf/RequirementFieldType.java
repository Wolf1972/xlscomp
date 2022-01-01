package com.example.wolf;

import java.util.HashMap;

public enum RequirementFieldType {
    RQ_LEVEL,        // A(0): Requirement level
    RQ_NAME,         // B(1): Requirement
    RQ_PRIORITY,     // C(2): Requirement priority
    RQ_DONE,         // D(3): Requirement has realised
    RQ_OTHER,        // E(4): Requirement from other source (mxWeb)
    RQ_NEW_REQ,      // F(5): New requirement flag
    RQ_INTEGRATION,  // G(6): Integration requirement
    RQ_SERVICE,      // H(7): Integration service requirement
    RQ_COMMENT,      // I(8): Comment for requirement
    RQ_LINKED,       // J(9): Linked requirement
    RQ_CURR_STATUS,  // K(10): Current status
    RQ_TYPE,         // L(11): Requirement type
    RQ_SOURCE,       // M(12): Requirement source
    RQ_FOUNDATION,   // N(13): Requirement foundation
    // Private columns
    RQ_VERSION,      // O(14): Plan to realised in version
    RQ_RELEASE,      // P(15): Plan to realized in release
    RQ_QUESTIONS,    // Q(16): Work questions for requirement
    RQ_OTHER_REL,    // R(17): Release in other source (mxWeb)
    RQ_TT,           // S(18): Team track task
    RQ_TRELLO,       // T(19): Trello task
    RQ_PRIMARY,      // U(20): Primary responsible
    RQ_SECONDARY,    // V(21): Secondary responsible
    RQ_RISK,         // W(22): Risk
    RQ_RISK_DESC,    // X(23): Risk description
    // Development & testing columns
    RQ_CONSOLE,      // Y(24): Console
    RQ_CLIENT,       // Z(25): Client part
    RQ_MOBILE,       // AA(26): Mobile application
    RQ_NOTE_NEW,     // AB(27): Note for new application
    RQ_EXIST_OLD,    // AC(28): Has is old application
    RQ_NOTE_OLD      // AD(29): Note for old application
}
