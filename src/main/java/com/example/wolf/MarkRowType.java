package com.example.wolf;

public enum MarkRowType {
    PARENT,  // Parent rows - mark as 50% gray
    ADDED,   // Added rows - mark as bold red
    DELETED, // Deleted rows - mark as strikeout
    CHANGED  // Changed columns - mark as bold blue
}
