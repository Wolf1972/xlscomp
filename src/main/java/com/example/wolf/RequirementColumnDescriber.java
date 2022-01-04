package com.example.wolf;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class RequirementColumnDescriber {

    public HashMap<RequirementFieldType, Integer> map = new HashMap<>(); // Describer: fieldtype - column
    public HashMap<Integer, RequirementFieldType> reverse = new HashMap<>(); // Reverse describer: column - fieldType

    public RequirementColumnDescriber(HashMap<RequirementFieldType, Integer> map) {
        this.map = map;
        for (Map.Entry<RequirementFieldType, Integer> item : map.entrySet()) {
            RequirementFieldType key = item.getKey();
            Integer value = item.getValue();
            reverse.put(value, key);
        }
    }

    public Integer getColumn(RequirementFieldType column) {
        return map.get(column);
    }

    public RequirementFieldType getField(int column) {
        return reverse.get(column);
    }

    /**
     * Builds mapping between two structures: this and specified (target)
     * @param targetDescriber - target describer
     * @return - array with mapping: list of all this columns with indexes in target describer (or null, if column doesn't exist in target describer)
     */
    public ArrayList<Integer> structureMappingBuilder(RequirementColumnDescriber targetDescriber) {
        ArrayList<Integer> mapping = new ArrayList<>();
        // TODO
        return mapping;
    }
}
