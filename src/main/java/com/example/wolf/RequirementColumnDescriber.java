package com.example.wolf;

import java.util.HashMap;
import java.util.Map;

public class RequirementColumnDescriber {

    public HashMap<RequirementFieldType, Integer> map = new HashMap<>();
    public HashMap<Integer, RequirementFieldType> reverse = new HashMap<>();

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
}
