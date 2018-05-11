package net.matedata.utils;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.SortedMap;
import java.util.TreeMap;

public class MapUtils
{
    private MapUtils()
    {
    }

    public static SortedMap sortKey(Map<String, Object> map)
    {
        List<String> keyList = new ArrayList<>(map.keySet());
        Collections.sort(keyList);
        SortedMap<String, Object> sortedMap = new TreeMap<>();
        for(String k : keyList)
        {
            sortedMap.put(k, map.get(k));
        }
        return sortedMap;
    }
}
