package com.jxls.plus.transform.jexcel;

import com.jxls.plus.common.Context;

import java.util.Map;

/**
 * @author Leonid Vysochyn
 */
public class JexcelContext extends Context {
    public static final String JEXCEL_OBJECT_KEY = "util";

    public JexcelContext() {
        varMap.put(JEXCEL_OBJECT_KEY, new JexcelUtil());
    }

    public JexcelContext(Map<String, Object> map) {
        super(map);
        varMap.put(JEXCEL_OBJECT_KEY, new JexcelUtil());
    }
}
