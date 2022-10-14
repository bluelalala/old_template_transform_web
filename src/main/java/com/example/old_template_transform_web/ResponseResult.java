package com.example.old_template_transform_web;

import lombok.Data;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

@Data
public class ResponseResult implements Serializable {

    private static final long serialVersionUID = -4923551872770678496L;

    private boolean success;

    private String errorMsg;

    private List<Map<String, String>> data;

    public ResponseResult(boolean success, String errorMsg, List<Map<String, String>> data) {
        this.success = success;
        this.errorMsg = errorMsg;
        this.data = data;
    }

    public static ResponseResult success(List<Map<String, String>> data) {
        return new ResponseResult(true, null, data);
    }

    public static ResponseResult fail(String errorMsg) {
        return new ResponseResult(false, errorMsg, null);
    }
}
