package com.bochao.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 编码规范对应实体类
 * */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class DeviceInfo {

    /**
     * 区域
     * */
    private String area;

    /**
     * 设备名称
     * */
    private String deviceName;

    /**
     * 部件
     * */
    private String parts;

    /**
     * 编码
     * */
    private String code;

    /**
     * 模型名称
     * */
    private String modelName;
}
