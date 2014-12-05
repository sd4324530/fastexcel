package com.github.sd4324530.fastexcel.entity;

import com.github.sd4324530.fastexcel.annotation.MapperCell;

/**
 * @author peiyu
 */
public class MyTest {

    @MapperCell(cellName = "名称")
    private String name;

    @MapperCell(cellName = "联系电话")
    private String phone;

    @MapperCell(cellName = "地址")
    private String address;

    @MapperCell(cellName = "一级分类ID")
    private int type;

    @MapperCell(cellName = "经度")
    private double lat;

    @Override
    public String toString() {
        return "MyTest{" +
                "name='" + name + '\'' +
                ", phone='" + phone + '\'' +
                ", address='" + address + '\'' +
                ", type=" + type +
                ", lat=" + lat +
                '}';
    }
}
