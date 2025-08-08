package com.yupi.yupicturebackend.bean;

import lombok.Data;
import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;

@Data
@Setter
@Getter
public class Process {
    private String id;        // 工序ID


    private List<Step> steps = new ArrayList<>();  // 包含的工步
    // 添加工步
    public void addStep(Step step) {
        this.steps.add(step);
    }
}