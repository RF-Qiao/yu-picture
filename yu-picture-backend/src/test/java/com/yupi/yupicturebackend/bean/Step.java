package com.yupi.yupicturebackend.bean;

import lombok.Data;
import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.List;

/**
 * 工步类
 */
@Data
@Setter
@Getter
public class Step {
    private String stepId;           // 工步ID

    
    private List<Component> components = new ArrayList<>();       // 关联的部件
    private List<ProcessResource> resources = new ArrayList<>();  // 关联的工艺资源


    // 添加部件
    public void addComponent(Component component) {
        this.components.add(component);
    }
    
    // 添加工艺资源
    public void addResource(ProcessResource resource) {
        this.resources.add(resource);
    }
    
    // 其他构造方法、getter和setter省略...
}