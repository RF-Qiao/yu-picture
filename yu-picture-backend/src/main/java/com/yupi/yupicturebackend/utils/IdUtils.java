package com.yupi.yupicturebackend.utils;

import java.util.UUID;


public class IdUtils {

    public static String generateId() {
        return UUID.randomUUID().toString().replaceAll("-", "");
    }
}
