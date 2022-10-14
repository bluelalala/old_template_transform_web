package com.example.old_template_transform_web;

import javax.servlet.MultipartConfigElement;

import org.springframework.boot.web.servlet.MultipartConfigFactory;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.util.unit.DataSize;

@Configuration
public class MulterFileConfig {
    @Bean
    public MultipartConfigElement multipartConfigElement() {
        MultipartConfigFactory factory = new MultipartConfigFactory();
        // 最大文件10M
        factory.setMaxFileSize(DataSize.ofBytes(10485760));
        // 总上传数据总大小100M
        factory.setMaxRequestSize(DataSize.ofBytes(104857600));
        return factory.createMultipartConfig();
    }
}