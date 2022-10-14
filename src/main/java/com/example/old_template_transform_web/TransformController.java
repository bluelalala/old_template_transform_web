package com.example.old_template_transform_web;

import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 文件上传测试类
 */
@Controller
@RequestMapping("/transform")
public class TransformController {

    @RequestMapping(value = "/index")
    public String index() {
        return "index";
    }

    @ResponseBody
    @RequestMapping(value = "/upload")
    public ResponseResult upload(@RequestParam("file_upload") MultipartFile[] files) throws IOException {
        // 新模板输出文件夹
        String basePath = System.getProperty("user.dir") + File.separator + "newTemplates";
        Date date = new Date();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-hh-mm-ss");
        String datePath = basePath + File.separator + dateFormat.format(date);
        File dateFile = new File(datePath);
        if (!dateFile.exists()) {
            dateFile.mkdirs();
        }

        List<Map<String, String>> list = new ArrayList<>();
        for (int i = 0; i < files.length; i++) {
            MultipartFile file = files[i];
            if (file.isEmpty()) {
                return ResponseResult.fail(file.getOriginalFilename() + "文件为空！");
            }
            byte[] bytes = TransformUtil.procCustomDocument(file.getInputStream());
            String path = datePath + File.separator + file.getOriginalFilename() + "x";
            TransformUtil.procEditAbleRange(bytes, path);

            Map<String, String> map = new HashMap<>();
            map.put("name", file.getOriginalFilename() + "x");
            map.put("folder", dateFormat.format(date));
            list.add(map);
        }
        return ResponseResult.success(list);
    }

    @ResponseBody
    @RequestMapping("/download")
    public ResponseEntity<InputStreamResource> download(@RequestParam("name") String name, @RequestParam("folder") String folder) throws IOException {
        String basePath = System.getProperty("user.dir") + File.separator + "newTemplates";
        String path = basePath + File.separator + folder + File.separator + name;
        FileSystemResource file = new FileSystemResource(path);
        HttpHeaders headers = new HttpHeaders();
        headers.add("Cache-Control", "no-cache, no-store, must-revalidate");
        String fileName = URLEncoder.encode(file.getFilename(), StandardCharsets.UTF_8.toString());
        headers.add("Content-Disposition", "attachment; filename=" + fileName);
        headers.add("Pragma", "no-cache");
        headers.add("Expires", "0");

        return ResponseEntity
                .ok()
                .headers(headers)
                .contentLength(file.contentLength())
                .contentType(MediaType.parseMediaType("application/octet-stream"))
                .body(new InputStreamResource(file.getInputStream()));
    }
}