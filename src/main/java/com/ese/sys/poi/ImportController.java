package com.ese.sys.poi;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

@Controller
public class ImportController {

    @RequestMapping("/import")
    public String importExcel(@RequestParam MultipartFile upfile) {

        try {
            // 获取上传文件的输入流对象
            InputStream inputStream = upfile.getInputStream();
            // 获取上传文件的文件名
            String filename = upfile.getOriginalFilename();

            List<Map<String, Object>> list =ExcelUtils.readExcel(filename,inputStream);

            ObjectMapper objectMapper = new ObjectMapper();
            String jsonStr = objectMapper.writeValueAsString(list);

            List<User> userList = objectMapper.readValue(jsonStr, new TypeReference<List<User>>() {
            });

        } catch (IOException e) {
            e.printStackTrace();
        }


        return "redirect:/index.html";
    }
}
