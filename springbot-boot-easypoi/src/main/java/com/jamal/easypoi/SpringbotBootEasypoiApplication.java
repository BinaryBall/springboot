package com.jamal.easypoi;

import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.util.PoiPublicUtil;
import com.jamal.easypoi.Entity.UserEntity;
import com.jamal.easypoi.utils.EasyPoiUtils;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.util.ClassUtils;

import java.io.File;
import java.io.IOException;
import java.util.*;

@SpringBootApplication
public class SpringbotBootEasypoiApplication {

    public static void main(String[] args) throws IOException {
//        SpringApplication.run(SpringbotBootEasypoiApplication.class, args);
//        testExportExcel();
//        testImportExcel();
        testExportExcels();

    }

    /**
     * 测试单sheet导出
     * @throws IOException
     */
    public static void testExportExcel() throws IOException {
        List<UserEntity> list = new ArrayList<>();
        int i = 0;
        while (i < 10){
            UserEntity user = new UserEntity();
            user.setId(i+1);
            user.setAge(20+i);
            user.setEmail("abc@163.com");
            user.setName("张三"+i);
            user.setSex(i%2==0?1:2);
            user.setTime(new Date());
            list.add(user);
            i++;
        }
        EasyPoiUtils.exportExcel(UserEntity.class,list,"src/main/resources/excel/","user.xls");
    }

    /**
     * 测试多sheet导出
     * @throws IOException
     */
    public static void testExportExcels() throws IOException {
        List<Map<String, Object>> list = new ArrayList<>();
        for(int n=1;n<4;n++){
            ExportParams exportParams = new ExportParams("用户信息"+n,"用户信息"+n);
            Object entity = UserEntity.class;
            List<UserEntity> data = new ArrayList<>();
            int i = 0;
            while (i < 10){
                UserEntity user = new UserEntity();
                user.setId(i*n+1);
                user.setAge(20+i);
                user.setEmail("abc@163.com");
                user.setName("张三"+i*n);
                user.setSex(i%2==0?1:2);
                user.setTime(new Date());
                data.add(user);
                i++;
            }
            Map<String,Object> map = new HashMap<>();
            map.put("title",exportParams);
            map.put("entity",entity);
            map.put("data",data);
            list.add(map);
        }
        EasyPoiUtils.exportExcel(list,"src/main/resources/excel/","user1.xls");
    }

    /**
     * 测试导入
     */
    public static void testImportExcel(){
        List<UserEntity> list = EasyPoiUtils.importExcel(

                new File("src/main/resources/excel/user.xls"),
                UserEntity.class, new ImportParams());
        list.forEach((user)->{
            System.out.println(user);
        });

    }
}
