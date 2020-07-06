package com.github.kdm1jkm.easypoint;

import javafx.fxml.LoadException;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.json.simple.JSONArray;

import javax.naming.NameNotFoundException;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class Manager {
    private static final String ZIP_PPTX_NAME = "template.pptx";
    private static final String ZIP_JSON_NAME = "data.json";
    private static final int BUFFER_SIZE = 2048;
    public final List<TemplateSlide> templateSlides = new ArrayList<>();
    public final List<ModifiedSlide> modifiedSlides = new ArrayList<>();
    private XMLSlideShow templateSlideShow = null;

    Manager(File file) throws IOException {
        String fileName = file.getName();
        String[] arr = fileName.split("\\.");
        switch (arr[arr.length - 1]) {
            case "pptx":
                templateSlideShow = new XMLSlideShow(new FileInputStream(file));
                break;
            case "eptx":
                ZipInputStream zipInputStream = new ZipInputStream(new FileInputStream(file));
                getFromZip(zipInputStream);
                break;
        }

        if (templateSlideShow == null) {
            throw new LoadException("File Loading Error.");
        }

        for (XSLFSlide slide : templateSlideShow.getSlides()) {
            templateSlides.add(new TemplateSlide(null, slide));
        }

    }

    private void getFromZip(ZipInputStream zipInputStream) throws IOException {
        ZipEntry zipEntry;
        // 압축파일 속 각 파일에 대해
        while ((zipEntry = zipInputStream.getNextEntry()) != null) {

            // 디렉토리는 건너뜀
            if (zipEntry.isDirectory())
                continue;

            // 파일명
            String fileName = zipEntry.getName();

            // 스트림을 옮겨줌
            PipedInputStream pipedInputStream = new PipedInputStream();
            PipedOutputStream pipedOutputStream = new PipedOutputStream(pipedInputStream);

            // 데이터 저장?공간
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();

            // 읽어서 스트림에 넣기
            int size;
            byte[] buffer = new byte[BUFFER_SIZE];
            while ((size = zipInputStream.read(buffer, 0, BUFFER_SIZE)) != -1) {
                byteArrayOutputStream.write(buffer, 0, size);
            }

            // pos write해주는 부분
            new Thread(
                    () -> {
                        try {
                            byteArrayOutputStream.writeTo(pipedOutputStream);
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
            ).start();

            // 파일명에 따라 다르게 행동
            switch (fileName) {
                case ZIP_PPTX_NAME:
                    templateSlideShow = new XMLSlideShow(pipedInputStream);
                    break;
                case ZIP_JSON_NAME:
                    StringBuilder jsonStr = new StringBuilder();
                    byte[] bytes = new byte[BUFFER_SIZE];
                    while ((size = pipedInputStream.read(bytes, 0, BUFFER_SIZE)) != -1) {
                        for (int i = 0; i < size; i++) {
                            jsonStr.append((char) bytes[i]);
                        }
                        if (size < BUFFER_SIZE) break;
                    }
                    // TODO: jsonStr 가져온걸로 parsing 진행 후 List<ModifiedSlide>제작
                    break;
            }

        }
        zipInputStream.closeEntry();
    }

    private void parse() {
        for (XSLFSlide slide : templateSlideShow.getSlides()) {
            templateSlides.add(new TemplateSlide(null, slide));
        }
    }

    public void append(int index) {
        TemplateSlide templateSlide = templateSlides.get(index);

        modifiedSlides.add(new ModifiedSlide(templateSlide));
    }

    public void append(String name) throws NameNotFoundException {
        for (int i = 0; i < templateSlides.size(); i++) {
            if (templateSlides.get(i).name.equals(name)) {
                append(i);
                return;
            }
        }

        throw new NameNotFoundException(String.format("Name %s is not found in TemplateSlides", name));
    }

    public void export(File file) throws IOException {
        // 폴더 생성
        file.getAbsoluteFile().getParentFile().mkdirs();

        FileOutputStream out = new FileOutputStream(file);

        XMLSlideShow newSlideshow = new XMLSlideShow(templateSlideShow.getPackage());

        // 슬라이드 모두 삭제
        while (newSlideshow.getSlides().size() != 0) newSlideshow.removeSlide(0);

        for (ModifiedSlide modifiedSlide : modifiedSlides) {
            XSLFSlide newSlide = newSlideshow.createSlide(modifiedSlide.parent.original.getSlideLayout());
            modifiedSlide.apply(newSlide);
        }

        newSlideshow.write(out);
        out.close();
    }

    public void save(File file) throws IOException {
        file.getAbsoluteFile().getParentFile().mkdirs();
        FileOutputStream fout = new FileOutputStream(file);
        ZipOutputStream zout = new ZipOutputStream(fout);

        // JSON파일
        zout.putNextEntry(new ZipEntry(ZIP_JSON_NAME));
        zout.write(getModifiedJSON().toJSONString().getBytes());
        zout.closeEntry();
        // pptx파일
        zout.putNextEntry(new ZipEntry(ZIP_PPTX_NAME));
        XMLSlideShow out = new XMLSlideShow(templateSlideShow.getPackage());
        out.write(zout);
        /*
         {
         File temp = File.createTempFile("TEMP_", ".pptx");
         temp.deleteOnExit();
         FileOutputStream foutTemplate = new FileOutputStream(temp);
         System.out.println("temp = " + temp);
         XMLSlideShow out = new XMLSlideShow(templateSlideShow.getPackage());
         foutTemplate.close();
         FileInputStream fin = new FileInputStream(temp);
         int length;
         byte[] buffer = new byte[BUFFER_SIZE];

         while ((length = fin.read(buffer, 0, BUFFER_SIZE)) != -1) {
         zout.write(buffer, 0, length);
         }
         }
         */

        zout.close();
        fout.close();
    }

    public JSONArray getModifiedJSON() {
        JSONArray result = new JSONArray();
        for (ModifiedSlide modified : modifiedSlides) {
            result.add(modified.getJSON());
        }

        return result;
    }


}
