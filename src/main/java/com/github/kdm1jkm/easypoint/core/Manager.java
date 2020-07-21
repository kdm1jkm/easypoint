package com.github.kdm1jkm.easypoint.core;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import javax.naming.NameNotFoundException;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * 전체적으로 모든 작업을 총괄하는 클래스. {@link Manager} 한 객체는 하나의 템플릿 파일과 대응된다.
 */
public class Manager {
    private static final String ZIP_PPTX_NAME = "template.pptx";
    private static final String ZIP_JSON_NAME = "data.json";
    private static final int BUFFER_SIZE = 2048;

    /**
     * 원본 {@link XMLSlideShow}를 {@link TemplateSlide}로 파싱한 정보를 모아놓는 곳.
     */
    public final List<TemplateSlide> templateSlides = new ArrayList<>();
    /**
     * 실제 수정사항 정보를 모아놓는 곳
     */
    public final List<ModifiedSlide> modifiedSlides = new ArrayList<>();

    private XMLSlideShow templateSlideShow = null;
    private String JSONString;

    /**
     * 기본 생성자. {@link File}. 파일 확장자에 따라 파싱을 수행한다.
     *
     * @param file 파싱을 수행할 파일.
     * @throws IOException    파일을 열고 닫는 과정에서 생기는 예외
     * @throws ParseException JSON파일을 파싱하는 과정에서 생기는 예외
     */
    public Manager(File file) throws IOException, ParseException {
        String fileName = file.getName();
        String[] arr = fileName.split("\\.");
        switch (arr[arr.length - 1]) {
            case "pptx":
                templateSlideShow = new XMLSlideShow(new FileInputStream(file));
                parse();
                break;
            case "eptx":
                ZipInputStream zipInputStream = new ZipInputStream(new FileInputStream(file));
                getFromZip(zipInputStream);
                parse();
                parseJSON();
                break;
            default:
                throw new FileNotFoundException(String.format("%s is not pptx or eptx file.", fileName));
        }

        if (templateSlideShow == null) {
            throw new FileNotFoundException("Can't load file");
        }


    }


    /**
     * {@link #templateSlides}의 index번째 슬라이드를 {@link #modifiedSlides}에 추가합니다.
     *
     * @param index {@link #templateSlides}의 몇 번째 슬라이드를 추가할 것인지를 지정합니다.
     */
    public ModifiedSlide append(int index) {
        TemplateSlide templateSlide = templateSlides.get(index);
        ModifiedSlide new_modified = new ModifiedSlide(templateSlide);

        modifiedSlides.add(new_modified);

        return new_modified;
    }

    /**
     * {@link #templateSlides}에서 슬라이드 이름이 name인 슬라이드를 {@link #modifiedSlides}에 추가합니다.
     *
     * @param name {@link #templateSlides}에서 찾을 슬라이드 이름입니다.
     * @throws NameNotFoundException {@link #templateSlides}에서 name이라는 이름의 슬라이드를 찾지 못했을 경우 발생합니다.
     */
    public ModifiedSlide append(String name) throws NameNotFoundException {
        for (int i = 0; i < templateSlides.size(); i++) {
            if (templateSlides.get(i).name.equals(name)) {
                return append(i);
            }
        }

        throw new NameNotFoundException(String.format("Name %s is not found in TemplateSlides", name));
    }

    /**
     * 지금까지의 변경사항을 pptx파일로 내보냅니다.
     *
     * @param file 파일을 내보낼 위치를 나타냅니다.
     * @throws IOException 파일 출력 과정에서 발생하는 오류입니다.
     */
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

        newSlideshow.getSlides().forEach(slide->{
            slide.getShapes().forEach(shape->{
                if(shape instanceof XSLFTextShape)
                    System.out.println("((XSLFTextShape) shape).getText() = " + ((XSLFTextShape) shape).getText());
            });
        });
        System.out.println("file.getAbsolutePath() = " + file.getAbsolutePath());

        newSlideshow.write(out);
        out.close();
    }

    /**
     * 현재까지의 저장사항을 원본 템플릿 슬라이드쇼와 JSON파일의 압축파일 형태로 저장합니다.
     *
     * @param file 파일을 내보낼 위치를 나타냅니다.
     * @throws IOException 파일 출력 과정에서 발생하는 오류입니다.
     */
    public void save(File file) throws IOException {
        file.getAbsoluteFile().getParentFile().mkdirs();
        FileOutputStream fout = new FileOutputStream(file);
        ZipOutputStream zout = new ZipOutputStream(fout);

        // JSON 파일
        zout.putNextEntry(new ZipEntry(ZIP_JSON_NAME));
        zout.write(getModifiedJSON().toJSONString().getBytes());
        zout.closeEntry();
        // pptx 파일
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

    /**
     * 현재까지의 변경사항을 {@link JSONArray}객체로 정리해서 내보냅니다.
     *
     * @return 변경사항을 정리한 내용입니다.
     */
    public JSONArray getModifiedJSON() {
        JSONArray result = new JSONArray();
        for (ModifiedSlide modified : modifiedSlides) {
            result.add(modified.getJSON());
        }

        return result;
    }

    /**
     * 원본 XMLSlideShow객체를 복제해서 리턴합니다.
     * @return 원본 XMLSlideShow객체의 복사본
     */
    public XMLSlideShow getOriginalSlideshow(){
        return new XMLSlideShow(templateSlideShow.getPackage());
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
//            PipedInputStream pipedInputStream = new PipedInputStream();
//            PipedOutputStream pipedOutputStream = new PipedOutputStream(pipedInputStream);

            // 데이터 저장?공간
            ByteArrayOutputStream bout = new ByteArrayOutputStream();

            // 읽어서 스트림에 넣기
            {
                int size;
                byte[] buffer = new byte[BUFFER_SIZE];
                while ((size = zipInputStream.read(buffer, 0, BUFFER_SIZE)) != -1) {
                    bout.write(buffer, 0, size);
                }
            }

//             pos write해주는 부분

            // 파일명에 따라 다르게 행동
            switch (fileName) {
                case ZIP_PPTX_NAME: {
                    /*
                    new Thread(
                            () -> {
                                try {
                                    bout.writeTo(pipedOutputStream);
                                } catch (IOException e) {
                                    e.printStackTrace();
                                }
                            }
                    ).start();
                     */
                    ByteArrayInputStream bin = new ByteArrayInputStream(bout.toByteArray());
                    templateSlideShow = new XMLSlideShow(bin);
                    break;
                }
                case ZIP_JSON_NAME:
                    JSONString = new String(bout.toByteArray());
            }
        }
        zipInputStream.closeEntry();
    }

    private void parse() {
        for (XSLFSlide slide : templateSlideShow.getSlides()) {
            templateSlides.add(new TemplateSlide(null, slide));
        }
    }

    private void parseJSON() throws ParseException {
        JSONParser parser = new JSONParser();
        JSONArray array = (JSONArray) parser.parse(JSONString);
        for (Object o : array) {
            JSONArray inarr = (JSONArray) o;
            String name = (String) inarr.get(0);
            for (TemplateSlide templateSlide : templateSlides) {
                if (templateSlide.name.equals(name)) {
                    modifiedSlides.add(new ModifiedSlide(templateSlide, (JSONObject) inarr.get(1)));
                    break;
                }
            }
        }
    }

}
