/*
package com.github.kdm1jkm.easypoint;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.json.simple.JSONArray;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class OldManager {
    private final XMLSlideShow templateSlideShow;
    private final XMLSlideShow newSlideShow ;
    private final List<TemplateSlide> templateSlides = new ArrayList<>();
    private final List<TemplateSlideInfo> modified = new ArrayList<>();


    public OldManager(FileInputStream templateFile) throws IOException {
        templateSlideShow = new XMLSlideShow(templateFile);
        newSlideShow = new XMLSlideShow(templateSlideShow.getPackage());
//        newSlideShow = new XMLSlideShow();
        newSlideShow.setPageSize(templateSlideShow.getPageSize());
    }

    public void parse() {
        for (int i = 0; i < templateSlideShow.getSlides().size(); i++) {
            TemplateSlide st = new TemplateSlide(String.format("Untitled %d", i + 1), templateSlideShow.getSlides().get(i));
            st.parse();
            templateSlides.add(st);
        }
    }

    public List<TemplateSlideInfo> getInfo() {
        List<TemplateSlideInfo> result = new ArrayList<>();

        for (TemplateSlide slide : templateSlides) {
            result.add(slide.getInfo());
        }

        return result;
    }

    private XSLFSlide copySlide(int index) {
        XSLFSlide originalSlide = templateSlideShow.getSlides().get(index);

        XSLFSlide newSlide = newSlideShow.createSlide(originalSlide.getSlideLayout());
        newSlide.importContent(originalSlide);

        return newSlide;
    }

    public void appendSlide(int index){
        copySlide(index);

    }

    public void save(OutputStream out) throws IOException {
        newSlideShow.write(out);
    }

    public String getJSON() {
        JSONArray result = new JSONArray();
        for (TemplateSlide slide : templateSlides) {
            result.add(slide.getInfo().getJSON());
        }
        return result.toJSONString();
    }

}
*/