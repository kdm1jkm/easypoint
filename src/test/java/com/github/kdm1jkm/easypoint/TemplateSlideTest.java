package com.github.kdm1jkm.easypoint;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import static org.junit.Assert.*;


public class TemplateSlideTest {

    public static XSLFSlide slide;
    public static TemplateSlide templateSlide;

    @BeforeClass
    public static void beforeClass() throws Exception {
        String fileName = "test.pptx";
        FileInputStream in = new FileInputStream(new File(fileName));

        XMLSlideShow show = new XMLSlideShow(in);
        slide = show.getSlides().get(0);
        in.close();

        templateSlide = new TemplateSlide(null, slide);
    }

    @Test
    public void constructTest() {
        assertEquals("제목 슬라이드", templateSlide.name);
    }

    @Test
    public void originalTest() {
        assertSame(slide, templateSlide.original);
    }

    @Test
    public void JSONTest() {
        assertEquals(templateSlide.getInfo().getJSON().toJSONString(), "[\"제목 슬라이드\",[\"제목\",\"부제목\"]]");
    }

}