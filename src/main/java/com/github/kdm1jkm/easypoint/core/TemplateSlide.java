package com.github.kdm1jkm.easypoint.core;

import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 템플릿에서 사용할 pptx의 각 슬라이드의 파싱 정보를 담아 둘 클래스
 */
public class TemplateSlide {
    /**
     * pptx에서 내용 텍스트상자 찾기 위한 정규표현식 패턴
     */
    public static final Pattern REGEX_CONTENT = Pattern.compile("!.+");

    /**
     * pptx에서 제목 텍스트상자 찾기 위한 정규표현식 패턴
     */
    public static final Pattern REGEX_TITLE = Pattern.compile("#.+");

    private static int count = 0;

    /**
     * 원본 슬라이드 참조
     */
    public final XSLFSlide original;

    /**
     * 파싱한 정보를 저장하는 리스트
     * 텍스트 상자의 이름 목록이다.
     */
    public final LinkedHashSet<String> data = new LinkedHashSet<>();

    /**
     * 슬라이드의 제목
     */
    public String name;

    /**
     * 기본 생성자이다. 원본 슬라이드를 받아서 파싱한다.
     *
     * @param name     슬라이드의 이름이다. 파싱 시 나온 이름이 있을경우 이 이름은 무시된다.
     * @param original 원본 슬라이드이다.
     */
    public TemplateSlide(String name, XSLFSlide original) {
        this.original = original;
        if (name == null) {
            this.name = String.format("Untitled %d", count);
            count++;
        } else {
            this.name = name;
        }
        parse();
    }

    /**
     * 이 슬라이드의 정보를 담은 {@link TemplateSlideInfo}객체를 반환한다.
     *
     * @return 슬라이드의 정보
     */
    public TemplateSlideInfo getInfo() {
        TemplateSlideInfo result = new TemplateSlideInfo(this);
        result.values.addAll(data);
        return result;
    }

    private void parse() {
        List<XSLFShape> deleteList = new ArrayList<>();

        for (XSLFShape shape : original.getShapes()) {
            if (!(shape instanceof XSLFTextShape)) {
                continue;
            }

            String text = ((XSLFTextShape) shape).getText();

            Matcher contentMatcher = REGEX_CONTENT.matcher(text);
            Matcher titleMatcher = REGEX_TITLE.matcher(text);

            if (contentMatcher.find()) {
                String content = contentMatcher.group();
                data.add(content.substring(1));
            }

            if (titleMatcher.find()) {
                name = titleMatcher.group().substring(1);
                deleteList.add(shape);
            }

        }

        for (XSLFShape shape : deleteList) {
            original.removeShape(shape);
        }
    }
}