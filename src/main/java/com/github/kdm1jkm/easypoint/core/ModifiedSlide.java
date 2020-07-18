package com.github.kdm1jkm.easypoint.core;

import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 한 슬라이드의 변경 사항을 저장하고 관리하는 클래스이다.
 */
public class ModifiedSlide {
    private static final Pattern REGEX_CONTENT = Pattern.compile("!.+");
    private static final Pattern REGEX_TITLE = Pattern.compile("#.+");

    /**
     * 원본 {@link TemplateSlide}의 참조이다.
     */
    public final TemplateSlide parent;
    /**
     * 변경사항을 저장할 Map이다. Key에는 텍스트상자명, Value에는 변경될 사항이 들어간다.
     * 단순한 Map이고, 실제 슬라이드 정보로 내보낼 때 정보를 변경하기 위해 사용하므로 어떤 상관없는 값이 들어가더라도 오류를 일으키지는 않는다.
     */
    public final LinkedHashMap<String, String> data = new LinkedHashMap<>();
    /**
     * 슬라이드명이다.
     */
    public String name;

    /**
     * 변경사항이 없을 때 사용하는 생성자이다.
     *
     * @param parent 원본 {@link TemplateSlide}이다. 이곳에서 이름과 원본 슬라이드 정보를 받아온다.
     */
    ModifiedSlide(TemplateSlide parent) {
        this.parent = parent;
        name = parent.name;
        parse(new HashMap<>());
    }

    /**
     * 변경사항을 적용하는 생성자이다.
     *
     * @param parent 원본 {@link TemplateSlide}이다. 이곳에서 이름과 원본 슬라이드 정보를 받아온다.
     * @param obj    JSONArray객체를 받아와 이미 수정된 정보를 받아오는데 사용한다.
     */
    ModifiedSlide(TemplateSlide parent, JSONObject obj) {
        this.parent = parent;

        name = parent.name;

        Map<String, String> keys = new HashMap<>();

        for (Object o : obj.keySet()) {
            String k = (String) o;
            keys.put(k, (String) obj.get(k));
        }

        parse(keys);
    }

    /**
     * 파라미터로 받은 {@link XSLFSlide}객체에 지금까지의 변경사항을 적용한다.
     *
     * @param newSlide 변경사항을 적용할 {@link XSLFSlide}객체입니다.
     */
    public void apply(XSLFSlide newSlide) {
        newSlide.importContent(parent.original);
        List<XSLFShape> delList = new ArrayList<>();

        doWithTextShape(newSlide,
                (XSLFTextShape textShape, String content) -> {
                    String newText = data.getOrDefault(content, "");
                    if (newText.equals("")) {
                        delList.add(textShape);
                    } else {
                        textShape.setText(newText);
                    }
                },
                (XSLFTextShape textShape, String title) ->
                        delList.add(textShape));

        for (XSLFShape shape : delList) {
            newSlide.removeShape(shape);
        }
    }

    /**
     * 현재까지의 변경사항을 JSONObject의 JSONArray로 변환해 반환하는 함수
     *
     * @return 변경사항을 JSONArray로 변환한 것
     */
    public JSONArray getJSON() {
        JSONArray jarr = new JSONArray();
        jarr.add(name);
        JSONObject obj = new JSONObject();
        obj.putAll(data);
        jarr.add(obj);
        return jarr;
    }

    private void parse(Map<String, String> names) {
        doWithTextShape(parent.original,
                (XSLFTextShape textShape, String content) ->
                        data.put(content, names.getOrDefault(content, "")),
                (XSLFTextShape textShape, String title) -> name = title);
    }

    private void doWithTextShape(XSLFSlide slide, ShapeStrLambda runWhenContent, ShapeStrLambda runWhenTitle) {
        for (XSLFShape shape : slide.getShapes()) {
            if (!(shape instanceof XSLFTextShape)) continue;

            String text = ((XSLFTextShape) shape).getText();

            Matcher contentMatcher = REGEX_CONTENT.matcher(text);
            Matcher titleMatcher = REGEX_TITLE.matcher(text);

            if (contentMatcher.find()) {
                runWhenContent.run((XSLFTextShape) shape, contentMatcher.group().substring(1));
            } else if (titleMatcher.find()) {
                runWhenTitle.run((XSLFTextShape) shape, titleMatcher.group().substring(1));
            }
        }
    }

    @FunctionalInterface
    private interface ShapeStrLambda {
        void run(XSLFTextShape textShape, String str);
    }
}
