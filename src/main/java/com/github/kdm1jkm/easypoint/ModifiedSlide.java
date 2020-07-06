package com.github.kdm1jkm.easypoint;

import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.ParseException;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ModifiedSlide {
    public static final Pattern REGEX_CONTENT = Pattern.compile("!.+");
    public static final Pattern REGEX_TITLE = Pattern.compile("#.+");
    public final TemplateSlide parent;
    public final LinkedHashMap<String, String> data = new LinkedHashMap<>();
    public String name;


    ModifiedSlide(TemplateSlide parent, JSONArray array) throws ParseException {
        this.parent = parent;

        name = (String) array.get(0);
        JSONObject obj = (JSONObject) array.get(1);

        Map<String, String> keys = new HashMap<>();

        for (Object o : obj.keySet()) {
            String k = (String) o;
            keys.put(k, (String) obj.get(k));
        }

        parse(keys);
    }

    ModifiedSlide(TemplateSlide parent) {
        this.parent = parent;
        name = parent.name;
        parse(new HashMap<>());
    }

    private void parse(Map<String, String> names) {
        doWithTextShape(parent.original,
                (XSLFTextShape textShape, String content) ->
                        data.put(content, names.getOrDefault(content, "")),
                (XSLFTextShape textShape, String title) -> name = title);
    }

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
     * 슬라이드의 모든 Shape에 대해 REGEX_CONTENT와 REGEX_TITLE에 맞는 것들을 각각 runWhenContent와 runWhenTitle로 실행하는 함수.
     *
     * @param slide          모든 Shape를 찾아낼 슬라이드
     * @param runWhenContent REGEX_CONTENT에 맞는 TextShape에 대해 실행할 Lambda
     * @param runWhenTitle   REGEX_TITLE에 맞는 TextShape에 대해 실행할 Lambda
     */
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

    public JSONArray getJSON() {
        JSONArray jarr = new JSONArray();
        jarr.add(name);
        JSONObject obj = new JSONObject();
        obj.putAll(data);
        jarr.add(obj);
        return jarr;
    }

    /**
     * TextShape와 파싱한 글자를 이용하는 인터페이스. 이름을 뭐로 지어야 할 지를 모르겠다.
     */
    @FunctionalInterface
    private interface ShapeStrLambda {
        void run(XSLFTextShape textShape, String str);
    }

}
