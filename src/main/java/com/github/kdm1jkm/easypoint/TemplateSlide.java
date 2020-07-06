package com.github.kdm1jkm.easypoint;

import com.sun.istack.internal.Nullable;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.json.simple.JSONArray;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TemplateSlide {
    public static final Pattern REGEX_CONTENT = Pattern.compile("!.+");
    public static final Pattern REGEX_TITLE = Pattern.compile("#.+");
    private static int count = 0;
    public final XSLFSlide original;
    public final LinkedHashSet<String> data = new LinkedHashSet<>();
    public String name;

    public TemplateSlide(@Nullable String name, XSLFSlide original) {
        this.original = original;
        if (name == null) {
            this.name = String.format("Untitled %d", count);
            count++;
        } else {
            this.name = name;
        }
        parse();
    }

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
//            System.out.println(text);
//            System.out.println("text");

            if (contentMatcher.find()) {
//                System.out.println("contentm");
                String content = contentMatcher.group();
                data.add(content.substring(1));
            }

            if (titleMatcher.find()) {
//                System.out.println("title");
                name = titleMatcher.group().substring(1);
                deleteList.add(shape);
            }

        }

        for (XSLFShape shape : deleteList) {
            original.removeShape(shape);
        }
    }

    public void copy(XSLFSlide newSlide) {
        newSlide.importContent(original);
    }
}

class TemplateSlideInfo {
    public final String name;
    public final List<String> values = new ArrayList<>();
    public final TemplateSlide parent;

    public TemplateSlideInfo(TemplateSlide parent) {
        this.name = parent.name;
        this.parent = parent;
    }

    public JSONArray getJSON() {
//        System.out.println(name);
//        System.out.println(values.toString());

        JSONArray array = new JSONArray();
        array.addAll(values);

        JSONArray result = new JSONArray();
        result.add(name);
        result.add(array);

        return result;
    }
}