package com.github.kdm1jkm.easypoint;

import org.json.simple.JSONArray;

import java.util.ArrayList;
import java.util.List;

/**
 * {@link TemplateSlide}객체에서 정보를 요청했을 시 반환되는 클래스.
 */
public class TemplateSlideInfo {
    /**
     * 슬라이드의 이름
     */
    public final String name;
    /**
     * 텍스트 상자의 이름 목록
     */
    public final List<String> values = new ArrayList<>();
    /**
     * 정보를 담고 있는 원본 {@link TemplateSlide}의 참조.
     */
    public final TemplateSlide parent;

    /**
     * 기본 생성자. 원본 클래스의 참조를 가지고 있다.
     *
     * @param parent 원본 클래스 {@link TemplateSlide}
     */
    public TemplateSlideInfo(TemplateSlide parent) {
        this.name = parent.name;
        this.parent = parent;
    }

    /**
     * 정보를 JSON형식으로 변환해 반환한다.
     *
     * @return 변환된 정보
     */
    public JSONArray getJSON() {
        JSONArray array = new JSONArray();
        array.addAll(values);

        JSONArray result = new JSONArray();
        result.add(name);
        result.add(array);

        return result;
    }
}