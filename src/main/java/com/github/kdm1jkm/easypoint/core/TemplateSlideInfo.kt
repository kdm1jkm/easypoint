package com.github.kdm1jkm.easypoint.core


import java.util.*

/**
 * [TemplateSlide]객체에서 정보를 요청했을 시 반환되는 클래스.
 *
 * @param parent 원래 [TemplateSlide] 객체의 참조
 */
class TemplateSlideInfo(val parent: TemplateSlide) {
    /**
     * 슬라이드의 이름
     */
    val name: String = parent.name

    /**
     * 텍스트 상자의 이름 목록
     */
    val values: MutableList<String> = ArrayList()

}