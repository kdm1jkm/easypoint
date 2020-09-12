package com.github.kdm1jkm.easypoint.core

import org.apache.poi.xslf.usermodel.XSLFShape
import org.apache.poi.xslf.usermodel.XSLFSlide
import org.apache.poi.xslf.usermodel.XSLFTextShape
import java.util.*
import java.util.regex.Pattern

/**
 * 템플릿에서 사용할 pptx의 각 슬라이드의 파싱 정보를 담아 둘 클래스
 *
 * @param name     슬라이드의 이름. 파싱 시 나온 이름이 있을경우 이 이름은 무시된다.
 * @param original 원본 슬라이드.
 */
class TemplateSlide(var name: String, val original: XSLFSlide) {
    /**
     * 파싱한 정보를 저장하는 리스트
     * 텍스트 상자의 이름 목록이다.
     */
    val data = LinkedHashSet<String>()

    /**
     * 이 슬라이드의 정보를 담은 [TemplateSlideInfo]객체를 반환한다.
     *
     * @return 슬라이드의 정보
     */
    val info: TemplateSlideInfo
        get() {
            val result = TemplateSlideInfo(this)
            result.values.addAll(data)
            return result
        }

    private fun parse() {
        val deleteList: MutableList<XSLFShape> = ArrayList()
        for (shape in original.shapes) {
            if (shape !is XSLFTextShape) {
                continue
            }
            val text = shape.text
            val contentMatcher = REGEX_CONTENT.matcher(text)
            val titleMatcher = REGEX_TITLE.matcher(text)
            if (contentMatcher.find()) {
                val content = contentMatcher.group()
                data.add(content.substring(1))
            }
            if (titleMatcher.find()) {
                name = titleMatcher.group().substring(1)
                deleteList.add(shape)
            }
        }
        deleteList.forEach { original.removeShape(it) }
    }

    companion object {
        /**
         * pptx에서 내용 텍스트상자 찾기 위한 정규표현식 패턴
         */
        val REGEX_CONTENT: Pattern = Pattern.compile("!.+")

        /**
         * pptx에서 제목 텍스트상자 찾기 위한 정규표현식 패턴
         */
        val REGEX_TITLE: Pattern = Pattern.compile("#.+")
    }

    init {
        parse()
    }
}