package com.github.kdm1jkm.easypoint.core

import org.apache.poi.xslf.usermodel.*
import org.json.simple.JSONArray
import org.json.simple.JSONObject
import java.util.*
import java.util.function.BiConsumer
import java.util.function.Consumer
import java.util.regex.Pattern

/**
 * 한 슬라이드의 변경 사항을 저장하고 관리하는 클래스이다.
 * @param parent 원본 [TemplateSlide]의 참조
 * @param obj 적용시킬 수정 정보
 */
class ModifiedSlide(val parent: TemplateSlide, obj: JSONObject?) {
    /**
     * 변경사항을 저장할 Map. Key에는 텍스트상자명, Value에는 변경될 사항이 들어간다.
     * 단순한 Map이고, 실제 슬라이드 정보로 내보낼 때 정보를 변경하기 위해 사용하므로 어떤 값이 들어가더라도 오류를 일으키지는 않는다.
     */
    val data = LinkedHashMap<String, String>()

    /**
     * 슬라이드명
     */
    var name: String = parent.name

    /**
     * 파라미터로 받은 [XSLFSlide]객체에 지금까지의 변경사항을 적용한다.
     *
     * @param newSlide 변경사항을 적용할 [XSLFSlide]객체
     */
    fun apply(newSlide: XSLFSlide) {
        newSlide.importContent(parent.original)
        val delList: MutableList<XSLFShape> = ArrayList()
        doWithTextShape(newSlide,
                { textShape, content ->
                    val newText = data[content]
                    if (newText == null) {
                        delList.add(textShape)
                    } else {
                        val original = textShape.textParagraphs[0].textRuns[0]
                        textShape.text = newText
                        textShape.forEach(Consumer { textParagraph: XSLFTextParagraph ->
                            textParagraph.forEach(Consumer { textRun: XSLFTextRun ->
                                textRun.fontColor = original.fontColor
                                textRun.fontSize = original.fontSize
                                textRun.isBold = original.isBold
                                textRun.isItalic = original.isItalic
                                textRun.fontFamily = original.fontFamily
                                textRun.isUnderlined = original.isUnderlined
                            })
                        })
                    }
                }
        ) { textShape, _ -> delList.add(textShape) }
        delList.forEach { newSlide.removeShape(it) }
    }

    /**
     * 현재까지의 변경사항을 JSONObject의 JSONArray로 변환해 반환하는 함수
     *
     * @return 변경사항을 JSONArray로 변환한 것
     */
    val json: JSONArray
        get() {
            val resultArr = JSONArray()
            resultArr.add(name)
            val obj = JSONObject()
            obj.putAll(data)
            resultArr.add(obj)
            return resultArr
        }

    private fun parse(names: Map<String, String>) {
        doWithTextShape(parent.original,
                { _: XSLFTextShape?, content: String ->
                    data[content] = names[content] ?: ""
                }, { _: XSLFTextShape?, title: String -> name = title })
    }

    private fun doWithTextShape(slide: XSLFSlide,
                                runWhenContent: BiConsumer<XSLFTextShape, String>,
                                runWhenTitle: BiConsumer<XSLFTextShape, String>) {
        for (shape in slide.shapes) {
            if (shape !is XSLFTextShape) continue
            val text = shape.text
            val contentMatcher = REGEX_CONTENT.matcher(text)
            val titleMatcher = REGEX_TITLE.matcher(text)
            if (contentMatcher.find()) {
                runWhenContent.accept(shape, contentMatcher.group().substring(1))
            } else if (titleMatcher.find()) {
                runWhenTitle.accept(shape, titleMatcher.group().substring(1))
            }
        }
    }

    companion object {
        private val REGEX_CONTENT = Pattern.compile("!.+")
        private val REGEX_TITLE = Pattern.compile("#.+")
    }

    init {
        val keys: MutableMap<String, String> = HashMap()
        if (obj != null) {
            for (o in obj.keys) {
                val k = o as String
                keys[k] = obj[k] as String
            }
        }
        parse(keys)
    }
}