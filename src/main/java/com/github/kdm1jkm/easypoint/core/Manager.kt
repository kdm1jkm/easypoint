package com.github.kdm1jkm.easypoint.core

import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xslf.usermodel.XSLFShape
import org.apache.poi.xslf.usermodel.XSLFSlide
import org.apache.poi.xslf.usermodel.XSLFTextShape
import org.json.simple.JSONArray
import org.json.simple.JSONObject
import org.json.simple.parser.JSONParser
import java.io.*
import java.util.*
import java.util.zip.ZipEntry
import java.util.zip.ZipInputStream
import java.util.zip.ZipOutputStream
import javax.naming.NameNotFoundException

/**
 * 전체적으로 모든 작업을 총괄하는 클래스. [Manager] 한 객체는 하나의 템플릿 파일과 대응된다.
 */
class Manager(file: File) {
    /**
     * 원본 [XMLSlideShow]를 [TemplateSlide]로 파싱한 정보를 모아놓는 곳.
     */
    val templateSlides: MutableList<TemplateSlide> = ArrayList()

    /**
     * 실제 수정사항 정보를 모아놓는 곳
     */
    val modifiedSlides: MutableList<ModifiedSlide> = ArrayList()

    private var templateSlideShow: XMLSlideShow? = null
    private var jsonString: String? = null

    init {
        val arr = file.name.split("\\.").toTypedArray()
        when (arr[arr.size - 1]) {
            "pptx" -> {
                templateSlideShow = XMLSlideShow(FileInputStream(file))
                parse()
            }
            "eptx" -> {
                val zipInputStream = ZipInputStream(FileInputStream(file))
                getFromZip(zipInputStream)
                parse()
                parseJSON()
            }
            else -> throw FileNotFoundException(String.format("%s is not pptx or eptx file.", file.name))
        }
        if (templateSlideShow == null) {
            throw FileNotFoundException("Can't load file")
        }
    }

    /**
     * [templateSlides]의 index번째 슬라이드를 [modifiedSlides]에 추가합니다.
     *
     * @param index [templateSlides]의 몇 번째 슬라이드를 추가할 것인지를 지정합니다.
     */
    fun append(index: Int): ModifiedSlide {
        val templateSlide = templateSlides[index]
        val newModified = ModifiedSlide(templateSlide, null)
        modifiedSlides.add(newModified)
        return newModified
    }

    /**
     * [templateSlides]에서 슬라이드 이름이 name인 슬라이드를 [modifiedSlides]에 추가합니다.
     *
     * @param name [templateSlides]에서 찾을 슬라이드 이름입니다.
     */
    fun append(name: String): ModifiedSlide {
        for ((i, templateSlide) in templateSlides.withIndex()) {
            if (templateSlide.name == name) {
                return append(i)
            }
        }
        throw NameNotFoundException(String.format("Name %s is not found in TemplateSlides", name))
    }

    /**
     * 지금까지의 변경사항을 pptx파일로 내보냅니다.
     *
     * @param file 파일을 내보낼 위치를 나타냅니다.
     */
    fun export(file: File) {
        // 폴더 생성
        file.absoluteFile.parentFile.mkdirs()
        val out = FileOutputStream(file)
        val newSlideshow = XMLSlideShow(templateSlideShow!!.getPackage())

        // 슬라이드 모두 삭제
        while (newSlideshow.slides.size != 0) newSlideshow.removeSlide(0)

        for (modifiedSlide in modifiedSlides) {
            val newSlide = newSlideshow.createSlide(modifiedSlide.parent.original.slideLayout)
            modifiedSlide.apply(newSlide)
        }
        newSlideshow.slides.forEach { slide: XSLFSlide ->
            slide.shapes.forEach { shape: XSLFShape? ->
                if (shape is XSLFTextShape)
                    println("shape.text = ${shape.text}")
            }
        }
        println("file.absolutePath = ${file.absolutePath}")
        newSlideshow.write(out)
        out.close()
    }

    /**
     * 현재까지의 저장사항을 원본 템플릿 슬라이드쇼와 JSON파일의 압축파일 형태로 저장합니다.
     *
     * @param file 파일을 내보낼 위치를 나타냅니다.
     */
    fun save(file: File) {
        file.absoluteFile.parentFile.mkdirs()
        val fileOutputStream = FileOutputStream(file)
        val zipOutputStream = ZipOutputStream(fileOutputStream)

        // JSON 파일
        zipOutputStream.putNextEntry(ZipEntry(ZIP_JSON_NAME))
        zipOutputStream.write(modifiedJSON.toJSONString().toByteArray())
        zipOutputStream.closeEntry()
        // pptx 파일
        zipOutputStream.putNextEntry(ZipEntry(ZIP_PPTX_NAME))
        val out = XMLSlideShow(templateSlideShow!!.getPackage())
        out.write(zipOutputStream)

        zipOutputStream.close()
        fileOutputStream.close()
    }

    /**
     * 현재까지의 변경사항을 [JSONArray]객체로 정리해서 내보냅니다.
     *
     * @return 변경사항을 정리한 내용입니다.
     */
    val modifiedJSON: JSONArray
        get() {
            val result = JSONArray()
            for (modified in modifiedSlides) {
                result.add(modified.json)
            }
            return result
        }

    /**
     * 원본 XMLSlideShow객체를 복제해서 리턴합니다.
     * @return 원본 XMLSlideShow객체의 복사본
     */
    val originalSlideshow: XMLSlideShow
        get() = XMLSlideShow(templateSlideShow!!.getPackage())

    private fun getFromZip(zipInputStream: ZipInputStream) {
        var zipEntry: ZipEntry
        // 압축파일 속 각 파일에 대해
        while (zipInputStream.nextEntry.also { zipEntry = it } != null) {

            // 디렉토리는 건너뜀
            if (zipEntry.isDirectory) continue

            // 파일명
            val fileName = zipEntry.name

            // 데이터 저장?공간
            val byteArrayOutputStream = ByteArrayOutputStream()

            // 읽어서 스트림에 넣기
            var size: Int
            val buffer = ByteArray(BUFFER_SIZE)
            while (zipInputStream.read(buffer, 0, BUFFER_SIZE).also { size = it } != -1) {
                byteArrayOutputStream.write(buffer, 0, size)
            }
            when (fileName) {
                ZIP_PPTX_NAME -> {
                    val bin = ByteArrayInputStream(byteArrayOutputStream.toByteArray())
                    templateSlideShow = XMLSlideShow(bin)
                }
                ZIP_JSON_NAME -> jsonString = String(byteArrayOutputStream.toByteArray())
            }
        }
        zipInputStream.closeEntry()
    }

    private fun parse() {
        for (slide in templateSlideShow!!.slides) {
            templateSlides.add(TemplateSlide(null, slide))
        }
    }

    private fun parseJSON() {
        val parser = JSONParser()
        val array = parser.parse(jsonString) as JSONArray
        array.forEach {
            val innerArray = it as JSONArray
            val name = innerArray[0] as String
            for (templateSlide in templateSlides) {
                if (templateSlide.name == name) {
                    modifiedSlides.add(ModifiedSlide(templateSlide, innerArray[1] as JSONObject?))
                    break
                }
            }
        }
    }

    companion object {
        private const val ZIP_PPTX_NAME = "template.pptx"
        private const val ZIP_JSON_NAME = "data.json"
        private const val BUFFER_SIZE = 2048
    }
}