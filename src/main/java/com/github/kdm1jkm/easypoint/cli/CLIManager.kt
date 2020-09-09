package com.github.kdm1jkm.easypoint.cli

import com.github.kdm1jkm.easypoint.core.Manager
import com.github.kdm1jkm.easypoint.core.ModifiedSlide
import org.apache.commons.cli.*
import java.io.*
import java.nio.charset.StandardCharsets

class CLIManager(args: Array<String>) {
    private lateinit var manager: Manager
    var isEnd = false
        private set

    companion object {
        @JvmStatic
        val HELP_STRING = "help"

        @JvmStatic
        val FILE_STRING = "file"

        @JvmStatic
        val TEMPLATES_STRING = "templates"

        @JvmStatic
        val CURRENT_STRING = "current"

        @JvmStatic
        val ORIGINAL_PPTX_STRING = "o_pptx"

        @JvmStatic
        val PPTX_STRING = "pptx"

//        @JvmStatic
//        val OUTPUT_STRING = "out"

        @JvmStatic
        val FROM_STRING = "from"

        @JvmStatic
        val COMMAND_LINE_SYNTAX_STRING = "java -jar easypoint.jar"

        @JvmStatic
        val FILE_ARGUMENT_NOT_FOUND_STRING = "Argument \"file\" is missing."

        @JvmStatic
        val FILE_NOT_FOUND_STRING = "Can't find %s"

        @JvmStatic
        val DIRECTORY_NOT_FILE_STRING = "%s is directory, not a file."
    }

    private fun parseOption(args: Array<String>) {
        val options = Options()
        addOptions(options)
        val parser: CommandLineParser = DefaultParser()
        val commandLine = parser.parse(options, args)
        val formatter = HelpFormatter()

        if (commandLine.hasOption(HELP_STRING)) {
            formatter.printHelp(COMMAND_LINE_SYNTAX_STRING, options)
            isEnd = true
            return
        }

        // 파일 파싱
        if (!commandLine.hasOption(FILE_STRING)) {
            formatter.printHelp(COMMAND_LINE_SYNTAX_STRING, options)
            throw ParseException(FILE_ARGUMENT_NOT_FOUND_STRING)
        }
        val file = File(commandLine.getOptionValue(FILE_STRING))
        if (!file.exists()) {
            throw FileNotFoundException(String.format(FILE_NOT_FOUND_STRING, file.name))
        } else if (file.isDirectory) {
            throw FileNotFoundException(String.format(DIRECTORY_NOT_FILE_STRING, file.name))
        }

        manager = Manager(file)
        val path = getOnlyPath(file)
        if (file.name.endsWith(".pptx")) {
            manager.save(File("$path.eptx"))
        }
        if (commandLine.hasOption(TEMPLATES_STRING)) {
            isEnd = true
            val templateFile = File(path + "_templates.txt")
            val writer = FileWriter(templateFile)
            for (templateSlide in manager.templateSlides) {
                writer.write(String.format("# %s\n", templateSlide.name))
                for (v in templateSlide.info.values) {
                    writer.write(String.format("%s\n", v))
                }
            }
            writer.flush()
            writer.close()
        }
        if (commandLine.hasOption(FROM_STRING)) {
            val sourceFile = File(commandLine.getOptionValue(FROM_STRING))
            val bufferedReader = BufferedReader(InputStreamReader(FileInputStream(sourceFile), StandardCharsets.UTF_8))
            var line: String
            var title = ""
            var currentSlide: ModifiedSlide? = null
            manager.modifiedSlides.clear()
            while (bufferedReader.readLine().also { line = it } != null) {
                if (line.startsWith("# ")) {
                    title = line.substring(2)
                    println(String.format("# %s", title))
                    currentSlide = manager.append(title)
                } else {
                    println(String.format("(%s) %s", title, line))
                    val splitLine = line.split(": ").toTypedArray()
                    if (splitLine.size > 1) {
                        val key = splitLine[0]
                        val value = splitLine[1]
                        if (currentSlide == null) throw Exception("File format is wrong")
                        currentSlide.data[key] = value
                    }
                }
            }
            manager.save(sourceFile)
        }
        if (commandLine.hasOption(CURRENT_STRING)) {
            isEnd = true
            val currentFile = File(path + "_currentState.txt")
            val writer = BufferedWriter(OutputStreamWriter(FileOutputStream(currentFile), StandardCharsets.UTF_8))
            for (modifiedSlide in manager.modifiedSlides) {
                writer.write(String.format("# %s\n", modifiedSlide.name))
                for ((key, value) in modifiedSlide.data) {
                    writer.write(String.format("%s: %s\n", key, value))
                }
            }
            writer.flush()
            writer.close()
        }
        if (commandLine.hasOption(ORIGINAL_PPTX_STRING)) {
            manager.originalSlideshow.write(FileOutputStream(File(path + "_original.pptx")))
        }
        if (commandLine.hasOption(PPTX_STRING)) {
            manager.export(File("$path.pptx"))
        }
    }

    private fun getOnlyPath(file: File): String {
        val path = file.absolutePath
        return path.substring(0, path.length - 5)
    }

    private fun addOptions(options: Options) {
        val help = Option(HELP_STRING, false, "Show help.")
        val templates = Option(TEMPLATES_STRING, false, "Make list of template slides.")
        val current = Option(CURRENT_STRING, false, "Make list of modified slides.")
        val originalPptx = Option(ORIGINAL_PPTX_STRING, false, "Make Original pptx file.")
        val pptx = Option(PPTX_STRING, false, "Convert .eptx file into .pptx file.")
        val file = Option.builder(FILE_STRING).argName("filename").hasArg().desc("file to get data from.").build()
//        val output = Option.builder(OUTPUT_STRING).argName("filename").hasArg().desc("file to export data.").build()
        val from = Option.builder(FROM_STRING).argName("filename").hasArg().desc("file to get data and apply to modified slides.").build()
        options.addOption(help)
        options.addOption(templates)
        options.addOption(current)
        options.addOption(originalPptx)
        options.addOption(pptx)
        options.addOption(file)
//        options.addOption(output)
        options.addOption(from)
    }

    init {
        parseOption(args)
    }
}