package com.github.kdm1jkm.easypoint.cli;

import com.github.kdm1jkm.easypoint.core.Manager;
import com.github.kdm1jkm.easypoint.core.ModifiedSlide;
import com.github.kdm1jkm.easypoint.core.TemplateSlide;
import org.apache.commons.cli.*;
import org.json.simple.parser.ParseException;

import javax.naming.NameNotFoundException;
import java.io.*;
import java.nio.charset.StandardCharsets;

public class CLIManager {
//    private final Scanner scanner = new Scanner(System.in);
    private Manager manager;
    private boolean isEnd = false;

    public CLIManager(String[] args) throws IOException, ParseException, org.apache.commons.cli.ParseException, NameNotFoundException {
        parseOption(args);
    }

    private void parseOption(String[] args) throws org.apache.commons.cli.ParseException, IOException, ParseException, NameNotFoundException {
        Options options = new Options();

        {
            Option help = new Option("help", false, "Show help.");
            Option templates = new Option("templates", false, "Make list of template slides.");
            Option current = new Option("current", false, "Make list of modified slides.");
            Option o_pptx = new Option("o_pptx", false, "Make Original pptx file.");
            Option pptx = new Option("pptx", false, "Convert .eptx file into .pptx file.");

            Option file = Option.builder("file").argName("filename").hasArg().desc("file to get data from.").build();
            Option output = Option.builder("out").argName("filename").hasArg().desc("file to export data.").build();
            Option from = Option.builder("from").argName("filename").hasArg().desc("file to get data and apply to modified slides.").build();

            options.addOption(help);
            options.addOption(templates);
            options.addOption(current);
            options.addOption(o_pptx);
            options.addOption(pptx);
            options.addOption(file);
//            options.addOption(output);
            options.addOption(from);
        }

        CommandLineParser parser = new DefaultParser();
        CommandLine line = parser.parse(options, args);
        HelpFormatter formatter = new HelpFormatter();

        if (line.hasOption("help")) {
            formatter.printHelp("java -jar easypoint.jar", options);
            isEnd = true;
            return;
        }

        // 파일 파싱
        if (!line.hasOption("file")) {
            formatter.printHelp("java -jar easypoint.jar", options);
            throw new org.apache.commons.cli.ParseException("argument \"file\" is missing.");
        }

        String fileName = line.getOptionValue("file");
        File file = new File(line.getOptionValue("file"));

        if (!file.exists()) {
//            System.out.println(String.format("Can't find %s", fileName));
            throw new FileNotFoundException(String.format("Can't find %s", fileName));
        } else if (file.isDirectory()) {
//            System.out.println(String.format("%s is directory, not a file.", fileName));
            throw new FileNotFoundException(String.format("%s is directory, not a file.", fileName));
        }

        manager = new Manager(file);
        String path = file.getAbsolutePath();
        path = path.substring(0, path.length() - 5);

        if(file.getName().endsWith(".pptx")){
            manager.save(new File(path + ".eptx"));
        }

        if (line.hasOption("templates")) {
            isEnd = true;
            File templateFile = new File(path + "_templates.txt");
            FileWriter writer = new FileWriter(templateFile);
            for (TemplateSlide templateSlide : manager.templateSlides) {
//                System.out.println(String.format("# %s", templateSlide.name));
                writer.write(String.format("# %s\n", templateSlide.name));
                for (String v : templateSlide.getInfo().values) {
//                    System.out.println(v);
                    writer.write(String.format("%s\n", v));
                }
            }
            writer.flush();
            writer.close();
        }

        if (line.hasOption("from")) {
            File fromFile = new File(line.getOptionValue("from"));
            FileReader fileReader = new FileReader(fromFile);
//            BufferedReader bufferedReader = new BufferedReader(fileReader);
            BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(new FileInputStream(fromFile),
                    StandardCharsets.UTF_8));

            String l;
            String title = "";
            ModifiedSlide current = null;

            manager.modifiedSlides.clear();

            while ((l = bufferedReader.readLine()) != null) {
                if (l.startsWith("# ")) {
                    title = l.substring(2);
                    System.out.println(String.format("# %s", title));
                    current = manager.append(title);
                } else {
                    System.out.println(String.format("(%s) %s", title, l));
                    String[] split_l = l.split(": ");
                    String k = split_l[0];
                    String v = split_l[1];
                    current.data.put(k, v);
                }
            }
            manager.save(file);
            fileReader.close();
        }

        if (line.hasOption("current")) {
            isEnd = true;
            File currentFile = new File(path + "_currentState.txt");
            FileWriter writer = new FileWriter(currentFile);

            for (ModifiedSlide modifiedSlide : manager.modifiedSlides) {
                writer.write(String.format("# %s\n", modifiedSlide.name));
                modifiedSlide.data.forEach((k, v) -> {
                    try {
                        writer.write(String.format("%s: %s\n", k, v));
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                });
            }

            writer.flush();
            writer.close();
        }


        if(line.hasOption("o_pptx")){
            manager.getOriginalSlideshow().write(new FileOutputStream(new File(path + "_original.pptx")));
        }

        if(line.hasOption("pptx")){
            manager.export(new File(path + ".pptx"));
        }
    }

    public boolean isEnd() {
        return isEnd;
    }

    /*
    public void temp() {

        while (true) {
            System.out.println("---Current State---");
            for (ModifiedSlide slide : manager.modifiedSlides) {
                System.out.println(String.format("# %s", slide.name));
                slide.data.forEach((k, v) -> System.out.println(String.format("%s: %s", k, v)));
                for (String k : slide.data.keySet()) {
                    String v = slide.data.get(k);

                    System.out.println(String.format("%s: %s", k, v));
                }
            }

            System.out.print(
                    "---Commands---\n" +
                            "-[M]odify    -[A]dd       -[D]elete\n" +
                            "-[R]earrange -{E]xport    -[S]ave and Exit\n" +
                            "Input: "
            );
            String input = scanner.nextLine();
            switch (input) {
                case "m":
                case "M":
                case "modify":
                case "Modify": {
                    System.out.println("Modify");
                    break;
                }

                case "a":
                case "A":
                case "add":
                case "Add": {
                    System.out.println("Add");
                    break;
                }

                case "d":
                case "D":
                case "delete":
                case "Delete": {
                    System.out.println("Delete");
                    break;
                }

                case "r":
                case "R":
                case "rearrange":
                case "Rearrange": {
                    System.out.println("Rearrange");
                    break;
                }

                case "e":
                case "E":
                case "export":
                case "Export": {
                    System.out.println("Export");
                    break;
                }

                case "s":
                case "S":
                case "Save":
                case "save":
                case "Save and exit": {
                    System.out.println("Save and exit");
                    break;
                }
            }
            break;

        }
    }

     */
}
