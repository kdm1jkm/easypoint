package com.github.kdm1jkm.easypoint;


import com.github.kdm1jkm.easypoint.core.Manager;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.json.simple.parser.ParseException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) throws IOException, ParseException {
        /*
        System.out.println("Hello, world!");

        System.out.print("Enter filename: ");
        Scanner scanner = new Scanner(System.in);
        String filename = scanner.nextLine();

        File file = new File(filename);

        XMLSlideShow show = new XMLSlideShow(new FileInputStream(file));

        XSLFSlide slide=  show.getSlides().get(0);
        XSLFTextShape shape = (XSLFTextShape) slide.getShapes().get(0);

        shape.setText("This is TEXT");

        FileOutputStream out = new FileOutputStream(new File("output.pptx"));
        show.write(out);
        out.close();
        /**/


//        /*
        Manager m = new Manager(new File("test.pptx"));
        m.append(0);
        for (String s : m.modifiedSlides.get(0).data.keySet()) {
            System.out.println("s = " + s);
        }
        m.modifiedSlides.get(0).data.put("제목", "title");
        m.save(new File("output.eptx"));

        Manager other = new Manager(new File("output.eptx"));
        other.export(new File("output.pptx"));
        /**/


        /*
        String fileName;
        if (args.length != 1) {
            System.out.println("Usage: java -jar easypoint.jar <FileName>");
            fileName = "test.pptx";
//            return;
        } else {
            fileName = args[0];
        }
        File file = new File(fileName);

        if (!file.exists()) {
            System.out.println(String.format("Can't find %s", fileName));
            return;
        } else if (file.isDirectory()) {
            System.out.println(String.format("%s is directory, not a file.", fileName));
            return;
        }

//        CLIManager
        /**/

    }
}
