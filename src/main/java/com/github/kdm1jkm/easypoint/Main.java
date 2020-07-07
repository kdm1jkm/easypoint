package com.github.kdm1jkm.easypoint;


import org.json.simple.parser.ParseException;

import java.io.File;
import java.io.IOException;

class Main {
    public static void main(String[] args) throws IOException, ParseException {
        Manager m = new Manager(new File("test.pptx"));
        m.append(0);
        for (String s : m.modifiedSlides.get(0).data.keySet()) {
            System.out.println("s = " + s);
        }
        m.modifiedSlides.get(0).data.put("제목", "이거슨 제목일겁니다.");
        m.save(new File("output.eptx"));

        Manager other = new Manager(new File("output.eptx"));
        other.export(new File("output.pptx"));
    }
}
