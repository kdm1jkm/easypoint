package com.github.kdm1jkm.easypoint;


import java.io.File;
import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {

//        OldManager m = new OldManager(new FileInputStream("test.pptx"));
//        m.parse();
//        m.copySlide(0);
//        System.out.println(m.getJSON());
//        m.save(new FileOutputStream("test2.pptx"));

        Manager m = new Manager(new File("test.pptx"));
        m.append(0);
        for(String s: m.modifiedSlides.get(0).data.keySet()){
            System.out.println("s = " + s);
        }
        m.modifiedSlides.get(0).data.put("제목", "이거슨 제목일겁니다.");
        m.export(new File("output.pptx"));

    }
}
