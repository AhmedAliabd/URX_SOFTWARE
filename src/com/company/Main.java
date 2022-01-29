package com.company;

import java.io.*;
import java.nio.file.Files;
import static java.nio.file.StandardOpenOption.CREATE;

public class Main {

    public static void main(String[] args) {
        final File folder = new File("/home/aaa/Downloads/");
        final File exported = new File("/home/aaa/Downloads/Exported/test.txt");
        listFilesForFolder(folder,exported,1);

    }

    public static void listFilesForFolder(final File folder,final File exported, int skipLine) {
        try {
            OutputStream outputStream = new BufferedOutputStream(Files.newOutputStream(exported.toPath(), CREATE));
            BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(outputStream));
            for (final File fileEntry : folder.listFiles()) {
                String ext = fileEntry.getName().substring(fileEntry.getName().lastIndexOf('.') + 1);
                if (fileEntry.isDirectory()) {
                    continue;
                } else if(ext.equals("txt")) {
                    InputStream inputStream = new BufferedInputStream(Files.newInputStream(fileEntry.toPath()));
                    BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream));
                    String line = null;
                    for(int i = 0; i <skipLine; i++){
                        reader.readLine();
                    }
                    line = reader.readLine();
                    System.out.println(line);
                    writer.write(line);
                    writer.newLine();
                    writer.flush();
                    reader.close();
                }
            }
            writer.close();
        }catch (Exception e){
            System.out.println(e.getMessage());
        }
        }
    }





