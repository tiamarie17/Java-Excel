package org.excel;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

public class Utils {
    //Convert txt file to array of strings
    public static String[] ConvertTxtFile(Path path) throws IOException {
        //TODO: Make this method accept any path
        List<String> lines = Files.readAllLines(path);
        String[] arr = lines.toArray(new String[lines.size()]);
        System.out.println("arr is " + Arrays.toString(arr));
        return arr;
    }
}
