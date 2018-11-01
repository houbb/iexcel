package com.github.houbb.iexcel.util;

import java.io.File;

/**
 *
 * @author jipengfei
 * @author binbin.hou
 */
public class PoiFileUtil {

    private static final String JAVA_IO_TMPDIR = "java.io.tmpdir";

    private static final String POIFILES = "poifiles";

    /**
     */
    public static void createPOIFilesDirectory() {

        String tmpDir = System.getProperty(JAVA_IO_TMPDIR);
        if (tmpDir == null) {
            throw new RuntimeException(
                "Systems temporary directory not defined - set the -D" + JAVA_IO_TMPDIR + " jvm property!");
        }
        File directory = new File(tmpDir, POIFILES);
        syncCreatePOIFilesDirectory(directory);
    }

    /**
     *
     * @param directory
     */
    private static synchronized void syncCreatePOIFilesDirectory(File directory) {
        if (!directory.exists()) {
            directory.mkdirs();
        }
    }
}
