package com.allinmoney.platform;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;

/**
 * Created by Chris on 1/12/2018.
 * @author Chris
 */

public class ScaUtil {
    private static final Logger logger = LoggerFactory.getLogger(ScaUtil.class);


    /**
     * Return the file list within given parent directory. If @param recursive is
     * true, the list will consists of all files of each directory.
     * @param parent The file list root directory
     * @param recursive Identify if recursively walk through the given directory
     * @return file list
     */
    public static List<File> listFiles(final File parent, final boolean recursive) {
        if (!parent.exists()) {
            logger.debug("The parent directory: " + parent.getName() + " is not existing!");
            return Collections.emptyList();
        }

        List<File> fileList = new LinkedList<>();

        if (parent.isDirectory()) {
            File[] files = parent.listFiles();
            if (recursive) {
                if (files != null) {
                    for (File file : files) {
                        if (file.isDirectory()) {
                            fileList.addAll(listFiles(file, recursive));
                        } else {
                            fileList.add(file);
                        }
                    }
                }
            }
        }

        return fileList;
    }
}
