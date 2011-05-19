//
// JODConverter - Java OpenDocument Converter
// Copyright 2009 Art of Solving Ltd
// Copyright 2004-2009 Mirko Nasato
//
// JODConverter is free software: you can redistribute it and/or
// modify it under the terms of the GNU Lesser General Public License
// as published by the Free Software Foundation, either version 3 of
// the License, or (at your option) any later version.
//
// JODConverter is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
// Lesser General Public License for more details.
//
// You should have received a copy of the GNU Lesser General
// Public License along with JODConverter.  If not, see
// <http://www.gnu.org/licenses/>.
//
package org.artofsolving.jodconverter.util;

import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

public class PlatformUtils {

    private static final Logger logger = Logger.getLogger(PlatformUtils.class.getName());

    private static final String OS_NAME = System.getProperty("os.name").toLowerCase();

    private static final String WINDOWS = "windows";

    private static final String MAC = "mac";

    private static final String LINUX = "linux";

    private static final String[] LINUX_OO_HOME_PATH = { "/usr/lib/openoffice",
            "/usr/lib/libreoffice", "/usr/lib/openoffice.org",
            "/usr/lib/openoffice.org3", "/opt/openoffice.org3",
            "/opt/libreoffice", "/usr/lib/ooo" };

    private static final String[] MAC_OO_HOME_PATH = {
            "/Applications/LibreOffice.app/Contents",
            "/Applications/OpenOffice.org.app/Contents" };

    private static final String[] WINDOWS_OO_HOME_PATH = {
            System.getenv("ProgramFiles") + File.separator
                    + "OpenOffice.org 3",
            System.getenv("ProgramFiles") + File.separator
                    + "LibreOffice 3",
            System.getenv("ProgramFiles(x86)") + File.separator
                    + "OpenOffice.org 3",
            System.getenv("ProgramFiles(x86)") + File.separator
                    + "LibreOffice 3" };

    private static final String[] LINUX_OO_PROFILE_PATH = {
            System.getProperty("user.home") + File.separator
                    + ".openoffice.org/3",
            System.getProperty("user.home") + File.separator
                    + ".libreoffice/3" };

    private static final String[] MAC_OO_PROFILE_PATH = {
            System.getProperty("user.home") + File.separator
                    + "Library/Application Support/OpenOffice.org/3",
            System.getProperty("user.home") + File.separator
                    + "Library/Application Support/LibreOffice.org/3" };

    private static final String[] WINDOWS_OO_PROFILE_PATH = {
            System.getenv("APPDATA") + File.separator + "OpenOffice.org/3",
            System.getenv("APPDATA") + File.separator + "LibreOffice.org/3" };

    public static String OO_HOME_PATH = findOfficeHome();

    public static String OO_PROFILE_DIR_PATH = findOfficeProfileDir();

    private PlatformUtils() {
        throw new AssertionError("utility class must not be instantiated");
    }

    public static boolean isLinux() {
        return OS_NAME.startsWith(LINUX);
    }

    public static boolean isMac() {
        return OS_NAME.startsWith(MAC);
    }

    public static boolean isWindows() {
        return OS_NAME.startsWith(WINDOWS);
    }

    private static String findOfficeHome() {
        String[] homeList = null;
        if (isLinux()) {
            homeList = LINUX_OO_HOME_PATH;
        } else if (isMac()) {
            homeList = MAC_OO_HOME_PATH;
        } else if (isWindows()){
            homeList = WINDOWS_OO_HOME_PATH;
        }
        return searchExistingfile(Arrays.asList(homeList));
    }

    private static String findOfficeProfileDir() {
        String[] profileDirList = null;
        if (isLinux()) {
            profileDirList = LINUX_OO_PROFILE_PATH;
        } else if (isMac()) {
            profileDirList = MAC_OO_PROFILE_PATH;
        } else if (isWindows()){
            profileDirList = WINDOWS_OO_PROFILE_PATH;
        }
        return searchExistingfile(Arrays.asList(profileDirList));
    }

    protected static String searchExistingfile(List<String> pathList) {
        File officeHome = null;
        for (String path : pathList) {
            officeHome = new File(path);
            if (officeHome.exists()) {
                logger.log(Level.INFO, "Jod will be using " + path);
                return path;
            }
        }
        return null;
    }

}
