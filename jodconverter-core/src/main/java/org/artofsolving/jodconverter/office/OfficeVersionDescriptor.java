package org.artofsolving.jodconverter.office;

import java.util.logging.Logger;

public class OfficeVersionDescriptor {

    protected String productName;

    protected String version;

    protected boolean useGnuStyleLongOptions = false;

    private final Logger logger = Logger.getLogger(getClass().getName());

    public OfficeVersionDescriptor(String checkString) {
        logger.info("Building " + this.getClass().getSimpleName() + ": "
                + checkString);
        String versionLabel = checkString;
        String[] lines = checkString.split("\\n");
        if (lines.length > 1) {
            if (lines[0].contains("--version")) {
                useGnuStyleLongOptions = true;
                versionLabel = lines[1];
            }
        }
        versionLabel = versionLabel.trim();
        String[] parts = versionLabel.split(" ");
        if (parts.length > 1) {
            productName = parts[0];
            version = parts[1];
        } else {
            productName = versionLabel;
            version = "???";
        }
        logger.info("soffice info : " + productName + " version " + version
                + " useGnuStyleLongOptions:"
                + Boolean.toString(useGnuStyleLongOptions));
    }

    public String getProductName() {
        return productName;
    }

    public void setProductName(String productName) {
        this.productName = productName;
    }

    public String getVersion() {
        return version;
    }

    public void setVersion(String version) {
        this.version = version;
    }

    public boolean useGnuStyleLongOptions() {
        return useGnuStyleLongOptions;
    }

    public boolean hasUserEnvBug() {
        if (productName.toLowerCase().contains("libreoffice")) {
            if (version.contains("3.4") || version.contains("3.5")) {
                return true;
            }
        }
        return false;
    }

    @Override
    public String toString() {
        StringBuffer sb = new StringBuffer();
        sb.append(this.getClass().getSimpleName());
        sb.append("\n" + getProductName());
        sb.append("\n" + getVersion());
        sb.append("\n useGnuStyleLongOptions :" + useGnuStyleLongOptions());
        return sb.toString();
    }
}
