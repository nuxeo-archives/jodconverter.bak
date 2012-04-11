package org.artofsolving.jodconverter.office;

import java.util.logging.Logger;

public class OfficeVersionDescriptor {

    protected String productName;

    protected String version;

    protected boolean useGnuStyleLongOptions = false;

    private final Logger logger = Logger.getLogger(getClass().getName());

    public OfficeVersionDescriptor(String checkString) {
        logger.fine("Building " + this.getClass().getSimpleName() + ": "
                + checkString.trim());
        String versionLabel = checkString;
        String[] lines = checkString.split("\\n");
        if (lines.length > 1) {
            if (lines[0].contains("--help")) {
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
        logger.info("soffice info: " + toString());
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

    @Override
    public String toString() {
        return String.format(
            "Product: %s - Version: %s - useGnuStyleLongOptions: %s",
            getProductName(), getVersion(), useGnuStyleLongOptions());
    }
}
