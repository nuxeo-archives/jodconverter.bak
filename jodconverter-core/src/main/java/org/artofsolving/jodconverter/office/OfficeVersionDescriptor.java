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
        String productLine = null;
        String[] lines = checkString.split("\\n");
        for (String line : lines) {
            if (line.contains("--help")) {
                useGnuStyleLongOptions = true;
            }
            String lowerLine = line.trim().toLowerCase();
            if (lowerLine.startsWith("openoffice") || lowerLine.startsWith("libreoffice")) {
                productLine = line.trim();
            }
        }
        if (productLine != null) {
            String[] parts = productLine.split(" ");
            if (parts.length > 0) {
                productName = parts[0];
            } else {
                productName = "???";
            }
            if (parts.length > 1) {
                version = parts[1];
            } else {
                version = "???";
            }
        } else {
            productName = "???";
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
