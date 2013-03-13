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
package org.artofsolving.jodconverter;

import static org.artofsolving.jodconverter.office.OfficeUtils.cast;

import java.io.File;
import java.io.Serializable;
import java.util.HashMap;
import java.util.Map;

import org.artofsolving.jodconverter.document.DocumentFamily;
import org.artofsolving.jodconverter.document.DocumentFormat;

import com.sun.star.container.XIndexAccess;
import com.sun.star.lang.XComponent;
import com.sun.star.text.XDocumentIndex;
import com.sun.star.text.XDocumentIndexesSupplier;
import com.sun.star.text.XTextDocument;
import com.sun.star.util.XRefreshable;

/**
 * 
 * Added Map<String, Object> params to control DocumentIndex update
 * 
 * @author <a href="mailto:tdelprat@nuxeo.com">Tiry</a>
 * 
 */
public class StandardConversionTask extends AbstractConversionTask {

    public static final String UPDATE_DOCUMENT_INDEX = "updateDocumentIndex";

    private final DocumentFormat outputFormat;

    private Map<String, ?> defaultLoadProperties;

    private DocumentFormat inputFormat;

    protected final Map<String, Serializable> params;

    public StandardConversionTask(File inputFile, File outputFile,
            DocumentFormat outputFormat, Map<String, Serializable> params) {
        super(inputFile, outputFile);
        this.outputFormat = outputFormat;
        if (params == null) {
            params = new HashMap<String, Serializable>();
        }
        this.params = params;
    }

    public StandardConversionTask(File inputFile, File outputFile,
            DocumentFormat outputFormat) {
        this(inputFile, outputFile, outputFormat, null);
    }

    public void setDefaultLoadProperties(Map<String, ?> defaultLoadProperties) {
        this.defaultLoadProperties = defaultLoadProperties;
    }

    public void setInputFormat(DocumentFormat inputFormat) {
        this.inputFormat = inputFormat;
    }

    @Override
    protected Map<String, ?> getLoadProperties(File inputFile) {
        Map<String, Object> loadProperties = new HashMap<String, Object>();
        if (defaultLoadProperties != null) {
            loadProperties.putAll(defaultLoadProperties);
        }
        if (inputFormat != null && inputFormat.getLoadProperties() != null) {
            loadProperties.putAll(inputFormat.getLoadProperties());
        }
        return loadProperties;
    }

    @Override
    protected Map<String, ?> getStoreProperties(File outputFile,
            XComponent document) {
        DocumentFamily family = OfficeDocumentUtils.getDocumentFamily(document);
        return outputFormat.getStoreProperties(family);
    }

    @Override
    protected void handleDocumentLoaded(XComponent document) {
        if (updateDocumentIndexes()) {
            doUpdateDocumentIndexes(document);
        }
        super.handleDocumentLoaded(document);
    }

    protected boolean updateDocumentIndexes() {
        Serializable flag = params.get(UPDATE_DOCUMENT_INDEX);
        if (flag != null && flag.toString().equalsIgnoreCase("true")) {
            return true;
        }
        return false;
    }

    protected void doUpdateDocumentIndexes(XComponent document) {
        XTextDocument xDocument = cast(XTextDocument.class, document);
        if (xDocument != null) {
            XDocumentIndexesSupplier indexSupplier = cast(
                    XDocumentIndexesSupplier.class, xDocument);
            XDocumentIndex index = null;

            XRefreshable xRefreshable = cast(XRefreshable.class, document);
            if (xRefreshable != null) {
                // This refresh operation solves issues with ToC update operations,
                // which could lead to bad page numbers in some scenarios (specific
                // hosts, specific documents, conversion to non-single-file document 
                // format like ODT - PDF not affected).
                // References:
                //   * http://www.oooforum.org/forum/viewtopic.phtml?t=7826
                //   * https://issues.apache.org/ooo/show_bug.cgi?id=29165
                xRefreshable.refresh();
            }
            
            if (indexSupplier != null) {
                XIndexAccess ia = indexSupplier.getDocumentIndexes();
                for (int i = 0; i < ia.getCount(); i++) {
                    Object idx = null;
                    try {
                        idx = ia.getByIndex(i);
                        index = cast(XDocumentIndex.class, idx);
                        if (index != null) {
                            index.update();
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }
        }
    }

}
